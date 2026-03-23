"""
Word 文档读取与章节提取工具

支持：
- 读取 .docx 文件全文（含段落、表格）
- 按标题层级结构化解析出章节树
- 按章节名称模糊/语义匹配查找内容
- 按 token 数量进行滑动窗口分块，保留重叠上下文
- 提供文档元数据（标题层级、节标题等）

底层转换：mammoth (docx→HTML) + markdownify (HTML→Markdown)，输出纯正 Markdown。
"""

from __future__ import annotations

import io
import os
import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import mammoth
import markdownify
import tiktoken
from markdownify import MarkdownConverter


# ── 分块配置 ──────────────────────────────────────────────────────────────────
CHUNK_SIZE: int = int(os.getenv("CHUNK_SIZE", "6000"))
CHUNK_OVERLAP: int = int(os.getenv("CHUNK_OVERLAP", "500"))
FUZZY_THRESHOLD: float = float(os.getenv("FUZZY_THRESHOLD", "0.45"))
HTTP_DOC_TIMEOUT: float = float(os.getenv("HTTP_DOC_TIMEOUT", "20"))

_ENCODER = tiktoken.get_encoding("cl100k_base")
logger = logging.getLogger("pharma-mcp.document")

# mammoth 样式映射：保留标题层级，其余不转样式名
_MAMMOTH_STYLE_MAP = """
p[style-name='heading 1'] => h1:fresh
p[style-name='heading 2'] => h2:fresh
p[style-name='heading 3'] => h3:fresh
p[style-name='heading 4'] => h4:fresh
p[style-name='heading 5'] => h5:fresh
p[style-name='heading 6'] => h6:fresh
p[style-name='标题 1'] => h1:fresh
p[style-name='标题 2'] => h2:fresh
p[style-name='标题 3'] => h3:fresh
p[style-name='标题 4'] => h4:fresh
p[style-name='标题 5'] => h5:fresh
p[style-name='标题 6'] => h6:fresh
"""


def _is_http_url(file_path: str) -> bool:
    parsed = urlparse(file_path)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _cache_key_for_source(file_path: str) -> str:
    src = file_path.strip()
    if _is_http_url(src):
        return f"url::{src}"
    local = Path(src).expanduser().resolve()
    return f"file::{local}"


def _load_docx_bytes(file_path: str) -> bytes:
    """加载 DOCX 原始字节：支持本地路径和 HTTP(S) URL。"""
    src = file_path.strip()
    if _is_http_url(src):
        logger.info("从 HTTP(S) 下载 Word 文档: %s", src)
        req = Request(src, headers={"User-Agent": "pharma-doc-reader/1.0"})
        try:
            with urlopen(req, timeout=HTTP_DOC_TIMEOUT) as resp:
                data = resp.read()
                content_type = (resp.headers.get("Content-Type") or "").lower()
        except HTTPError as e:
            raise ValueError(f"HTTP 文档下载失败：{src}，状态码：{getattr(e, 'code', 'unknown')}") from e
        except URLError as e:
            raise ValueError(f"HTTP 文档下载失败：{src}，原因：{getattr(e, 'reason', 'unknown')}") from e
        if not data:
            raise ValueError(f"HTTP 文档为空：{src}")
        if not src.lower().endswith(".docx") and "officedocument.wordprocessingml.document" not in content_type:
            logger.warning("URL 可能不是 docx: %s content-type=%s", src, content_type)
        return data

    path = Path(src).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"文件不存在：{file_path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"仅支持 .docx 格式，当前文件：{path.suffix}")
    return path.read_bytes()


class _HtmlTableConverter(MarkdownConverter):
    """
    与标准 MarkdownConverter 完全相同，但将 <table> 元素直接输出为
    HTML 原始格式，而非 Markdown 管道表格。
    HTML 表格对 LLM 更友好：清晰支持合并单元格、嵌套表格、换行内容。
    """

    def convert_table(self, el, text, **kwargs) -> str:  # type: ignore[override]
        # el 是 BeautifulSoup 元素，decode() 还原为 HTML 字符串
        return "\n\n" + el.decode() + "\n\n"


def _convert_to_markdown(file_path: str) -> str:
    """
    docx → HTML (mammoth) → Markdown (markdownify，表格保留 HTML)。

    - mammoth 按 Word 样式正确识别标题层级，输出语义化 HTML
    - 段落、列表、标题转为标准 Markdown
    - <table> 保留 HTML 原始格式：更清晰，支持合并单元格和嵌套表格
    """
    raw = _load_docx_bytes(file_path)
    result = mammoth.convert_to_html(
        io.BytesIO(raw),
        style_map=_MAMMOTH_STYLE_MAP,
        convert_image=mammoth.images.img_element(lambda _: {}),
    )
    if result.messages:
        for msg in result.messages:
            logger.debug("mammoth: %s", msg)

    md = _HtmlTableConverter(
        heading_style="ATX",        # # ## ### 风格
        bullets="-",                # 统一用 -
        newline_style="backslash",  # 换行用 \
        strip=["img"],              # 去掉图片
    ).convert(result.value)
    # 清理多余空行（连续 3 个以上空行压缩为 2 个）
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return md


def _count_tokens(text: str) -> int:
    return len(_ENCODER.encode(text))


# ── 数据结构 ──────────────────────────────────────────────────────────────────

@dataclass
class DocChunk:
    index: int
    text: str
    token_count: int
    is_last: bool = False


@dataclass
class DocMeta:
    file_path: str
    total_chars: int
    total_tokens: int
    chunk_count: int
    section_titles: list[str] = field(default_factory=list)


# ── Markdown 块解析 ───────────────────────────────────────────────────────────

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+)$")


def _parse_md_blocks(
    markdown: str,
) -> list[tuple[str, int | None, str]]:
    """
    将 Markdown 文本切分为类型化的块列表。

    返回：
        [('heading', level, title), ('table', None, markdown_table), ('text', None, text), ...]
    """
    lines = markdown.split("\n")
    blocks: list[tuple[str, int | None, str]] = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # 标题行
        m = _HEADING_RE.match(line)
        if m:
            blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        # 表格行（以 | 开头，收集连续行）
        if line.strip().startswith("|"):
            table_lines: list[str] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            blocks.append(("table", None, "\n".join(table_lines)))
            continue

        # 空行跳过
        if not line.strip():
            i += 1
            continue

        # 普通段落（收集连续非空、非标题、非表格的行）
        para_lines: list[str] = []
        while (
            i < len(lines)
            and lines[i].strip()
            and not lines[i].strip().startswith("|")
            and not _HEADING_RE.match(lines[i])
        ):
            para_lines.append(lines[i])
            i += 1
        if para_lines:
            blocks.append(("text", None, "\n".join(para_lines)))

    return blocks


def extract_document_text(file_path: str) -> tuple[str, list[str]]:
    """
    提取 Word 文档全部 Markdown 文本。

    Returns:
        (full_markdown, section_titles)
    """
    logger.info("开始读取 Word 文档（mammoth）: %s", file_path)
    md = _convert_to_markdown(file_path)
    titles = [m.group(2).strip() for m in (_HEADING_RE.match(l) for l in md.split("\n")) if m]
    logger.info("文档读取完成: chars=%d headings=%d", len(md), len(titles))
    return md, titles


# ── 长文本分块 ────────────────────────────────────────────────────────────────

def split_into_chunks(
    text: str,
    chunk_size: int = CHUNK_SIZE,
    overlap: int = CHUNK_OVERLAP,
) -> list[DocChunk]:
    """
    将长文本按 token 数量进行滑动窗口分块。

    策略：
    1. 优先在段落边界（\\n\\n）切分，保持语义完整
    2. 若单段落超过 chunk_size，按句子切分（。！？.!?）
    3. 相邻块保留 overlap token 的重叠内容，避免信息断层
    """
    tokens = _ENCODER.encode(text)
    total_tokens = len(tokens)
    logger.info("开始分块: total_tokens=%d chunk_size=%d overlap=%d", total_tokens, chunk_size, overlap)

    if total_tokens <= chunk_size:
        logger.info("无需分块，单块返回")
        return [DocChunk(index=0, text=text, token_count=total_tokens, is_last=True)]

    chunks: list[DocChunk] = []
    start = 0
    idx = 0

    while start < total_tokens:
        end = min(start + chunk_size, total_tokens)
        chunk_tokens = tokens[start:end]
        chunk_text = _ENCODER.decode(chunk_tokens)

        # 尝试在段落边界对齐（向前查找最近的 \n\n）
        if end < total_tokens:
            better_end = _find_paragraph_boundary(chunk_text)
            if better_end > len(chunk_text) // 2:
                chunk_text = chunk_text[:better_end]
                chunk_tokens = _ENCODER.encode(chunk_text)

        chunks.append(DocChunk(
            index=idx,
            text=chunk_text,
            token_count=len(chunk_tokens),
        ))

        # 下一块起点 = 当前块末尾 - overlap
        advance = max(len(chunk_tokens) - overlap, 1)
        start += advance
        idx += 1

    if chunks:
        chunks[-1].is_last = True

    logger.info("分块完成: chunk_count=%d", len(chunks))
    return chunks


def _find_paragraph_boundary(text: str) -> int:
    """在文本末尾 20% 范围内查找最近的段落边界位置。"""
    search_from = int(len(text) * 0.8)
    pos = text.rfind("\n\n", search_from)
    if pos == -1:
        # 退而求其次：查找句末标点
        for punct in ("。", "！", "？", ".", "!", "?", "\n"):
            pos = text.rfind(punct, search_from)
            if pos != -1:
                return pos + 1
    return pos if pos != -1 else len(text)


# ── 公开接口 ──────────────────────────────────────────────────────────────────

def load_document(file_path: str) -> tuple[DocMeta, list[DocChunk]]:
    """
    加载 Word 文档，返回元数据和分块列表。

    Args:
        file_path: .docx 文件的绝对或相对路径

    Returns:
        (DocMeta, List[DocChunk])
    """
    full_text, section_titles = extract_document_text(file_path)
    chunks = split_into_chunks(full_text)
    total_tokens = sum(c.token_count for c in chunks)

    meta = DocMeta(
        file_path=file_path,
        total_chars=len(full_text),
        total_tokens=_count_tokens(full_text),
        chunk_count=len(chunks),
        section_titles=section_titles,
    )
    logger.info(
        "文档加载完成: path=%s total_chars=%d total_tokens=%d chunks=%d",
        file_path,
        meta.total_chars,
        meta.total_tokens,
        meta.chunk_count,
    )
    return meta, chunks


# ── 章节数据结构 ──────────────────────────────────────────────────────────────

@dataclass
class Section:
    """一个章节节点。"""
    title: str
    level: int          # 标题层级（1/2/3/…）
    content: str        # 该章节自身的正文（不含子章节）
    aliases: list[str] = field(default_factory=list)
    children: list[Section] = field(default_factory=list)

    def full_text(self) -> str:
        """递归获取该章节及其所有子章节的完整 Markdown 文本。"""
        parts: list[str] = []
        if self.level > 0:
            prefix = "#" * self.level
            heading_lines = [f"{prefix} {self.title}"]
            for alias in self.aliases:
                heading_lines.append(f"{prefix} {alias}")
            parts.append("\n".join(heading_lines))
        if self.content.strip():
            parts.append(self.content.strip())
        for child in self.children:
            parts.append(child.full_text())
        return "\n\n".join(p for p in parts if p)


# ── 章节结构化解析 ────────────────────────────────────────────────────────────

def extract_sections(file_path: str) -> tuple[list[Section], str, list[str]]:
    """
    解析 Word 文档为章节树（基于 mammoth+markdownify 输出的纯 Markdown）。

    Returns:
        (top_level_sections, full_markdown, all_section_titles)
    """
    logger.info("开始解析文档章节结构: %s", file_path)
    full_text = _convert_to_markdown(file_path)
    blocks = _parse_md_blocks(full_text)

    all_titles: list[str] = []
    root_sections: list[Section] = []
    stack: list[Section] = []
    current_body_parts: list[str] = []
    last_heading_section: Section | None = None
    last_was_heading = False

    def _heading_number(text: str) -> str:
        m = re.match(r"\s*(\d+(?:\.\d+)*)", text)
        return m.group(1) if m else ""

    def _flush_body() -> None:
        if stack and current_body_parts:
            block_text = "\n\n".join(current_body_parts)
            if stack[-1].content.strip():
                stack[-1].content += "\n\n" + block_text
            else:
                stack[-1].content = block_text
            current_body_parts.clear()

    preamble_parts: list[str] = []

    for btype, level, text in blocks:
        if btype == "heading":
            _flush_body()
            all_titles.append(text)

            # 中英文并行文档：同层级同编号相邻标题合并为别名
            if (
                last_was_heading
                and last_heading_section is not None
                and last_heading_section.level == level
                and _heading_number(last_heading_section.title)
                and _heading_number(last_heading_section.title) == _heading_number(text)
                and not last_heading_section.content.strip()
                and not last_heading_section.children
            ):
                logger.info("合并双语标题: '%s' / '%s'", last_heading_section.title, text)
                last_heading_section.aliases.append(text)
                last_was_heading = True
                continue

            section = Section(title=text, level=level, content="")
            while stack and stack[-1].level >= level:
                stack.pop()
            if stack:
                stack[-1].children.append(section)
            else:
                root_sections.append(section)
            stack.append(section)
            last_heading_section = section
            last_was_heading = True
        else:
            # 表格块前后留空行，保证 Markdown 渲染正确
            item = ("\n" + text + "\n") if btype == "table" else text
            if stack:
                current_body_parts.append(item)
            else:
                preamble_parts.append(item)
            last_was_heading = False

    _flush_body()

    if preamble_parts:
        preamble = Section(title="（前言/封面）", level=0, content="\n\n".join(preamble_parts))
        root_sections.insert(0, preamble)

    logger.info(
        "章节解析完成: root_sections=%d all_titles=%d full_chars=%d",
        len(root_sections), len(all_titles), len(full_text),
    )
    return root_sections, full_text, all_titles


# ── 章节匹配与查找 ────────────────────────────────────────────────────────────

def _normalize(text: str) -> str:
    """归一化文本：去空格、标点、转小写，用于比较。"""
    text = text.lower().strip()
    text = re.sub(r"[\s\u3000]+", "", text)          # 去除空白
    text = re.sub(r"[,，.。、:：；;！!？?()（）\-—–]", "", text)  # 去除常见标点
    # 去除常见编号前缀：1. / 1.1 / 一、 / (一) / 第一章 等
    text = re.sub(r"^[\d.]+", "", text)
    text = re.sub(r"^[一二三四五六七八九十百千]+[、.]?", "", text)
    text = re.sub(r"^第[一二三四五六七八九十\d]+[章节条款篇]", "", text)
    text = re.sub(r"^[\(（][\d一二三四五六七八九十]+[\)）]", "", text)
    return text


def _fuzzy_score(query: str, title: str) -> float:
    """计算两个字符串的模糊匹配分数（0~1）。"""

    # 0) 精确数字编号匹配（如 5.1、6.2），需在归一化前判断
    query_num = re.search(r"\d+(?:\.\d+)+", query)
    title_num = re.search(r"\d+(?:\.\d+)+", title)
    if query_num and title_num and query_num.group() == title_num.group():
        return 1.0

    nq = _normalize(query)
    nt = _normalize(title)
    if not nq or not nt:
        return 0.0

    # 1) 精确包含
    if nq in nt or nt in nq:
        return 1.0

    # 2) SequenceMatcher 相似度
    ratio = SequenceMatcher(None, nq, nt).ratio()

    # 3) 关键词覆盖率加分
    q_chars = set(nq)
    t_chars = set(nt)
    if q_chars:
        overlap = len(q_chars & t_chars) / len(q_chars)
        ratio = max(ratio, overlap * 0.9)

    return ratio


def _best_section_score(query: str, section: Section) -> float:
    """对主标题和别名一起计算最佳模糊匹配分数。"""
    scores = [_fuzzy_score(query, section.title)]
    scores.extend(_fuzzy_score(query, alias) for alias in section.aliases)
    return max(scores)


def _flatten_sections(sections: list[Section]) -> list[Section]:
    """递归展平章节树为列表。"""
    result: list[Section] = []
    for s in sections:
        result.append(s)
        result.extend(_flatten_sections(s.children))
    return result


def _split_queries(section_name: str) -> list[str]:
    """支持用逗号/顿号/分号/竖线分隔多个章节关键词。"""
    parts = re.split(r"[，,、;；|]+", section_name)
    return [p.strip() for p in parts if p.strip()]


def find_sections(
    sections: list[Section],
    query: str,
    threshold: float = FUZZY_THRESHOLD,
    top_k: int = 1,
) -> list[tuple[Section, float]]:
    """
    在章节树中按名称模糊匹配查找。

    Args:
        sections: 顶层章节列表（通常来自 extract_sections）
        query: 要查找的章节名称
        threshold: 匹配分数阈值（0~1），默认 0.45

    Returns:
        按匹配分数降序排列的 [(Section, score), ...]
    """
    all_sections = _flatten_sections(sections)
    logger.info("开始章节匹配: query='%s' total_sections=%d threshold=%.2f", query, len(all_sections), threshold)

    scored: list[tuple[Section, float]] = []
    for section in all_sections:
        score = _best_section_score(query, section)
        if score >= threshold:
            scored.append((section, score))
            logger.debug(
                "匹配: '%s' aliases=%s -> score=%.3f",
                section.title,
                section.aliases,
                score,
            )

    scored.sort(key=lambda x: x[1], reverse=True)
    if top_k > 0:
        scored = scored[:top_k]
    logger.info("匹配结果: %d 个章节符合阈值并返回前 %d 个", len(scored), top_k)
    return scored


# ── 文档解析缓存 ──────────────────────────────────────────────────────────────
_DOC_CACHE = {}


def get_section_content(file_path: str, section_name: str) -> dict:
    """
    根据文件路径和章节名称获取章节内容。

    Args:
        file_path: .docx 文件路径
        section_name: 章节名称（支持模糊匹配），"*" 表示返回全部内容

    Returns:
        {
            "status": "success" | "error",
            "file_path": str,
            "section_name": str,
            "matched_sections": [
                {"title": str, "level": int, "score": float, "content": str}
            ],
            "all_titles": [...],       # 文档所有章节标题
            "total_chars": int,
            "message": str | None,
        }
    """
    logger.info("get_section_content: file=%s section='%s'", file_path, section_name)

    cache_key = _cache_key_for_source(file_path)
    if cache_key in _DOC_CACHE:
        sections, full_text, all_titles = _DOC_CACHE[cache_key]
        logger.info("命中文档缓存: %s", file_path)
    else:
        try:
            sections, full_text, all_titles = extract_sections(file_path)
            _DOC_CACHE[cache_key] = (sections, full_text, all_titles)
            logger.info("文档解析结果已缓存: %s", file_path)
        except (FileNotFoundError, ValueError) as e:
            logger.error("文档解析失败: %s", e)
            return {"status": "error", "message": str(e)}

    # 返回全部内容
    if section_name.strip() == "*":
        logger.info("返回全部文档内容, chars=%d", len(full_text))
        return {
            "status": "success",
            "file_path": file_path,
            "section_name": "*",
            "matched_sections": [
                {
                    "title": "（全文）",
                    "level": 0,
                    "score": 1.0,
                    "content": full_text,
                }
            ],
            "all_titles": all_titles,
            "total_chars": len(full_text),
            "message": None,
        }

    queries = _split_queries(section_name)
    if not queries:
        return {
            "status": "error",
            "message": "section_name 不能为空。",
        }

    merged_matches: dict[int, tuple[Section, float]] = {}
    for query in queries:
        current_matches = find_sections(sections, query, top_k=1)
        for section, score in current_matches:
            key = id(section)
            existed = merged_matches.get(key)
            if existed is None or score > existed[1]:
                merged_matches[key] = (section, score)

    matches = sorted(merged_matches.values(), key=lambda x: x[1], reverse=True)

    if not matches:
        logger.warning("未匹配到章节: query='%s'", section_name)
        return {
            "status": "success",
            "file_path": file_path,
            "section_name": section_name,
            "matched_sections": [],
            "all_titles": all_titles,
            "total_chars": 0,
            "message": f"未找到与 '{section_name}' 匹配的章节。文档包含以下章节：{', '.join(all_titles[:20])}",
        }

    matched_out: list[dict] = []
    for section, score in matches:
        content = section.full_text()
        matched_out.append({
            "title": section.title,
            "level": section.level,
            "score": round(score, 3),
            "content": content,
        })

    total_chars = sum(len(m["content"]) for m in matched_out)
    logger.info(
        "章节匹配完成: query='%s' matched=%d total_chars=%d",
        section_name,
        len(matched_out),
        total_chars,
    )
    return {
        "status": "success",
        "file_path": file_path,
        "section_name": section_name,
        "matched_sections": matched_out,
        "all_titles": all_titles,
        "total_chars": total_chars,
        "message": None,
    }


if __name__ == "__main__":
    import argparse
    import json
    parser = argparse.ArgumentParser(description="Word章节提取测试")
    parser.add_argument("file", help="Word文档路径（.docx）")
    parser.add_argument("section", help="章节名称，支持模糊匹配，*为全文")
    parser.add_argument("--json", action="store_true", help="输出JSON格式")
    args = parser.parse_args()

    result = get_section_content(args.file, args.section)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        if result["status"] != "success":
            print("❌", result.get("message", "未知错误"))
        else:
            for m in result["matched_sections"]:
                print(f"\n=== {m['title']} (H{m['level']}, 匹配度 {m['score']*100:.0f}%) ===\n")
                print(m["content"])
            if not result["matched_sections"]:
                print("⚠️ 未找到匹配章节。章节列表：")
                for t in result["all_titles"]:
                    print("  -", t)
