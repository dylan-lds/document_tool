"""
Word 文档读取与章节提取工具

支持：
- 读取 .docx 文件全文（含段落、表格）
- 按标题层级结构化解析出章节树
- 按章节名称模糊/语义匹配查找内容
- 按 token 数量进行滑动窗口分块，保留重叠上下文
- 提供文档元数据（标题层级、节标题等）
"""

from __future__ import annotations

import io
import os
import logging
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterator
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import tiktoken
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


# ── 分块配置 ──────────────────────────────────────────────────────────────────
CHUNK_SIZE: int = int(os.getenv("CHUNK_SIZE", "6000"))       # 每块最大 token 数
CHUNK_OVERLAP: int = int(os.getenv("CHUNK_OVERLAP", "500"))  # 块间重叠 token 数
FUZZY_THRESHOLD: float = float(os.getenv("FUZZY_THRESHOLD", "0.45"))  # 模糊匹配阈值
HTTP_DOC_TIMEOUT: float = float(os.getenv("HTTP_DOC_TIMEOUT", "20"))

_ENCODER = tiktoken.get_encoding("cl100k_base")
logger = logging.getLogger("pharma-mcp.document")


def _is_http_url(file_path: str) -> bool:
    parsed = urlparse(file_path)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _cache_key_for_source(file_path: str) -> str:
    src = file_path.strip()
    if _is_http_url(src):
        return f"url::{src}"
    local = Path(src).expanduser().resolve()
    return f"file::{local}"


def _load_docx_document(file_path: str) -> Document:
    """加载 DOCX：支持本地路径和 HTTP(S) URL。"""
    src = file_path.strip()
    if _is_http_url(src):
        logger.info("从 HTTP(S) 下载 Word 文档: %s", src)
        req = Request(src, headers={"User-Agent": "pharma-doc-reader/1.0"})
        try:
            with urlopen(req, timeout=HTTP_DOC_TIMEOUT) as resp:
                data = resp.read()
                content_type = (resp.headers.get("Content-Type") or "").lower()
        except HTTPError as e:
            logger.error("下载文档失败(HTTP): %s status=%s", src, getattr(e, "code", "unknown"))
            raise ValueError(f"HTTP 文档下载失败：{src}，状态码：{getattr(e, 'code', 'unknown')}") from e
        except URLError as e:
            logger.error("下载文档失败(URL): %s reason=%s", src, getattr(e, "reason", "unknown"))
            raise ValueError(f"HTTP 文档下载失败：{src}，原因：{getattr(e, 'reason', 'unknown')}") from e

        if not data:
            logger.error("下载文档为空: %s", src)
            raise ValueError(f"HTTP 文档为空：{src}")

        if not src.lower().endswith(".docx") and "officedocument.wordprocessingml.document" not in content_type:
            logger.warning("URL 可能不是 docx 文件: %s content-type=%s", src, content_type)

        try:
            return Document(io.BytesIO(data))
        except Exception as e:
            logger.error("HTTP 文档解析失败: %s", src)
            raise ValueError(f"HTTP 文档解析失败（非有效 .docx）：{src}") from e

    path = Path(src).expanduser()
    logger.info("读取本地 Word 文档: %s", path)
    if not path.exists():
        logger.error("文档不存在: %s", src)
        raise FileNotFoundError(f"文件不存在：{file_path}")
    if path.suffix.lower() != ".docx":
        logger.error("文档格式不支持: %s", path.suffix)
        raise ValueError(f"仅支持 .docx 格式，当前文件：{path.suffix}")
    return Document(str(path))


def _count_tokens(text: str) -> int:
    logger.debug("统计文本 token 数, 长度=%d", len(text))
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


# ── 文档内容提取 ──────────────────────────────────────────────────────────────

def _iter_block_items(doc: Document) -> Iterator[Paragraph | Table]:
    """按文档顺序迭代段落和表格（保持原始顺序）。"""
    logger.debug("开始遍历文档块")
    from docx.oxml import OxmlElement  # noqa: F401

    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def _table_to_text(table: Table) -> str:
    """将表格转换为 Markdown 风格文本。"""
    logger.debug("表格转文本, 行数=%d", len(table.rows))
    rows_text: list[str] = []
    for i, row in enumerate(table.rows):
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        rows_text.append("| " + " | ".join(cells) + " |")
        if i == 0:
            # 表头分隔线
            rows_text.append("| " + " | ".join(["---"] * len(cells)) + " |")
    return "\n".join(rows_text)


def extract_document_text(file_path: str) -> tuple[str, list[str]]:
    """
    提取 Word 文档全部文本内容。

    Returns:
        (full_text, section_titles) 全文字符串 + 所有标题列表
    """
    logger.info("开始读取 Word 文档: %s", file_path)
    doc = _load_docx_document(file_path)
    parts: list[str] = []
    section_titles: list[str] = []

    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            # 识别标题段落
            style_name = block.style.name if block.style else ""
            if style_name.startswith("Heading") or style_name.startswith("标题"):
                section_titles.append(text)
                parts.append(f"\n{'#' * _heading_level(style_name)} {text}\n")
            else:
                parts.append(text)
        elif isinstance(block, Table):
            parts.append("\n" + _table_to_text(block) + "\n")

    full_text = "\n".join(parts)
    logger.info("文档读取完成: chars=%d sections=%d", len(full_text), len(section_titles))
    return full_text, section_titles


def _heading_level(style_name: str) -> int:
    """从样式名称解析标题层级，默认返回 2。"""
    for char in style_name:
        if char.isdigit():
            return int(char)
    return 2


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
        """递归获取该章节及其所有子章节的完整文本。"""
        parts = []
        prefix = "#" * self.level
        parts.append(f"{prefix} {self.title}\n")
        for alias in self.aliases:
            parts.append(f"{prefix} {alias}\n")
        if self.content.strip():
            parts.append(self.content.strip())
        for child in self.children:
            parts.append(child.full_text())
        return "\n\n".join(parts)


# ── 章节结构化解析 ────────────────────────────────────────────────────────────

def extract_sections(file_path: str) -> tuple[list[Section], str, list[str]]:
    """
    解析 Word 文档为章节树。

    Returns:
        (top_level_sections, full_text, all_section_titles)
    """
    logger.info("开始解析文档章节结构: %s", file_path)
    doc = _load_docx_document(file_path)

    # 先线性提取所有块
    blocks: list[tuple[str, int | None, str]] = []  # (type, heading_level, text)
    all_titles: list[str] = []
    full_parts: list[str] = []

    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            style_name = block.style.name if block.style else ""
            if style_name.startswith("Heading") or style_name.startswith("标题"):
                level = _heading_level(style_name)
                blocks.append(("heading", level, text))
                all_titles.append(text)
                full_parts.append(f"\n{'#' * level} {text}\n")
            else:
                blocks.append(("text", None, text))
                full_parts.append(text)
        elif isinstance(block, Table):
            table_text = _table_to_text(block)
            blocks.append(("table", None, table_text))
            full_parts.append("\n" + table_text + "\n")

    full_text = "\n".join(full_parts)

    # 构建章节树
    root_sections: list[Section] = []
    # 栈中保存 (level, Section) 用于构建父子关系
    stack: list[Section] = []
    current_body_parts: list[str] = []
    last_heading_section: Section | None = None
    last_was_heading = False

    def _heading_number(text: str) -> str:
        match = re.match(r"\s*(\d+(?:\.\d+)*)", text)
        return match.group(1) if match else ""

    def _flush_body() -> None:
        """把累积的正文冲入最近的章节。"""
        if stack and current_body_parts:
            stack[-1].content += "\n".join(current_body_parts)
            current_body_parts.clear()

    # 文档开头无标题的前言内容
    preamble_parts: list[str] = []

    for btype, level, text in blocks:
        if btype == "heading":
            _flush_body()

            # 中英文并行文档：若连续两个同层级、同编号标题相邻，则视为同一章节的别名
            if (
                last_was_heading
                and last_heading_section is not None
                and last_heading_section.level == level
                and _heading_number(last_heading_section.title)
                and _heading_number(last_heading_section.title) == _heading_number(text)
                and not last_heading_section.content.strip()
                and not last_heading_section.children
            ):
                logger.info(
                    "合并双语同编号标题: primary='%s' alias='%s'",
                    last_heading_section.title,
                    text,
                )
                last_heading_section.aliases.append(text)
                last_was_heading = True
                continue

            section = Section(title=text, level=level, content="")
            # 找到合适的父节点
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
            if stack:
                current_body_parts.append(text)
            else:
                preamble_parts.append(text)
            last_was_heading = False

    _flush_body()

    # 如有前言，插入一个虚拟章节
    if preamble_parts:
        preamble = Section(title="（前言/封面）", level=0, content="\n".join(preamble_parts))
        root_sections.insert(0, preamble)

    logger.info(
        "章节解析完成: root_sections=%d all_titles=%d full_chars=%d",
        len(root_sections),
        len(all_titles),
        len(full_text),
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
