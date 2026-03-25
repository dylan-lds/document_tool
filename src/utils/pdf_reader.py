"""
pdf_reader.py - PDF内容读取与表格结构化提取

功能：
- 读取PDF全文内容，保留表格结构（HTML风格，适合大模型输入）
- 按章节名称模糊匹配提取特定章节内容（类似Word章节提取）
- 支持本地和HTTP(S)路径
- 适配大模型喂入需求
"""
from __future__ import annotations

import io
import importlib
import logging
import os
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

logger = logging.getLogger("pharma-mcp.pdf")
PDF_HTTP_TIMEOUT: float = float(os.getenv("PDF_HTTP_TIMEOUT", os.getenv("HTTP_DOC_TIMEOUT", "30")))

_PDF_CACHE: dict[str, dict[str, Any]] = {}


def _load_pymupdf() -> Any:
    try:
        return importlib.import_module("fitz")
    except ImportError as e:
        raise RuntimeError("未安装 PyMuPDF，请先安装 pymupdf>=1.24.0") from e


def _is_http_url(file_path: str) -> bool:
    parsed = urlparse(file_path)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _cache_key_for_source(file_path: str) -> str:
    src = file_path.strip()
    if _is_http_url(src):
        return f"url::{src}"
    return f"file::{Path(src).expanduser().resolve()}"


def _download_pdf(url: str) -> bytes:
    req = Request(url, headers={"User-Agent": "pharma-pdf-reader/1.0"})
    try:
        with urlopen(req, timeout=PDF_HTTP_TIMEOUT) as resp:
            data = resp.read()
    except HTTPError as e:
        raise ValueError(f"HTTP PDF 下载失败：{url}，状态码：{getattr(e, 'code', 'unknown')}") from e
    except URLError as e:
        raise ValueError(f"HTTP PDF 下载失败：{url}，原因：{getattr(e, 'reason', 'unknown')}") from e
    if not data:
        raise ValueError(f"HTTP PDF 为空：{url}")
    return data


def _rows_to_markdown(rows: list[list[Any]]) -> str:
    normalized_rows: list[list[str]] = []
    max_cols = max((len(row) for row in rows), default=0)
    if max_cols == 0:
        return ""

    for row in rows:
        normalized = [str(cell or "").strip().replace("\n", " <br> ") for cell in row]
        if len(normalized) < max_cols:
            normalized.extend([""] * (max_cols - len(normalized)))
        normalized_rows.append([cell.replace("|", r"\|") for cell in normalized])

    lines = []
    for idx, row in enumerate(normalized_rows):
        lines.append("| " + " | ".join(row) + " |")
        if idx == 0:
            lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    return "\n".join(lines)


def _table_to_markdown(table: Any) -> str:
    to_markdown = getattr(table, "to_markdown", None)
    if callable(to_markdown):
        try:
            md = to_markdown()
            if md:
                return md.strip()
        except Exception:
            logger.debug("PDF表格 to_markdown 失败，回退到 extract", exc_info=True)

    extract = getattr(table, "extract", None)
    if callable(extract):
        rows = extract() or []
        return _rows_to_markdown(rows)

    return ""


def _extract_page_tables(page: Any) -> list[str]:
    find_tables = getattr(page, "find_tables", None)
    if not callable(find_tables):
        return []

    try:
        found = find_tables()
    except Exception:
        logger.warning("PDF表格检测失败，已跳过当前页", exc_info=True)
        return []

    tables = getattr(found, "tables", found) or []
    table_markdowns: list[str] = []
    for table in tables:
        md = _table_to_markdown(table)
        if md:
            table_markdowns.append(md)
    return table_markdowns


def _normalize_page_text(text: str) -> str:
    text = text.replace("\x00", "")
    text = "\n".join(line.rstrip() for line in text.splitlines())
    text = "\n\n".join(part.strip() for part in text.split("\n\n") if part.strip())
    return text.strip()


def extract_pdf_content(file_path: str) -> dict[str, Any]:
    """
    读取PDF全文内容，表格优先结构化为Markdown，适合大模型输入。
    Returns: {
        'status': 'success'|'error',
        'file_path': str,
        'content': str,  # 全文内容，表格已结构化
        'tables': List[str],  # 每个表格的Markdown文本
        'message': str  # 错误时返回
    }
    """
    cache_key = _cache_key_for_source(file_path)
    cached = _PDF_CACHE.get(cache_key)
    if cached is not None:
        logger.info("命中PDF缓存: %s", file_path)
        return dict(cached)

    fitz = _load_pymupdf()
    doc: Any = None
    try:
        if _is_http_url(file_path):
            logger.info("下载PDF: %s", file_path)
            pdf_bytes = _download_pdf(file_path)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        else:
            path = Path(file_path).expanduser().resolve()
            if not path.exists():
                raise FileNotFoundError(f"文件不存在：{file_path}")
            if path.suffix.lower() != ".pdf":
                raise ValueError(f"仅支持 .pdf 格式，当前文件：{path.suffix}")
            doc = fitz.open(str(path))

        page_sections: list[str] = []
        tables_md: list[str] = []

        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            page_parts = [f"## 第 {page_num + 1} 页"]

            table_markdowns = _extract_page_tables(page)
            if table_markdowns:
                page_parts.append("[表格内容]")
                for idx, md in enumerate(table_markdowns, 1):
                    labeled_md = f"### 表格 {idx}\n{md.strip()}"
                    tables_md.append(f"第 {page_num + 1} 页 / 表格 {idx}\n{md.strip()}")
                    page_parts.append(labeled_md)

            text = _normalize_page_text(page.get_text("text"))
            if text:
                page_parts.append("[正文内容]")
                page_parts.append(text)

            page_sections.append("\n\n".join(part for part in page_parts if part))

        result = {
            "status": "success",
            "file_path": file_path,
            "content": "\n\n".join(section for section in page_sections if section).strip(),
            "tables": tables_md,
            "message": "",
        }
        _PDF_CACHE[cache_key] = dict(result)
        return result
    except Exception as e:
        logger.error("PDF加载失败: %s", e)
        return {
            "status": "error",
            "file_path": file_path,
            "content": "",
            "tables": [],
            "message": str(e),
        }
    finally:
        if doc is not None:
            doc.close()


# ── PDF 章节解析与匹配 ────────────────────────────────────────────────────────

# 匹配 PDF 中的编号标题行：'1. 标题' / '2.1 标题' / '4.3.2 预验证结果'
# 严格：编号必须含小数点（如 1. / 2.1 / 4.3.2），或者整数后跟 '.'，
# 后面跟至少一个中文字符或英文字母。
# 排除 '3 针LOD' / '22.5 - 27.5 mg' 这类非标题行。
_PDF_NUM_SEC_RE = re.compile(
    r"^(\d{1,2}(?:\.\d{1,2})+\.?)\s+(?=[\u4e00-\u9fff\u3400-\u4dbfA-Za-z])(.+)$"
    r"|"
    r"^(\d{1,2}\.)\s+(?=[\u4e00-\u9fff\u3400-\u4dbfA-Za-z])(.+)$"
)

FUZZY_THRESHOLD: float = float(os.getenv("FUZZY_THRESHOLD", "0.45"))


@dataclass
class PDFSection:
    """PDF 中的一个章节。"""
    title: str
    level: int
    page_start: int       # 起始页码（1-based）
    content: str = ""     # 该章节正文（含表格 HTML）
    children: list[PDFSection] = field(default_factory=list)

    def full_text(self) -> str:
        """递归获取该章节及其所有子章节的完整文本。"""
        parts: list[str] = []
        if self.level > 0:
            prefix = "#" * self.level
            parts.append(f"{prefix} {self.title}")
        if self.content.strip():
            parts.append(self.content.strip())
        for child in self.children:
            parts.append(child.full_text())
        return "\n\n".join(p for p in parts if p)


def _extract_pdf_text_with_tables(file_path: str) -> tuple[list[tuple[int, str]], list[str]]:
    """
    提取 PDF 每页的文本块，表格转为 HTML 格式。

    Returns:
        ([(page_num_1based, text_block), ...], [table_html, ...])
    """
    fitz = _load_pymupdf()
    doc: Any = None
    try:
        if _is_http_url(file_path):
            pdf_bytes = _download_pdf(file_path)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        else:
            path = Path(file_path).expanduser().resolve()
            if not path.exists():
                raise FileNotFoundError(f"文件不存在：{file_path}")
            if path.suffix.lower() != ".pdf":
                raise ValueError(f"仅支持 .pdf 格式，当前文件：{path.suffix}")
            doc = fitz.open(str(path))

        blocks: list[tuple[int, str]] = []
        all_tables: list[str] = []

        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pn = page_num + 1

            # 表格区域
            table_markdowns = _extract_page_tables(page)
            for idx, md in enumerate(table_markdowns, 1):
                # 将 Markdown 表格转成 HTML table
                html_table = _markdown_table_to_html(md)
                blocks.append((pn, html_table))
                all_tables.append(f"第 {pn} 页 / 表格 {idx}\n{md.strip()}")

            # 正文
            text = _normalize_page_text(page.get_text("text"))
            if text:
                for line in text.split("\n"):
                    stripped = line.strip()
                    if stripped:
                        blocks.append((pn, stripped))

        return blocks, all_tables
    finally:
        if doc is not None:
            doc.close()


def _markdown_table_to_html(md: str) -> str:
    """将 Markdown 表格转为 HTML <table>，方便大模型阅读。"""
    lines = [l.strip() for l in md.strip().splitlines() if l.strip()]
    if len(lines) < 2:
        return md

    def parse_row(line: str) -> list[str]:
        # 去掉首尾 |
        line = line.strip()
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]
        return [cell.strip() for cell in line.split("|")]

    # 检测分隔行
    rows: list[list[str]] = []
    for line in lines:
        cells = parse_row(line)
        # 跳过 |---|---|
        if all(re.match(r"^-+$", c.strip()) for c in cells if c.strip()):
            continue
        rows.append(cells)

    if not rows:
        return md

    parts = ["<table>"]
    for i, row in enumerate(rows):
        tag = "th" if i == 0 else "td"
        parts.append("  <tr>")
        for cell in row:
            cell_clean = cell.replace("<br>", "\n").replace(" <br> ", "\n").strip()
            parts.append(f"    <{tag}>{cell_clean}</{tag}>")
        parts.append("  </tr>")
    parts.append("</table>")
    return "\n".join(parts)


def _parse_pdf_sections(
    blocks: list[tuple[int, str]],
) -> tuple[list[PDFSection], str, list[str]]:
    """
    将 PDF 文本块解析为章节树。

    编号标题（'N. 标题' / 'N.N 标题'）识别为章节，
    其余内容挂载到最近的章节下。

    Returns:
        (root_sections, full_text, all_titles)
    """
    root_sections: list[PDFSection] = []
    stack: list[PDFSection] = []
    current_body_parts: list[str] = []
    all_titles: list[str] = []
    full_parts: list[str] = []

    def _flush_body() -> None:
        if stack and current_body_parts:
            block_text = "\n\n".join(current_body_parts)
            if stack[-1].content.strip():
                stack[-1].content += "\n\n" + block_text
            else:
                stack[-1].content = block_text
            current_body_parts.clear()

    preamble_parts: list[str] = []

    for page_num, text in blocks:
        full_parts.append(text)

        # 检查是否为编号标题
        m = _PDF_NUM_SEC_RE.match(text)
        if m and not text.strip().startswith("<table"):
            _flush_body()
            # 两个替代分支：group(1) 或 group(3) 为编号部分
            num_str = (m.group(1) or m.group(3)).rstrip(".")
            level = len(num_str.split("."))
            title = text.strip()
            all_titles.append(title)

            section = PDFSection(title=title, level=level, page_start=page_num)
            while stack and stack[-1].level >= level:
                stack.pop()
            if stack:
                stack[-1].children.append(section)
            else:
                root_sections.append(section)
            stack.append(section)
        else:
            if stack:
                current_body_parts.append(text)
            else:
                preamble_parts.append(text)

    _flush_body()

    if preamble_parts:
        preamble = PDFSection(
            title="（前言/封面）", level=0, page_start=1,
            content="\n\n".join(preamble_parts),
        )
        root_sections.insert(0, preamble)

    full_text = "\n\n".join(full_parts)
    return root_sections, full_text, all_titles


# ── 模糊匹配（复用 document_reader 的算法逻辑）───────────────────────────────

def _normalize_for_match(text: str) -> str:
    """归一化文本：去空格、标点、转小写。"""
    text = text.lower().strip()
    text = re.sub(r"[\s\u3000]+", "", text)
    text = re.sub(r"[,，.。、:：；;！!？?()（）\-—–]", "", text)
    text = re.sub(r"^[\d.]+", "", text)
    return text


def _fuzzy_score_pdf(query: str, title: str) -> float:
    # 精确数字编号匹配
    query_num = re.search(r"\d+(?:\.\d+)+", query)
    title_num = re.search(r"\d+(?:\.\d+)+", title)
    if query_num and title_num and query_num.group() == title_num.group():
        return 1.0

    nq = _normalize_for_match(query)
    nt = _normalize_for_match(title)
    if not nq or not nt:
        return 0.0

    if nq in nt or nt in nq:
        return 1.0

    ratio = SequenceMatcher(None, nq, nt).ratio()
    q_chars = set(nq)
    t_chars = set(nt)
    if q_chars:
        overlap = len(q_chars & t_chars) / len(q_chars)
        ratio = max(ratio, overlap * 0.9)

    return ratio


def _flatten_pdf_sections(sections: list[PDFSection]) -> list[PDFSection]:
    result: list[PDFSection] = []
    for s in sections:
        result.append(s)
        result.extend(_flatten_pdf_sections(s.children))
    return result


_PDF_SECTION_CACHE: dict[str, tuple[list[PDFSection], str, list[str]]] = {}


def get_pdf_section_content(file_path: str, section_name: str) -> dict[str, Any]:
    """
    按章节名称从 PDF 中提取内容。

    Args:
        file_path: PDF 文件路径
        section_name: 章节名称（模糊匹配），'*' 返回全文

    Returns:
        与 Word 版 get_section_content 返回结构一致：
        {
            "status": "success"|"error",
            "file_path": str,
            "section_name": str,
            "matched_sections": [{title, level, score, content}, ...],
            "all_titles": [...],
            "total_chars": int,
            "message": str|None,
        }
    """
    logger.info("get_pdf_section_content: file=%s section='%s'", file_path, section_name)

    cache_key = _cache_key_for_source(file_path)
    if cache_key in _PDF_SECTION_CACHE:
        sections, full_text, all_titles = _PDF_SECTION_CACHE[cache_key]
        logger.info("命中PDF章节缓存: %s", file_path)
    else:
        try:
            blocks, _tables = _extract_pdf_text_with_tables(file_path)
            sections, full_text, all_titles = _parse_pdf_sections(blocks)
            logger.info("解析PDF完成: %d sections, %d chars, %d tables", all_titles, len(full_text), len(_tables))
            _PDF_SECTION_CACHE[cache_key] = (sections, full_text, all_titles)
        except (FileNotFoundError, ValueError) as e:
            logger.error("PDF解析失败: %s", e)
            return {"status": "error", "message": str(e)}

    # 全文
    if section_name.strip() == "*":
        return {
            "status": "success",
            "file_path": file_path,
            "section_name": "*",
            "matched_sections": [
                {"title": "（全文）", "level": 0, "score": 1.0, "content": full_text}
            ],
            "all_titles": all_titles,
            "total_chars": len(full_text),
            "message": None,
        }

    # 支持多关键词
    queries = [q.strip() for q in re.split(r"[，,、;；|]+", section_name) if q.strip()]
    if not queries:
        return {"status": "error", "message": "section_name 不能为空。"}

    all_flat = _flatten_pdf_sections(sections)
    merged: dict[int, tuple[PDFSection, float]] = {}
    for query in queries:
        for sec in all_flat:
            score = _fuzzy_score_pdf(query, sec.title)
            if score >= FUZZY_THRESHOLD:
                key = id(sec)
                if key not in merged or score > merged[key][1]:
                    merged[key] = (sec, score)

    matches = sorted(merged.values(), key=lambda x: x[1], reverse=True)[:3]

    if not matches:
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
    for sec, score in matches:
        content = sec.full_text()
        matched_out.append({
            "title": sec.title,
            "level": sec.level,
            "score": round(score, 3),
            "content": content,
        })

    total_chars = sum(len(m["content"]) for m in matched_out)
    return {
        "status": "success",
        "file_path": file_path,
        "section_name": section_name,
        "matched_sections": matched_out,
        "all_titles": all_titles,
        "total_chars": total_chars,
        "message": None,
    }


# ── PDF 按页码提取 ─────────────────────────────────────────────────────────────

def _parse_page_numbers(page_numbers_str: str, total_pages: int) -> list[int]:
    """
    解析页码字符串，返回有效页码列表（1-based，已去重排序）。
    支持格式：
      - 单页：  "3"
      - 范围：  "2-5"
      - 组合：  "1,3-5,7"
    """
    pages: set[int] = set()
    for part in re.split(r"[，,]+", page_numbers_str.strip()):
        part = part.strip()
        if not part:
            continue
        range_match = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", part)
        if range_match:
            start, end = int(range_match.group(1)), int(range_match.group(2))
            for p in range(start, end + 1):
                if 1 <= p <= total_pages:
                    pages.add(p)
        elif re.fullmatch(r"\d+", part):
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p)
    return sorted(pages)


def get_pdf_page_content(file_path: str, page_numbers: str) -> dict[str, Any]:
    """
    按页码从 PDF 中提取内容。

    Args:
        file_path:    PDF 文件路径或 URL
        page_numbers: 页码字符串，支持单页 "3"、范围 "2-5"、组合 "1,3-5,7"

    Returns:
        {
            "status":       "success"|"error",
            "file_path":    str,
            "page_numbers": str,          # 原始输入
            "pages":        [             # 每页结果
                {
                    "page": int,          # 页码（1-based）
                    "content": str,       # 正文文本
                    "tables": [str],      # 该页表格 Markdown 列表
                }
            ],
            "total_pages":  int,          # PDF 总页数
            "total_chars":  int,          # 所有返回页内容字符数之和
            "message":      str|None,
        }
    """
    logger.info("get_pdf_page_content: file=%s pages='%s'", file_path, page_numbers)

    fitz = _load_pymupdf()
    doc: Any = None
    try:
        if _is_http_url(file_path):
            logger.info("下载PDF: %s", file_path)
            pdf_bytes = _download_pdf(file_path)
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        else:
            path = Path(file_path).expanduser().resolve()
            if not path.exists():
                raise FileNotFoundError(f"文件不存在：{file_path}")
            if path.suffix.lower() != ".pdf":
                raise ValueError(f"仅支持 .pdf 格式，当前文件：{path.suffix}")
            doc = fitz.open(str(path))

        total_pages = doc.page_count
        requested = _parse_page_numbers(page_numbers, total_pages)

        if not requested:
            return {
                "status": "error",
                "file_path": file_path,
                "page_numbers": page_numbers,
                "pages": [],
                "total_pages": total_pages,
                "total_chars": 0,
                "message": (
                    f"页码 '{page_numbers}' 无效或超出范围，"
                    f"文档共 {total_pages} 页，请输入 1~{total_pages} 之间的页码。"
                ),
            }

        pages_out: list[dict[str, Any]] = []
        total_chars = 0

        for pn in requested:
            page = doc.load_page(pn - 1)  # 0-based index
            table_markdowns = _extract_page_tables(page)
            text = _normalize_page_text(page.get_text("text"))
            total_chars += len(text) + sum(len(t) for t in table_markdowns)
            pages_out.append({
                "page": pn,
                "content": text,
                "tables": table_markdowns,
            })

        return {
            "status": "success",
            "file_path": file_path,
            "page_numbers": page_numbers,
            "pages": pages_out,
            "total_pages": total_pages,
            "total_chars": total_chars,
            "message": None,
        }

    except (FileNotFoundError, ValueError) as e:
        logger.error("PDF按页提取失败: %s", e)
        return {
            "status": "error",
            "file_path": file_path,
            "page_numbers": page_numbers,
            "pages": [],
            "total_pages": 0,
            "total_chars": 0,
            "message": str(e),
        }
    except Exception as e:
        logger.exception("PDF按页提取异常: %s", e)
        return {
            "status": "error",
            "file_path": file_path,
            "page_numbers": page_numbers,
            "pages": [],
            "total_pages": 0,
            "total_chars": 0,
            "message": str(e),
        }
    finally:
        if doc is not None:
            doc.close()