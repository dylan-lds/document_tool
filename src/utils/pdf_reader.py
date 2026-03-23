"""
pdf_reader.py - PDF内容读取与表格结构化提取

功能：
- 读取PDF全文内容，保留表格结构（优先Markdown/CSV风格，适合大模型输入）
- 支持本地和HTTP(S)路径
- 适配大模型喂入需求
"""
from __future__ import annotations

import io
import importlib
import logging
import os
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
