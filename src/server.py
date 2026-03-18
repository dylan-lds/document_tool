"""
制药行业文档内容提取服务

支持两种暴露方式：
1. MCP
2. OpenAPI（适配 Dify 等基于 OpenAPI Schema 的工具接入）

运行方式：
        python -m src.server --mode mcp --transport stdio
        python -m src.server --mode mcp --transport streamable-http
        python -m src.server --mode openapi
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
import uuid
from typing import Any
from typing import Annotated

from dotenv import load_dotenv
from mcp.server.fastmcp import Context, FastMCP
from pydantic import BaseModel, Field

try:
    from fastapi import Request as FastAPIRequest
except Exception:
    FastAPIRequest = Any

load_dotenv()

from src.utils.document_reader import get_section_content as _get_section_content

MCP_HOST = os.getenv("MCP_HOST", "127.0.0.1")
MCP_PORT = int(os.getenv("MCP_PORT", "8000"))
OPENAPI_HOST = os.getenv("OPENAPI_HOST", "127.0.0.1")
OPENAPI_PORT = int(os.getenv("OPENAPI_PORT", "8001"))

mcp = FastMCP("pharma-doc-reader", host=MCP_HOST, port=MCP_PORT)

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
)
logger = logging.getLogger("pharma-mcp")


class SectionMatch(BaseModel):
    title: str = Field(description="匹配到的章节主标题")
    level: int = Field(description="标题层级")
    score: float = Field(description="模糊匹配分数")
    content: str = Field(description="章节完整内容")


class SectionContentRequest(BaseModel):
    file_path: str = Field(description="Word 文档路径，支持本地路径或 HTTP(S) URL")
    section_name: str = Field(description="章节名称，支持模糊匹配；传 * 返回全部内容")


class SectionContentResponse(BaseModel):
    status: str = Field(description="success 或 error")
    file_path: str | None = Field(default=None, description="输入文件路径")
    section_name: str | None = Field(default=None, description="输入章节名")
    matched_sections: list[SectionMatch] = Field(default_factory=list, description="匹配结果")
    all_titles: list[str] = Field(default_factory=list, description="文档所有章节标题")
    total_chars: int = Field(default=0, description="返回内容总字符数")
    message: str | None = Field(default=None, description="错误或提示信息")


def _run_section_query(file_path: str, section_name: str) -> dict:
    request_id = str(uuid.uuid4())[:8]
    start_at = time.perf_counter()
    logger.info(
        "[章节提取][%s] 收到请求: file_path=%s, section_name='%s'",
        request_id,
        file_path,
        section_name,
    )
    result = _get_section_content(file_path, section_name)
    elapsed = time.perf_counter() - start_at
    logger.info(
        "[章节提取][%s] 结束: status=%s matched=%d elapsed=%.2fs",
        request_id,
        result.get("status"),
        len(result.get("matched_sections", [])),
        elapsed,
    )
    return result


def _format_mcp_output(result: dict, file_path: str, section_name: str, elapsed: float) -> str:
    if result.get("status") == "error":
        return f"❌ 提取失败：{result['message']}"

    matched = result.get("matched_sections", [])
    all_titles = result.get("all_titles", [])
    if not matched:
        msg = result.get("message", "未找到匹配章节")
        output = f"⚠️ {msg}\n\n"
        output += "📑 文档章节列表：\n"
        for i, title in enumerate(all_titles, 1):
            output += f"  {i}. {title}\n"
        return output

    header = (
        f"📄 文档: {file_path}\n"
        f"🔍 查询: {section_name}\n"
        f"✅ 匹配章节数: {len(matched)}\n"
        f"📊 总字符数: {result.get('total_chars', 0):,}\n"
        f"⏱️ 耗时: {elapsed:.2f}s\n"
    )
    if all_titles:
        header += f"📑 文档共 {len(all_titles)} 个章节\n"

    parts = [header, "---"]
    for i, m in enumerate(matched, 1):
        score_pct = m.get("score", 0) * 100
        parts.append(
            f"\n### 匹配 {i}: {m['title']}  "
            f"(层级 H{m['level']}, 匹配度 {score_pct:.0f}%)\n"
        )
        parts.append(m["content"])
    return "\n".join(parts)


def create_openapi_app() -> Any:
    from fastapi import FastAPI, HTTPException, Query

    app = FastAPI(
        title="Pharma Doc Reader API",
        description="根据文件路径与章节名称提取 Word 文档章节内容。适合 Dify 等 OpenAPI 工具接入。",
        version="1.0.0",
        openapi_version="3.1.0",
    )

    @app.get("/health", tags=["system"])
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.get(
        "/section-content",
        response_model=SectionContentResponse,
        tags=["document"],
        summary="按章节提取文档内容",
        operation_id="GetSectionContent",
    )
    async def get_section_content_http(
        file_path: str = Query(description="Word 文档路径，支持本地路径或 HTTP(S) URL"),
        section_name: str = Query(description="章节名称，支持模糊匹配；传 * 返回全部内容"),
    ) -> SectionContentResponse:
        result = _run_section_query(file_path, section_name)
        if result.get("status") == "error":
            raise HTTPException(status_code=400, detail=result.get("message", "unknown error"))
        return SectionContentResponse(**result)

    @app.post(
        "/section-content",
        response_model=SectionContentResponse,
        tags=["document"],
        summary="按章节提取文档内容",
        operation_id="PostSectionContent",
    )
    async def post_section_content_http(request: FastAPIRequest) -> SectionContentResponse:
        content_type = (request.headers.get("content-type") or "").lower()

        file_path: str | None = None
        section_name: str | None = None

        if "application/json" in content_type:
            try:
                payload = await request.json()
            except json.JSONDecodeError as e:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "请求体 JSON 格式错误，请检查是否使用英文双引号并正确闭合。"
                        f" 解析错误: {e.msg}"
                    ),
                ) from e
            file_path = payload.get("file_path")
            section_name = payload.get("section_name")
        elif "application/x-www-form-urlencoded" in content_type or "multipart/form-data" in content_type:
            form = await request.form()
            file_path = form.get("file_path")
            section_name = form.get("section_name")
        else:
            file_path = request.query_params.get("file_path")
            section_name = request.query_params.get("section_name")

        if not file_path or not section_name:
            raise HTTPException(
                status_code=400,
                detail="缺少必填参数 file_path 或 section_name",
            )

        result = _run_section_query(str(file_path), str(section_name))
        if result.get("status") == "error":
            raise HTTPException(status_code=400, detail=result.get("message", "unknown error"))
        return SectionContentResponse(**result)

    return app


# ── 工具：获取文档章节内容 ────────────────────────────────────────────────────

@mcp.tool()
async def get_section_content(
    file_path: Annotated[str, "Word文档(.docx)路径，支持本地路径或 HTTP(S) URL"],
    section_name: Annotated[str, "章节名称，支持模糊匹配。传 * 返回全部文档内容"],
    ctx: Context,
) -> str:
    """
    从 Word 文档中提取指定章节的完整内容。

    - section_name 支持模糊匹配（如 "实验方法" 可匹配到 "3.2 实验方法与步骤"）
    - section_name 为 "*" 时返回整个文档的全部内容
    - 返回匹配到的章节标题、匹配分数和内容；未匹配时列出文档全部章节标题供参考
    """
    start_at = time.perf_counter()
    request_id = str(uuid.uuid4())[:8]
    await ctx.info(f"[章节提取][{request_id}] 开始解析文档")

    try:
        result = _run_section_query(file_path, section_name)
    except Exception as e:
        logger.exception("[章节提取][%s] 异常: %s", request_id, e)
        await ctx.error(f"[章节提取][{request_id}] 执行异常: {e}")
        return f"❌ 提取失败：{e}"

    elapsed = time.perf_counter() - start_at
    await ctx.info(
        f"[章节提取][{request_id}] 完成，匹配 {len(result.get('matched_sections', []))} 个章节，耗时 {elapsed:.2f}s"
    )
    return _format_mcp_output(result, file_path, section_name, elapsed)


# ── 入口 ──────────────────────────────────────────────────────────────────────

def main_mcp(transport: str) -> None:
    if transport == "auto":
        transport = "streamable-http" if sys.stdin.isatty() else "stdio"

    if transport == "stdio" and sys.stdin.isatty():
        print("检测到交互终端：stdio 模式需要 MCP 客户端以管道方式启动。")
        print("如需本地常驻服务，请使用：python -m src.server --mode mcp --transport streamable-http")
        return

    if transport in {"sse", "streamable-http"}:
        if transport == "streamable-http":
            endpoint = mcp.settings.streamable_http_path
        else:
            endpoint = mcp.settings.sse_path
        print(
            f"MCP server 已启动: transport={transport}, "
            f"url=http://{mcp.settings.host}:{mcp.settings.port}{endpoint}"
        )

    mcp.run(transport=transport)


def main_openapi() -> None:
    try:
        import uvicorn
    except ImportError as e:
        raise RuntimeError(
            "未安装 OpenAPI 运行依赖，请先安装 fastapi 和 uvicorn。"
        ) from e

    app = create_openapi_app()
    print(
        f"OpenAPI server 已启动: url=http://{OPENAPI_HOST}:{OPENAPI_PORT} "
        f"openapi=http://{OPENAPI_HOST}:{OPENAPI_PORT}/openapi.json"
    )
    uvicorn.run(app, host=OPENAPI_HOST, port=OPENAPI_PORT)


def main():
    parser = argparse.ArgumentParser(description="Pharma document reader server")
    parser.add_argument(
        "--mode",
        choices=["mcp", "openapi"],
        default="mcp",
        help="启动模式：mcp / openapi",
    )
    parser.add_argument(
        "--transport",
        choices=["auto", "stdio", "sse", "streamable-http"],
        default="auto",
        help="MCP 传输模式：auto(默认) / stdio / sse / streamable-http",
    )
    args = parser.parse_args()

    if args.mode == "openapi":
        main_openapi()
        return

    main_mcp(args.transport)


if __name__ == "__main__":
    main()
