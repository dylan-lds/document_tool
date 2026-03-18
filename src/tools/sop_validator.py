"""
SOP 逻辑校验工具

功能：
- 读取制药行业分析方法文档 (.docx)
- 检测文档内部逻辑错误，包括但不限于：
  * 浓度计算错误（如两份10%溶液混合不应得到20%）
  * 体积/质量计算不一致
  * 稀释/混合操作前后浓度逻辑错误
  * 步骤顺序矛盾
  * 数值引用前后不一致
  * 单位换算错误
"""

from __future__ import annotations

import inspect
import logging
from collections.abc import Callable
from typing import Any

from src.utils.document_reader import load_document
from src.utils.llm_client import analyze_chunks_map_reduce_async, analyze_single_chunk_async

logger = logging.getLogger("pharma-mcp.sop")

# ── Prompt 模板 ────────────────────────────────────────────────────────────────

_MAP_SYSTEM = """你是一位资深制药行业QC专家，专注于分析方法SOP文档的逻辑审核。
你的任务是审查文档片段中的逻辑错误，包括但不限于：
1. 浓度计算错误（混合、稀释、配制过程中浓度前后矛盾）
   - 例：两份10%溶液等体积混合后，结果不应是20%，而应仍是10%
   - 例：将100mL 1mg/mL溶液稀释至1000mL后，浓度应为0.1mg/mL，不是1mg/mL
2. 质量/体积换算错误
3. 步骤前后矛盾（前面说A操作，后面引用时变成了B操作）
4. 仪器参数前后不一致
5. 时间/温度等条件矛盾
6. 数值引用与计算结果不符

输出要求：
- 如果发现问题，逐条列出，注明问题类型、问题描述、涉及的原文内容
- 如果该片段无逻辑错误，输出"本段落未发现明显逻辑错误"
- 不要捏造问题，要基于文档原文进行判断"""

_MAP_USER = """请审查以下文档片段（第 {chunk_index}/{total_chunks} 块）中的逻辑错误：

---文档内容开始---
{chunk_text}
---文档内容结束---

请列出所有发现的逻辑错误（含原文引用），如无问题请说明。"""

_REDUCE_SYSTEM = """你是一位资深制药行业QC专家，负责汇总多段文档分析结果，出具最终的SOP逻辑校验报告。
报告要求结构清晰，按严重程度分级（严重/一般/建议），并给出整体评估结论。"""

_REDUCE_USER = """以下是对同一份SOP文档各部分进行逻辑审核的结果（共 {total_chunks} 块）：

{all_chunk_results}

请综合以上分析，生成最终的SOP逻辑校验报告，格式如下：

## SOP 逻辑校验报告

### 一、总体评估
（通过/不通过，及简要说明）

### 二、发现的逻辑问题清单
按严重程度分级列出：
- 🔴 严重问题（影响实验结果有效性）
- 🟡 一般问题（可能影响重现性或理解）  
- 🔵 建议改进（表述不规范或可优化点）

每条问题包含：问题编号、问题类型、问题描述、原文引用

### 三、修改建议
针对严重和一般问题给出具体修改建议

### 四、结论
"""


async def _emit_progress(
    callback: Callable[[str, dict[str, Any]], Any] | None,
    stage: str,
    payload: dict[str, Any],
) -> None:
    if callback is None:
        return
    result = callback(stage, payload)
    if inspect.isawaitable(result):
        await result


async def run_sop_validation(
    file_path: str,
    progress_callback: Callable[[str, dict[str, Any]], Any] | None = None,
) -> dict:
    """
    对 Word 格式的 SOP 文档进行逻辑校验。

    Args:
        file_path: .docx 文件路径

    Returns:
        {
          "status": "success" | "error",
          "file_path": str,
          "doc_meta": {...},
          "report": str  # Markdown 格式的校验报告
        }
    """
    logger.info("开始 SOP 校验: file=%s", file_path)
    try:
        meta, chunks = load_document(file_path)
    except (FileNotFoundError, ValueError) as e:
        logger.error("SOP 文档加载失败: %s", e)
        return {"status": "error", "message": str(e)}

    chunk_texts = [c.text for c in chunks]
    logger.info("SOP 文档分块完成: chunks=%d", len(chunk_texts))

    if len(chunk_texts) == 1:
        await _emit_progress(progress_callback, "single_start", {"total": 1})
        stream_chars = 0

        async def _on_delta(delta: str) -> None:
            nonlocal stream_chars
            stream_chars += len(delta)
            if stream_chars % 300 < len(delta):
                await _emit_progress(
                    progress_callback,
                    "single_stream",
                    {"chars": stream_chars},
                )

        report = await analyze_single_chunk_async(
            system_prompt=_MAP_SYSTEM,
            user_template=_MAP_USER.replace(
                "{chunk_index}/{total_chunks}", "1/1"
            ).replace("{chunk_index}", "1").replace("{total_chunks}", "1"),
            chunk_text=chunk_texts[0],
            on_delta=_on_delta,
        )
        await _emit_progress(progress_callback, "single_done", {"total": 1})
    else:
        await _emit_progress(progress_callback, "map_start", {"total": len(chunk_texts)})
        reduce_chars = 0

        async def _on_map_progress(completed: int, total: int) -> None:
            await _emit_progress(
                progress_callback,
                "map_progress",
                {"completed": completed, "total": total},
            )

        async def _on_reduce_delta(delta: str) -> None:
            nonlocal reduce_chars
            reduce_chars += len(delta)
            if reduce_chars % 300 < len(delta):
                await _emit_progress(
                    progress_callback,
                    "reduce_stream",
                    {"chars": reduce_chars},
                )

        report = await analyze_chunks_map_reduce_async(
            chunks=chunk_texts,
            map_system_prompt=_MAP_SYSTEM,
            map_user_template=_MAP_USER,
            reduce_system_prompt=_REDUCE_SYSTEM,
            reduce_user_template=_REDUCE_USER,
            on_map_progress=_on_map_progress,
            on_reduce_delta=_on_reduce_delta,
        )
        await _emit_progress(progress_callback, "reduce_done", {"chunks": len(chunk_texts)})

    logger.info("SOP 校验完成: file=%s report_chars=%d", file_path, len(report))

    return {
        "status": "success",
        "file_path": file_path,
        "doc_meta": {
            "total_chars": meta.total_chars,
            "total_tokens": meta.total_tokens,
            "chunk_count": meta.chunk_count,
            "section_titles": meta.section_titles,
        },
        "report": report,
    }
