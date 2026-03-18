"""
ELN 校验工具

功能：
- 读取分析方法文档 (.docx) 中的实验结果
- 与 ELN（电子实验记录）系统中的数据进行对比
- 检测文档与 ELN 记录之间的差异（数值不符、操作步骤不一致等）

ELN 数据接口：
- 当前版本使用 Mock 数据（src/mock_data/eln_records.json）
- 接口预留：可通过替换 _fetch_eln_records() 函数对接真实 ELN 系统
"""

from __future__ import annotations

import inspect
import json
import logging
from collections.abc import Callable
from pathlib import Path
from typing import Any

from src.utils.document_reader import load_document
from src.utils.llm_client import (
    analyze_chunks_map_reduce_async,
    analyze_single_chunk_async,
    chat_completion_stream,
)

logger = logging.getLogger("pharma-mcp.eln")

# ── Mock ELN 数据加载 ──────────────────────────────────────────────────────────

_ELN_DATA_PATH = Path(__file__).parent.parent / "mock_data" / "eln_records.json"


def _fetch_eln_records(experiment_name: str | None = None) -> list[dict]:
    """
    获取 ELN 记录。

    ⚠️ 接口预留点：将此函数替换为真实 ELN API 调用即可对接生产环境。
    例如：
        response = requests.get(f"{ELN_API_BASE}/records", params={"name": experiment_name})
        return response.json()["records"]

    Args:
        experiment_name: 可选，按实验名称过滤。为 None 时返回所有记录。

    Returns:
        ELN 记录列表
    """
    with open(_ELN_DATA_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    records = data.get("records", [])
    if experiment_name:
        records = [
            r for r in records
            if experiment_name.lower() in r.get("experiment_name", "").lower()
        ]
    logger.info("加载 ELN 记录完成: experiment_name=%s records=%d", experiment_name, len(records))
    return records


# ── Prompt 模板 ────────────────────────────────────────────────────────────────

_MAP_SYSTEM = """你是制药行业QA审核专家，负责核查分析方法文档与ELN（电子实验记录）数据的一致性。
请从文档片段中提取所有具体的实验数据和结论，包括：
- 数值结果（含量、pH、溶出量、RSD等）
- 试剂规格与批号
- 仪器型号
- 溶液配制参数（浓度、体积、稀释倍数）
- 实验条件（温度、时间等）
- 判断结论

输出格式：列出从文档中提取到的所有可量化/可比对的信息，用于后续与ELN记录对比。
如果该片段没有相关实验数据，请说明"本段落无实验数据"。"""

_MAP_USER = """请从以下文档片段（第 {chunk_index}/{total_chunks} 块）中提取所有实验数据和结论：

---文档内容开始---
{chunk_text}
---文档内容结束---

请列出所有可与ELN记录比对的数据项。"""

_COMPARE_SYSTEM = """你是制药行业QA审核专家，负责将分析方法文档中的数据与ELN电子实验记录进行逐项对比，
找出所有不一致之处，并出具差异报告。

对比维度：
1. 数值一致性（允许合理的测量误差，如含量±0.5%视为一致）
2. 操作参数一致性（浓度、体积、稀释倍数等）
3. 仪器设备一致性
4. 判断结论一致性
5. 文档中记录了但ELN中缺失的信息
6. ELN中有但文档未体现的关键数据"""

_COMPARE_USER = """请将以下从文档中提取的数据与ELN记录进行对比分析：

## 文档中提取的数据
{doc_extracted_data}

## ELN 系统记录（共 {eln_count} 条相关记录）
```json
{eln_records_json}
```

请生成详细的 ELN 对比校验报告，格式如下：

## ELN 校验报告

### 一、匹配的 ELN 记录
列出与文档最相关的ELN记录（ELN ID、实验名称、操作者、日期）

### 二、一致项目
列出文档与ELN完全一致的数据项

### 三、差异项目
按严重程度分级：
- 🔴 严重差异（数值明显不符或结论相反）
- 🟡 一般差异（微小偏差或格式不同）
- 🔵 仅文档有/仅ELN有（信息缺失）

### 四、结论
（通过/不通过 + 说明）
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


async def run_eln_validation(
    file_path: str,
    experiment_name: str | None = None,
    progress_callback: Callable[[str, dict[str, Any]], Any] | None = None,
) -> dict:
    """
    对 Word 文档进行 ELN 数据对比校验。

    Args:
        file_path: .docx 文件路径
        experiment_name: 可选，指定要对比的实验名称（用于过滤ELN记录）

    Returns:
        {
          "status": "success" | "error",
          "file_path": str,
          "eln_records_count": int,
          "doc_meta": {...},
          "report": str
        }
    """
    logger.info("开始 ELN 校验: file=%s experiment_name=%s", file_path, experiment_name)
    try:
        meta, chunks = load_document(file_path)
    except (FileNotFoundError, ValueError) as e:
        logger.error("ELN 校验文档加载失败: %s", e)
        return {"status": "error", "message": str(e)}

    # 拉取 ELN 记录
    eln_records = _fetch_eln_records(experiment_name)

    chunk_texts = [c.text for c in chunks]

    # ── 阶段1：从文档中提取实验数据 ──────────────────────────────────────────
    if len(chunk_texts) == 1:
        await _emit_progress(progress_callback, "extract_single_start", {"total": 1})
        extract_chars = 0

        async def _on_extract_delta(delta: str) -> None:
            nonlocal extract_chars
            extract_chars += len(delta)
            if extract_chars % 300 < len(delta):
                await _emit_progress(
                    progress_callback,
                    "extract_single_stream",
                    {"chars": extract_chars},
                )

        doc_extracted = await analyze_single_chunk_async(
            system_prompt=_MAP_SYSTEM,
            user_template=_MAP_USER.replace("{chunk_index}", "1").replace("{total_chunks}", "1"),
            chunk_text=chunk_texts[0],
            on_delta=_on_extract_delta,
        )
        await _emit_progress(progress_callback, "extract_single_done", {"total": 1})
    else:
        await _emit_progress(progress_callback, "extract_map_start", {"total": len(chunk_texts)})
        reduce_chars = 0

        async def _on_map_progress(completed: int, total: int) -> None:
            await _emit_progress(
                progress_callback,
                "extract_map_progress",
                {"completed": completed, "total": total},
            )

        async def _on_reduce_delta(delta: str) -> None:
            nonlocal reduce_chars
            reduce_chars += len(delta)
            if reduce_chars % 300 < len(delta):
                await _emit_progress(
                    progress_callback,
                    "extract_reduce_stream",
                    {"chars": reduce_chars},
                )

        doc_extracted = await analyze_chunks_map_reduce_async(
            chunks=chunk_texts,
            map_system_prompt=_MAP_SYSTEM,
            map_user_template=_MAP_USER,
            reduce_system_prompt="你是制药行业数据整理专家，请将多个片段中提取的实验数据汇总为简洁完整的列表，去除重复信息。",
            reduce_user_template="请汇总以下各片段中提取的实验数据：\n\n{all_chunk_results}",
            on_map_progress=_on_map_progress,
            on_reduce_delta=_on_reduce_delta,
        )
        await _emit_progress(progress_callback, "extract_reduce_done", {"chunks": len(chunk_texts)})

    # ── 阶段2：与 ELN 数据对比 ────────────────────────────────────────────────
    eln_json_str = json.dumps(eln_records, ensure_ascii=False, indent=2)
    compare_user_msg = _COMPARE_USER.format(
        doc_extracted_data=doc_extracted,
        eln_count=len(eln_records),
        eln_records_json=eln_json_str,
    )
    compare_chars = 0

    async def _on_compare_delta(delta: str) -> None:
        nonlocal compare_chars
        compare_chars += len(delta)
        if compare_chars % 300 < len(delta):
            await _emit_progress(
                progress_callback,
                "compare_stream",
                {"chars": compare_chars},
            )

    report = await chat_completion_stream(
        system_prompt=_COMPARE_SYSTEM,
        user_message=compare_user_msg,
        on_delta=_on_compare_delta,
    )
    await _emit_progress(progress_callback, "compare_done", {"chars": len(report)})

    logger.info(
        "ELN 校验完成: file=%s records=%d report_chars=%d",
        file_path,
        len(eln_records),
        len(report),
    )

    return {
        "status": "success",
        "file_path": file_path,
        "eln_records_count": len(eln_records),
        "doc_meta": {
            "total_chars": meta.total_chars,
            "total_tokens": meta.total_tokens,
            "chunk_count": meta.chunk_count,
            "section_titles": meta.section_titles,
        },
        "report": report,
    }
