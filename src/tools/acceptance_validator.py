"""
验收标准校验工具

功能：
- 读取分析方法文档 (.docx)
- 依据制药行业验收标准（ICH Q2、中国药典等）校验文档内容是否满足要求
- 检测性能指标是否达标、实验设计是否合规

验收标准数据接口：
- 当前版本使用 Mock 数据（src/mock_data/acceptance_criteria.json）
- 接口预留：可通过替换 _fetch_acceptance_criteria() 函数对接企业标准库
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

logger = logging.getLogger("pharma-mcp.acceptance")

# ── Mock 验收标准加载 ──────────────────────────────────────────────────────────

_CRITERIA_DATA_PATH = Path(__file__).parent.parent / "mock_data" / "acceptance_criteria.json"


def _fetch_acceptance_criteria(method_type: str | None = None) -> list[dict]:
    """
    获取验收标准。

    ⚠️ 接口预留点：将此函数替换为企业标准数据库查询即可对接生产环境。
    例如：
        response = requests.get(f"{STANDARDS_API_BASE}/criteria", params={"type": method_type})
        return response.json()["standards"]

    Args:
        method_type: 可选，按方法类型过滤。
                     支持: "HPLC", "UV", "溶出度", "无菌", "pH", "通用"

    Returns:
        验收标准列表
    """
    with open(_CRITERIA_DATA_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    standards = data.get("standards", [])
    if method_type:
        type_map = {
            "HPLC": "AC-HPLC-001",
            "UV": "AC-UV-001",
            "溶出度": "AC-DISS-001",
            "无菌": "AC-STERILE-001",
            "pH": "AC-PH-001",
            "通用": "AC-METHOD-VAL-001",
        }
        target_id = type_map.get(method_type)
        if target_id:
            standards = [s for s in standards if s.get("id") == target_id]

    logger.info("加载验收标准完成: method_type=%s standards=%d", method_type, len(standards))
    return standards


# ── Prompt 模板 ────────────────────────────────────────────────────────────────

_MAP_SYSTEM = """你是制药行业分析方法验证专家，熟悉ICH Q2(R1)、中国药典2020版相关规定。
你的任务是从文档片段中提取所有方法验证相关的实验数据，包括：
- 分析方法类型（HPLC、UV、溶出度、无菌、pH等）
- 性能指标数值（精密度、准确度、线性、专属性等）
- 实验设计（重复次数、浓度点数量等）
- 实测结果（含量、RSD、R²、回收率等具体数值）
- 接受标准（文档自带的判断依据）

如果该片段不包含方法验证数据，请说明"本段落无方法验证数据"。"""

_MAP_USER = """请从以下文档片段（第 {chunk_index}/{total_chunks} 块）中提取所有方法验证数据：

---文档内容开始---
{chunk_text}
---文档内容结束---

请详细列出所有验证数据，用于后续标准符合性评估。"""

_REDUCE_SYSTEM = """你是制药行业QA专家，负责汇总文档各段分析结果，提炼出完整的方法验证数据清单。"""

_REDUCE_USER = """请综合以下 {total_chunks} 个片段中提取的方法验证数据，整合为完整的数据清单：

{all_chunk_results}

请输出一份结构化的数据汇总，按验证项目分类整理。"""

_ASSESS_SYSTEM = """你是制药行业分析方法验证专家，负责根据行业验收标准对分析方法进行符合性评估。
你需要逐项对照验收标准，判断文档中的每项指标是否满足要求，并给出清晰的通过/不通过判断和理由。"""

_ASSESS_USER = """请根据以下验收标准，对从文档中提取的方法验证数据进行逐项评估：

## 从文档中提取的验证数据
{doc_data}

## 适用的验收标准（共 {criteria_count} 项标准）
```json
{criteria_json}
```

请生成详细的验收标准符合性评估报告，格式如下：

## 验收标准校验报告

### 一、文档方法类型识别
（识别文档属于哪类分析方法，适用哪些验收标准）

### 二、逐项符合性评估
按验收标准逐项评估，每项包含：
| 验收项目 | 标准要求 | 文档记录值 | 是否符合 | 备注 |

### 三、不符合项汇总
- 🔴 不符合项（列出具体项目、标准要求、实际值、差距）
- 🟡 部分符合项（数据不完整或描述不清晰）
- 🔵 未评估项（文档未提供相关数据）

### 四、整体评估结论
（通过/不通过 + 总结，建议补充的内容）

### 五、整改建议
针对不符合项和部分符合项，给出具体整改措施
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


async def run_acceptance_validation(
    file_path: str,
    method_type: str | None = None,
    progress_callback: Callable[[str, dict[str, Any]], Any] | None = None,
) -> dict:
    """
    对 Word 文档进行验收标准符合性校验。

    Args:
        file_path: .docx 文件路径
        method_type: 可选，指定分析方法类型以获取对应验收标准
                     支持: "HPLC", "UV", "溶出度", "无菌", "pH", "通用"
                     为 None 时使用全部标准

    Returns:
        {
          "status": "success" | "error",
          "file_path": str,
          "criteria_count": int,
          "doc_meta": {...},
          "report": str
        }
    """
    logger.info("开始验收标准校验: file=%s method_type=%s", file_path, method_type)
    try:
        meta, chunks = load_document(file_path)
    except (FileNotFoundError, ValueError) as e:
        logger.error("验收标准校验文档加载失败: %s", e)
        return {"status": "error", "message": str(e)}

    # 拉取验收标准
    criteria = _fetch_acceptance_criteria(method_type)

    chunk_texts = [c.text for c in chunks]

    # ── 阶段1：从文档提取验证数据 ─────────────────────────────────────────────
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

        doc_data = await analyze_single_chunk_async(
            system_prompt=_MAP_SYSTEM,
            user_template=_MAP_USER.replace("{chunk_index}", "1").replace("{total_chunks}", "1"),
            chunk_text=chunk_texts[0],
            on_delta=_on_extract_delta,
        )
        await _emit_progress(progress_callback, "extract_single_done", {"total": 1})
    else:
        await _emit_progress(progress_callback, "extract_map_start", {"total": len(chunk_texts)})
        extract_reduce_chars = 0

        async def _on_map_progress(completed: int, total: int) -> None:
            await _emit_progress(
                progress_callback,
                "extract_map_progress",
                {"completed": completed, "total": total},
            )

        async def _on_reduce_delta(delta: str) -> None:
            nonlocal extract_reduce_chars
            extract_reduce_chars += len(delta)
            if extract_reduce_chars % 300 < len(delta):
                await _emit_progress(
                    progress_callback,
                    "extract_reduce_stream",
                    {"chars": extract_reduce_chars},
                )

        doc_data = await analyze_chunks_map_reduce_async(
            chunks=chunk_texts,
            map_system_prompt=_MAP_SYSTEM,
            map_user_template=_MAP_USER,
            reduce_system_prompt=_REDUCE_SYSTEM,
            reduce_user_template=_REDUCE_USER,
            on_map_progress=_on_map_progress,
            on_reduce_delta=_on_reduce_delta,
        )
        await _emit_progress(progress_callback, "extract_reduce_done", {"chunks": len(chunk_texts)})

    # ── 阶段2：对照验收标准进行评估 ───────────────────────────────────────────
    criteria_json_str = json.dumps(criteria, ensure_ascii=False, indent=2)
    assess_user_msg = _ASSESS_USER.format(
        doc_data=doc_data,
        criteria_count=len(criteria),
        criteria_json=criteria_json_str,
    )
    assess_chars = 0

    async def _on_assess_delta(delta: str) -> None:
        nonlocal assess_chars
        assess_chars += len(delta)
        if assess_chars % 300 < len(delta):
            await _emit_progress(
                progress_callback,
                "assess_stream",
                {"chars": assess_chars},
            )

    report = await chat_completion_stream(
        system_prompt=_ASSESS_SYSTEM,
        user_message=assess_user_msg,
        on_delta=_on_assess_delta,
    )
    await _emit_progress(progress_callback, "assess_done", {"chars": len(report)})

    logger.info(
        "验收标准校验完成: file=%s criteria=%d report_chars=%d",
        file_path,
        len(criteria),
        len(report),
    )

    return {
        "status": "success",
        "file_path": file_path,
        "criteria_count": len(criteria),
        "method_type": method_type or "全部",
        "doc_meta": {
            "total_chars": meta.total_chars,
            "total_tokens": meta.total_tokens,
            "chunk_count": meta.chunk_count,
            "section_titles": meta.section_titles,
        },
        "report": report,
    }
