"""
Qwen 大模型客户端（兼容 OpenAI SDK）

支持：
- 阿里云 DashScope OpenAI 兼容接口
- qwen-long / qwen-max / qwen-plus 等系列模型
- 长文本 Map-Reduce 聚合分析：将多块分析结果合并为最终报告
- 流式与非流式调用
"""

from __future__ import annotations

import asyncio
import inspect
import logging
import os
from collections.abc import Awaitable, Callable
from typing import Any

from dotenv import load_dotenv
from openai import AsyncOpenAI, OpenAI

load_dotenv()

# ── 配置 ──────────────────────────────────────────────────────────────────────
QWEN_API_KEY: str = os.getenv("DASHSCOPE_API_KEY", "")
QWEN_BASE_URL: str = os.getenv(
    "QWEN_BASE_URL",
    "https://dashscope.aliyuncs.com/compatible-mode/v1",
)
QWEN_MODEL: str = os.getenv("QWEN_MODEL", "qwen-long")

# 长文本 qwen-long 最大输出 token
MAX_OUTPUT_TOKENS: int = 4096
MAP_MAX_CONCURRENCY: int = int(os.getenv("MAP_MAX_CONCURRENCY", "4"))

logger = logging.getLogger("pharma-mcp.llm")


def _get_client() -> OpenAI:
    if not QWEN_API_KEY:
        logger.error("未配置 DASHSCOPE_API_KEY，无法初始化同步 Qwen 客户端")
        raise RuntimeError(
            "未配置 DASHSCOPE_API_KEY，请在 .env 文件中设置或导出环境变量。"
        )
    logger.debug("初始化同步 Qwen 客户端，base_url=%s", QWEN_BASE_URL)
    return OpenAI(api_key=QWEN_API_KEY, base_url=QWEN_BASE_URL)


def _get_async_client() -> AsyncOpenAI:
    if not QWEN_API_KEY:
        logger.error("未配置 DASHSCOPE_API_KEY，无法初始化异步 Qwen 客户端")
        raise RuntimeError(
            "未配置 DASHSCOPE_API_KEY，请在 .env 文件中设置或导出环境变量。"
        )
    logger.debug("初始化异步 Qwen 客户端，base_url=%s", QWEN_BASE_URL)
    return AsyncOpenAI(api_key=QWEN_API_KEY, base_url=QWEN_BASE_URL)


async def _invoke_callback(
    callback: Callable[..., Any] | None,
    *args: Any,
    **kwargs: Any,
) -> None:
    if callback is None:
        return
    result = callback(*args, **kwargs)
    if inspect.isawaitable(result):
        await result


# ── 基础调用 ──────────────────────────────────────────────────────────────────

def chat_completion(
    system_prompt: str,
    user_message: str,
    model: str = QWEN_MODEL,
    temperature: float = 0.1,
    max_tokens: int = MAX_OUTPUT_TOKENS,
) -> str:
    """单轮对话，返回模型输出文本。"""
    logger.info(
        "同步调用 chat_completion: model=%s temperature=%.2f max_tokens=%d",
        model,
        temperature,
        max_tokens,
    )
    client = _get_client()
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    content = response.choices[0].message.content or ""
    logger.info("同步调用完成: 返回长度=%d", len(content))
    return content


async def chat_completion_async(
    system_prompt: str,
    user_message: str,
    model: str = QWEN_MODEL,
    temperature: float = 0.1,
    max_tokens: int = MAX_OUTPUT_TOKENS,
) -> str:
    """异步单轮对话，返回模型输出文本。"""
    logger.info(
        "异步调用 chat_completion_async: model=%s temperature=%.2f max_tokens=%d",
        model,
        temperature,
        max_tokens,
    )
    client = _get_async_client()
    response = await client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    content = response.choices[0].message.content or ""
    logger.info("异步调用完成: 返回长度=%d", len(content))
    return content


async def chat_completion_stream(
    system_prompt: str,
    user_message: str,
    model: str = QWEN_MODEL,
    temperature: float = 0.1,
    max_tokens: int = MAX_OUTPUT_TOKENS,
    on_delta: Callable[[str], Any] | None = None,
) -> str:
    """异步流式调用；返回完整文本，并在增量到达时触发 on_delta。"""
    logger.info(
        "异步流式调用 chat_completion_stream: model=%s temperature=%.2f max_tokens=%d",
        model,
        temperature,
        max_tokens,
    )
    client = _get_async_client()
    stream = await client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
        stream=True,
    )

    parts: list[str] = []
    async for event in stream:
        choices = getattr(event, "choices", None)
        if not choices:
            continue
        delta = getattr(choices[0], "delta", None)
        content = getattr(delta, "content", "") if delta else ""
        if not content:
            continue
        parts.append(content)
        await _invoke_callback(on_delta, content)

    result = "".join(parts)
    logger.info("异步流式调用完成: 返回长度=%d", len(result))
    return result


# ── Map-Reduce 长文本分析 ──────────────────────────────────────────────────────

def analyze_chunks_map_reduce(
    chunks: list[str],
    map_system_prompt: str,
    map_user_template: str,
    reduce_system_prompt: str,
    reduce_user_template: str,
    model: str = QWEN_MODEL,
) -> str:
    """
    Map-Reduce 策略处理长文档：

    Map 阶段：对每个文本块独立分析，提取局部问题/信息
    Reduce 阶段：将所有局部结果汇总，生成最终综合报告

    Args:
        chunks: 文本块列表
        map_system_prompt: Map 阶段系统提示词
        map_user_template: Map 阶段用户提示词模板，需含 {chunk_index}、{total_chunks}、{chunk_text} 占位符
        reduce_system_prompt: Reduce 阶段系统提示词
        reduce_user_template: Reduce 阶段用户提示词模板，需含 {all_chunk_results} 占位符
        model: 使用的模型名称

    Returns:
        最终综合分析报告字符串
    """
    logger.info("同步 analyze_chunks_map_reduce 开始: chunks=%d", len(chunks))
    try:
        return asyncio.run(
            analyze_chunks_map_reduce_async(
                chunks=chunks,
                map_system_prompt=map_system_prompt,
                map_user_template=map_user_template,
                reduce_system_prompt=reduce_system_prompt,
                reduce_user_template=reduce_user_template,
                model=model,
            )
        )
    except RuntimeError:
        logger.warning("检测到运行中事件循环，回退到同步串行处理")
        total = len(chunks)
        chunk_results: list[str] = []
        for i, chunk_text in enumerate(chunks):
            user_msg = map_user_template.format(
                chunk_index=i + 1,
                total_chunks=total,
                chunk_text=chunk_text,
            )
            result = chat_completion(
                system_prompt=map_system_prompt,
                user_message=user_msg,
                model=model,
            )
            chunk_results.append(f"【第 {i + 1}/{total} 块分析结果】\n{result}")

        all_chunk_results = "\n\n".join(chunk_results)
        reduce_user_msg = reduce_user_template.format(
            all_chunk_results=all_chunk_results,
            total_chunks=total,
        )
        return chat_completion(
            system_prompt=reduce_system_prompt,
            user_message=reduce_user_msg,
            model=model,
            max_tokens=MAX_OUTPUT_TOKENS,
        )


async def analyze_chunks_map_reduce_async(
    chunks: list[str],
    map_system_prompt: str,
    map_user_template: str,
    reduce_system_prompt: str,
    reduce_user_template: str,
    model: str = QWEN_MODEL,
    max_concurrency: int = MAP_MAX_CONCURRENCY,
    on_map_progress: Callable[[int, int], Any] | None = None,
    on_reduce_delta: Callable[[str], Any] | None = None,
) -> str:
    """异步 Map-Reduce：Map 并发聚合 + Reduce 流式生成。"""
    total = len(chunks)
    if total == 0:
        logger.warning("analyze_chunks_map_reduce_async 收到空 chunks")
        return ""

    concurrency = max(1, min(max_concurrency, total))
    logger.info(
        "异步 Map-Reduce 开始: chunks=%d max_concurrency=%d model=%s",
        total,
        concurrency,
        model,
    )

    semaphore = asyncio.Semaphore(concurrency)
    chunk_results: list[str] = [""] * total
    completed = 0

    async def _map_one(index: int, chunk_text: str) -> None:
        nonlocal completed
        async with semaphore:
            user_msg = map_user_template.format(
                chunk_index=index + 1,
                total_chunks=total,
                chunk_text=chunk_text,
            )
            result = await chat_completion_async(
                system_prompt=map_system_prompt,
                user_message=user_msg,
                model=model,
            )
            chunk_results[index] = f"【第 {index + 1}/{total} 块分析结果】\n{result}"
            completed += 1
            logger.info("Map 阶段进度: %d/%d", completed, total)
            await _invoke_callback(on_map_progress, completed, total)

    await asyncio.gather(*[_map_one(i, text) for i, text in enumerate(chunks)])

    all_chunk_results = "\n\n".join(chunk_results)
    reduce_user_msg = reduce_user_template.format(
        all_chunk_results=all_chunk_results,
        total_chunks=total,
    )
    logger.info("Map 阶段完成，进入 Reduce 阶段")
    return await chat_completion_stream(
        system_prompt=reduce_system_prompt,
        user_message=reduce_user_msg,
        model=model,
        max_tokens=MAX_OUTPUT_TOKENS,
        on_delta=on_reduce_delta,
    )


def analyze_single_chunk(
    system_prompt: str,
    user_template: str,
    chunk_text: str,
    model: str = QWEN_MODEL,
) -> str:
    """对单个文本块进行分析（文档较短时直接调用，无需 Map-Reduce）。"""
    user_msg = user_template.format(chunk_text=chunk_text)
    return chat_completion(system_prompt=system_prompt, user_message=user_msg, model=model)


async def analyze_single_chunk_async(
    system_prompt: str,
    user_template: str,
    chunk_text: str,
    model: str = QWEN_MODEL,
    on_delta: Callable[[str], Any] | None = None,
) -> str:
    """异步单块分析；提供 on_delta 时使用流式输出。"""
    user_msg = user_template.format(chunk_text=chunk_text)
    if on_delta is not None:
        return await chat_completion_stream(
            system_prompt=system_prompt,
            user_message=user_msg,
            model=model,
            on_delta=on_delta,
        )
    return await chat_completion_async(
        system_prompt=system_prompt,
        user_message=user_msg,
        model=model,
    )
