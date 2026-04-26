from __future__ import annotations

"""逐页讲稿润色：用 openai SDK 按 Chat Completions 形态请求 OpenAI 兼容接口。"""

import json
import os
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import openai


class LlmPolishError(RuntimeError):
    pass


@dataclass(frozen=True)
class OpenAIConfig:
    base_url: str
    api_key: str
    model: str


def load_openai_config(
    *,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> OpenAIConfig:
    base_url = (base_url or os.environ.get("OPENAI_BASE_URL") or "").strip()
    api_key = (api_key or os.environ.get("OPENAI_API_KEY") or "").strip()
    model = (model or os.environ.get("OPENAI_MODEL") or "").strip()
    for name, val in (("OPENAI_API_KEY", api_key), ("OPENAI_MODEL", model), ("OPENAI_BASE_URL", base_url)):
        if val is not None and not isinstance(val, str):
            raise LlmPolishError(
                f"{name} 必须是字符串。是否多写了逗号变成元组？"
                "正确：api_key = 'xxx'；错误：api_key = 'OPENAI_API_KEY', 'xxx'。"
            )
    if not base_url:
        raise LlmPolishError("未设置 OPENAI_BASE_URL（OpenAI 兼容接口服务地址，如 https://host 或 https://host/v1）。")
    if not api_key:
        raise LlmPolishError("未设置 OPENAI_API_KEY（鉴权密钥）。")
    if not model:
        raise LlmPolishError("未设置 OPENAI_MODEL（模型名 / id）。")
    return OpenAIConfig(base_url=base_url.rstrip("/"), api_key=api_key, model=model)


def _sdk_base_url(base_url: str) -> str:
    b = base_url.rstrip("/")
    if b.endswith("/v1"):
        return b
    return f"{b}/v1"


def _openai_client(cfg: OpenAIConfig) -> openai.OpenAI:
    return openai.OpenAI(api_key=cfg.api_key, base_url=_sdk_base_url(cfg.base_url))


def polish_slide_scripts(
    *,
    slides: List[Dict[str, Any]],
    config: OpenAIConfig,
    seconds_target: str = "medium",
    add_transitions: bool = True,
    temperature: float = 0.2,
) -> Tuple[List[Dict[str, Any]], str]:
    if seconds_target == "short":
        length_rule = "每页口播控制在约 20-40 秒。"
    elif seconds_target == "long":
        length_rule = "每页口播控制在约 70-120 秒。"
    else:
        length_rule = "每页口播控制在约 40-80 秒。"

    system = (
        "你是一位专业的演讲稿编辑。你将对逐页讲解稿进行全局梳理、术语统一、去重、补充必要过渡。"
        "你必须保持事实与 PPT 内容一致，不要编造具体数值/结论。"
        "只输出一个合法 JSON 对象本体：不要 Markdown 代码围栏；所有字符串值里的英文双引号必须写成 \\\" 或改用中文引号「」。"
    )
    user = {
        "task": "请对逐页讲解稿进行全局梳理与校验，输出严格 JSON。",
        "rules": [
            "不要新增 PPT 中不存在的关键事实；如不确定，用更保守表述。",
            "跨页去重：同一概念不要在每页重复定义。",
            "术语统一：同一对象在全篇用同一叫法。",
            "风格：口语化但专业，适合现场讲解。",
            length_rule,
            "输出字段：{\"slides\": [{\"slide_index\": number, \"speaker_notes\": string}], \"global_outline_md\": string}",
            "speaker_notes 以段落组织，避免超长句；其中禁止出现未转义的英文双引号 \"，需要引号时用「」或『』。",
            "如果 add_transitions=true，则在相邻页之间加入自然衔接（但不要每页都重复“下一页我们将…”）。",
        ],
        "add_transitions": add_transitions,
        "slides": slides,
    }

    resp = openai_chat_completions(
        config=config,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": json.dumps(user, ensure_ascii=False)}],
        temperature=temperature,
        max_tokens=16000,
        json_mode=True,
    )

    parsed = _parse_model_json(resp)
    polished = parsed.get("slides")
    outline = parsed.get("global_outline_md", "")
    if not isinstance(polished, list):
        raise LlmPolishError("LLM 输出解析失败：未找到 slides 列表。")

    polished_slides: List[Dict[str, Any]] = []
    by_index = {int(s.get("slide_index")): s for s in polished if isinstance(s, dict) and "slide_index" in s}
    for s in slides:
        idx = int(s.get("slide_index"))
        new_notes = by_index.get(idx, {}).get("speaker_notes", s.get("speaker_notes", ""))
        new_s = dict(s)
        new_s["speaker_notes"] = str(new_notes).strip()
        polished_slides.append(new_s)
    return polished_slides, str(outline)


def openai_chat_completions(
    *,
    config: OpenAIConfig,
    messages: List[Dict[str, Any]],
    temperature: float = 0.2,
    max_tokens: int = 2048,
    json_mode: bool = False,
) -> str:
    client = _openai_client(config)
    kwargs: Dict[str, Any] = {
        "model": config.model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "timeout": 300.0,
    }
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    try:
        completion = client.chat.completions.create(**kwargs)
    except Exception as e:
        if json_mode:
            kwargs.pop("response_format", None)
            try:
                completion = client.chat.completions.create(**kwargs)
            except Exception as e2:
                raise LlmPolishError(f"调用 LLM 接口失败：{type(e2).__name__}: {e2}") from e2
        else:
            raise LlmPolishError(f"调用 LLM 接口失败：{type(e).__name__}: {e}") from e

    try:
        content = completion.choices[0].message.content
    except (AttributeError, IndexError, KeyError) as e:
        raise LlmPolishError(f"解析接口返回失败：{type(e).__name__}: {e}") from e

    if content is None or not str(content).strip():
        raise LlmPolishError("模型返回空内容。")
    return str(content)


def _parse_model_json(text: str) -> Dict[str, Any]:
    """优先整段 json.loads；兼容 ```json 围栏；失败再退回括号扫描。"""
    t = (text or "").strip()
    if t.startswith("```"):
        first_nl = t.find("\n")
        if first_nl != -1:
            t = t[first_nl + 1 :]
        if t.rstrip().endswith("```"):
            t = t.rstrip()[:-3].rstrip()
    try:
        obj = json.loads(t)
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    return _extract_first_json(t)


def _extract_first_json(text: str) -> Dict[str, Any]:
    text = (text or "").strip()
    start = text.find("{")
    if start == -1:
        raise LlmPolishError("LLM 输出不包含 JSON 对象。")

    depth = 0
    for i in range(start, len(text)):
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                candidate = text[start : i + 1]
                try:
                    return json.loads(candidate)
                except Exception as e:
                    raise LlmPolishError(
                        f"LLM JSON 解析失败：{type(e).__name__}: {e}\nCandidate: {candidate[:500]}"
                    ) from e

    raise LlmPolishError("LLM 输出包含 '{' 但未形成完整 JSON。")

