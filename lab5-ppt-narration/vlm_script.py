from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class SlideDraft:
    slide_index: int
    image_path: Path
    slide_title: Optional[str]
    key_points: List[str]
    speaker_notes: str
    do_not_say: List[str]
    raw_text: str


DEFAULT_VLM_SYSTEM = (
    "你是一位专业的中文技术讲师。你将根据 PPT 单页内容，为现场听众生成可口播的讲解稿。"
)


def build_slide_prompt(
    *,
    slide_index: int,
    total_slides: Optional[int] = None,
    seconds_target: str = "medium",
    add_transition: bool = True,
    audience: str = "一般技术听众",
) -> str:
    # time budget guidance
    if seconds_target == "short":
        length_rule = "控制在约 20-40 秒的口播长度。"
    elif seconds_target == "long":
        length_rule = "控制在约 70-120 秒的口播长度。"
    else:
        length_rule = "控制在约 40-80 秒的口播长度。"

    pos = f"第 {slide_index} 页" + (f" / 共 {total_slides} 页" if total_slides else "")

    return (
        f"你正在讲解一份 PPT 的 {pos}。\n"
        f"听众类型：{audience}。\n"
        "任务：根据这一页的内容生成“演讲稿”，不是逐字念 PPT。\n"
        "规则：\n"
        "- 忽略页角 logo、装饰性背景、页眉页脚、版权声明、重复的公司介绍等冗余信息。\n"
        "- 若出现图表/数据：先给一句话结论，再解释关键趋势/对比，最后给 1 条 takeaway。\n"
        "- 若出现流程/架构图：先讲目的，再按顺序讲关键节点与输入输出。\n"
        "- 若出现列表：不要逐条机械复读，合并成 3-6 个要点。\n"
        f"- {length_rule}\n"
        "- 输出必须是严格 JSON（不要包含 ```），字段为：\n"
        "  {\"slide_title\": string|null, \"key_points\": [string], \"speaker_notes\": string, \"do_not_say\": [string]}\n"
        + (
            "- speaker_notes 末尾加 1 句自然过渡到下一页的引导语。\n" if add_transition else ""
        )
    )


def generate_slide_drafts(
    *,
    model: Any,
    processor: Any,
    slide_image_paths: List[Path],
    seconds_target: str = "medium",
    add_transition: bool = True,
    audience: str = "一般技术听众",
    max_new_tokens: int = 700,
) -> List[SlideDraft]:
    """
    Use Qwen3-VL (OpenVINO + Optimum Intel) to generate structured narration drafts per slide.
    """
    drafts: List[SlideDraft] = []
    total = len(slide_image_paths)

    for i, img_path in enumerate(slide_image_paths, start=1):
        prompt = build_slide_prompt(
            slide_index=i, total_slides=total, seconds_target=seconds_target, add_transition=add_transition, audience=audience
        )

        messages = [
            {"role": "system", "content": [{"type": "text", "text": DEFAULT_VLM_SYSTEM}]},
            {
                "role": "user",
                "content": [
                    {"type": "image", "image": str(img_path)},
                    {"type": "text", "text": prompt},
                ],
            },
        ]

        inputs = processor.apply_chat_template(
            messages, tokenize=True, add_generation_prompt=True, return_dict=True, return_tensors="pt"
        )
        out_ids = model.generate(**inputs, max_new_tokens=max_new_tokens)

        # Decode only the newly generated tokens to avoid including the prompt/template.
        input_len = int(inputs["input_ids"].shape[-1])
        gen_ids = out_ids[0][input_len:]
        raw_text = processor.tokenizer.decode(gen_ids, skip_special_tokens=True)

        parsed = _best_effort_parse_json(raw_text)
        drafts.append(
            SlideDraft(
                slide_index=i,
                image_path=Path(img_path),
                slide_title=parsed.get("slide_title"),
                key_points=list(parsed.get("key_points") or []),
                speaker_notes=str(parsed.get("speaker_notes") or "").strip(),
                do_not_say=list(parsed.get("do_not_say") or []),
                raw_text=raw_text,
            )
        )

    return drafts


def save_drafts_json(drafts: List[SlideDraft], path: Path) -> None:
    obj: List[Dict[str, Any]] = []
    for d in drafts:
        obj.append(
            {
                "slide_index": d.slide_index,
                "image_path": str(d.image_path),
                "slide_title": d.slide_title,
                "key_points": d.key_points,
                "speaker_notes": d.speaker_notes,
                "do_not_say": d.do_not_say,
                "raw_text": d.raw_text,
            }
        )
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def _best_effort_parse_json(text: str) -> Dict[str, Any]:
    """
    The model may include extra tokens. This function extracts the first JSON object.
    """
    text = text.strip()
    # Extract the first valid JSON object by scanning braces.
    start = text.find("{")
    if start == -1:
        return {"slide_title": None, "key_points": [], "speaker_notes": text, "do_not_say": []}

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
                except Exception:
                    break

    return {"slide_title": None, "key_points": [], "speaker_notes": text, "do_not_say": []}

