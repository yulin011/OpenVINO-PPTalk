from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import gradio as gr

from ppt_utils import PptRenderError, embed_slide_wavs_to_pptx, export_pptx_to_pngs
from tts_narration import package_outputs_zip, save_slides_json, synthesize_slides_to_wavs
from vlm_script import generate_slide_drafts, save_drafts_json
from llm_polish import OpenAIConfig, load_openai_config, polish_slide_scripts


def _run_id() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def make_demo(
    *,
    vlm_model: Any,
    vlm_processor: Any,
    tts_model: Any,
    default_outputs_dir: Path,
) -> gr.Blocks:
    default_outputs_dir = Path(default_outputs_dir)
    default_outputs_dir.mkdir(parents=True, exist_ok=True)

    def pipeline(
        ppt_file,
        render_backend,
        seconds_target,
        add_transition,
        audience,
        openai_base_url,
        openai_api_key,
        openai_model,
        speaker,
        language,
        tts_style,
        progress=gr.Progress(track_tqdm=True),
    ):
        if ppt_file is None:
            raise gr.Error("请先上传 PPT/PPTX。")

        ppt_path = Path(ppt_file.name if hasattr(ppt_file, "name") else str(ppt_file))
        run_id = _run_id()
        run_dir = default_outputs_dir / run_id
        slides_dir = run_dir / "slides"
        audio_dir = run_dir / "audio"
        scripts_dir = run_dir / "scripts"
        scripts_dir.mkdir(parents=True, exist_ok=True)

        # A) PPT -> PNGs
        progress(0.05, desc="导出 PPT 为图片")
        try:
            slides, manifest_path = export_pptx_to_pngs(ppt_path, slides_dir, backend=render_backend, prefer_dpi=200)
        except PptRenderError as e:
            raise gr.Error(str(e))

        slide_paths = [s.path for s in slides]

        # B) VLM per-slide draft
        progress(0.25, desc="逐页生成讲解稿（初稿）")
        drafts = generate_slide_drafts(
            model=vlm_model,
            processor=vlm_processor,
            slide_image_paths=slide_paths,
            seconds_target=seconds_target,
            add_transition=add_transition,
            audience=audience,
        )
        drafts_json = scripts_dir / "slides_drafts.json"
        save_drafts_json(drafts, drafts_json)

        # C) LLM polish (OpenAI compatible)
        progress(0.55, desc="全局梳理校验（LLM）")
        cfg = load_openai_config(base_url=openai_base_url, api_key=openai_api_key, model=openai_model)
        slides_in = [
            {
                "slide_index": d.slide_index,
                "slide_title": d.slide_title,
                "key_points": d.key_points,
                "speaker_notes": d.speaker_notes,
            }
            for d in drafts
        ]
        polished_slides, global_outline_md = polish_slide_scripts(
            slides=slides_in,
            config=cfg,
            seconds_target=seconds_target,
            add_transitions=add_transition,
        )
        polished_json = scripts_dir / "slides_polished.json"
        save_slides_json(polished_slides, polished_json)
        outline_path = scripts_dir / "global_outline.md"
        outline_path.write_text(global_outline_md, encoding="utf-8")

        # D) TTS per-slide
        progress(0.75, desc="逐页生成语音（TTS）")
        audios = synthesize_slides_to_wavs(
            ov_tts_model=tts_model,
            slides=polished_slides,
            out_dir=audio_dir,
            language=language,
            speaker=speaker,
            style_instruct=tts_style,
        )

        # E) Embed audio into PPTX (PowerPoint COM, Windows-only)
        narrated_pptx: Optional[Path] = None
        try:
            narrated_pptx = embed_slide_wavs_to_pptx(
                ppt_path=ppt_path,
                audio_dir=audio_dir,
                out_pptx=run_dir / f"narrated_{ppt_path.stem}_{run_id}.pptx",
                autoplay=True,
                hide_icon=True,
            )
        except Exception:
            # Best-effort: keep pipeline usable even if PowerPoint is not available.
            narrated_pptx = None

        # F) zip
        progress(0.95, desc="打包输出")
        out_zip = run_dir / f"package_{run_id}.zip"
        extra = [manifest_path, outline_path]
        if narrated_pptx is not None and narrated_pptx.exists():
            extra.append(narrated_pptx)
        package_outputs_zip(
            slides_dir=slides_dir,
            scripts_json=polished_json,
            audio_dir=audio_dir,
            out_zip=out_zip,
            extra_files=extra,
        )

        # UI outputs
        gallery = [str(p) for p in slide_paths]
        scripts_text = json.dumps(polished_slides, ensure_ascii=False, indent=2)

        # per-slide audio selection: return first audio if exists
        first_audio = None
        if audios:
            first_audio = str(audios[0].wav_path)

        return gallery, scripts_text, global_outline_md, first_audio, str(out_zip)

    with gr.Blocks(title="Lab5：PPT 讲解稿 + 逐页语音") as demo:
        gr.Markdown(
            """
## Lab5：PPT 讲解稿 + 逐页语音

流程：上传 PPT → 导出每页图片 → VLM 逐页生成讲解稿（初稿）→ OpenAI 兼容 LLM 全局梳理 → Qwen3-TTS 逐页生成语音 → 打包下载。
"""
        )

        with gr.Row():
            ppt_file = gr.File(label="上传 PPT / PPTX", file_types=[".ppt", ".pptx"])
            render_backend = gr.Dropdown(
                label="PPT 渲染后端",
                choices=["auto", "powerpoint", "libreoffice"],
                value="auto",
            )

        with gr.Row():
            seconds_target = gr.Dropdown(label="每页讲解长度", choices=["short", "medium", "long"], value="medium")
            add_transition = gr.Checkbox(label="加入页间自然过渡", value=True)
            audience = gr.Textbox(label="听众类型", value="一般技术听众")

        gr.Markdown("### 3.3 全局梳理校验（OpenAI 兼容接口）")
        with gr.Row():
            openai_base_url = gr.Textbox(
                label="OPENAI_BASE_URL",
                value="",
                placeholder="https://host 或 https://host/v1",
            )
            openai_model = gr.Textbox(label="OPENAI_MODEL", value="", placeholder="模型名 / id")
        openai_api_key = gr.Textbox(label="OPENAI_API_KEY", value="", placeholder="如不填则从环境变量读取", type="password")

        gr.Markdown("### 3.4 文转语音（Qwen3-TTS）")
        with gr.Row():
            speaker = gr.Textbox(label="Speaker", value="vivian")
            language = gr.Dropdown(label="Language", choices=["Chinese", "English", "Japanese", "Korean", "French", "German", "Spanish", "Portuguese", "Russian", "Auto"], value="Chinese")
        tts_style = gr.Textbox(label="TTS 风格指令（instruct）", value="用清晰、自然、适合讲解的语气朗读。", lines=2)

        run_btn = gr.Button("一键生成", variant="primary")

        with gr.Row():
            gallery = gr.Gallery(label="逐页图片预览", columns=3, height=320)
        with gr.Row():
            scripts = gr.Code(label="逐页讲解稿（JSON，可复制/二次编辑）", language="json")
        outline = gr.Markdown(label="全局讲解提纲（Markdown）")

        with gr.Row():
            audio_preview = gr.Audio(label="音频预览（默认第 1 页）", type="filepath")
            zip_out = gr.File(label="打包下载（zip）")

        run_btn.click(
            pipeline,
            inputs=[
                ppt_file,
                render_backend,
                seconds_target,
                add_transition,
                audience,
                openai_base_url,
                openai_api_key,
                openai_model,
                speaker,
                language,
                tts_style,
            ],
            outputs=[gallery, scripts, outline, audio_preview, zip_out],
        )

    return demo

