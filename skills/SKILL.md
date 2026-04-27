---
name: pptalk
description: Runs the PPTalk (lab5-ppt-narration) pipeline in a Linux-only environment (no Microsoft PowerPoint). Downloads required OpenVINO model artifacts from ModelScope, converts PPT/PPTX to per-slide PNGs via LibreOffice, generates per-slide narration drafts with Qwen3-VL, performs global polishing using an OpenAI-compatible API, synthesizes per-slide WAVs with Qwen3-TTS, and packages outputs into a zip for download. Use when the user uploads a PPT/PPTX and asks for slide-by-slide scripts + audio + a downloadable zip package.
---

# PPTalk（Lab5）Linux 一键旁白打包 Skill

## 目标

在**纯 Linux 环境**（无 Microsoft PowerPoint）中，实现与你的项目 `lab5-ppt-narration/` 一致的交付流程：

- 输入：用户上传的 `.ppt` / `.pptx`
- 输出：逐页图片、逐页讲解稿（JSON）、全局提纲（Markdown）、逐页音频（WAV）、最终打包 `zip`

关键约束：

- **不依赖 PowerPoint**：PPT 渲染必须走 `LibreOffice (soffice) -> PDF -> PNG` 的兜底链路（见 `lab5-ppt-narration/ppt_utils.py`）。
- **必须使用仓库现有脚本**：优先调用 `lab5-ppt-narration/*.py` 内的实现，而不是重写核心逻辑。
- **模型很大**：需要从 ModelScope 下载 OpenVINO IR 模型目录，下载可能很慢，需提示用户耐心等待。
- **全局梳理必须可用**：本 Skill 需要用户提供 OpenAI 兼容接口配置（`OPENAI_BASE_URL/OPENAI_API_KEY/OPENAI_MODEL`），否则视为不可运行。

---

## 运行前准备（一次性）

### 0) 目录约定

以下步骤假设工作目录为 `$WORKDIR`。

---

### 1) 拉取你的仓库代码（必须）

```bash
git clone https://github.com/yulin011/OpenVINO-PPTalk.git
cd OpenVINO-PPTalk
```

> 说明：仓库里包含 `lab5-ppt-narration/` 的全部 Python 脚本，但**不包含**大模型目录（它们被 `.gitignore` 忽略）。

---

### 2) 安装系统依赖（Linux，无 PowerPoint）

必须安装：
- `libreoffice`：提供 `soffice`，用于 PPT->PDF
- `poppler-utils`：供 `pdf2image` 使用（Linux 通常通过 `pdftoppm`）

Ubuntu / Debian：

```bash
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

快速自检：

```bash
command -v soffice
command -v pdftoppm
```

---

### 3) Python 环境与依赖安装

建议用 venv（Python 3.10+）：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
```

安装项目依赖（按仓库为准）：

```bash
pip install -r requirements.txt
```

额外确保（Linux 渲染链路需要）：

```bash
pip install pdf2image
```

---

## 模型准备（必须，从 ModelScope 下载）

PPTalk 默认需要两个本地目录（与 `lab5-ppt-narration/run_pptalk_gradio.py` 一致）：

- `lab5-ppt-narration/Qwen3-VL-4B-Instruct-int4-ov/`
- `lab5-ppt-narration/Qwen3-TTS-CustomVoice-0.6B-fp16-ov/`

使用 ModelScope 下载（模型很大，可能需要较长时间；网络不稳定时建议重试）：

```bash
python - <<'PY'
from modelscope.hub.snapshot_download import snapshot_download
from pathlib import Path

root = Path("lab5-ppt-narration")
root.mkdir(parents=True, exist_ok=True)

print("Downloading VLM (OpenVINO IR) ...")
snapshot_download(
    "snake7gun/Qwen3-VL-4B-Instruct-int4-ov",
    local_dir=str(root / "Qwen3-VL-4B-Instruct-int4-ov"),
    local_dir_use_symlinks=False,
)

print("Downloading TTS (OpenVINO IR) ...")
snapshot_download(
    "snake7gun/Qwen3-TTS-CustomVoice-0.6B-fp16-ov",
    local_dir=str(root / "Qwen3-TTS-CustomVoice-0.6B-fp16-ov"),
    local_dir_use_symlinks=False,
)
print("Done.")
PY
```

下载完成后应能看到（示例）：

```bash
ls -la lab5-ppt-narration/Qwen3-VL-4B-Instruct-int4-ov | head
ls -la lab5-ppt-narration/Qwen3-TTS-CustomVoice-0.6B-fp16-ov | head
```

---

## 必需的外部接口（OpenAI 兼容，全局梳理阶段）

本 Skill **强制要求**设置以下环境变量（用于 `lab5-ppt-narration/llm_polish.py`）：

```bash
export OPENAI_BASE_URL="https://your-host-or-gateway/v1"
export OPENAI_API_KEY="your_key"
export OPENAI_MODEL="your_model_id"
```

若缺失其中任意一个，停止执行并提示用户补齐配置。

---

## 执行流程（每次处理一个 PPT/PPTX）

### 输入约定

假设用户上传到本机路径：

- `INPUT_PPT=/path/to/input.pptx`

---

### 方案 A（推荐，纯命令行 / Headless）：直接跑“脚本流水线”并产出 zip

该方案不会启动 Gradio，适合“智能体在 Linux 上跑完就交付产物”的场景。

1) 在仓库根目录创建一个临时运行脚本 `lab5-ppt-narration/run_headless.py`（只做编排，不重写核心逻辑）：

```bash
cat > lab5-ppt-narration/run_headless.py <<'PY'
from __future__ import annotations

import os
import time
from pathlib import Path

from transformers import AutoProcessor
from optimum.intel.openvino import OVModelForVisualCausalLM

from ppt_utils import export_pptx_to_pngs, PptRenderError
from vlm_script import generate_slide_drafts, save_drafts_json
from llm_polish import load_openai_config, polish_slide_scripts
from tts_narration import synthesize_slides_to_wavs, save_slides_json, package_outputs_zip
from qwen_3_tts_helper import OVQwen3TTSModel


def _require_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env: {name}")
    return v


def main() -> None:
    import argparse

    p = argparse.ArgumentParser()
    p.add_argument("--ppt", required=True, help="Path to .ppt/.pptx")
    p.add_argument("--device", default="CPU", help="OpenVINO device: CPU, GPU.0, AUTO ...")
    p.add_argument("--seconds-target", default="medium", choices=["short", "medium", "long"])
    p.add_argument("--add-transition", action="store_true")
    p.add_argument("--audience", default="一般技术听众")
    p.add_argument("--speaker", default="vivian")
    p.add_argument("--language", default="Chinese")
    p.add_argument("--tts-style", default="用清晰、自然、适合讲解的语气朗读。")
    p.add_argument("--render-backend", default="libreoffice", choices=["auto", "powerpoint", "libreoffice"])
    p.add_argument("--outputs-dir", default=str(Path(__file__).parent / "outputs"))
    p.add_argument("--vlm-dir", default=str(Path(__file__).parent / "Qwen3-VL-4B-Instruct-int4-ov"))
    p.add_argument("--tts-dir", default=str(Path(__file__).parent / "Qwen3-TTS-CustomVoice-0.6B-fp16-ov"))
    args = p.parse_args()

    # Require OpenAI-compatible config for global polish
    base_url = _require_env("OPENAI_BASE_URL")
    api_key = _require_env("OPENAI_API_KEY")
    model = _require_env("OPENAI_MODEL")

    ppt_path = Path(args.ppt).expanduser().resolve()
    if not ppt_path.exists():
        raise FileNotFoundError(str(ppt_path))

    run_id = time.strftime("%Y%m%d_%H%M%S")
    outputs_dir = Path(args.outputs_dir).expanduser().resolve() / run_id
    slides_dir = outputs_dir / "slides"
    scripts_dir = outputs_dir / "scripts"
    audio_dir = outputs_dir / "audio"
    scripts_dir.mkdir(parents=True, exist_ok=True)

    # A) PPT -> PNGs (Linux uses LibreOffice backend)
    try:
        slides, manifest_path = export_pptx_to_pngs(ppt_path, slides_dir, backend=args.render_backend, prefer_dpi=200)
    except PptRenderError as e:
        raise RuntimeError(str(e))

    slide_paths = [s.path for s in slides]

    # B) VLM per-slide draft
    vlm_dir = Path(args.vlm_dir)
    vlm_model = OVModelForVisualCausalLM.from_pretrained(vlm_dir, device=str(args.device))
    vlm_processor = AutoProcessor.from_pretrained(
        vlm_dir,
        min_pixels=256 * 28 * 28,
        max_pixels=1280 * 28 * 28,
        fix_mistral_regex=True,
    )
    drafts = generate_slide_drafts(
        model=vlm_model,
        processor=vlm_processor,
        slide_image_paths=slide_paths,
        seconds_target=args.seconds_target,
        add_transition=bool(args.add_transition),
        audience=args.audience,
    )
    drafts_json = scripts_dir / "slides_drafts.json"
    save_drafts_json(drafts, drafts_json)

    # C) Global polish (OpenAI compatible)
    cfg = load_openai_config(base_url=base_url, api_key=api_key, model=model)
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
        seconds_target=args.seconds_target,
        add_transitions=bool(args.add_transition),
    )
    polished_json = scripts_dir / "slides_polished.json"
    save_slides_json(polished_slides, polished_json)
    outline_path = scripts_dir / "global_outline.md"
    outline_path.write_text(global_outline_md, encoding="utf-8")

    # D) TTS per-slide
    tts_dir = Path(args.tts_dir)
    tts_model = OVQwen3TTSModel.from_pretrained(model_dir=tts_dir, device=str(args.device))
    synthesize_slides_to_wavs(
        ov_tts_model=tts_model,
        slides=polished_slides,
        out_dir=audio_dir,
        language=args.language,
        speaker=args.speaker,
        style_instruct=args.tts_style,
    )

    # F) zip (E: embed to PPTX is Windows-only; skip on Linux)
    out_zip = outputs_dir / f"package_{run_id}.zip"
    package_outputs_zip(
        slides_dir=slides_dir,
        scripts_json=polished_json,
        audio_dir=audio_dir,
        out_zip=out_zip,
        extra_files=[manifest_path, outline_path],
    )

    print("OK")
    print(f"Outputs dir: {outputs_dir}")
    print(f"Zip: {out_zip}")


if __name__ == "__main__":
    main()
PY
```

2) 运行（示例）：

```bash
python lab5-ppt-narration/run_headless.py --ppt "$INPUT_PPT" --device CPU --render-backend libreoffice
```

3) 交付给用户：

- 输出目录：`lab5-ppt-narration/outputs/<run_id>/`
- 最终包：`lab5-ppt-narration/outputs/<run_id>/package_<run_id>.zip`

> 要求：把 `package_*.zip` 作为最终可下载产物返回给用户；如用户还要逐页文件，也可同时返回 `slides/`、`scripts/`、`audio/`。

---

### 方案 B（可选，Gradio）：启动网页上传并下载 zip

```bash
python lab5-ppt-narration/run_pptalk_gradio.py --host 0.0.0.0 --port 7860 --device CPU
```

注意：
- Linux 下不要选择 `powerpoint` 渲染后端；应使用 `auto`（会自动落到 LibreOffice）或直接用 `libreoffice`。
- 该方案依赖“有人访问 Web UI 并上传 PPT”，不适合完全离线/无人交互的智能体运行形态。

---

## 常见问题与排查（按优先级）

### 1) PPT 无法导出图片（最常见）

报错里若出现 “未检测到可用渲染后端” 或 “LibreOffice 转 PDF 失败”：

- 确认 `soffice` 存在：`command -v soffice`
- 确认 LibreOffice 可 headless 转 PDF：

```bash
soffice --headless --convert-to pdf --outdir /tmp "$INPUT_PPT"
ls -la /tmp/*.pdf
```

- 若 `pdf2image` 报 Poppler 缺失，确认 `pdftoppm`：`command -v pdftoppm`

---

### 2) 模型下载很慢 / 失败

提示用户：
- 模型目录体积大，首次下载耗时长是正常现象
- 网络波动可重试下载命令
- 确保磁盘空间充足

---

### 3) OpenAI 兼容接口不可用

本 Skill 强制依赖全局梳理阶段：
- 缺少 `OPENAI_BASE_URL/OPENAI_API_KEY/OPENAI_MODEL`：立即停止并提示补齐
- HTTP/鉴权错误：把错误日志回传给用户，让用户修复 endpoint / key / model

---

### 4) TTS 生成失败或超时

建议：
- 先用更短 PPT 验证端到端可行
- 适当降低并发（本项目是逐页串行生成）
- 若某一页过长，可先人工缩短该页 `speaker_notes` 再重跑

---

## 验收标准（必须满足）

- 成功生成 `package_<run_id>.zip`
- zip 内至少包含：
  - `slides/slide_0001.png ...`
  - `scripts/slides_polished.json`
  - `audio/slide_0001.wav ...`
- `global_outline.md` 作为额外文件包含在 zip 根目录（或与 scripts 同步返回给用户）

