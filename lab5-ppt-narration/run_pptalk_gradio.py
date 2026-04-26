from __future__ import annotations

import argparse
from pathlib import Path

from transformers import AutoProcessor


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="PPTalk Gradio demo (Lab5)")
    p.add_argument("--device", default="CPU", help="OpenVINO device, e.g. CPU, GPU.0, AUTO")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=7860)
    p.add_argument("--share", action="store_true", help="Enable Gradio share link")
    p.add_argument("--outputs-dir", default=str(Path(__file__).parent / "outputs"), help="Outputs directory")
    p.add_argument(
        "--vlm-dir",
        default=str(Path(__file__).parent / "Qwen3-VL-4B-Instruct-int4-ov"),
        help="Path to Qwen3-VL OpenVINO model directory",
    )
    p.add_argument(
        "--tts-dir",
        default=str(Path(__file__).parent / "Qwen3-TTS-CustomVoice-0.6B-fp16-ov"),
        help="Path to Qwen3-TTS OpenVINO model directory",
    )
    return p.parse_args()


def main() -> None:
    args = _parse_args()

    # Import-time noise control:
    # - torch/optimum may print environment probing messages during import
    import contextlib
    import io

    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        from optimum.intel.openvino import OVModelForVisualCausalLM
        from qwen_3_tts_helper import OVQwen3TTSModel

    from gradio_app import make_demo

    vlm_dir = Path(args.vlm_dir)
    tts_dir = Path(args.tts_dir)
    outputs_dir = Path(args.outputs_dir)

    vlm_model = OVModelForVisualCausalLM.from_pretrained(vlm_dir, device=str(args.device))

    min_pixels = 256 * 28 * 28
    max_pixels = 1280 * 28 * 28
    vlm_processor = AutoProcessor.from_pretrained(
        vlm_dir,
        min_pixels=min_pixels,
        max_pixels=max_pixels,
        fix_mistral_regex=True,
    )

    tts_model = OVQwen3TTSModel.from_pretrained(model_dir=tts_dir, device=str(args.device))

    demo = make_demo(vlm_model=vlm_model, vlm_processor=vlm_processor, tts_model=tts_model, default_outputs_dir=outputs_dir)
    demo.launch(server_name=args.host, server_port=args.port, share=bool(args.share))


if __name__ == "__main__":
    main()

