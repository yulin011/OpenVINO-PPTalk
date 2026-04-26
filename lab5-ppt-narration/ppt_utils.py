from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple
from contextlib import contextmanager


class PptRenderError(RuntimeError):
    pass


@dataclass(frozen=True)
class SlideImage:
    slide_index: int  # 1-based
    path: Path
    width: Optional[int] = None
    height: Optional[int] = None


def _now_run_id() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def _zero_pad_name(slide_index: int, width: int = 4, ext: str = ".png") -> str:
    return f"slide_{slide_index:0{width}d}{ext}"


def detect_powerpoint_com() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # type: ignore

        with _com_initialized():
            _ = win32com.client.Dispatch("PowerPoint.Application")
            try:
                _.Quit()
            except Exception:
                pass
            return True
    except Exception:
        return False


@contextmanager
def _com_initialized():
    """
    Ensure COM is initialized for the current thread.
    Gradio runs user functions in worker threads; PowerPoint COM requires CoInitialize per-thread.
    """
    if sys.platform != "win32":
        yield
        return
    try:
        import pythoncom  # type: ignore
    except Exception:
        # pywin32 not installed; caller will raise a clearer error later
        yield
        return

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def detect_soffice() -> Optional[str]:
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    return exe


def export_pptx_to_pngs(
    ppt_path: Path,
    out_dir: Path,
    *,
    backend: str = "auto",
    prefer_dpi: int = 200,
) -> Tuple[List[SlideImage], Path]:
    """
    Export PPT/PPTX pages to ordered PNGs.

    Returns: (slide_images, manifest_path)
    """
    ppt_path = Path(ppt_path).expanduser().resolve()
    if not ppt_path.exists():
        raise FileNotFoundError(str(ppt_path))
    if ppt_path.suffix.lower() not in {".ppt", ".pptx"}:
        raise ValueError(f"Unsupported file type: {ppt_path.suffix}")

    out_dir = Path(out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if backend not in {"auto", "powerpoint", "libreoffice"}:
        raise ValueError("backend must be one of: auto|powerpoint|libreoffice")

    tried = []
    if backend in {"auto", "powerpoint"}:
        tried.append("powerpoint")
        if detect_powerpoint_com():
            slide_paths = _export_with_powerpoint_com(ppt_path, out_dir)
            return _finalize_and_manifest(slide_paths, out_dir)

    if backend in {"auto", "libreoffice"}:
        tried.append("libreoffice")
        soffice = detect_soffice()
        if soffice:
            slide_paths = _export_with_libreoffice_pdf_pipeline(ppt_path, out_dir, soffice=soffice, dpi=prefer_dpi)
            return _finalize_and_manifest(slide_paths, out_dir)

    raise PptRenderError(
        "无法将 PPT 导出为图片：未检测到可用渲染后端。\n"
        f"已尝试: {', '.join(tried) or '无'}\n"
        "可选解决方案：\n"
        "- 安装 Microsoft PowerPoint（推荐，渲染一致性最好），并 `pip install pywin32`\n"
        "- 或安装 LibreOffice（提供 soffice 命令），并安装 Poppler + `pip install pdf2image`\n"
    )


def _export_with_powerpoint_com(ppt_path: Path, out_dir: Path) -> List[Path]:
    if sys.platform != "win32":
        raise PptRenderError("PowerPoint COM 仅支持 Windows。")
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        raise PptRenderError("未安装 pywin32，无法使用 PowerPoint COM。请先 `pip install pywin32`。") from e

    # PowerPoint export uses file names: Slide1.PNG, Slide2.PNG, ...
    app = None
    pres = None
    try:
        with _com_initialized():
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = 1
            pres = app.Presentations.Open(str(ppt_path), WithWindow=False)
            pres.Export(str(out_dir), "PNG")

        # On Windows, the filesystem is case-insensitive, so "*.png" and "*.PNG" can return
        # the same files. Use a single case-insensitive match and deduplicate by real path.
        candidates = list(out_dir.glob("*.png")) + list(out_dir.glob("*.PNG"))
        uniq: List[Path] = []
        seen = set()
        for p in candidates:
            key = str(p.resolve()).lower()
            if key in seen:
                continue
            seen.add(key)
            uniq.append(p)
        slide_paths = sorted(uniq)

        if not slide_paths:
            raise PptRenderError("PowerPoint 导出未产生任何 PNG 文件。")
        return slide_paths
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def embed_slide_wavs_to_pptx(
    ppt_path: Path,
    audio_dir: Path,
    out_pptx: Path,
    *,
    autoplay: bool = True,
    hide_icon: bool = True,
) -> Path:
    """
    Embed per-slide WAVs into a PPTX using PowerPoint COM automation.

    Expected WAV names: slide_0001.wav, slide_0002.wav, ...
    """
    if sys.platform != "win32":
        raise PptRenderError("嵌入音频到 PPTX 仅支持 Windows + PowerPoint（COM）。")
    ppt_path = Path(ppt_path).expanduser().resolve()
    audio_dir = Path(audio_dir).expanduser().resolve()
    out_pptx = Path(out_pptx).expanduser().resolve()
    out_pptx.parent.mkdir(parents=True, exist_ok=True)

    if not ppt_path.exists():
        raise FileNotFoundError(str(ppt_path))
    if not audio_dir.exists():
        raise FileNotFoundError(str(audio_dir))

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        raise PptRenderError("未安装 pywin32，无法使用 PowerPoint COM。请先 `pip install pywin32`。") from e

    app = None
    pres = None
    try:
        with _com_initialized():
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = 1
            pres = app.Presentations.Open(str(ppt_path), WithWindow=False)

            slide_count = int(pres.Slides.Count)
            missing: List[int] = []

            # Place audio icon off-canvas if requested.
            left, top, width, height = (0, 0, 16, 16)
            if hide_icon:
                left, top = (-1000, -1000)

            for i in range(1, slide_count + 1):
                wav_path = audio_dir / f"slide_{i:04d}.wav"
                if not wav_path.exists():
                    missing.append(i)
                    continue

                slide = pres.Slides(i)
                # AddMediaObject2: FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height
                # SaveWithDocument=True ensures the audio is embedded.
                shape = slide.Shapes.AddMediaObject2(str(wav_path), False, True, left, top, width, height)

                if autoplay:
                    try:
                        shape.AnimationSettings.PlaySettings.PlayOnEntry = True
                        shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = True
                    except Exception:
                        # Some PPT versions behave differently; keep best-effort.
                        pass

            if missing:
                raise PptRenderError(
                    "未找到部分页的语音文件，无法生成带旁白 PPT："
                    + ", ".join(str(i) for i in missing)
                    + f"。请确认 {audio_dir} 下存在 slide_XXXX.wav。"
                )

            # Save as PPTX (24)
            pres.SaveAs(str(out_pptx), 24)

        return out_pptx
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def _export_with_libreoffice_pdf_pipeline(
    ppt_path: Path,
    out_dir: Path,
    *,
    soffice: str,
    dpi: int,
) -> List[Path]:
    pdf_dir = out_dir / "_pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    # 1) PPT -> PDF
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(pdf_dir),
        str(ppt_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise PptRenderError(f"LibreOffice 转 PDF 失败：{proc.stderr or proc.stdout}")

    pdfs = sorted(pdf_dir.glob("*.pdf"))
    if not pdfs:
        raise PptRenderError("LibreOffice 转 PDF 结果为空（未生成 pdf）。")
    pdf_path = pdfs[0]

    # 2) PDF -> PNG (requires pdf2image + poppler)
    try:
        from pdf2image import convert_from_path  # type: ignore
    except Exception as e:
        raise PptRenderError("缺少依赖 pdf2image：请先 `pip install pdf2image`，并在系统安装 Poppler。") from e

    images = convert_from_path(str(pdf_path), dpi=dpi)
    slide_paths: List[Path] = []
    for i, img in enumerate(images, start=1):
        p = out_dir / _zero_pad_name(i, ext=".png")
        img.save(p, "PNG")
        slide_paths.append(p)
    if not slide_paths:
        raise PptRenderError("PDF 转图片结果为空（未生成 png）。")
    return slide_paths


def _finalize_and_manifest(slide_paths: List[Path], out_dir: Path) -> Tuple[List[SlideImage], Path]:
    # Normalize names to slide_0001.png ...
    normalized: List[Path] = []

    # Some backends export Slide1.PNG naming; sort by numeric suffix if present
    def _sort_key(p: Path):
        stem = p.stem
        digits = "".join(ch for ch in stem if ch.isdigit())
        return int(digits) if digits else 10**9

    for idx, src in enumerate(sorted(slide_paths, key=_sort_key), start=1):
        dst = out_dir / _zero_pad_name(idx, ext=".png")
        if src.resolve() != dst.resolve():
            try:
                if dst.exists():
                    dst.unlink()
                src.replace(dst)
            except Exception:
                shutil.copy2(src, dst)
        normalized.append(dst)

    slides: List[SlideImage] = []
    for i, p in enumerate(normalized, start=1):
        w = h = None
        try:
            from PIL import Image

            with Image.open(p) as im:
                w, h = im.size
        except Exception:
            pass
        slides.append(SlideImage(slide_index=i, path=p, width=w, height=h))

    manifest = {
        "slides_dir": str(out_dir),
        "count": len(slides),
        "slides": [
            {"slide_index": s.slide_index, "path": str(s.path), "width": s.width, "height": s.height} for s in slides
        ],
    }
    manifest_path = out_dir / "slides_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return slides, manifest_path

