from __future__ import annotations

import json
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
from scipy.io.wavfile import write as wav_write


@dataclass(frozen=True)
class SlideAudio:
    slide_index: int
    wav_path: Path
    sample_rate: int
    seconds: float


def split_text_for_tts(text: str, max_chars: int = 260) -> List[str]:
    """
    Very simple paragraph/sentence split to keep TTS stable for long slides.
    """
    t = (text or "").strip()
    if not t:
        return []

    parts: List[str] = []
    for para in t.splitlines():
        p = para.strip()
        if not p:
            continue
        if len(p) <= max_chars:
            parts.append(p)
            continue
        # sentence-ish split
        buf = ""
        for ch in p:
            buf += ch
            if ch in "。！？!?" and len(buf) >= max_chars * 0.5:
                parts.append(buf.strip())
                buf = ""
        if buf.strip():
            parts.append(buf.strip())

    # final pass: hard chunk
    final: List[str] = []
    for p in parts:
        if len(p) <= max_chars:
            final.append(p)
        else:
            for i in range(0, len(p), max_chars):
                final.append(p[i : i + max_chars])
    return final


def synthesize_slides_to_wavs(
    *,
    ov_tts_model: Any,
    slides: List[Dict[str, Any]],
    out_dir: Path,
    language: str = "Chinese",
    speaker: str = "vivian",
    style_instruct: str = "用清晰、自然、适合讲解的语气朗读。",
    per_slide_split: bool = True,
    max_new_tokens: int = 2048,
) -> List[SlideAudio]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    results: List[SlideAudio] = []

    for s in slides:
        idx = int(s.get("slide_index"))
        notes = str(s.get("speaker_notes", "")).strip()
        if not notes:
            continue

        chunks = split_text_for_tts(notes) if per_slide_split else [notes]
        wav_pieces: List[np.ndarray] = []
        sr_final: Optional[int] = None

        for chunk in chunks:
            wavs, sr = ov_tts_model.generate_custom_voice(
                text=chunk,
                speaker=str(speaker).lower().replace(" ", "_"),
                language=language,
                instruct=style_instruct,
                non_streaming_mode=True,
                max_new_tokens=max_new_tokens,
            )
            if wavs is None or not wavs:
                raise RuntimeError(f"TTS 生成失败：slide {idx}")
            sr_final = int(sr)
            wav_pieces.append(np.asarray(wavs[0], dtype=np.float32))

        wav = _concat_wavs(wav_pieces, sr=sr_final or 22050)
        wav_path = out_dir / f"slide_{idx:04d}.wav"
        wav_write(str(wav_path), int(sr_final or 22050), wav)
        results.append(SlideAudio(slide_index=idx, wav_path=wav_path, sample_rate=int(sr_final or 22050), seconds=len(wav) / float(sr_final or 22050)))

    return results


def package_outputs_zip(
    *,
    slides_dir: Path,
    scripts_json: Path,
    audio_dir: Path,
    out_zip: Path,
    extra_files: Optional[Iterable[Path]] = None,
) -> Path:
    out_zip = Path(out_zip)
    out_zip.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Slides
        slides_dir = Path(slides_dir)
        for p in sorted(slides_dir.glob("slide_*.png")):
            z.write(p, arcname=f"slides/{p.name}")

        # Scripts
        z.write(Path(scripts_json), arcname=f"scripts/{Path(scripts_json).name}")

        # Audio
        audio_dir = Path(audio_dir)
        for p in sorted(audio_dir.glob("slide_*.wav")):
            z.write(p, arcname=f"audio/{p.name}")

        for p in extra_files or []:
            p = Path(p)
            if p.exists() and p.is_file():
                z.write(p, arcname=p.name)

    return out_zip


def save_slides_json(slides: List[Dict[str, Any]], path: Path) -> None:
    Path(path).write_text(json.dumps(slides, ensure_ascii=False, indent=2), encoding="utf-8")


def _concat_wavs(pieces: List[np.ndarray], sr: int) -> np.ndarray:
    if not pieces:
        return np.zeros((0,), dtype=np.int16)
    # 100ms silence between chunks
    gap = np.zeros((int(sr * 0.10),), dtype=np.float32)
    merged: List[np.ndarray] = []
    for i, p in enumerate(pieces):
        if i > 0:
            merged.append(gap)
        merged.append(p)
    wav = np.concatenate(merged, axis=0)
    # Normalize to int16
    wav = np.clip(wav, -1.0, 1.0)
    return (wav * 32767.0).astype(np.int16)

