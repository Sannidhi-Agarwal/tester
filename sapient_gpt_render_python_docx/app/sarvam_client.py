"""Sarvam STT helper for the Sapient Minutes GPT backend.

This module intentionally keeps Sarvam logic separate from FastAPI so it can be
unit-tested or reused from scripts.
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

import requests

SARVAM_API_KEY = os.environ.get("SARVAM_API_KEY")
SARVAM_REST_URL = os.environ.get("SARVAM_REST_URL", "https://api.sarvam.ai/speech-to-text")
SARVAM_MODEL = os.environ.get("SARVAM_MODEL", "saaras:v3")
SARVAM_DIARIZATION = os.environ.get("SARVAM_DIARIZATION", "false").lower() == "true"

LANGUAGE_ALIASES = {
    "en_IN": "en-IN",
    "en-IN": "en-IN",
    "english": "en-IN",
    "hi_IN": "hi-IN",
    "hi-IN": "hi-IN",
    "hindi": "hi-IN",
}


def normalize_language(language: str | None) -> str:
    lang = (language or "en-IN").strip()
    return LANGUAGE_ALIASES.get(lang, lang.replace("_", "-"))


def resolve_mode(language: str, requested_mode: str | None = None) -> str:
    """Business rule: English is transcribed; Hindi is translated."""
    lang = normalize_language(language)
    if lang == "en-IN":
        return "transcribe"
    if lang == "hi-IN":
        return "translate"
    return (requested_mode or "transcribe").strip().lower()


def extract_transcript(obj: Any) -> str:
    """Best-effort extraction from Sarvam REST or batch-style responses."""
    if obj is None:
        return ""
    if isinstance(obj, str):
        return obj.strip()
    if isinstance(obj, dict):
        for key in (
            "transcript",
            "text",
            "transcription",
            "translated_text",
            "translation",
            "raw_text",
        ):
            value = obj.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()

        diarized = obj.get("diarized_transcript") or obj.get("diarizedTranscript")
        if isinstance(diarized, dict):
            entries = diarized.get("entries") or []
            parts: list[str] = []
            for entry in entries:
                if isinstance(entry, dict) and entry.get("transcript"):
                    speaker = entry.get("speaker_id") or entry.get("speaker")
                    prefix = f"Speaker {speaker}: " if speaker is not None else ""
                    parts.append(prefix + str(entry["transcript"]).strip())
            if parts:
                return "\n".join(parts)

        for value in obj.values():
            found = extract_transcript(value)
            if found:
                return found
    if isinstance(obj, list):
        parts = [extract_transcript(item) for item in obj]
        return "\n".join(part for part in parts if part).strip()
    return ""


def transcribe_rest(
    audio_path: Path,
    filename: str,
    content_type: str,
    language: str,
    mode: str,
) -> dict[str, Any]:
    if not SARVAM_API_KEY:
        raise RuntimeError("SARVAM_API_KEY is not set.")

    language_code = normalize_language(language)
    resolved_mode = resolve_mode(language_code, mode)

    with audio_path.open("rb") as audio_file:
        response = requests.post(
            SARVAM_REST_URL,
            headers={"api-subscription-key": SARVAM_API_KEY},
            files={"file": (filename, audio_file, content_type or "application/octet-stream")},
            data={
                "model": SARVAM_MODEL,
                "language_code": language_code,
                "mode": resolved_mode,
            },
            timeout=240,
        )

    if response.status_code >= 400:
        raise RuntimeError(f"Sarvam REST failed ({response.status_code}): {response.text}")

    try:
        raw = response.json()
    except Exception:
        raw = {"raw_text": response.text}

    return {
        "engine": "sarvam_rest",
        "file_name": filename,
        "language_code": language_code,
        "mode": resolved_mode,
        "transcript": extract_transcript(raw),
        "raw": raw,
    }


def transcribe_batch(audio_path: Path, language: str, mode: str) -> dict[str, Any]:
    """Batch path for longer files. Requires the sarvamai Python package."""
    if not SARVAM_API_KEY:
        raise RuntimeError("SARVAM_API_KEY is not set.")

    try:
        from sarvamai import SarvamAI
    except Exception as exc:
        raise RuntimeError(f"sarvamai SDK is not importable: {exc}") from exc

    language_code = normalize_language(language)
    resolved_mode = resolve_mode(language_code, mode)
    output_dir = Path(tempfile.mkdtemp(prefix="sarvam_batch_out_"))

    try:
        client = SarvamAI(api_subscription_key=SARVAM_API_KEY)
        job = client.speech_to_text_job.create_job(
            model=SARVAM_MODEL,
            mode=resolved_mode,
            language_code=language_code,
            with_diarization=SARVAM_DIARIZATION,
        )
        job.upload_files(file_paths=[str(audio_path)])
        job.start()
        job.wait_until_complete()
        file_results = job.get_file_results()
        job.download_outputs(output_dir=str(output_dir))

        raw_outputs: list[Any] = []
        downloaded_files: list[str] = []
        for path in output_dir.rglob("*"):
            if not path.is_file():
                continue
            downloaded_files.append(path.name)
            try:
                if path.suffix.lower() == ".json":
                    raw_outputs.append(json.loads(path.read_text(encoding="utf-8")))
                else:
                    raw_outputs.append(path.read_text(encoding="utf-8"))
            except Exception:
                pass

        return {
            "engine": "sarvam_batch",
            "language_code": language_code,
            "mode": resolved_mode,
            "file_results": file_results,
            "downloaded_files": downloaded_files,
            "transcript": extract_transcript(raw_outputs),
            "raw": raw_outputs,
        }
    finally:
        shutil.rmtree(output_dir, ignore_errors=True)
