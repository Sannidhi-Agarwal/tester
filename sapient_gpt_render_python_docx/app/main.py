from __future__ import annotations

import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from app.sarvam_client import normalize_language, resolve_mode, transcribe_batch, transcribe_rest
from app.scripts.generate_minutes import build_document

app = FastAPI(
    title="Sapient Minutes GPT Backend",
    version="1.0.0",
    description="Native Render Python backend for Custom GPT: audio to Sarvam transcript, JSON to Sapient DOCX.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "assets" / "owl_logo.png"
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "/tmp/sapient_minutes_outputs"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

AUDIO_EXTS = {".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac", ".webm", ".mp4"}
USE_BATCH_BY_DEFAULT = os.environ.get("USE_BATCH_BY_DEFAULT", "false").lower() == "true"


class MinutesPayload(BaseModel):
    client: str = Field("Not specified", description="Client name")
    attendees: str = Field("Not specified", description="Meeting attendees")
    location: str = Field("Not specified", description="Meeting location")
    date: str = Field("Not specified", description="Meeting date/time as text")
    sections: list[dict[str, Any]] = Field(default_factory=list)
    portfolio_table: dict[str, Any] | None = None


class GenerateDocxRequest(BaseModel):
    minutes: MinutesPayload


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", (value or "").strip()).strip("._-")
    return cleaned or "file"


@app.get("/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "sapient-minutes-gpt-backend"}


@app.post("/process-audio")
async def process_audio(
    file: UploadFile = File(..., description="Meeting audio file"),
    name: str = Form(..., description="Original/user-facing file name"),
    language: str = Form("en-IN", description="Language code, e.g. en-IN or hi_IN"),
    mode: str | None = Form(None, description="Requested mode. en-IN becomes transcribe; hi-IN becomes translate."),
    use_batch: bool = Form(False, description="Use Sarvam batch path for longer audio."),
):
    """Upload audio and get a transcript/translation back from Sarvam.

    Business rule:
    - en-IN / English -> transcribe
    - hi-IN / hi_IN / Hindi -> translate
    """
    language_code = normalize_language(language)
    sarvam_mode = resolve_mode(language_code, mode)

    original_name = name or file.filename or "audio"
    suffix = Path(original_name).suffix.lower() or Path(file.filename or "").suffix.lower()
    if suffix and suffix not in AUDIO_EXTS:
        raise HTTPException(status_code=400, detail=f"Unsupported audio extension: {suffix}")

    tmpdir = Path(tempfile.mkdtemp(prefix="sapient_audio_"))
    audio_path = tmpdir / safe_filename(original_name)
    if suffix and not audio_path.suffix:
        audio_path = audio_path.with_suffix(suffix)

    try:
        with audio_path.open("wb") as output_file:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                output_file.write(chunk)

        try:
            if use_batch or USE_BATCH_BY_DEFAULT:
                result = transcribe_batch(audio_path, language_code, sarvam_mode)
            else:
                result = transcribe_rest(
                    audio_path=audio_path,
                    filename=original_name,
                    content_type=file.content_type or "application/octet-stream",
                    language=language_code,
                    mode=sarvam_mode,
                )
        except Exception as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse(result)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@app.post("/generate-docx")
def generate_docx(request: GenerateDocxRequest):
    """Generate the Sapient minutes DOCX from GPT-created strict JSON."""
    minutes = request.minutes.model_dump(exclude_none=True)
    client_name = safe_filename(minutes.get("client", "Client"))
    output_path = OUTPUT_DIR / f"Minutes_{client_name}_{uuid.uuid4().hex[:8]}.docx"

    try:
        build_document(minutes, output_path, LOGO_PATH if LOGO_PATH.exists() else None)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"DOCX generation failed: {exc}") from exc

    return FileResponse(
        path=str(output_path),
        filename=f"Minutes_{client_name}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
