# Sapient Minutes GPT Backend — Render Python Version

This is the non-Docker GitHub + Render setup.

It exposes two APIs:

1. `POST /process-audio`  
   Receives an uploaded meeting audio file and sends it to Sarvam.

2. `POST /generate-docx`  
   Receives strict minutes JSON from the Custom GPT and returns a branded `.docx` file.

PDF generation is intentionally not included in this version.

---

## Folder structure

```text
app/
  main.py                         # FastAPI routes
  sarvam_client.py                # Sarvam REST/batch helper
  assets/
    owl_logo.png                  # Sapient logo used in DOCX
  scripts/
    generate_minutes.py           # DOCX generation logic
examples/
  sample_minutes.json             # Test payload for /generate-docx
requirements.txt
render.yaml
openapi_gpt_action.yaml
```

---

## Deploy on Render using GitHub

### 1. Create a GitHub repo

Upload this folder to a GitHub repository.

### 2. Create a Render Web Service

In Render:

- New → Web Service
- Connect your GitHub repo
- Runtime: `Python 3`
- Build Command:

```bash
pip install -r requirements.txt
```

- Start Command:

```bash
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

### 3. Add environment variables

In Render → Environment:

```text
SARVAM_API_KEY=your_sarvam_api_key
SARVAM_MODEL=saaras:v3
USE_BATCH_BY_DEFAULT=false
```

For longer audio, set:

```text
USE_BATCH_BY_DEFAULT=true
```

or pass `use_batch=true` in the `/process-audio` request.

---

## API 1: Process audio

Endpoint:

```http
POST /process-audio
```

Form fields:

| Field | Type | Notes |
|---|---|---|
| `file` | file | Audio upload |
| `name` | string | User-facing/original file name |
| `language` | string | `en-IN`, `hi_IN`, or `hi-IN` |
| `mode` | string | Optional |
| `use_batch` | boolean | Optional; useful for longer files |

Business rule:

```text
en-IN  -> transcribe
hi_IN  -> translate
hi-IN  -> translate
```

Example curl:

```bash
curl -X POST "https://YOUR-RENDER-APP.onrender.com/process-audio" \
  -F "file=@meeting.mp3" \
  -F "name=meeting.mp3" \
  -F "language=en-IN" \
  -F "use_batch=false"
```

Response shape:

```json
{
  "engine": "sarvam_rest",
  "file_name": "meeting.mp3",
  "language_code": "en-IN",
  "mode": "transcribe",
  "transcript": "...",
  "raw": {}
}
```

---

## API 2: Generate DOCX

Endpoint:

```http
POST /generate-docx
```

Request JSON:

```json
{
  "minutes": {
    "client": "ABC Family Office",
    "attendees": "Client Team; Sapient Team",
    "location": "Zoom",
    "date": "25 May 2026",
    "sections": [
      {
        "heading": "Discussion Summary",
        "bullets": [
          "Reviewed current portfolio allocation.",
          "Discussed upcoming liquidity needs."
        ]
      }
    ]
  }
}
```

Example curl:

```bash
curl -X POST "https://YOUR-RENDER-APP.onrender.com/generate-docx" \
  -H "Content-Type: application/json" \
  --data @examples/sample_minutes.json \
  --output minutes.docx
```

---

## Custom GPT Action setup

Use `openapi_gpt_action.yaml` as the Action schema.

Before uploading it to the GPT Action editor, replace:

```text
https://YOUR-RENDER-APP.onrender.com
```

with your actual Render URL.

Suggested GPT instruction:

```text
When the user uploads meeting audio:
1. Call processAudio with the audio file, file name, and language.
2. Use the returned transcript only as source material.
3. Convert the transcript into the strict Sapient minutes JSON.
4. Call generateDocx with that JSON.
5. Return the generated DOCX to the user.

Do not invent details. If client, date, location, or attendees are unclear, write "Not specified".
```

---

## Local testing

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
export SARVAM_API_KEY="your_key"
uvicorn app.main:app --reload
```

Open:

```text
http://127.0.0.1:8000/docs
```
