#!/usr/bin/env python3
"""
USCIS Certified Translation Web App
====================================
FastAPI backend that accepts PDF uploads, sends them to Gemini for translation,
renders the result to DOCX via render_translation.py, converts to PDF, and
returns both files for download.

Uses httpx for direct Gemini REST API calls (avoids heavy google-generativeai SDK).
"""

import base64
import json
import os
import tempfile
import uuid
from pathlib import Path

import httpx
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from render_translation import render_document, validate_json

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
TEMPLATE_PATH = os.environ.get("DOCX_TEMPLATE_PATH", None)

# Load system prompt
PROMPT_PATH = Path(__file__).parent / "uscis-translation-prompt-v2.md"
SYSTEM_PROMPT = PROMPT_PATH.read_text(encoding="utf-8") if PROMPT_PATH.exists() else ""

# Temp directory for outputs
OUTPUT_DIR = Path(tempfile.mkdtemp(prefix="uscis_"))

GEMINI_API_BASE = "https://generativelanguage.googleapis.com/v1beta/models"

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(title="USCIS Certified Translation", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve static frontend
STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# In-memory job tracking
jobs: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# Gemini API (direct REST)
# ---------------------------------------------------------------------------

async def call_gemini(pdf_bytes: bytes) -> dict:
    """Send a PDF to Gemini via REST API and return parsed JSON translation."""
    if not GEMINI_API_KEY:
        raise HTTPException(500, "GEMINI_API_KEY environment variable not set.")

    url = f"{GEMINI_API_BASE}/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

    pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

    payload = {
        "system_instruction": {
            "parts": [{"text": SYSTEM_PROMPT}]
        },
        "contents": [
            {
                "parts": [
                    {
                        "inline_data": {
                            "mime_type": "application/pdf",
                            "data": pdf_b64,
                        }
                    },
                    {
                        "text": "Translate this document. Output ONLY valid JSON conforming to the schema in your instructions."
                    },
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }

    async with httpx.AsyncClient(timeout=300.0) as client:
        resp = await client.post(url, json=payload)

    if resp.status_code != 200:
        detail = resp.text[:500]
        raise HTTPException(502, f"Gemini API error ({resp.status_code}): {detail}")

    result = resp.json()

    # Extract text from response
    try:
        candidates = result["candidates"]
        parts = candidates[0]["content"]["parts"]
        raw_text = parts[0]["text"].strip()
    except (KeyError, IndexError) as e:
        raise HTTPException(502, f"Unexpected Gemini response structure: {e}")

    # Parse JSON
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise HTTPException(502, f"Gemini returned invalid JSON: {e}")


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index():
    index_path = STATIC_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(index_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>USCIS Translator</h1><p>Static files not found.</p>")


@app.post("/api/translate")
async def translate(file: UploadFile = File(...)):
    """
    Accept a PDF upload, translate via Gemini, render to DOCX + PDF.
    Returns a job object with download URLs.
    """
    # Validate file type
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are accepted.")

    # Read uploaded file
    contents = await file.read()
    if len(contents) > 50 * 1024 * 1024:  # 50MB limit
        raise HTTPException(400, "File too large. Maximum size is 50MB.")

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "processing", "filename": file.filename}

    try:
        # --- Step 1: Translate via Gemini ---
        translation_data = await call_gemini(contents)

        # --- Step 2: Validate ---
        errors = validate_json(translation_data)
        if errors:
            jobs[job_id]["validation_warnings"] = errors

        # --- Step 3: Render DOCX ---
        safe_name = file.filename.rsplit(".", 1)[0]
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in safe_name)
        docx_filename = f"{safe_name}_translated.docx"
        docx_path = OUTPUT_DIR / f"{job_id}_{docx_filename}"

        render_document(
            translation_data,
            template_path=TEMPLATE_PATH,
            output_path=str(docx_path),
        )

        # --- Step 4: Convert DOCX to PDF via LibreOffice ---
        pdf_path = docx_path.with_suffix(".pdf")
        os.system(
            f'libreoffice --headless --convert-to pdf --outdir "{OUTPUT_DIR}" "{docx_path}" 2>/dev/null'
        )

        # LibreOffice names the output based on input filename
        lo_pdf = OUTPUT_DIR / f"{job_id}_{safe_name}_translated.pdf"
        pdf_available = lo_pdf.exists()
        if pdf_available and lo_pdf != pdf_path:
            lo_pdf.rename(pdf_path)
            pdf_available = pdf_path.exists()

        # --- Step 5: Return result ---
        result = {
            "status": "complete",
            "job_id": job_id,
            "filename": file.filename,
            "docx_url": f"/api/download/{job_id}/docx",
            "metadata": translation_data.get("metadata", {}),
        }
        if pdf_available:
            result["pdf_url"] = f"/api/download/{job_id}/pdf"

        if errors:
            result["validation_warnings"] = errors

        jobs[job_id] = {
            **result,
            "_docx_path": str(docx_path),
            "_pdf_path": str(pdf_path) if pdf_available else None,
            "_docx_filename": docx_filename,
            "_pdf_filename": docx_filename.replace(".docx", ".pdf"),
            "_json": translation_data,
        }

        return result

    except HTTPException:
        raise
    except Exception as e:
        jobs[job_id] = {"status": "error", "error": str(e)}
        raise HTTPException(500, f"Translation failed: {e}")


@app.get("/api/download/{job_id}/{file_type}")
async def download(job_id: str, file_type: str):
    """Download the translated DOCX or PDF."""
    job = jobs.get(job_id)
    if not job or job.get("status") != "complete":
        raise HTTPException(404, "Job not found or not complete.")

    if file_type == "docx":
        path = job.get("_docx_path")
        filename = job.get("_docx_filename", "translation.docx")
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif file_type == "pdf":
        path = job.get("_pdf_path")
        filename = job.get("_pdf_filename", "translation.pdf")
        media_type = "application/pdf"
    elif file_type == "json":
        return job.get("_json", {})
    else:
        raise HTTPException(400, "Invalid file type. Use 'docx', 'pdf', or 'json'.")

    if not path or not Path(path).exists():
        raise HTTPException(404, f"{file_type.upper()} file not available.")

    return FileResponse(path, media_type=media_type, filename=filename)
