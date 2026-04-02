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
import logging
import os
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

import httpx
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

from render_translation import render_document, validate_json

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("uscis")

# ---------------------------------------------------------------------------
# Config — load .env if python-dotenv is available
# ---------------------------------------------------------------------------

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
TEMPLATE_PATH = os.environ.get("DOCX_TEMPLATE_PATH", None)

# Load system prompt
PROMPT_PATH = Path(__file__).parent / "uscis-translation-prompt-v2.md"
SYSTEM_PROMPT = PROMPT_PATH.read_text(encoding="utf-8") if PROMPT_PATH.exists() else ""

# Temp directory for outputs
OUTPUT_DIR = Path(tempfile.mkdtemp(prefix="uscis_"))
logger.info("Output directory: %s", OUTPUT_DIR)

# Check LibreOffice availability
LIBREOFFICE_BIN = shutil.which("libreoffice") or shutil.which("soffice")
if LIBREOFFICE_BIN:
    logger.info("LibreOffice found: %s — PDF export enabled", LIBREOFFICE_BIN)
else:
    logger.warning("LibreOffice not found — PDF export disabled (DOCX only)")

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

STATIC_DIR = Path(__file__).parent / "static"

# In-memory job tracking
jobs: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# Gemini API (direct REST)
# ---------------------------------------------------------------------------

async def call_gemini(pdf_bytes: bytes) -> dict:
    """Send a PDF to Gemini via REST API and return parsed JSON translation."""
    if not GEMINI_API_KEY:
        raise HTTPException(500, "GEMINI_API_KEY not set. Add it to .env or export it.")

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

    logger.info("Sending %d bytes to Gemini (%s)...", len(pdf_bytes), GEMINI_MODEL)

    async with httpx.AsyncClient(timeout=300.0) as client:
        resp = await client.post(url, json=payload)

    if resp.status_code != 200:
        detail = resp.text[:500]
        logger.error("Gemini API error %d: %s", resp.status_code, detail)
        raise HTTPException(502, f"Gemini API error ({resp.status_code}): {detail}")

    result = resp.json()

    # Extract text from response
    try:
        candidates = result["candidates"]
        parts = candidates[0]["content"]["parts"]
        raw_text = parts[0]["text"].strip()
    except (KeyError, IndexError) as e:
        logger.error("Unexpected Gemini response: %s", json.dumps(result)[:500])
        raise HTTPException(502, f"Unexpected Gemini response structure: {e}")

    # Parse JSON
    try:
        data = json.loads(raw_text)
        logger.info("Gemini returned valid JSON with %d pages", len(data.get("pages", [])))
        return data
    except json.JSONDecodeError as e:
        logger.error("Gemini returned invalid JSON: %s... | error: %s", raw_text[:200], e)
        raise HTTPException(502, f"Gemini returned invalid JSON: {e}")


def convert_docx_to_pdf(docx_path: Path) -> Path | None:
    """Convert DOCX to PDF via LibreOffice. Returns PDF path or None."""
    if not LIBREOFFICE_BIN:
        return None

    try:
        subprocess.run(
            [LIBREOFFICE_BIN, "--headless", "--convert-to", "pdf",
             "--outdir", str(docx_path.parent), str(docx_path)],
            capture_output=True, timeout=120,
        )
        pdf_path = docx_path.with_suffix(".pdf")
        if pdf_path.exists():
            logger.info("PDF generated: %s", pdf_path)
            return pdf_path
    except (subprocess.TimeoutExpired, OSError) as e:
        logger.warning("LibreOffice conversion failed: %s", e)

    return None


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the frontend."""
    index_path = STATIC_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(index_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>USCIS Translator</h1><p>static/index.html not found.</p>")


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
    if len(contents) > 50 * 1024 * 1024:
        raise HTTPException(400, "File too large. Maximum size is 50MB.")

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "processing", "filename": file.filename}
    logger.info("Job %s: translating %s (%d bytes)", job_id, file.filename, len(contents))

    try:
        # --- Step 1: Translate via Gemini ---
        translation_data = await call_gemini(contents)

        # --- Step 2: Validate ---
        errors = validate_json(translation_data)
        if errors:
            logger.warning("Job %s: schema validation warnings: %s", job_id, errors)

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
        logger.info("Job %s: DOCX rendered to %s", job_id, docx_path)

        # --- Step 4: Convert DOCX to PDF ---
        pdf_path = convert_docx_to_pdf(docx_path)

        # --- Step 5: Return result ---
        result = {
            "status": "complete",
            "job_id": job_id,
            "filename": file.filename,
            "docx_url": f"/api/download/{job_id}/docx",
            "metadata": translation_data.get("metadata", {}),
        }
        if pdf_path:
            result["pdf_url"] = f"/api/download/{job_id}/pdf"

        if errors:
            result["validation_warnings"] = errors

        jobs[job_id] = {
            **result,
            "_docx_path": str(docx_path),
            "_pdf_path": str(pdf_path) if pdf_path else None,
            "_docx_filename": docx_filename,
            "_pdf_filename": docx_filename.replace(".docx", ".pdf"),
            "_json": translation_data,
        }

        logger.info("Job %s: complete", job_id)
        return result

    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Job %s failed", job_id)
        jobs[job_id] = {"status": "error", "error": str(e)}
        raise HTTPException(500, f"Translation failed: {e}")


@app.get("/api/download/{job_id}/{file_type}")
async def download(job_id: str, file_type: str):
    """Download the translated DOCX, PDF, or raw JSON."""
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


@app.get("/api/health")
async def health():
    """Health check endpoint."""
    return {
        "status": "ok",
        "gemini_configured": bool(GEMINI_API_KEY),
        "gemini_model": GEMINI_MODEL,
        "libreoffice_available": bool(LIBREOFFICE_BIN),
        "template_configured": bool(TEMPLATE_PATH),
    }
