"""
PPT Generator API — FastAPI Backend
myrobo.in — Full Stack PPT Generator
"""

import os
import uuid
import json
import shutil
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import uvicorn

from parser import extract_document_content
from template_reader import analyze_template
from ppt_generator import generate_ppt

# ── App Setup ──────────────────────────────────────────────
app = FastAPI(title="PPT Generator API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_DOC  = {".pdf", ".docx", ".txt", ".rtf"}
ALLOWED_PPT  = {".pptx"}
MAX_SIZE_MB  = 50


# ── Health Check ───────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "message": "PPT Generator API is running"}


# ── Main Route ─────────────────────────────────────────────
@app.post("/generate-ppt")
async def generate_ppt_route(
    document: UploadFile = File(...),
    template: UploadFile = File(...),
):
    """
    Accept a document + template PPT, generate a new PPT
    that mirrors the template's design with the document's content.
    """
    job_id = str(uuid.uuid4())[:8]
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    try:
        # ── Validate file types ──
        doc_ext = Path(document.filename).suffix.lower()
        ppt_ext = Path(template.filename).suffix.lower()

        if doc_ext not in ALLOWED_DOC:
            raise HTTPException(400, f"Document type '{doc_ext}' not supported. Use: {ALLOWED_DOC}")
        if ppt_ext not in ALLOWED_PPT:
            raise HTTPException(400, "Template must be a .pptx file")

        # ── Save uploaded files ──
        doc_path = job_dir / f"document{doc_ext}"
        ppt_path = job_dir / f"template.pptx"
        out_path = OUTPUT_DIR / f"generated_{job_id}.pptx"

        with open(doc_path, "wb") as f:
            content = await document.read()
            if len(content) > MAX_SIZE_MB * 1024 * 1024:
                raise HTTPException(400, f"File too large. Max {MAX_SIZE_MB}MB allowed.")
            f.write(content)

        with open(ppt_path, "wb") as f:
            content = await template.read()
            f.write(content)

        print(f"\n[{job_id}] Document: {document.filename} | Template: {template.filename}")

        # ── Step 1: Extract document content ──
        print(f"[{job_id}] Extracting document content...")
        doc_data = extract_document_content(str(doc_path))
        print(f"[{job_id}] Found {len(doc_data.get('sections', []))} sections")

        # ── Step 2: Analyze template ──
        print(f"[{job_id}] Analyzing template...")
        template_data = analyze_template(str(ppt_path))
        print(f"[{job_id}] Template has {template_data.get('slide_count', 0)} slides, "
              f"{len(template_data.get('layouts', []))} layouts")

        # ── Step 3: Generate PPT ──
        print(f"[{job_id}] Generating PPT...")
        generate_ppt(doc_data, template_data, str(ppt_path), str(out_path))
        print(f"[{job_id}] Done! Output: {out_path}")

        # ── Return file ──
        safe_name = (doc_data.get("title") or "presentation").replace(" ", "_")[:40]
        return FileResponse(
            path=str(out_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=f"{safe_name}.pptx",
        )

    except HTTPException:
        raise
    except Exception as e:
        print(f"[{job_id}] ERROR: {e}")
        import traceback; traceback.print_exc()
        raise HTTPException(500, f"Generation failed: {str(e)}")
    finally:
        # Cleanup upload folder
        shutil.rmtree(job_dir, ignore_errors=True)


# ── Run ────────────────────────────────────────────────────
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
