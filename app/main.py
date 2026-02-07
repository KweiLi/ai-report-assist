"""FastAPI application – routes for the debiasing pipeline."""

from __future__ import annotations

import json
import logging
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import UPLOAD_DIR
from app.pipeline.analyzer import analyze_text
from app.pipeline.debiaser import BIAS_COLORS, debias_text, highlight_debiased, highlight_original
from app.pipeline.exporter import export_docx, export_pdf
from app.pipeline.extractor import extract_text
from app.pipeline.masker import mask_text
from app.pipeline.unmasker import unmask_text

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(name)s] %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="AI Report Assist", description="Law enforcement report debiasing tool")

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))
app.mount("/static", StaticFiles(directory=str(Path(__file__).parent / "static")), name="static")

# In-memory job store (for MVP – swap for DB/Redis in production)
jobs: dict[str, dict] = {}


@app.get("/", response_class=HTMLResponse)
async def upload_page(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})


@app.post("/upload")
async def upload_file(request: Request, file: UploadFile = File(...)):
    """Accept a PDF, extract text, and redirect to review."""
    job_id = str(uuid.uuid4())[:8]
    file_path = UPLOAD_DIR / f"{job_id}_{file.filename}"

    content = await file.read()
    file_path.write_bytes(content)

    # Step 1: Extract text
    extraction = extract_text(file_path)

    jobs[job_id] = {
        "filename": file.filename,
        "file_path": str(file_path),
        "original_text": extraction.total_text,
        "is_scanned": extraction.is_scanned,
        "pages": len(extraction.pages),
        "status": "extracted",
    }

    return RedirectResponse(url=f"/review/{job_id}", status_code=303)


@app.get("/review/{job_id}", response_class=HTMLResponse)
async def review_page(request: Request, job_id: str):
    job = jobs.get(job_id)
    if not job:
        return HTMLResponse("<h1>Job not found</h1>", status_code=404)
    return templates.TemplateResponse("review.html", {
        "request": request,
        "job_id": job_id,
        "job": job,
        "bias_colors": BIAS_COLORS,
    })


@app.post("/process/{job_id}")
async def process_report(request: Request, job_id: str):
    """Run the full pipeline: analyze → mask → debias → unmask."""
    job = jobs.get(job_id)
    if not job:
        return HTMLResponse("<h1>Job not found</h1>", status_code=404)

    original_text = job["original_text"]

    # Step 2: Analyze PII
    analysis = analyze_text(original_text)

    # Step 3: Mask
    mask_result = mask_text(original_text, analysis.entities)

    # Step 4: Debias (cloud – only masked text sent)
    debias_result = debias_text(mask_result.masked_text)

    # Step 5: Unmask
    unmask_result = unmask_text(debias_result.debiased_text, mask_result.entity_mapping)

    # Build highlighted HTML for side-by-side view
    original_highlighted = highlight_original(original_text, debias_result.changes)
    debiased_highlighted = highlight_debiased(unmask_result.final_text, debias_result.changes)

    # Serialize bias changes for the template
    bias_changes = [
        {
            "original_phrase": c.original_phrase,
            "replacement_phrase": c.replacement_phrase,
            "bias_type": c.bias_type,
            "explanation": c.explanation,
        }
        for c in debias_result.changes
    ]

    job.update({
        "masked_text": mask_result.masked_text,
        "entities_found": mask_result.entities_found,
        "entity_mapping": mask_result.entity_mapping,
        "acronyms_preserved": analysis.acronyms_preserved,
        "debiased_masked": debias_result.debiased_text,
        "debiased_text": unmask_result.final_text,
        "original_highlighted": original_highlighted,
        "debiased_highlighted": debiased_highlighted,
        "bias_changes": bias_changes,
        "changes_summary": debias_result.changes_summary,
        "unresolved_tokens": unmask_result.unresolved_tokens,
        "status": "processed",
    })

    return RedirectResponse(url=f"/review/{job_id}", status_code=303)


@app.get("/export/{job_id}")
async def export_report(job_id: str, format: str = "pdf"):
    """Export the debiased report as PDF or DOCX."""
    job = jobs.get(job_id)
    if not job or job.get("status") != "processed":
        return HTMLResponse("<h1>Job not found or not yet processed</h1>", status_code=404)

    text = job["debiased_text"]
    title = f"Debiased Report - {job['filename']}"
    entities_found = job.get("entities_found")
    changes_summary = job.get("changes_summary")
    acronyms_preserved = job.get("acronyms_preserved")

    if format == "docx":
        output_path = UPLOAD_DIR / f"{job_id}_debiased.docx"
        export_docx(text, output_path, title=title, entities_found=entities_found, changes_summary=changes_summary, acronyms_preserved=acronyms_preserved)
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        output_path = UPLOAD_DIR / f"{job_id}_debiased.pdf"
        export_pdf(text, output_path, title=title, entities_found=entities_found, changes_summary=changes_summary, acronyms_preserved=acronyms_preserved)
        media_type = "application/pdf"

    return FileResponse(
        path=str(output_path),
        filename=output_path.name,
        media_type=media_type,
    )


@app.get("/api/job/{job_id}")
async def get_job_status(job_id: str):
    """JSON endpoint for job data (used by HTMX)."""
    job = jobs.get(job_id)
    if not job:
        return {"error": "not found"}
    # Return a serializable subset
    return {
        "job_id": job_id,
        "filename": job.get("filename"),
        "status": job.get("status"),
        "is_scanned": job.get("is_scanned"),
        "pages": job.get("pages"),
        "entities_found": job.get("entities_found"),
        "changes_summary": job.get("changes_summary"),
    }
