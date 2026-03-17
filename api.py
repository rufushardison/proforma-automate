"""
api.py

Thin FastAPI wrapper around the proforma extraction + Excel generation pipeline.
Run alongside Streamlit on a separate port, e.g.:
    uvicorn api:app --host 0.0.0.0 --port 8502

Endpoints:
    POST /generate  — extract assumptions and produce a filled .xlsm file
    GET  /download/{file_id}  — download the generated file
"""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path

import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from extractor import extract_assumptions
from excel_writer import fill_template

app = FastAPI()

_BASE_DIR = Path(__file__).parent
_MANIFESTS_DIR = _BASE_DIR / "manifests"
_TEMPLATES_DIR = _BASE_DIR / "templates"
_DOWNLOADS_DIR = _BASE_DIR / "downloads"
_DOWNLOADS_DIR.mkdir(exist_ok=True)

_TEMPLATE_MAP = {
    "single": ("Single Tenant Model.xlsm", "template_a.json"),
    "multi":  ("Multi Tenant Model.xlsm",  "template_b.json"),
}


class GenerateRequest(BaseModel):
    deal_summary: str
    template: str = "single"  # "single" or "multi"


@app.post("/generate")
async def generate_proforma(req: GenerateRequest):
    template_key = req.template.lower()
    if template_key not in _TEMPLATE_MAP:
        raise HTTPException(status_code=400, detail=f"Invalid template '{req.template}'. Use 'single' or 'multi'.")

    template_file, manifest_file = _TEMPLATE_MAP[template_key]
    template_path = _TEMPLATES_DIR / template_file
    manifest_path = _MANIFESTS_DIR / manifest_file

    if not template_path.exists():
        raise HTTPException(status_code=500, detail=f"Template file not found: {template_file}")
    if not manifest_path.exists():
        raise HTTPException(status_code=500, detail=f"Manifest file not found: {manifest_file}")

    with open(manifest_path) as f:
        manifest = json.load(f)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured on proforma server.")

    client = anthropic.Anthropic(api_key=api_key)

    extraction = extract_assumptions(req.deal_summary, manifest, client)
    buffer = fill_template(template_path, manifest, extraction)

    file_id = str(uuid.uuid4())
    filename = f"proforma_{file_id}.xlsm"
    (_DOWNLOADS_DIR / filename).write_bytes(buffer.read())

    return {"file_id": file_id, "filename": filename}


@app.get("/download/{filename}")
async def download_file(filename: str):
    # Basic path traversal guard
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    path = _DOWNLOADS_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found or expired.")

    return FileResponse(
        path=path,
        filename="proforma.xlsm",
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
    )
