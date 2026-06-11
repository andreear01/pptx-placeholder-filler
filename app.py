from __future__ import annotations

import io
import json
import os
import zipfile
from dataclasses import asdict

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from powerpoint_replacer import fill_pptx_bytes, base64_to_pptx, pptx_to_base64


app = FastAPI(
    title="PowerPoint Placeholder Filler",
    version="1.0.0",
    description="Fills {{AI_...}} placeholders in PowerPoint files using AI-generated JSON insights."
)


def _check_api_key(x_api_key: str | None):
    expected = os.getenv("PPTX_TOOL_API_KEY")
    if expected and x_api_key != expected:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")


class FillPowerPointRequest(BaseModel):
    pptx_base64: str = Field(..., description="Base64 encoded .pptx file")
    replacements: dict[str, str] = Field(
        ...,
        description="JSON mapping where each key is a placeholder like {{AI_MX_OVERALL_INSIGHTS}} and each value is the exact text or image prompt."
    )
    remove_image_placeholders: bool = Field(
        default=True,
        description="If true, {{AI_IMAGE_...}} placeholders are removed from the PPTX and returned separately as image_prompts."
    )


class FillPowerPointResponse(BaseModel):
    completed_pptx_base64: str
    report: dict
    image_prompts: dict[str, str]


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/fill_powerpoint_base64", response_model=FillPowerPointResponse)
def fill_powerpoint_base64(request: FillPowerPointRequest, x_api_key: str | None = Header(default=None)):
    """
    Best for GPT Actions / API clients that can pass base64 JSON.
    """
    _check_api_key(x_api_key)

    try:
        pptx_bytes = base64_to_pptx(request.pptx_base64)
        completed_bytes, report = fill_pptx_bytes(
            pptx_bytes,
            request.replacements,
            remove_image_placeholders=request.remove_image_placeholders,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    return FillPowerPointResponse(
        completed_pptx_base64=pptx_to_base64(completed_bytes),
        report=asdict(report),
        image_prompts=report.image_prompts,
    )


@app.post("/fill_powerpoint_file")
async def fill_powerpoint_file(
    pptx_file: UploadFile = File(...),
    replacements_json: str = Form(...),
    remove_image_placeholders: bool = Form(True),
    x_api_key: str | None = Header(default=None),
):
    """
    Best for local/internal tools.
    Returns a ZIP containing:
    - completed_report.pptx
    - image_prompts.json
    - fill_report.json
    """
    _check_api_key(x_api_key)

    if not pptx_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Please upload a .pptx file.")

    try:
        pptx_bytes = await pptx_file.read()
        replacements = json.loads(replacements_json)
        completed_bytes, report = fill_pptx_bytes(
            pptx_bytes,
            replacements,
            remove_image_placeholders=remove_image_placeholders,
        )
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="replacements_json is not valid JSON.")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("completed_report.pptx", completed_bytes)
        zf.writestr("image_prompts.json", json.dumps(report.image_prompts, ensure_ascii=False, indent=2))
        zf.writestr("fill_report.json", json.dumps(asdict(report), ensure_ascii=False, indent=2))

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="filled_powerpoint_output.zip"'},
    )
