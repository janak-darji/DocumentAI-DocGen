"""SWMS document parse route."""

from __future__ import annotations

from fastapi import APIRouter, File, Form, UploadFile

from app.exceptions import UnsupportedFileTypeError
from app.schemas.parse import ParsedSwmsResponse
from app.services.docx_parser import parse_docx_bytes
from app.services.pdf_parser import parse_pdf_bytes

router = APIRouter(tags=["parse"])

SUPPORTED_EXTENSIONS = {".pdf", ".docx"}
SUPPORTED_MIME_TYPES = {
    "application/pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def _resolve_file_type(filename: str, content_type: str | None) -> str:
    """
    Determines supported file type from filename extension or MIME type.

    Args:
        filename: Uploaded file name.
        content_type: Optional MIME type from upload.

    Returns:
        Normalized file type (`pdf` or `docx`).

    Raises:
        UnsupportedFileTypeError: When the file type is not supported.
    """
    lowered = filename.lower()
    if lowered.endswith(".pdf"):
        return "pdf"
    if lowered.endswith(".docx"):
        return "docx"

    if content_type in SUPPORTED_MIME_TYPES:
        return "docx" if "wordprocessingml" in content_type else "pdf"

    raise UnsupportedFileTypeError("Only PDF and DOCX files are supported")


@router.post("/parse", response_model=ParsedSwmsResponse)
async def parse_swms_document(
    file: UploadFile = File(...),
    activity_type: str = Form(..., alias="activity_type"),
) -> ParsedSwmsResponse:
    """
    Parse an uploaded SWMS PDF or DOCX and return structured steps and hazards.

    Args:
        file: Uploaded SWMS document.
        activity_type: Activity type metadata echoed in the response.

    Returns:
        Structured SWMS parse result.
    """
    if not file.filename:
        raise UnsupportedFileTypeError("File name is required")

    extension = file.filename.lower()
    if not any(extension.endswith(ext) for ext in SUPPORTED_EXTENSIONS):
        raise UnsupportedFileTypeError("Only PDF and DOCX files are supported")

    file_type = _resolve_file_type(file.filename, file.content_type)
    content = await file.read()

    if file_type == "pdf":
        # PDF Risk Control cells use symbol-bullet parsing (•, -, ▪, etc.).
        steps = parse_pdf_bytes(content)
    else:
        steps = parse_docx_bytes(content)

    return ParsedSwmsResponse(activityType=activity_type, steps=steps)
