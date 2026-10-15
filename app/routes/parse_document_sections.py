"""Document sections and General Information field extraction route."""

from __future__ import annotations

from fastapi import APIRouter, Depends, File, UploadFile

from app.exceptions import PayloadTooLargeError, UnsupportedFileTypeError
from app.schemas.document_sections import ParsedDocumentSectionsResponse
from app.security import max_upload_bytes, verify_docgen_api_key
from app.services.document_section_parser import parse_docx_sections_bytes

router = APIRouter(tags=["sections"], dependencies=[Depends(verify_docgen_api_key)])

SUPPORTED_EXTENSIONS = {".docx", ".doc"}
SUPPORTED_MIME_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
}


@router.post("/parse-document-sections", response_model=ParsedDocumentSectionsResponse)
async def parse_document_sections_route(
    file: UploadFile = File(...),
) -> ParsedDocumentSectionsResponse:
    """
    Parses an uploaded DOC/DOCX document and returns structured sections,
    subsections hierarchy, and dynamic General Information fields.
    """
    if not file.filename:
        raise UnsupportedFileTypeError("File name is required")

    extension = file.filename.lower()
    if not any(extension.endswith(ext) for ext in SUPPORTED_EXTENSIONS):
        raise UnsupportedFileTypeError("Only DOC and DOCX files are supported; PDF and other formats are not accepted")

    content = await file.read()

    if len(content) > max_upload_bytes():
        raise PayloadTooLargeError(
            "Uploaded file is too large",
            detail=f"Maximum upload size is {max_upload_bytes()} bytes",
        )

    return parse_docx_sections_bytes(content)
