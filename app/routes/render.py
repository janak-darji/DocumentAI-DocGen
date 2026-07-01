"""SWMS Word document render route."""

from __future__ import annotations

from fastapi import APIRouter, Depends
from fastapi.responses import Response

from app.exceptions import DocumentRenderError
from app.schemas.render import RenderSwmsRequest
from app.security import verify_docgen_api_key
from app.services.swms_renderer import render_swms_document_bytes

router = APIRouter(tags=["render"], dependencies=[Depends(verify_docgen_api_key)])


@router.post("/render")
async def render_swms_document(request: RenderSwmsRequest) -> Response:
    """
    Render a SWMS Word document from the configured template and structured data.

    Returns:
        Generated DOCX bytes.
    """
    try:
        document_bytes = render_swms_document_bytes(request)
    except DocumentRenderError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise DocumentRenderError(
            "Document rendering failed",
            detail=f"{type(exc).__name__}: {exc}",
        ) from exc

    return Response(
        content=document_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="swms-generated.docx"'},
    )
