"""FastAPI application entry point for docgen-service."""

from __future__ import annotations

import uuid

from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from app.exceptions import AppError
from app.logger import configure_logging, logger
from app.routes.parse import router as parse_router
from app.routes.render import router as render_router
from app.security import configured_api_key, is_production_environment

load_dotenv()

configure_logging()

if not configured_api_key() and not is_production_environment():
    logger.warning(
        "DOCGEN_API_KEY is not set — /parse and /render are unauthenticated (development only)",
    )
elif not configured_api_key() and is_production_environment():
    logger.error("DOCGEN_API_KEY must be set in production")

app = FastAPI(title="SWMS Docgen Service", version="1.0.0")
app.include_router(parse_router)
app.include_router(render_router)


def _error_payload(
    code: str,
    message: str,
    request_id: str,
    *,
    detail: str | None = None,
) -> dict:
    error: dict[str, str] = {"code": code, "message": message, "requestId": request_id}
    if detail and detail != message:
        error["detail"] = detail
    return {"error": error}


@app.middleware("http")
async def attach_request_id(request: Request, call_next):
    """Attach a request ID to every request/response."""
    request_id = request.headers.get("x-request-id") or str(uuid.uuid4())
    request.state.request_id = request_id
    response = await call_next(request)
    response.headers["x-request-id"] = request_id
    return response


@app.exception_handler(AppError)
async def app_error_handler(request: Request, exc: AppError) -> JSONResponse:
    """Maps application errors to structured JSON responses and logs exact causes."""
    request_id = getattr(request.state, "request_id", str(uuid.uuid4()))
    exc.log(request_id, path=request.url.path)
    return JSONResponse(
        status_code=exc.status_code,
        content=_error_payload(
            exc.error_code,
            exc.message,
            request_id,
            detail=exc.detail,
        ),
    )


@app.exception_handler(RequestValidationError)
async def validation_error_handler(
    request: Request,
    exc: RequestValidationError,
) -> JSONResponse:
    """Maps request validation failures to structured JSON responses."""
    request_id = getattr(request.state, "request_id", str(uuid.uuid4()))
    message = "; ".join(
        f"{'.'.join(str(part) for part in error.get('loc', []))}: {error.get('msg')}"
        for error in exc.errors()
    )
    logger.error(
        "Validation error on %s: %s",
        request.url.path,
        message,
        extra={"request_id": request_id, "errors": exc.errors()},
    )
    return JSONResponse(
        status_code=422,
        content=_error_payload("VALIDATION_ERROR", message, request_id),
    )


@app.exception_handler(Exception)
async def unhandled_error_handler(request: Request, exc: Exception) -> JSONResponse:
    """Maps unexpected exceptions to a generic 500 response and logs the full traceback."""
    request_id = getattr(request.state, "request_id", str(uuid.uuid4()))
    logger.exception(
        "Unhandled error on %s",
        request.url.path,
        extra={"request_id": request_id},
    )
    return JSONResponse(
        status_code=500,
        content=_error_payload("INTERNAL_ERROR", "Internal server error", request_id),
    )


@app.get("/health")
async def health() -> dict[str, str]:
    """Health check endpoint."""
    return {"status": "ok"}
