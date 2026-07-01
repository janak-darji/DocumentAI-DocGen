"""Service-to-service authentication for docgen-service."""

from __future__ import annotations

import os
import secrets

from fastapi import Header

from app.exceptions import UnauthorizedError

API_KEY_HEADER = "X-Docgen-Api-Key"
DEFAULT_MAX_UPLOAD_BYTES = 25 * 1024 * 1024


def configured_api_key() -> str | None:
    """Returns the configured internal API key, if any."""
    key = os.getenv("DOCGEN_API_KEY", "").strip()
    return key or None


def is_production_environment() -> bool:
    """Returns True when the service runs in a production-like environment."""
    environment = os.getenv("ENVIRONMENT", os.getenv("NODE_ENV", "development")).strip().lower()
    return environment == "production"


def max_upload_bytes() -> int:
    """Returns the maximum accepted upload size for /parse."""
    raw = os.getenv("DOCGEN_MAX_UPLOAD_BYTES", "").strip()
    if not raw:
        return DEFAULT_MAX_UPLOAD_BYTES

    try:
        parsed = int(raw, 10)
    except ValueError:
        return DEFAULT_MAX_UPLOAD_BYTES

    return parsed if parsed > 0 else DEFAULT_MAX_UPLOAD_BYTES


async def verify_docgen_api_key(
    x_docgen_api_key: str | None = Header(default=None, alias=API_KEY_HEADER),
) -> None:
    """
    Validates the shared internal API key sent by the backend.

    When DOCGEN_API_KEY is unset in development, requests are allowed with a warning
    logged at startup. In production, the key must be configured and supplied.
    """
    configured = configured_api_key()

    if not configured:
        if is_production_environment():
            raise UnauthorizedError("DOCGEN_API_KEY is not configured on docgen-service")
        return

    if not x_docgen_api_key or not secrets.compare_digest(x_docgen_api_key, configured):
        raise UnauthorizedError("Invalid or missing docgen API key")
