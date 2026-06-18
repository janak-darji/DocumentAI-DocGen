"""Application-specific exceptions for structured API errors."""

from __future__ import annotations

import logging
from typing import Optional

logger = logging.getLogger("docgen-service")


class AppError(Exception):
    """Base exception with HTTP status and machine-readable code."""

    def __init__(
        self,
        message: str,
        status_code: int,
        error_code: str,
        *,
        detail: Optional[str] = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.status_code = status_code
        self.error_code = error_code
        self.detail = detail or message

    def log(self, request_id: str, *, path: str | None = None) -> None:
        """
        Logs the exact underlying error server-side.

        Uses the exception cause chain when present so wrapped errors are not lost.
        """
        context = {
            "request_id": request_id,
            "error_code": self.error_code,
            "client_message": self.message,
            "detail": self.detail,
        }
        if path:
            context["path"] = path

        if self.__cause__ is not None:
            logger.error(
                "Request failed: %s",
                context,
                exc_info=self.__cause__,
            )
            return

        logger.error("Request failed: %s", context, exc_info=self)


class UnsupportedFileTypeError(AppError):
    """Raised when the uploaded file type is not supported."""

    def __init__(self, message: str = "Unsupported file type", *, detail: str | None = None) -> None:
        super().__init__(
            message,
            status_code=415,
            error_code="UNSUPPORTED_FILE_TYPE",
            detail=detail,
        )


class CorruptFileError(AppError):
    """Raised when the uploaded file cannot be read or is corrupt."""

    def __init__(
        self,
        message: str = "Corrupt or unreadable file",
        *,
        detail: str | None = None,
    ) -> None:
        super().__init__(
            message,
            status_code=422,
            error_code="CORRUPT_FILE",
            detail=detail,
        )


class DocumentParseError(AppError):
    """Raised when SWMS table extraction fails."""

    def __init__(
        self,
        message: str = "Document parsing failed",
        *,
        detail: str | None = None,
    ) -> None:
        super().__init__(
            message,
            status_code=422,
            error_code="DOCUMENT_PARSE_ERROR",
            detail=detail,
        )
