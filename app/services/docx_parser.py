"""DOCX SWMS table extraction using python-docx."""

from __future__ import annotations

import io
from typing import BinaryIO

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from app.exceptions import CorruptFileError, DocumentParseError
from app.services.swms_table import group_swms_rows, map_table_rows_from_raw_table


def _table_to_matrix(table) -> list[list[str]]:
    """Converts a python-docx table to a string matrix."""
    matrix: list[list[str]] = []
    for row in table.rows:
        matrix.append([cell.text.strip() for cell in row.cells])
    return matrix


def parse_docx_swms(file_obj: BinaryIO) -> list[dict]:
    """
    Extracts SWMS steps from a DOCX file.

    Args:
        file_obj: Readable binary stream positioned at the start of the DOCX.

    Returns:
        Structured steps with nested hazards.

    Raises:
        CorruptFileError: When the DOCX cannot be opened.
        DocumentParseError: When no SWMS tables are found.
    """
    try:
        document = Document(file_obj)
    except PackageNotFoundError as exc:
        raise CorruptFileError(
            "Unable to read DOCX file",
            detail=f"{type(exc).__name__}: {exc}",
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise CorruptFileError(
            "Unable to read DOCX file",
            detail=f"{type(exc).__name__}: {exc}",
        ) from exc

    all_steps: list[dict] = []

    for table in document.tables:
        matrix = _table_to_matrix(table)
        if len(matrix) < 2:
            continue

        mapped_rows = map_table_rows_from_raw_table(matrix)
        if not mapped_rows:
            continue

        grouped = group_swms_rows(mapped_rows)
        for step in grouped:
            step["sequencePosition"] = len(all_steps) + 1
            all_steps.append(step)

    if not all_steps:
        raise DocumentParseError("No SWMS tables found in DOCX")

    return all_steps


def parse_docx_bytes(content: bytes) -> list[dict]:
    """
    Parses SWMS content from DOCX bytes.

    Args:
        content: Raw DOCX bytes.

    Returns:
        Structured steps with nested hazards.
    """
    return parse_docx_swms(io.BytesIO(content))
