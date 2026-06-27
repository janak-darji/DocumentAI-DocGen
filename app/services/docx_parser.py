"""DOCX SWMS table extraction using python-docx."""

from __future__ import annotations

import io
from typing import BinaryIO

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn

from app.exceptions import CorruptFileError, DocumentParseError
from app.services.swms_table import group_swms_rows, map_table_rows_from_raw_table
from app.services.swms_title import pick_swms_title_from_lines


def _table_to_matrix(table) -> list[list[str]]:
    """Converts a python-docx table to a string matrix."""
    matrix: list[list[str]] = []
    for row in table.rows:
        matrix.append([cell.text.strip() for cell in row.cells])
    return matrix


def _extract_docx_swms_title(document: Document) -> str | None:
    """Extracts paragraph text appearing before the first SWMS table."""
    lines: list[str] = []

    for element in document.element.body:
        if element.tag == qn("w:tbl"):
            break
        if element.tag == qn("w:p"):
            texts = [node.text for node in element.iter() if node.text]
            paragraph_text = "".join(texts).strip()
            if paragraph_text:
                lines.append(paragraph_text)

    return pick_swms_title_from_lines(lines)


def parse_docx_swms(file_obj: BinaryIO) -> dict:
    """
    Extracts SWMS steps and title from a DOCX file.

    Args:
        file_obj: Readable binary stream positioned at the start of the DOCX.

    Returns:
        Dict with `steps` and `swmsTitle`.

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

    swms_title = _extract_docx_swms_title(document)
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

    return {
        "steps": all_steps,
        "swmsTitle": swms_title or "",
    }


def parse_docx_bytes(content: bytes) -> dict:
    """
    Parses SWMS content from DOCX bytes.

    Args:
        content: Raw DOCX bytes.

    Returns:
        Dict with `steps` and `swmsTitle`.
    """
    return parse_docx_swms(io.BytesIO(content))
