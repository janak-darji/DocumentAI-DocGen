"""PDF SWMS table extraction using pdfplumber."""

from __future__ import annotations

import io
from typing import BinaryIO

import pdfplumber

from app.exceptions import CorruptFileError, DocumentParseError
from app.services.swms_table import (
    group_swms_rows,
    is_footer_only_table,
    map_table_rows_from_raw_table,
    merge_grouped_steps,
    normalize_pdf_control_cell,
    parse_pdf_controls,
    score_swms_table,
    select_best_swms_table,
)

PRIMARY_TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_tolerance": 5,
    "snap_tolerance": 3,
    "join_tolerance": 3,
}

FALLBACK_TABLE_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "snap_tolerance": 3,
    "join_tolerance": 3,
}

# Exclude repeating page headers/footers from table detection.
# Row-level footer filtering handles PPE disclaimers; keep crop modest so
# step headers at page breaks are not clipped.
PAGE_HEADER_RATIO = 0.05
PAGE_FOOTER_RATIO = 0.08
PAGE_CROP_ATTEMPTS = (
    (PAGE_HEADER_RATIO, PAGE_FOOTER_RATIO),
    (PAGE_HEADER_RATIO, 0.05),
    (0.02, 0.0),
)


def _crop_page_content_area(
    page,
    *,
    header_ratio: float = PAGE_HEADER_RATIO,
    footer_ratio: float = PAGE_FOOTER_RATIO,
):
    """Crops top/bottom page margins where headers and footers commonly appear."""
    x0, top, x1, bottom = page.bbox
    height = bottom - top
    if height <= 0:
        return page

    crop_top = top + height * header_ratio
    crop_bottom = bottom - (height * footer_ratio if footer_ratio > 0 else 0)
    if crop_bottom <= crop_top:
        return page

    return page.crop((x0, crop_top, x1, crop_bottom))


def _extract_page_tables(page) -> list[list[list]]:
    """Extracts SWMS tables, preferring line-based detection over text-based fallbacks."""
    seen: set[str] = set()
    tables: list[list[list]] = []

    def _collect(settings: dict | None) -> None:
        extracted = page.extract_tables(table_settings=settings) if settings else page.extract_tables()
        for table in extracted or []:
            if not table or is_footer_only_table(table):
                continue
            fingerprint = repr(table[:5])
            if fingerprint in seen:
                continue
            seen.add(fingerprint)
            tables.append(table)

    _collect(PRIMARY_TABLE_SETTINGS)
    line_best = select_best_swms_table(tables)
    if line_best is not None and score_swms_table(line_best) > 0:
        return tables

    _collect(FALLBACK_TABLE_SETTINGS)
    _collect({})
    return tables


def _normalize_pdf_control_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    """Applies PDF-specific control-cell normalization before bullet parsing."""
    for row in rows:
        if row.get("controls"):
            row["controls"] = normalize_pdf_control_cell(row["controls"])
    return rows


def _extract_mapped_rows_from_page(page) -> list[dict[str, str]]:
    """Extracts mapped SWMS rows from a page, retrying with lighter crops when needed."""
    for header_ratio, footer_ratio in PAGE_CROP_ATTEMPTS:
        content_page = _crop_page_content_area(
            page,
            header_ratio=header_ratio,
            footer_ratio=footer_ratio,
        )
        tables = _extract_page_tables(content_page)
        best_table = select_best_swms_table(tables)
        if best_table is None or len(best_table) < 2:
            continue

        mapped_rows = map_table_rows_from_raw_table(best_table)
        if mapped_rows:
            return mapped_rows

    return []


def parse_pdf_swms(file_obj: BinaryIO) -> list[dict]:
    """
    Extracts SWMS steps from a PDF file.

    Args:
        file_obj: Readable binary stream positioned at the start of the PDF.

    Returns:
        Structured steps with nested hazards.

    Raises:
        CorruptFileError: When the PDF cannot be opened.
        DocumentParseError: When no SWMS tables are found.
    """
    try:
        pdf = pdfplumber.open(file_obj)
    except Exception as exc:  # noqa: BLE001
        raise CorruptFileError(
            "Unable to read PDF file",
            detail=f"{type(exc).__name__}: {exc}",
        ) from exc

    all_rows: list[dict[str, str]] = []

    with pdf:
        for page in pdf.pages:
            mapped_rows = _extract_mapped_rows_from_page(page)
            if mapped_rows:
                all_rows.extend(mapped_rows)

    if not all_rows:
        raise DocumentParseError("No SWMS tables found in PDF")

    all_rows = _normalize_pdf_control_rows(all_rows)
    grouped = group_swms_rows(all_rows, control_parser=parse_pdf_controls)
    return merge_grouped_steps(grouped)


def parse_pdf_bytes(content: bytes) -> list[dict]:
    """
    Parses SWMS content from PDF bytes.

    Args:
        content: Raw PDF bytes.

    Returns:
        Structured steps with nested hazards.
    """
    return parse_pdf_swms(io.BytesIO(content))
