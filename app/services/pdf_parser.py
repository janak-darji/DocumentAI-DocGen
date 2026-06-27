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
from app.services.swms_title import pick_swms_title_from_text

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


def _extract_swms_title_from_page(page) -> str | None:
    """Extracts SWMS title text from the region directly above the best SWMS table."""
    for header_ratio, footer_ratio in PAGE_CROP_ATTEMPTS:
        content_page = _crop_page_content_area(
            page,
            header_ratio=header_ratio,
            footer_ratio=footer_ratio,
        )

        found_tables = content_page.find_tables(table_settings=PRIMARY_TABLE_SETTINGS)
        if not found_tables:
            found_tables = content_page.find_tables()

        best_table_obj = None
        best_score = 0
        for table_obj in found_tables:
            extracted = table_obj.extract()
            if not extracted or is_footer_only_table(extracted):
                continue
            table_score = score_swms_table(extracted)
            if table_score > best_score:
                best_score = table_score
                best_table_obj = table_obj

        if best_table_obj is None or best_score <= 0:
            continue

        page_x0, page_top, page_x1, _page_bottom = content_page.bbox
        _table_x0, table_top, _table_x1, _table_bottom = best_table_obj.bbox

        if table_top <= page_top + 5:
            continue

        title_region = content_page.crop((page_x0, page_top, page_x1, table_top))
        title = pick_swms_title_from_text(title_region.extract_text())
        if title:
            return title

    return None


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


def parse_pdf_swms(file_obj: BinaryIO) -> dict:
    """
    Extracts SWMS steps and title from a PDF file.

    Args:
        file_obj: Readable binary stream positioned at the start of the PDF.

    Returns:
        Dict with `steps` and `swmsTitle`.

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
    swms_title: str | None = None

    with pdf:
        for page in pdf.pages:
            if swms_title is None:
                swms_title = _extract_swms_title_from_page(page)

            mapped_rows = _extract_mapped_rows_from_page(page)
            if mapped_rows:
                all_rows.extend(mapped_rows)

    if not all_rows:
        raise DocumentParseError("No SWMS tables found in PDF")

    all_rows = _normalize_pdf_control_rows(all_rows)
    grouped = group_swms_rows(all_rows, control_parser=parse_pdf_controls)
    steps = merge_grouped_steps(grouped)

    return {
        "steps": steps,
        "swmsTitle": swms_title or "",
    }


def parse_pdf_bytes(content: bytes) -> dict:
    """
    Parses SWMS content from PDF bytes.

    Args:
        content: Raw PDF bytes.

    Returns:
        Dict with `steps` and `swmsTitle`.
    """
    return parse_pdf_swms(io.BytesIO(content))
