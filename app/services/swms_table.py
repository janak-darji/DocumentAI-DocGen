"""Shared SWMS table normalization and merged-cell grouping logic."""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Optional

ControlParser = Callable[[str], list[str]]

COLUMN_ALIASES: dict[str, list[str]] = {
    "step_no": [
        "step no",
        "step no.",
        "item no",
        "item no.",
        "no.",
        "no",
        "step",
    ],
    "job_task_element": [
        "job/task element",
        "job task element",
        "job / task element",
        "task element",
        "job element",
    ],
    "hazard": [
        "potential hazard",
        "potential hazards",
        "hazard",
        "hazards",
        "identified hazard",
    ],
    "risk_level": [
        "risk level",
        "risk rating",
        "risk score",
        "initial risk",
        "risk level (l)",
        "risk level before",
    ],
    "controls": [
        "risk control measures",
        "risk control measure",
        "risk controls",
        "risk control",
        "control measures",
        "control measure",
        "controls",
        "control",
        "hierarchy of control",
        "recommended control",
        "recommended controls",
        "what are the controls",
        "control measures to be implemented",
    ],
    "post_risk_level": [
        "post risk level",
        "residual risk",
        "post-risk level",
        "residual risk level",
        "risk level (post)",
        "risk level after",
    ],
    "responsible_person": [
        "responsible person",
        "person responsible",
        "responsible",
        "persons responsible",
    ],
}

REQUIRED_COLUMNS = {"step_no", "job_task_element", "hazard"}
OPTIONAL_COLUMNS = {"controls", "risk_level", "post_risk_level", "responsible_person"}

# Short aliases only match on exact header text, not as substrings.
_SHORT_ALIAS_MAX_LEN = 4

STANDARD_COLUMN_ORDER = [
    "step_no",
    "job_task_element",
    "hazard",
    "risk_level",
    "controls",
    "post_risk_level",
    "responsible_person",
]

_PAGE_FOOTER_PATTERN = re.compile(
    r"^(?:"
    r"page\s*\d+(?:\s*(?:of|/|-)\s*\d+)?|"
    r"\d+\s*/\s*\d+|"
    r"page\s+\d+.*"
    r")\s*$",
    re.IGNORECASE,
)

_PAGE_META_FOOTER_PATTERN = re.compile(
    r"(?:"
    r"construction\s+safety\s+wise|"
    r"swms\d+|"
    r"all\s+rights\s+reserved|"
    r"issue\s+date\s*:|"
    r"version\s+no\s*:|"
    r"authorised\s+by\s*:|"
    r"authori[sz]ed\s+by\s*:"
    r")",
    re.IGNORECASE,
)

_PPE_TABLE_FOOTER_PATTERN = re.compile(
    r"(?:"
    r"\*?\s*personal\s+protective\s+equipment|"
    r"\(ppe\)\s*:|"
    r"mandatory\s+ppe\s+before\s+entering|"
    r"wear\s+corre?ct\s+ppe\s+for\s+tasks"
    r")",
    re.IGNORECASE,
)

_HIERARCHY_LEGEND_PATTERN = re.compile(
    r"^\d+\.\s+"
    r"(?:elimination|substitute|substitution|engineering|administration|ppe|isolate)\b",
    re.IGNORECASE,
)

_HIERARCHY_HEADER_PATTERN = re.compile(
    r"\(\s*how\s+to\s+control\s+the\s+risk\s*\)",
    re.IGNORECASE,
)

_HIERARCHY_LEGEND_CONTROLS_PATTERN = re.compile(
    r"^\d+\.\s+\w+.*\(\s*(?:most|least)\s+effective\s*\)",
    re.IGNORECASE,
)

_CONTINUATION_PATTERN = re.compile(r"\bcontinued\s+(?:on|overleaf)\b", re.IGNORECASE)

_FOOTER_KEYWORD_PATTERN = re.compile(
    r"\b(?:"
    r"prepared\s+by|approved\s+by|reviewed\s+by|authorised\s+by|authorized\s+by|"
    r"signature|signed\s+by|date\s+signed|workers?\s+signature|employee\s+signature|"
    r"document\s+(?:no|number|#)|doc(?:ument)?\s*(?:no|rev|version)|"
    r"revision\s+(?:no|number|date)|issue\s+date|version\s+no|"
    r"this\s+swms|safe\s+work\s+method\s+statement\s+must|must\s+be\s+read\s+and\s+understood|"
    r"workers?\s+acknowledg|acknowledgement\s+of|"
    r"end\s+of\s+(?:swms|document)|"
    r"confidential|copyright|all\s+rights\s+reserved|"
    r"risk\s+(?:rating\s+)?legend|risk\s+matrix"
    r")\b",
    re.IGNORECASE,
)

_VALID_RISK_LEVEL_PATTERN = re.compile(
    r"^(?:"
    r"[elmh]|"
    r"low|medium|high|extreme|"
    r"c[1-5]|"
    r"\d{1,2}"
    r")$",
    re.IGNORECASE,
)

_DISCLAIMER_MIN_LENGTH = 80


def _row_combined_text(row: dict[str, str]) -> str:
    """Joins all mapped cell values into a single lowercase string."""
    return " ".join(value.strip() for value in row.values() if value and value.strip())


def _looks_like_page_footer_text(text: str) -> bool:
    """Returns True when text matches common PDF page footer formats."""
    stripped = text.strip()
    if not stripped:
        return False
    if _PAGE_FOOTER_PATTERN.match(stripped):
        return True
    if "©" in stripped:
        return True
    if _PAGE_META_FOOTER_PATTERN.search(stripped):
        return True
    return bool(re.search(r"\bpage\s+\d+\s+of\s+\d+\b", stripped, re.IGNORECASE))


def _extract_primary_risk_level(value: str) -> str:
    """Returns the first risk token when PDF cells contain duplicated values."""
    token = value.strip().split()[0] if value.strip() else ""
    return token


def _has_valid_risk_level(value: str) -> bool:
    """Returns True when a row has a recognisable SWMS risk rating."""
    if not value.strip():
        return False
    return _is_valid_risk_level(_extract_primary_risk_level(value))


def _looks_like_hierarchy_legend_row(row: dict[str, str]) -> bool:
    """Detects hierarchy-of-control legend rows printed below SWMS tables."""
    combined = _row_combined_text(row)
    if _HIERARCHY_HEADER_PATTERN.search(combined):
        return True

    controls = row.get("controls", "").strip()
    if not controls:
        return False

    if _HIERARCHY_LEGEND_PATTERN.match(controls):
        return True
    if _HIERARCHY_LEGEND_CONTROLS_PATTERN.match(controls):
        return True

    if not row.get("hazard", "").strip() and not row.get("step_no", "").strip():
        if re.search(r"\(\s*least\s+effective\s*\)", controls, re.IGNORECASE):
            return True

    return False


def _looks_like_ppe_footer_row(row: dict[str, str]) -> bool:
    """Detects the recurring PPE disclaimer row at the bottom of SWMS tables."""
    combined = _row_combined_text(row)
    return bool(_PPE_TABLE_FOOTER_PATTERN.search(combined))


def _row_data_quality(row: dict[str, str]) -> int:
    """Scores how complete a mapped SWMS row is from 0-4."""
    quality = 0
    if row.get("hazard", "").strip():
        quality += 1
    if _has_valid_risk_level(row.get("risk_level", "")):
        quality += 1
    if row.get("controls", "").strip():
        quality += 1
    if row.get("step_no", "").strip() or row.get("job_task_element", "").strip():
        quality += 1
    return quality


def _looks_like_fragment_row(row: dict[str, str]) -> bool:
    """
    Detects rows produced when pdfplumber splits footer or page-break text across cells.

    These rows often carry a numeric step number but only partial phrases in each column.
    """
    step_no = row.get("step_no", "").strip()
    if step_no and not re.match(r"^\d+$", step_no):
        return True

    hazard = row.get("hazard", "").strip()
    job_task = row.get("job_task_element", "").strip()
    controls = row.get("controls", "").strip()
    risk_level = row.get("risk_level", "").strip()
    has_risk = _has_valid_risk_level(risk_level)

    if hazard and not has_risk and not controls:
        return True

    if hazard and not has_risk and len(hazard) < 45:
        return True

    if not step_no:
        return False

    if has_risk:
        return False

    if hazard and len(hazard) >= 30:
        return False

    if job_task and len(job_task) >= 40 and hazard:
        return False

    if job_task and len(job_task) < 35 and not hazard:
        return True

    if hazard and len(hazard) < 28 and len(controls) < 35:
        return True

    if controls and len(controls) < 18 and not hazard:
        return True

    if risk_level and not has_risk and len(risk_level) < 12:
        return True

    return False


def _is_valid_risk_level(value: str) -> bool:
    """Returns True when a cell value looks like a SWMS risk rating."""
    return bool(_VALID_RISK_LEVEL_PATTERN.match(value.strip()))


def _looks_like_repeated_header_row(row: dict[str, str]) -> bool:
    """Detects header rows repeated at page/table breaks."""
    step_no = row.get("step_no", "").strip()
    if step_no and re.match(r"^\d+$", step_no):
        return False

    risk_level = row.get("risk_level", "").strip()
    if risk_level and _is_valid_risk_level(risk_level):
        return False

    exact_header_hits = 0
    for value in row.values():
        cell = value.strip()
        if not cell:
            continue

        normalized = _normalize_header_text(cell)
        canonical = normalize_header(cell)
        if not canonical:
            continue

        aliases = COLUMN_ALIASES.get(canonical, [])
        if normalized in aliases:
            exact_header_hits += 1

    return exact_header_hits >= 2


def _looks_like_risk_legend_row(row: dict[str, str]) -> bool:
    """Detects risk-rating legend rows often printed below SWMS tables."""
    combined = _row_combined_text(row).lower()
    if "legend" in combined or "risk matrix" in combined:
        return True

    cells = [value.strip() for value in row.values() if value and value.strip()]
    if len(cells) < 2:
        return False

    short_legend_tokens = {"l", "m", "h", "e", "low", "medium", "high", "extreme"}
    if all(cell.lower() in short_legend_tokens for cell in cells):
        return True

    return bool(
        re.search(
            r"\b[elmh]\s*[-=:]\s*(?:low|medium|high|extreme)\b",
            combined,
            re.IGNORECASE,
        )
    )


def _looks_like_disclaimer_row(row: dict[str, str]) -> bool:
    """Detects long disclaimer or note rows without step/hazard structure."""
    step_no = row.get("step_no", "").strip()
    hazard = row.get("hazard", "").strip()
    risk_level = row.get("risk_level", "").strip()
    job_task = row.get("job_task_element", "").strip()
    combined = _row_combined_text(row)

    if step_no or hazard:
        return False

    if risk_level and _is_valid_risk_level(risk_level):
        return False

    if len(combined) >= _DISCLAIMER_MIN_LENGTH:
        return True

    if job_task and (
        _FOOTER_KEYWORD_PATTERN.search(job_task)
        or _CONTINUATION_PATTERN.search(job_task)
    ):
        return True

    return False


def is_non_swms_data_row(row: dict[str, str]) -> bool:
    """
    Returns True when a mapped row is a page footer, table footer, legend, or signature block.

    These rows are common in SWMS PDF exports and must not become steps or hazards.
    """
    combined = _row_combined_text(row)
    if not combined:
        return True

    for value in row.values():
        cell = value.strip()
        if cell and _looks_like_page_footer_text(cell):
            return True

    if _FOOTER_KEYWORD_PATTERN.search(combined):
        hazard = row.get("hazard", "").strip()
        risk_level = row.get("risk_level", "").strip()
        if not hazard or not _is_valid_risk_level(risk_level):
            return True

    if _CONTINUATION_PATTERN.search(combined):
        return True

    if _looks_like_repeated_header_row(row):
        return True

    if _looks_like_risk_legend_row(row):
        return True

    if _looks_like_disclaimer_row(row):
        return True

    if _looks_like_hierarchy_legend_row(row):
        return True

    if _looks_like_ppe_footer_row(row):
        return True

    if _looks_like_fragment_row(row):
        return True

    return False


def filter_swms_data_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    """Removes footer, legend, and other non-data rows from mapped SWMS rows."""
    return [row for row in rows if not is_non_swms_data_row(row)]


def is_footer_only_table(table: list[list]) -> bool:
    """Returns True when a raw extracted table contains only page/footer text."""
    if not table:
        return True

    all_text = " ".join(str(cell or "").strip() for row in table for cell in row).strip()
    if not all_text:
        return True

    if resolve_table_layout(table) is not None:
        return False

    if len(table) <= 2 and (
        _looks_like_page_footer_text(all_text)
        or _PPE_TABLE_FOOTER_PATTERN.search(all_text) is not None
    ):
        return True

    return False


def score_swms_table(table: list[list]) -> int:
    """
    Scores an extracted PDF table for SWMS data quality.

    Higher scores indicate tables with more valid hazard rows and fewer footer artefacts.
    """
    if is_footer_only_table(table):
        return 0

    layout = resolve_table_layout(table)
    if layout is None:
        return 0

    data_start, column_map = layout
    score = _score_column_map(column_map)[0] * 1000
    data_rows = 0

    for row in table[data_start:]:
        cells = [str(cell or "") for cell in row]
        mapped = _map_cells_to_row(cells, column_map)

        if not any(mapped.values()) or is_non_swms_data_row(mapped):
            continue

        if _row_is_scorable_data(mapped):
            data_rows += 1
            score += _row_data_quality(mapped) * 6

    if data_rows == 0:
        return 0

    return score + (data_rows * 10)


def select_best_swms_table(tables: list[list[list]]) -> Optional[list[list]]:
    """Selects the highest-quality SWMS table from multiple pdfplumber extractions."""
    best_table: Optional[list[list]] = None
    best_score = 0

    for table in tables:
        score = score_swms_table(table)
        if score > best_score:
            best_score = score
            best_table = table

    if best_table is not None:
        return best_table

    best_mapped = 0
    for table in tables:
        mapped_count = len(map_table_rows_from_raw_table(table))
        if mapped_count > best_mapped:
            best_mapped = mapped_count
            best_table = table

    return best_table if best_mapped > 0 else None


def merge_grouped_steps(steps: list[dict]) -> list[dict]:
    """
    Merges steps that share the same step number across multi-page PDF tables.

    Keeps the longest job/task description and combines hazards without duplicates.
    """
    merged_order: list[str] = []
    merged_steps: dict[str, dict] = {}

    for step in steps:
        step_no = step.get("stepNo", "").strip()
        if not step_no:
            continue

        if step_no not in merged_steps:
            merged_steps[step_no] = {
                "stepNo": step_no,
                "jobTaskElement": step.get("jobTaskElement", ""),
                "sequencePosition": step.get("sequencePosition", 0),
                "hazards": list(step.get("hazards", [])),
            }
            merged_order.append(step_no)
            continue

        existing = merged_steps[step_no]
        incoming_task = step.get("jobTaskElement", "")
        if len(incoming_task) > len(existing.get("jobTaskElement", "")):
            existing["jobTaskElement"] = incoming_task

        seen = {
            (
                hazard.get("hazard", ""),
                hazard.get("riskLevel", ""),
                tuple(hazard.get("controls", [])),
            )
            for hazard in existing["hazards"]
        }
        for hazard in step.get("hazards", []):
            key = (
                hazard.get("hazard", ""),
                hazard.get("riskLevel", ""),
                tuple(hazard.get("controls", [])),
            )
            if key in seen:
                continue
            seen.add(key)
            existing["hazards"].append(hazard)

    def _step_sort_key(step_no: str) -> tuple[int, str]:
        if step_no.isdigit():
            return (0, f"{int(step_no):06d}")
        return (1, step_no)

    merged_order.sort(key=_step_sort_key)
    result = [merged_steps[step_no] for step_no in merged_order]
    for index, step in enumerate(result, start=1):
        step["sequencePosition"] = index
    return result


def _normalize_header_text(header: str) -> str:
    """Normalizes header text for alias matching."""
    text = header.strip().lower()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^\w\s/]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _alias_matches(normalized: str, alias: str) -> bool:
    """Returns True when an alias matches a normalized header string."""
    if normalized == alias:
        return True
    if len(alias) <= _SHORT_ALIAS_MAX_LEN:
        return False
    return alias in normalized


def normalize_header(header: str) -> Optional[str]:
    """
    Maps a table header cell to a canonical SWMS column key.

    Args:
        header: Raw header text from a PDF/DOCX table.

    Returns:
        Canonical column key or None when the header is not recognized.
    """
    normalized = _normalize_header_text(header)
    if not normalized:
        return None

    exact_matches: list[tuple[str, str]] = []
    partial_matches: list[tuple[int, str, str]] = []

    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if normalized == alias:
                exact_matches.append((canonical, alias))
            elif _alias_matches(normalized, alias):
                partial_matches.append((len(alias), canonical, alias))

    if exact_matches:
        exact_matches.sort(key=lambda item: len(item[1]), reverse=True)
        return exact_matches[0][0]

    if partial_matches:
        partial_matches.sort(key=lambda item: item[0], reverse=True)
        return partial_matches[0][1]

    return None


def build_column_map(headers: list[str]) -> dict[int, str]:
    """
    Builds a column index map from a header row.

    When multiple headers resolve to the same canonical key, the longest header wins.
    """
    candidates: list[tuple[int, str, str]] = []
    for index, header in enumerate(headers):
        canonical = normalize_header(header)
        if canonical:
            candidates.append((index, canonical, header))

    grouped: dict[str, list[tuple[int, str]]] = {}
    for index, canonical, header in candidates:
        grouped.setdefault(canonical, []).append((index, header))

    column_map: dict[int, str] = {}
    for canonical, options in grouped.items():
        best_index, _ = max(options, key=lambda item: len(item[1].strip()))
        column_map[best_index] = canonical

    return enrich_column_map(headers, column_map)


def enrich_column_map(headers: list[str], column_map: dict[int, str]) -> dict[int, str]:
    """
    Infers missing optional columns (especially controls) using standard SWMS order.

    PDF table extraction often misreads control headers; positional inference recovers them.
    """
    if "controls" in column_map.values():
        return column_map

    result = dict(column_map)
    num_columns = len(headers)

    hazard_indices = [index for index, key in result.items() if key == "hazard"]
    if hazard_indices:
        hazard_index = hazard_indices[0]
        for offset in (2, 1, 3):
            candidate = hazard_index + offset
            if candidate < num_columns and candidate not in result:
                result[candidate] = "controls"
                return result

    if num_columns >= len(STANDARD_COLUMN_ORDER):
        controls_index = STANDARD_COLUMN_ORDER.index("controls")
        if controls_index < num_columns and controls_index not in result:
            result[controls_index] = "controls"

    return result


def _merge_header_rows(header_rows: list[list]) -> list[str]:
    """Merges stacked header rows (common in SWMS PDF tables)."""
    if not header_rows:
        return []

    max_len = max(len(row) for row in header_rows)
    merged: list[str] = []

    for index in range(max_len):
        parts: list[str] = []
        for row in header_rows:
            if index < len(row):
                cell = str(row[index] or "").strip()
                if cell:
                    parts.append(cell)
        merged.append(" ".join(parts))

    return merged


def _score_column_map(column_map: dict[int, str]) -> tuple[int, int, int]:
    """Scores a column map — higher is better."""
    values = set(column_map.values())
    required = len(REQUIRED_COLUMNS & values)
    optional = len(OPTIONAL_COLUMNS & values)
    has_controls = 1 if "controls" in values else 0
    return required, optional, has_controls


def resolve_table_layout(table: list[list]) -> Optional[tuple[int, dict[int, str]]]:
    """
    Detects header row(s) and returns the first data row index plus column map.

    SWMS PDFs often use two-row headers, e.g.:
      | Risk | Risk      |
      | Level| Control   |
    """
    best: Optional[tuple[int, dict[int, str], tuple[int, int, int]]] = None

    scan_limit = min(5, len(table))
    for header_start in range(scan_limit):
        for header_row_count in (1, 2):
            header_end = header_start + header_row_count
            if header_end >= len(table):
                continue

            header_rows = table[header_start:header_end]
            headers = _merge_header_rows(header_rows)
            column_map = enrich_column_map(headers, build_column_map(headers))
            score = _score_column_map(column_map)

            if score[0] < len(REQUIRED_COLUMNS):
                continue

            data_start = header_end
            if best is None or score > best[2]:
                best = (data_start, column_map, score)

    if best is None:
        return None

    return best[0], best[1]


def _normalize_cell_value(value: str, canonical: str) -> str:
    """Normalizes a table cell, preserving line breaks for control lists."""
    text = str(value or "").strip()
    if canonical == "controls":
        return text
    return re.sub(r"\s+", " ", text)


_NUMBERED_BULLET_LINE = re.compile(r"^\d+[\.\)\]]\s*")
_NUMBERED_BULLET_SPLIT = re.compile(r"(?:^|\n)\s*\d+[\.\)\]]\s*")
_INLINE_NUMBERED_BULLET = re.compile(r"(?<!\w)\d+[\.\)\]]\s+")

# Symbol bullets common in SWMS PDF exports (Word/PDF encodings included).
_SYMBOL_CHAR_CLASS = (
    r"[\u2022\u00B7\u25AA\u25A0\u25CF\u25CB\u25E6\u2023\u2043\u2219"
    r"\uf0b7\uf076\uf0a7•●○◦▪■\*]"
)
_SYMBOL_BULLET_LINE = re.compile(
    rf"^\s*(?:{_SYMBOL_CHAR_CLASS}|[-–—])\s+",
)
_SYMBOL_BULLET_SPLIT = re.compile(
    rf"(?:^|[\n\r])\s*(?:{_SYMBOL_CHAR_CLASS}|[-–—])\s+",
)
_INLINE_SYMBOL_BULLET = re.compile(
    rf"(?:(?<=^)|(?<=[\n\r])|\s)(?:{_SYMBOL_CHAR_CLASS}|[-–—])\s+",
)


def normalize_pdf_control_cell(text: str) -> str:
    """
    Normalizes PDF-extracted control text before symbol-bullet parsing.

    PDF line breaks are often soft wraps inside a single bulleted item. This joins
    continuation lines and keeps breaks only before symbol bullets.

    Args:
        text: Raw control cell text from a PDF table cell.

    Returns:
        Control text with spurious line breaks removed.
    """
    lines = re.split(r"[\n\r]+", text.strip())
    if not lines:
        return ""

    merged: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        starts_new_item = bool(_SYMBOL_BULLET_LINE.match(stripped))

        if starts_new_item or not merged:
            merged.append(stripped)
            continue

        merged[-1] = f"{merged[-1]} {stripped}"

    return "\n".join(merged)


def _clean_control_text(text: str) -> str:
    """Collapses internal whitespace in a single control item."""
    return re.sub(r"\s+", " ", text.strip())


def parse_controls(raw_controls: str) -> list[str]:
    """
    Splits a risk-control cell into bullet items.

    Numbered bullets (1., 2., 1), etc.) are the primary delimiter. Newlines are
    only used to detect numbered items, not as standalone separators.

    Args:
        raw_controls: Raw control cell text, often numbered-list separated.

    Returns:
        List of individual control strings.
    """
    text = raw_controls.strip()
    if not text:
        return []

    inline_markers = list(_INLINE_NUMBERED_BULLET.finditer(text))
    if len(inline_markers) >= 2:
        items: list[str] = []
        for index, marker in enumerate(inline_markers):
            start = marker.end()
            end = (
                inline_markers[index + 1].start()
                if index + 1 < len(inline_markers)
                else len(text)
            )
            item = _clean_control_text(text[start:end])
            if item:
                items.append(item)
        if items:
            return items

    if _NUMBERED_BULLET_SPLIT.search(text):
        parts = _NUMBERED_BULLET_SPLIT.split(text)
        items = [_clean_control_text(part) for part in parts if part.strip()]
        if items:
            return items

    if _SYMBOL_BULLET_SPLIT.search(text):
        parts = _SYMBOL_BULLET_SPLIT.split(text)
        items = [_clean_control_text(part) for part in parts if part.strip()]
        if items:
            return items

    if ";" in text:
        items = [_clean_control_text(part) for part in text.split(";") if part.strip()]
        if len(items) > 1:
            return items

    return [_clean_control_text(text)]


def parse_pdf_controls(raw_controls: str) -> list[str]:
    """
    Splits PDF Risk Control cell text on symbol bullets (•, -, ▪, etc.).

    Symbol bullets are the primary delimiter for PDF SWMS documents. Numbered and
    semicolon formats are used only as fallbacks when no symbol bullets are found.

    Args:
        raw_controls: Normalized control cell text from a PDF table.

    Returns:
        List of individual control strings.
    """
    text = raw_controls.strip()
    if not text:
        return []

    inline_markers = list(_INLINE_SYMBOL_BULLET.finditer(text))
    if inline_markers:
        items: list[str] = []
        for index, marker in enumerate(inline_markers):
            start = marker.end()
            end = (
                inline_markers[index + 1].start()
                if index + 1 < len(inline_markers)
                else len(text)
            )
            item = _clean_control_text(text[start:end])
            if item:
                items.append(item)
        if items:
            return items

    if _SYMBOL_BULLET_SPLIT.search(text):
        parts = _SYMBOL_BULLET_SPLIT.split(text)
        items = [_clean_control_text(part) for part in parts if part.strip()]
        if items:
            return items

    if _SYMBOL_BULLET_LINE.match(text):
        without_bullet = _SYMBOL_BULLET_LINE.sub("", text, count=1)
        cleaned = _clean_control_text(without_bullet)
        if cleaned:
            return [cleaned]

    return parse_controls(text)


def _salvage_controls_from_row(cells: list[str], column_map: dict[int, str]) -> str:
    """
    Recovers control text when the mapped controls cell is empty.

    PDF merged cells often leave the controls column blank while text sits in an
  unmapped neighbouring column.
    """
    controls_indices = [index for index, key in column_map.items() if key == "controls"]
    if controls_indices:
        controls_index = controls_indices[0]
        if controls_index < len(cells) and cells[controls_index].strip():
            return cells[controls_index].strip()

    risk_index = next(
        (index for index, key in column_map.items() if key == "risk_level"),
        None,
    )
    post_index = next(
        (index for index, key in column_map.items() if key == "post_risk_level"),
        None,
    )

    start = risk_index + 1 if risk_index is not None else 0
    end = post_index if post_index is not None else len(cells)

    for index in range(start, end):
        if index in column_map:
            continue
        if index >= len(cells):
            continue
        text = cells[index].strip()
        if text:
            return text

    return ""


def _map_cells_to_row(cells: list[str], column_map: dict[int, str]) -> dict[str, str]:
    """Maps a raw table row to canonical SWMS fields, including salvaged controls."""
    mapped: dict[str, str] = {}
    for index, canonical in column_map.items():
        value = cells[index] if index < len(cells) else ""
        mapped[canonical] = _normalize_cell_value(value, canonical)

    salvaged_controls = _salvage_controls_from_row(cells, column_map)
    if salvaged_controls and not mapped.get("controls", "").strip():
        mapped["controls"] = salvaged_controls

    return mapped


def _row_is_scorable_data(row: dict[str, str]) -> bool:
    """Returns True when a mapped row contains usable SWMS hazard data."""
    if is_non_swms_data_row(row):
        return False

    hazard = row.get("hazard", "").strip()
    if not hazard:
        return False

    if _has_valid_risk_level(row.get("risk_level", "")):
        return True

    return bool(row.get("controls", "").strip())


def map_table_rows(headers: list[str], rows: list[list[str]]) -> list[dict[str, str]]:
    """
    Converts raw table rows into canonical SWMS row dictionaries.

    Args:
        headers: Header row values.
        rows: Data rows from the table.

    Returns:
        List of row dicts keyed by canonical column names.
    """
    column_map = enrich_column_map(headers, build_column_map(headers))

    if not REQUIRED_COLUMNS.issubset(set(column_map.values())):
        return []

    mapped_rows: list[dict[str, str]] = []
    for row in rows:
        cells = [str(cell or "") for cell in row]
        mapped = _map_cells_to_row(cells, column_map)

        if any(mapped.values()) and not is_non_swms_data_row(mapped):
            mapped_rows.append(mapped)

    return mapped_rows


def map_table_rows_from_raw_table(table: list[list]) -> list[dict[str, str]]:
    """
    Detects SWMS headers inside a raw PDF/DOCX table and maps data rows.

    Args:
        table: Full table matrix including header row(s).

    Returns:
        Canonical SWMS row dictionaries.
    """
    layout = resolve_table_layout(table)
    if layout is None:
        return []

    data_start, column_map = layout
    mapped_rows: list[dict[str, str]] = []

    for row in table[data_start:]:
        cells = [str(cell or "") for cell in row]
        mapped = _map_cells_to_row(cells, column_map)

        if any(mapped.values()) and not is_non_swms_data_row(mapped):
            mapped_rows.append(mapped)

    return filter_swms_data_rows(mapped_rows)


_SYMBOL_ONLY_CONTROL = re.compile(r"^[\uf0b7\u2022\u25aa\u2043\u2219•\-\s]+$")


def _is_meaningful_control(control: str) -> bool:
    """Returns False for bullet-only artefacts left by fragmented PDF extraction."""
    text = control.strip()
    return len(text) > 2 and not _SYMBOL_ONLY_CONTROL.match(text)


def _clean_parsed_controls(controls: list[str]) -> list[str]:
    """Removes empty and symbol-only control fragments."""
    return [control for control in controls if _is_meaningful_control(control)]


def group_swms_rows(
    rows: list[dict[str, str]],
    control_parser: ControlParser = parse_controls,
) -> list[dict]:
    """
    Groups raw SWMS table rows into steps with nested hazards.

    Handles merged cells by carrying forward the last non-empty step number and
    job/task element when subsequent hazard rows leave those cells blank.

    Args:
        rows: Canonical SWMS row dictionaries.
        control_parser: Function used to split Risk Control text into a list.

    Returns:
        Structured steps with hazards arrays suitable for API serialization.
    """
    steps: list[dict] = []
    current_step: Optional[dict] = None
    last_step_no = ""
    last_job_task = ""

    for row in rows:
        if is_non_swms_data_row(row):
            continue

        step_no = row.get("step_no", "").strip() or last_step_no
        job_task = row.get("job_task_element", "").strip() or last_job_task

        if row.get("step_no", "").strip():
            last_step_no = step_no
        if row.get("job_task_element", "").strip():
            last_job_task = job_task

        hazard = row.get("hazard", "").strip()
        risk_level = row.get("risk_level", "").strip()
        controls_raw = row.get("controls", "").strip()
        post_risk_level = row.get("post_risk_level", "").strip()
        responsible_person = row.get("responsible_person", "").strip()

        if not hazard and not risk_level and not controls_raw:
            continue

        parsed_controls = _clean_parsed_controls(control_parser(controls_raw))
        if not hazard and not _has_valid_risk_level(risk_level) and not parsed_controls:
            continue

        if not step_no and not job_task:
            continue

        step_changed = (
            current_step is None
            or (
                row.get("step_no", "").strip()
                and current_step["stepNo"] != step_no
            )
            or (
                row.get("job_task_element", "").strip()
                and not row.get("step_no", "").strip()
                and current_step["jobTaskElement"] != job_task
                and hazard
            )
        )

        if step_changed:
            if current_step is not None:
                steps.append(current_step)
            current_step = {
                "stepNo": step_no,
                "jobTaskElement": job_task,
                "sequencePosition": len(steps) + 1,
                "hazards": [],
            }

        assert current_step is not None
        current_step["hazards"].append(
            {
                "hazard": hazard,
                "riskLevel": risk_level,
                "controls": parsed_controls,
                "postRiskLevel": post_risk_level,
                "responsiblePerson": responsible_person,
            }
        )

    if current_step is not None:
        steps.append(current_step)

    return steps
