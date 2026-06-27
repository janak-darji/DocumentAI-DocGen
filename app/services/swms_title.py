"""SWMS document title extraction from content above the steps table."""

from __future__ import annotations

import re

# Lines that are boilerplate rather than the activity-specific SWMS title.
_TITLE_SKIP_PATTERNS = (
    re.compile(r"^safe work method statement$", re.IGNORECASE),
    re.compile(r"^swms\b", re.IGNORECASE),
    re.compile(r"^page\s+\d+", re.IGNORECASE),
    re.compile(r"^version\b", re.IGNORECASE),
    re.compile(r"^date\b", re.IGNORECASE),
    re.compile(r"^project\b", re.IGNORECASE),
    re.compile(r"^site\b", re.IGNORECASE),
    re.compile(r"^prepared by\b", re.IGNORECASE),
    re.compile(r"^approved by\b", re.IGNORECASE),
    re.compile(r"©", re.IGNORECASE),
    re.compile(r"^hierarchy of control", re.IGNORECASE),
    re.compile(r"^risk\s*rating", re.IGNORECASE),
)

_MAX_TITLE_LENGTH = 500


def _is_boilerplate_line(line: str) -> bool:
    """Returns True when a line is unlikely to be the SWMS activity title."""
    stripped = line.strip()
    if len(stripped) < 4:
        return True
    return any(pattern.search(stripped) for pattern in _TITLE_SKIP_PATTERNS)


def pick_swms_title_from_lines(lines: list[str]) -> str | None:
    """
    Picks the SWMS title from text lines immediately above the steps table.

    Prefers the longest non-boilerplate line closest to the table (bottom-up).
    """
    cleaned = [line.strip() for line in lines if line and line.strip()]
    if not cleaned:
        return None

    candidates: list[str] = []
    for line in reversed(cleaned[-12:]):
        if _is_boilerplate_line(line):
            continue
        candidates.append(line)
        if len(candidates) >= 3:
            break

    if not candidates:
        fallback = cleaned[-1].strip()
        return fallback[:_MAX_TITLE_LENGTH] if fallback else None

    title = max(candidates, key=len).strip()
    return title[:_MAX_TITLE_LENGTH] if title else None


def pick_swms_title_from_text(text: str | None) -> str | None:
    """Extracts a SWMS title from free-form text above a table."""
    if not text:
        return None
    lines = text.splitlines()
    return pick_swms_title_from_lines(lines)
