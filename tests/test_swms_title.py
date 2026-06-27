"""Unit tests for SWMS title extraction."""

from app.services.swms_title import pick_swms_title_from_lines, pick_swms_title_from_text


def test_pick_swms_title_prefers_activity_line_closest_to_table() -> None:
    lines = [
        "SAFE WORK METHOD STATEMENT",
        "Project: Site Alpha",
        "Floor Framing Work At Height Above 2m",
        "Prepared by: Supervisor",
    ]
    assert pick_swms_title_from_lines(lines) == "Floor Framing Work At Height Above 2m"


def test_pick_swms_title_from_text_handles_multiline_block() -> None:
    text = "SWMS\nConcrete Works Including Excavation\nVersion 1.0"
    assert pick_swms_title_from_text(text) == "Concrete Works Including Excavation"


def test_pick_swms_title_returns_none_for_empty_text() -> None:
    assert pick_swms_title_from_text(None) is None
    assert pick_swms_title_from_text("") is None
