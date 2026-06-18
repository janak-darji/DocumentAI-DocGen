"""Unit tests for SWMS table grouping and merged-cell detection."""

from app.services.swms_table import (
    build_column_map,
    filter_swms_data_rows,
    group_swms_rows,
    is_non_swms_data_row,
    map_table_rows,
    map_table_rows_from_raw_table,
    merge_grouped_steps,
    normalize_header,
    normalize_pdf_control_cell,
    parse_controls,
    parse_pdf_controls,
    select_best_swms_table,
)


SAMPLE_HEADERS = [
    "Step No.",
    "Job/Task Element",
    "Potential Hazard",
    "Risk Level",
    "Risk Control",
    "Post Risk Level",
    "Responsible Person",
]

SAMPLE_ROWS = [
    [
        "1",
        "Mobilisation to site",
        "Traffic collision",
        "High",
        "1. Use spotter\n2. Wear hi-vis",
        "Low",
        "Supervisor",
    ],
    [
        "",
        "",
        "Uneven ground",
        "Medium",
        "Inspect footing",
        "Low",
        "All workers",
    ],
    [
        "2",
        "Excavation",
        "Cave-in",
        "High",
        "Shoring installed",
        "Low",
        "Leading hand",
    ],
]


def test_normalize_header_maps_common_control_headers() -> None:
    """Risk control column headers map to the controls key."""
    assert normalize_header("Risk Control") == "controls"
    assert normalize_header("Risk Control Measures") == "controls"
    assert normalize_header("Control Measures") == "controls"
    assert normalize_header("Control") == "controls"


def test_normalize_header_does_not_map_risk_substrings_loosely() -> None:
    """Short aliases like 'risk' do not steal 'risk control' headers."""
    assert normalize_header("Risk Control") == "controls"
    assert normalize_header("Risk Level") == "risk_level"


def test_map_table_rows_recognizes_swms_headers() -> None:
    """Headers map to canonical SWMS column keys."""
    mapped = map_table_rows(SAMPLE_HEADERS, SAMPLE_ROWS)

    assert len(mapped) == 3
    assert mapped[0]["step_no"] == "1"
    assert mapped[0]["hazard"] == "Traffic collision"
    assert mapped[0]["controls"] == "1. Use spotter\n2. Wear hi-vis"
    assert mapped[1]["step_no"] == ""
    assert mapped[1]["hazard"] == "Uneven ground"


def test_group_swms_rows_merges_hazards_under_one_step() -> None:
    """Blank step/job cells inherit previous values (merged cells)."""
    mapped = map_table_rows(SAMPLE_HEADERS, SAMPLE_ROWS)
    steps = group_swms_rows(mapped)

    assert len(steps) == 2

    assert steps[0]["stepNo"] == "1"
    assert steps[0]["jobTaskElement"] == "Mobilisation to site"
    assert steps[0]["sequencePosition"] == 1
    assert len(steps[0]["hazards"]) == 2
    assert steps[0]["hazards"][0]["hazard"] == "Traffic collision"
    assert steps[0]["hazards"][0]["controls"] == ["Use spotter", "Wear hi-vis"]
    assert steps[0]["hazards"][1]["hazard"] == "Uneven ground"

    assert steps[1]["stepNo"] == "2"
    assert steps[1]["jobTaskElement"] == "Excavation"
    assert len(steps[1]["hazards"]) == 1
    assert steps[1]["hazards"][0]["controls"] == ["Shoring installed"]


def test_group_swms_rows_splits_controls_by_numbered_bullets() -> None:
    """Risk control cells split on numbered bullets, not bare newlines."""
    mapped = map_table_rows(
        SAMPLE_HEADERS,
        [
            [
                "1",
                "Setup",
                "Falling objects",
                "Medium",
                "1. Barricade area\n2. Hard hats mandatory",
                "Low",
                "Supervisor",
            ]
        ],
    )
    steps = group_swms_rows(mapped)

    assert steps[0]["hazards"][0]["controls"] == [
        "Barricade area",
        "Hard hats mandatory",
    ]


def test_group_swms_rows_does_not_split_unnumbered_newlines() -> None:
    """A single numbered control with a soft line wrap stays as one item."""
    controls = parse_controls("1. Barricade area and restrict\naccess to site")
    assert controls == ["Barricade area and restrict access to site"]


def test_normalize_pdf_control_cell_joins_soft_wraps() -> None:
    """PDF soft wraps inside a symbol-bulleted item are joined before parsing."""
    normalized = normalize_pdf_control_cell(
        "• Use spotter when\nreversing vehicle\n• Wear hi-vis"
    )
    assert normalized == "• Use spotter when reversing vehicle\n• Wear hi-vis"
    assert parse_pdf_controls(normalized) == [
        "Use spotter when reversing vehicle",
        "Wear hi-vis",
    ]


def test_parse_pdf_controls_splits_symbol_bullets_on_newlines() -> None:
    """PDF Risk Control cells split on line-start symbol bullets."""
    controls = parse_pdf_controls("• Isolate power\n• Lock out tag out")
    assert controls == ["Isolate power", "Lock out tag out"]


def test_parse_pdf_controls_splits_inline_symbol_bullets() -> None:
    """Multiple symbol bullets on one line become separate controls."""
    controls = parse_pdf_controls("• Use spotter • Wear hi-vis • Barricade area")
    assert controls == ["Use spotter", "Wear hi-vis", "Barricade area"]


def test_parse_pdf_controls_splits_dash_bullets() -> None:
    """Dash-led bullets at line start are treated as control separators."""
    controls = parse_pdf_controls("- Isolate power\n- Lock out tag out")
    assert controls == ["Isolate power", "Lock out tag out"]


def test_group_swms_rows_uses_pdf_symbol_control_parser() -> None:
    """PDF grouping splits controls by symbol bullets into separate array items."""
    mapped = map_table_rows(
        SAMPLE_HEADERS,
        [
            [
                "1",
                "Setup",
                "Falling objects",
                "Medium",
                "• Barricade area\n• Hard hats mandatory",
                "Low",
                "Supervisor",
            ]
        ],
    )
    steps = group_swms_rows(mapped, control_parser=parse_pdf_controls)

    assert steps[0]["hazards"][0]["controls"] == [
        "Barricade area",
        "Hard hats mandatory",
    ]


def test_parse_controls_handles_inline_numbered_bullets() -> None:
    """Inline numbered controls on one line are split correctly."""
    controls = parse_controls("1. Isolate power 2. Lock out tag out")
    assert controls == ["Isolate power", "Lock out tag out"]


def test_parse_controls_handles_symbol_bullets() -> None:
    """Symbol bullets split controls when no numbering is present."""
    controls = parse_controls("• Isolate power\n• Lock out tag out")
    assert controls == ["Isolate power", "Lock out tag out"]


def test_parse_controls_preserves_hyphenated_phrases() -> None:
    """Hyphens inside a numbered control are preserved."""
    controls = parse_controls("1. Lock-out tag out")
    assert controls == ["Lock-out tag out"]


def test_map_table_rows_from_raw_table_with_two_row_header() -> None:
    """Stacked SWMS headers are merged before column detection."""
    table = [
        ["Item", "Job/Task", "Potential", "Risk", "Risk", "Residual", "Responsible"],
        ["No.", "Element", "Hazard", "Level", "Control", "Risk", "Person"],
        [
            "1",
            "Setup plant",
            "Pinch points",
            "High",
            "Use guards",
            "Low",
            "Supervisor",
        ],
    ]

    mapped = map_table_rows_from_raw_table(table)
    steps = group_swms_rows(mapped)

    assert mapped[0]["controls"] == "Use guards"
    assert steps[0]["hazards"][0]["controls"] == ["Use guards"]


def test_build_column_map_detects_risk_control_measures_header() -> None:
    """Longer control header variants are recognized."""
    headers = [
        "Step No.",
        "Job Task Element",
        "Hazard",
        "Risk Level",
        "Risk Control Measures",
        "Post Risk Level",
        "Responsible Person",
    ]
    column_map = build_column_map(headers)

    assert column_map[4] == "controls"


def test_group_swms_rows_starts_new_step_when_step_number_changes() -> None:
    """A new step number starts a new grouped step."""
    mapped = map_table_rows(
        SAMPLE_HEADERS,
        [
            ["1", "Task A", "Hazard A", "High", "Control A", "Low", "Person A"],
            ["2", "Task B", "Hazard B", "Medium", "Control B", "Low", "Person B"],
        ],
    )
    steps = group_swms_rows(mapped)

    assert len(steps) == 2
    assert steps[0]["stepNo"] == "1"
    assert steps[0]["hazards"][0]["controls"] == ["Control A"]
    assert steps[1]["stepNo"] == "2"
    assert steps[1]["hazards"][0]["controls"] == ["Control B"]


def test_enrich_column_map_infers_controls_by_position() -> None:
    """Controls column is inferred when header text is missing or unrecognised."""
    headers = [
        "Step No.",
        "Job Task Element",
        "Hazard",
        "Risk Level",
        "Measures",
        "Post Risk Level",
        "Responsible Person",
    ]
    table = [
        headers,
        [
            "1",
            "Setup",
            "Crush injury",
            "High",
            "Use machine guards",
            "Low",
            "Supervisor",
        ],
    ]

    mapped = map_table_rows_from_raw_table(table)
    steps = group_swms_rows(mapped)

    assert mapped[0]["controls"] == "Use machine guards"
    assert steps[0]["hazards"][0]["controls"] == ["Use machine guards"]


def test_salvage_controls_from_unmapped_column() -> None:
    """Control text is recovered from an unmapped column between risk and post-risk."""
    headers = [
        "Step No.",
        "Job Task Element",
        "Hazard",
        "Risk Level",
        "Unknown Header",
        "Post Risk Level",
        "Responsible Person",
    ]
    table = [
        headers,
        [
            "1",
            "Setup",
            "Falling objects",
            "High",
            "Barricade work area",
            "Low",
            "Supervisor",
        ],
    ]

    mapped = map_table_rows_from_raw_table(table)
    steps = group_swms_rows(mapped)

    assert steps[0]["hazards"][0]["controls"] == ["Barricade work area"]


def test_parse_controls_handles_semicolon_separated_values() -> None:
    """Semicolon-separated controls are split into list items."""
    assert parse_controls("Isolate power; Lock out tag out") == [
        "Isolate power",
        "Lock out tag out",
    ]


def test_is_non_swms_data_row_filters_page_footer() -> None:
    """Page number footers are excluded from SWMS data rows."""
    assert is_non_swms_data_row({"step_no": "", "job_task_element": "", "hazard": "Page 2 of 5"})
    assert is_non_swms_data_row({"step_no": "", "job_task_element": "Page 3", "hazard": ""})


def test_is_non_swms_data_row_filters_table_disclaimer() -> None:
    """Long disclaimer rows without hazard structure are excluded."""
    row = {
        "step_no": "",
        "job_task_element": (
            "This SWMS must be read and understood by all workers before commencing work "
            "on site and reviewed when conditions change."
        ),
        "hazard": "",
        "risk_level": "",
        "controls": "",
    }
    assert is_non_swms_data_row(row)


def test_is_non_swms_data_row_filters_signature_block() -> None:
    """Signature and approval rows are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "Prepared by",
            "hazard": "",
            "risk_level": "",
            "responsible_person": "Supervisor",
        }
    )


def test_is_non_swms_data_row_filters_repeated_header_row() -> None:
    """Repeated header rows at page breaks are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "Step No.",
            "job_task_element": "Job/Task Element",
            "hazard": "Potential Hazard",
            "risk_level": "Risk Level",
            "controls": "Risk Control",
        }
    )


def test_is_non_swms_data_row_filters_risk_legend_row() -> None:
    """Risk legend rows below tables are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "L",
            "job_task_element": "Low",
            "hazard": "M",
            "risk_level": "Medium",
            "controls": "H",
            "post_risk_level": "High",
        }
    )


def test_is_non_swms_data_row_keeps_valid_hazard_row() -> None:
    """Valid hazard rows with step and risk data are retained."""
    row = {
        "step_no": "1",
        "job_task_element": "Mobilisation",
        "hazard": "Traffic collision",
        "risk_level": "High",
        "controls": "Use spotter",
    }
    assert not is_non_swms_data_row(row)


def test_map_table_rows_from_raw_table_filters_footer_rows() -> None:
    """Footer rows at the end of a table are not mapped as SWMS steps."""
    table = [
        SAMPLE_HEADERS,
        *SAMPLE_ROWS,
        [
            "",
            "This SWMS must be read and understood by all workers before commencing work on site.",
            "",
            "",
            "",
            "",
            "",
        ],
        ["", "", "", "", "", "", "Page 2 of 5"],
    ]

    mapped = map_table_rows_from_raw_table(table)
    steps = group_swms_rows(mapped)

    assert len(mapped) == 3
    assert len(steps) == 2
    assert steps[-1]["stepNo"] == "2"


def test_filter_swms_data_rows_removes_continuation_footer() -> None:
    """Continuation footers are removed from mapped rows."""
    rows = [
        {
            "step_no": "1",
            "job_task_element": "Setup",
            "hazard": "Pinch points",
            "risk_level": "High",
            "controls": "Use guards",
        },
        {
            "step_no": "",
            "job_task_element": "Continued on next page",
            "hazard": "",
            "risk_level": "",
            "controls": "",
        },
    ]

    filtered = filter_swms_data_rows(rows)
    assert len(filtered) == 1
    assert filtered[0]["hazard"] == "Pinch points"


def test_is_non_swms_data_row_filters_qld_ppe_footer() -> None:
    """QLD SWMS template PPE disclaimer rows are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "*PERSONAL PROTECTIVE EQUIPMENT (PPE): Workers must",
            "hazard": "",
            "risk_level": "",
            "controls": "",
        }
    )
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "",
            "hazard": "",
            "risk_level": "",
            "controls": "5. PPE (least effective)",
        }
    )
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "",
            "hazard": "",
            "risk_level": "",
            "controls": "2. Isolate / Substitute",
        }
    )
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "",
            "hazard": "",
            "risk_level": "",
            "controls": "(How to control the risk)",
        }
    )


def test_is_non_swms_data_row_filters_page_meta_footer() -> None:
    """Construction Safety Wise page metadata footers are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "",
            "job_task_element": "Page 9 of 15 SWMS030 ©Construction Safety Wise - All rights reserved",
            "hazard": "",
            "risk_level": "",
            "controls": "",
        }
    )


def test_is_non_swms_data_row_filters_split_ppe_fragment() -> None:
    """Footer text split across table cells is excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "*PERS",
            "job_task_element": "ONAL PROTECTIVE EQUIPMENT",
            "hazard": "(PPE): Workers must",
            "risk_level": "ing all m",
            "controls": "andatory PPE before enter",
        }
    )


def test_is_non_swms_data_row_filters_step_fragment_row() -> None:
    """Partial page-break rows with numeric step numbers are excluded."""
    assert is_non_swms_data_row(
        {
            "step_no": "12",
            "job_task_element": "wall frames to concrete slab or",
            "hazard": "Building materials and",
            "risk_level": "",
            "controls": "techniques and working",
        }
    )


def test_select_best_swms_table_prefers_complete_table() -> None:
    """The highest-quality table is selected when pdfplumber returns duplicates."""
    good_table = [
        SAMPLE_HEADERS,
        [
            "12",
            "Fixing tie-downs/securing Steel wall frames to concrete slab or masonry walls",
            "Falling objects Building materials and off-cuts falling on workers below",
            "High",
            "Train workers in correct manual handling techniques",
            "Low",
            "Supervisor",
        ],
    ]
    bad_table = [
        ["*PERS", "ONAL PROTECTIVE EQUIPMENT", "(PPE): Workers must", "be wear", "ing all m"],
        ["12", "Fixing tie-downs/securing Steel", "Falling objects", "", "Train workers in correct"],
    ]

    selected = select_best_swms_table([bad_table, good_table])
    assert selected == good_table


def test_is_non_swms_data_row_keeps_alphanumeric_step_numbers() -> None:
    """Alphanumeric step numbers such as 8a and 8b are valid SWMS rows."""
    row = {
        "step_no": "8a",
        "job_task_element": "Excavate trenches using excavator",
        "hazard": "Striking live electrical services",
        "risk_level": "1",
        "controls": "Dial before you dig",
    }
    assert not is_non_swms_data_row(row)


def test_merge_grouped_steps_sorts_alphanumeric_step_numbers() -> None:
    """Steps 8a and 8b are ordered between 8 and 9 when merging pages."""
    steps = [
        {"stepNo": "9", "jobTaskElement": "Nine", "hazards": [], "sequencePosition": 1},
        {"stepNo": "8b", "jobTaskElement": "Eight B", "hazards": [], "sequencePosition": 2},
        {"stepNo": "8a", "jobTaskElement": "Eight A", "hazards": [], "sequencePosition": 3},
        {"stepNo": "7", "jobTaskElement": "Seven", "hazards": [], "sequencePosition": 4},
    ]

    merged = merge_grouped_steps(steps)
    assert [step["stepNo"] for step in merged] == ["7", "8a", "8b", "9"]
