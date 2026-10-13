"""Tests for SWMS Word rendering."""

from __future__ import annotations

import io

from docx import Document
from docx.oxml.ns import qn

from app.schemas.render import RenderSwmsRequest
from app.services.swms_renderer import (
    BULLET_HANGING_INDENT,
    BULLET_LEFT_INDENT,
    render_swms_document_bytes,
    resolve_template_path,
)


def _sample_request() -> RenderSwmsRequest:
    return RenderSwmsRequest(
        title="Switchboard replacement",
        siteLocation="123 George St, Brisbane",
        jobActivities=["Isolate supply", "Install switchboard"],
        plantEquipment=["Multimeter", "Insulated tools"],
        tradeName="Electrician",
        activityType="Not Applicable",
        steps=[
            {
                "stepNo": "1",
                "jobTaskElement": "Isolate and lock out electrical supply",
                "sequencePosition": 1,
                "hazards": [
                    {
                        "hazard": "Electric shock",
                        "riskLevel": "1",
                        "controls": ["Lockout/tagout", "Test for dead"],
                        "postRiskLevel": "3",
                        "responsiblePerson": "Licensed electrician",
                    },
                    {
                        "hazard": "Arc flash",
                        "riskLevel": "1",
                        "controls": ["Wear arc-rated PPE"],
                        "postRiskLevel": "2",
                        "responsiblePerson": "Licensed electrician",
                    },
                ],
            },
            {
                "stepNo": "2",
                "jobTaskElement": "Install new switchboard",
                "sequencePosition": 2,
                "hazards": [
                    {
                        "hazard": "Manual handling injury",
                        "riskLevel": "2",
                        "controls": ["Use mechanical aids", "Team lift"],
                        "postRiskLevel": "3",
                        "responsiblePerson": "Leading hand",
                    }
                ],
            },
        ],
    )


def test_resolve_template_path_finds_bundled_template() -> None:
    path = resolve_template_path()
    assert path.name == "swms-hrcw-risk-columns.docx"


def test_render_swms_document_bytes_produces_steps_with_vertical_merge() -> None:
    rendered = render_swms_document_bytes(_sample_request())
    assert rendered.startswith(b"PK")

    document = Document(io.BytesIO(rendered))
    steps_table = document.tables[4]

    assert len(steps_table.rows) == 6  # 3 header rows + 3 hazard rows

    first_step_top = steps_table.rows[3].cells[0].text
    first_step_continuation_tc = steps_table.rows[4]._tr.findall(qn("w:tc"))[0]
    first_step_continuation_vmerge = first_step_continuation_tc.xpath("./w:tcPr/w:vMerge")
    assert first_step_top == "1"
    assert len(first_step_continuation_vmerge) == 1

    assert "Electric shock" in steps_table.rows[3].cells[2].text
    assert "Arc flash" in steps_table.rows[4].cells[2].text
    assert steps_table.rows[3].cells[1].text == "Isolate and lock out electrical supply"
    second_task_tc = steps_table.rows[4]._tr.findall(qn("w:tc"))[1]
    assert second_task_tc.xpath("./w:tcPr/w:vMerge")

    assert steps_table.rows[5].cells[0].text == "2"
    assert "Manual handling injury" in steps_table.rows[5].cells[2].text

    risk_control_tc = steps_table.rows[3]._tr.findall(qn("w:tc"))[4]
    grid_span = risk_control_tc.xpath("./w:tcPr/w:gridSpan")
    assert grid_span and grid_span[0].get(qn("w:val")) == "2"
    assert "Lockout/tagout" in steps_table.rows[3].cells[4].text
    assert "Hierarchy (informative):" not in steps_table.rows[3].cells[4].text


def test_risk_control_bullets_use_hanging_indent() -> None:
    rendered = render_swms_document_bytes(_sample_request())
    document = Document(io.BytesIO(rendered))
    risk_control_cell = document.tables[4].rows[3].cells[4]

    bullet_paragraphs = [
        paragraph
        for paragraph in risk_control_cell.paragraphs
        if paragraph.text.startswith("• ")
    ]
    assert len(bullet_paragraphs) == 2
    assert bullet_paragraphs[0].text == "• Lockout/tagout"
    assert bullet_paragraphs[1].text == "• Test for dead"

    for paragraph in bullet_paragraphs:
        assert paragraph.paragraph_format.left_indent == BULLET_LEFT_INDENT
        assert paragraph.paragraph_format.first_line_indent == BULLET_HANGING_INDENT

    assert all(
        not paragraph.text.startswith("Hierarchy (informative):")
        for paragraph in risk_control_cell.paragraphs
    )


def test_responsible_person_merges_trailing_micro_column() -> None:
    rendered = render_swms_document_bytes(_sample_request())
    document = Document(io.BytesIO(rendered))
    first_data_row = document.tables[4].rows[3]
    raw_cells = first_data_row._tr.findall(qn("w:tc"))

    # After merging risk-control+hierarchy and responsible+trailing micro-column,
    # the row should have 7 physical cells (not a separate blank trailing column).
    assert len(raw_cells) == 7

    responsible_tc = raw_cells[6]
    grid_span = responsible_tc.xpath("./w:tcPr/w:gridSpan")
    assert grid_span and grid_span[0].get(qn("w:val")) == "2"
    assert first_data_row.cells[7].text == "Licensed electrician"
    # Trailing micro-column is absorbed into Responsible Person (same text via grid span).
    assert first_data_row.cells[8].text == "Licensed electrician"
    assert first_data_row.cells[8]._tc is first_data_row.cells[7]._tc


def test_render_swms_document_bytes_fills_header_fields() -> None:
    rendered = render_swms_document_bytes(_sample_request())
    document = Document(io.BytesIO(rendered))
    header_table = document.tables[0]

    assert "123 George St, Brisbane" in header_table.rows[5].cells[2].text
    assert header_table.rows[7].cells[2].text == "Isolate supply, Install switchboard"
    assert header_table.rows[9].cells[2].text == "Multimeter, Insulated tools"
    assert "Switchboard replacement" in document.paragraphs[0].text


def test_render_swms_document_bytes_binds_legislation_table() -> None:
    request = _sample_request()
    request.state = "Queensland"
    request.legislation = {
        "state": "Queensland",
        "acts": ["Work Health and Safety Act 2011", "Electrical Safety Act 2002"],
        "regulations": ["Work Health and Safety Regulation 2011", "Electrical Safety Regulation 2013"],
        "codesOfPractice": [
            "Concrete pumping 2019",
            "Hazardous manual tasks 2021",
            "Managing risks of hazardous chemicals in the workplace 2021",
            "Steel Construction 2004",
            "Scaffolding 2021",
        ],
        "australianStandards": [
            "AS/NZS 3000:2018 Wiring Rules",
            "AS/NZS 3012:2019 Construction and demolition sites",
        ],
    }

    rendered = render_swms_document_bytes(request)
    assert rendered.startswith(b"PK")

    document = Document(io.BytesIO(rendered))

    # Table 0: Header table
    # Table 1: Legislation table (inserted right after Table 0 on page 2)
    assert len(document.tables) == 7
    leg_table = document.tables[1]

    # Verify Header
    assert leg_table.rows[0].cells[0].text == "Queensland Work Health and Safety Legislation"

    # Verify Acts and Regulations are present
    table_text = "\n".join(" | ".join(c.text for c in row.cells) for row in leg_table.rows)
    assert "Work Health and Safety Act 2011" in table_text
    assert "Electrical Safety Act 2002" in table_text
    assert "Work Health and Safety Regulation 2011" in table_text
    assert "Electrical Safety Regulation 2013" in table_text

    # Verify Codes of Practice subheader and items
    assert "Worksafe Queensland Codes of Practice" in table_text
    assert "Concrete pumping 2019" in table_text
    assert "Scaffolding 2021" in table_text

    # Verify Australian Standards subheader and items
    assert "Australian Standards" in table_text
    assert "AS/NZS 3000:2018 Wiring Rules" in table_text

    # Verify steps table is still correctly located and filled at its new index (Table 5)
    steps_table = document.tables[5]
    assert len(steps_table.rows) == 6
    assert steps_table.rows[3].cells[0].text == "1"
    assert "Electric shock" in steps_table.rows[3].cells[2].text


def test_render_swms_document_handles_empty_legislation_records_gracefully() -> None:
    request = _sample_request()
    request.state = "Tasmania"
    request.legislation = {
        "state": "Tasmania",
        "acts": [],
        "regulations": [],
        "codesOfPractice": [],
        "australianStandards": [],
    }

    rendered = render_swms_document_bytes(request)
    assert rendered.startswith(b"PK")

    document = Document(io.BytesIO(rendered))
    assert len(document.tables) == 7
    leg_table = document.tables[1]

    table_text = "\n".join(" | ".join(c.text for c in row.cells) for row in leg_table.rows)
    assert "Tasmania Work Health and Safety Legislation" in table_text
    assert "No specific legislation records currently catalogued for Tasmania" in table_text

    # Verify steps table is still populated
    steps_table = document.tables[5]
    assert len(steps_table.rows) == 6


def test_render_swms_legislation_font_consistency_and_deduplication() -> None:
    from docx.oxml.ns import qn

    request = _sample_request()
    request.state = "Queensland"
    request.legislation = {
        "state": "Queensland",
        "acts": [
            "Electrical Safety Act 2002",
            "Electrical Safety Act 2002 (Qld)",
            "Work Health and Safety Act 2011",
            "Work Health and Safety Act 2011 (Qld)",
        ],
        "regulations": [
            "Electrical Safety Regulation 2013",
            "Electrical Safety Regulation 2013 (Qld)",
            "Work Health and Safety Regulation 2011",
            "Work Health and Safety Regulation 2011 (Qld)",
        ],
        "codesOfPractice": ["Confined spaces Code of Practice 2021"],
        "australianStandards": ["AS/NZS 3000:2018 Wiring Rules"],
    }

    rendered = render_swms_document_bytes(request)
    document = Document(io.BytesIO(rendered))
    leg_table = document.tables[1]

    # Deduplication ensures that 2 Acts and 2 Regulations appear (not 4 each with state suffixes)
    act_reg_texts = [row.cells[0].paragraphs[0].text for row in leg_table.rows[1:5]]
    assert len(act_reg_texts) == 4
    assert act_reg_texts[0] == "Electrical Safety Act 2002"
    assert act_reg_texts[1] == "Work Health and Safety Act 2011"
    assert act_reg_texts[2] == "Electrical Safety Regulation 2013"
    assert act_reg_texts[3] == "Work Health and Safety Regulation 2011"

    # Verify uniform font and OpenXML bindings (Arial, bold, w:bCs, w:cs, w:eastAsia)
    for row in leg_table.rows[1:5]:
        p = row.cells[0].paragraphs[0]
        assert len(p.runs) == 1
        run = p.runs[0]
        assert run.font.name == "Arial"
        assert run.font.bold is True

        rPr = run._r.rPr
        assert rPr is not None
        rFonts = rPr.rFonts
        assert rFonts.get(qn("w:ascii")) == "Arial"
        assert rFonts.get(qn("w:hAnsi")) == "Arial"
        assert rFonts.get(qn("w:cs")) == "Arial"
        assert rFonts.get(qn("w:eastAsia")) == "Arial"
        assert rPr.find(qn("w:bCs")) is not None


def test_render_swms_nsw_legislation_format() -> None:
    request = _sample_request()
    request.title = "Steel shed construction"
    request.state = "New South Wales"
    request.legislation = {
        "state": "New South Wales",
        "acts": ["Work Health and Safety Act 2011"],
        "regulations": [
            "Work Health and Safety (General) Regulations 2017",
            "Work Health and Safety Amendment (Crystalline silica Substances) Amendment Regulation 2024",
        ],
        "codesOfPractice": [
            "Confined Spaces 2022",
            "Construction Work 2019",
            "Demolition Work 2022",
            "Excavation Work 2020",
            "First Aid in the Workplace 2020",
            "Formwork 2020",
            "Hazardous Manual Tasks 2019",
            "How to Manage and Control Asbestos in the Workplace 2022",
            "How to Manage Work Health and Safety Risks 2019",
            "How to Safely Remove Asbestos 2022",
            "Labelling of workplace hazardous chemicals 2022",
            "Managing electrical risks 2019",
            "Managing noise and preventing hearing loss at work 2022",
            "Managing risks of hazardous chemicals in the workplace 2022",
            "Managing the risk of falls in housing construction 2019",
            "Managing the risks of respirable crystalline silica from engineered stone in the workplace 2022",
            "Managing respirable crystalline silica dust exposure in construction and manufacturing of construction elements 2022 (QLD)",
            "Managing the work environment and facilities 2019",
            "Mobile crane COP (QLD) 2006",
            "Scaffolding industry Safety Standard 2022",
            "Preparation of Safety Data Sheets for Hazardous Chemicals 2022",
            "Safe Design of Structures 2019",
            "Traffic Management for work on Roads 2021 (Main Roads WA)",
            "Welding Processes 2022",
            "Work health and safety consultation, cooperation and coordination 2022",
            "Mentally Healthy Workplaces Codes of Practice",
            "Sexual and gender based harassment 2024",
            "Managing Psychosocial Hazards at work",
        ],
        "guidelines": [
            "Scaffolding near overhead electric lines information sheet (SWA)",
            "Working with crystalline silica substances guide 2024 (SWA)",
            "Working in the vicinity of overground and underground electric lines (SWA)",
            "Operating cranes and mobile plant near overhead electric lines (SWA)",
        ],
        "australianStandards": [
            "AS 3610-1995 Formwork for concrete",
            "AS-NZS 4576-1995 Guidelines for scaffolding",
            "AS-1269.3 2005 Occupational Hearing Protection",
            "AS1336:2014 Occupational Eye Protection",
        ],
    }

    rendered = render_swms_document_bytes(request)
    assert rendered.startswith(b"PK")

    document = Document(io.BytesIO(rendered))
    # In NSW format, 2 legislation tables are inserted (Table 1: 4 columns, Table 2: 3 columns)
    assert len(document.tables) == 8

    # Table 1: 4 columns (Acts, Regulations, Standards, Guidelines)
    t1 = document.tables[1]
    assert len(t1.columns) == 4
    assert "NSW Acts" in t1.rows[0].cells[0].text
    assert "NSW Regulations" in t1.rows[0].cells[1].text
    assert "Australian Standards" in t1.rows[0].cells[2].text
    assert "Guidelines, Technical Notes" in t1.rows[0].cells[3].text

    t1_text = "\n".join(" | ".join(c.text for c in row.cells) for row in t1.rows)
    assert "Work Health and Safety Act 2011" in t1_text
    assert "Work Health and Safety (General) Regulations 2017" in t1_text
    assert "AS 3610-1995 Formwork for concrete" in t1_text
    assert "Scaffolding near overhead electric lines information sheet (SWA)" in t1_text

    # Table 2: 3 columns (Checkbox | Left COPs | Right COPs)
    t2 = document.tables[2]
    assert len(t2.columns) == 3
    assert t2.rows[0].cells[0].text == "Codes of Practice (Please Tick)"

    # Check 14 data rows + 1 header row = 15 rows
    assert len(t2.rows) == 15
    t2_text = "\n".join(" | ".join(c.text for c in row.cells) for row in t2.rows)
    assert "Confined Spaces 2022" in t2_text
    assert "Construction Work 2019" in t2_text
    assert "Managing risks of hazardous chemicals in the workplace 2022" in t2_text
    assert "Managing the risk of falls in housing construction 2019" in t2_text
    assert "Mentally Healthy Workplaces Codes of Practice" in t2_text
    assert "Sexual and gender based harassment 2024" in t2_text
    assert "Managing Psychosocial Hazards at work" in t2_text

    # Verify "Mentally Healthy Workplaces Codes of Practice" run is bold
    found_bold_header = False
    for row in t2.rows[1:]:
        p = row.cells[2].paragraphs[0]
        if "Mentally Healthy Workplaces Codes of Practice" in p.text:
            assert len(p.runs) >= 1
            assert p.runs[0].font.bold is True
            found_bold_header = True
            break
    assert found_bold_header is True

    # Check paragraph headings
    para_texts = [p.text for p in document.paragraphs if p.text.strip()]
    assert "Steel shed construction" in para_texts
    assert "New South Wales Legislation" in para_texts
    assert any("part of Australian harmonised Work Health and Safety Laws" in t for t in para_texts)

    # Verify steps table is preserved and located
    steps_table = document.tables[6]
    assert len(steps_table.rows) == 6
    assert steps_table.rows[3].cells[0].text == "1"


def test_render_swms_nsw_legislation_defaults_when_empty() -> None:
    request = _sample_request()
    request.state = "NSW"
    request.legislation = {
        "state": "NSW",
        "acts": [],
        "regulations": [],
        "codesOfPractice": [],
        "guidelines": [],
        "australianStandards": [],
    }

    rendered = render_swms_document_bytes(request)
    assert rendered.startswith(b"PK")

    document = Document(io.BytesIO(rendered))
    assert len(document.tables) == 8

    t2 = document.tables[2]
    t2_text = "\n".join(" | ".join(c.text for c in row.cells) for row in t2.rows)
    assert "Confined Spaces 2022" in t2_text
    assert "Mentally Healthy Workplaces Codes of Practice" in t2_text
    assert "Managing Psychosocial Hazards at work" in t2_text

