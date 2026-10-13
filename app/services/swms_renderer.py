import io
import os
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont

from app.exceptions import DocumentRenderError
from app.schemas.render import RenderLegislation, RenderSwmsRequest

DEFAULT_TEMPLATE_PATH = (
    Path(__file__).resolve().parents[2] / "templates" / "swms-hrcw-risk-columns.docx"
)

HEADER_TABLE_INDEX = 0
STEPS_TABLE_INDEX = 4
STEPS_DATA_START_ROW = 3

STEP_COLUMNS = {
    "step_no": 0,
    "job_task": 1,
    "hazard": 2,
    "risk_level": 3,
    "risk_control": 4,
    "post_risk": 6,
    "responsible": 7,
}

RISK_CONTROL_COLUMN = STEP_COLUMNS["risk_control"]
HIERARCHY_COLUMN = 5
# Template includes a trailing micro-column that header/data rows merge into Responsible Person.
RESPONSIBLE_TRAILING_COLUMN = 8

# Hanging indent so wrapped control text aligns under the first text character, not the bullet.
BULLET_LEFT_INDENT = Inches(0.2)
BULLET_HANGING_INDENT = Inches(-0.15)
BULLET_SPACE_AFTER = Pt(2)

COMBINED_RISK_CONTROL_HEADER = (
    "RISK CONTROL\n"
    "How will the risk be controlled effectively?\n\n"
    "Hierarchy of Risk Control (informative):\n"
    "1. Eliminate the hazard\n"
    "2. Substitute the hazard\n"
    "3. Isolate the hazard from people\n"
    "4. Engineering controls\n"
    "5. Administrative controls\n"
    "6. Personal protective equipment (PPE)"
)

HEADER_CELL_MAP: dict[tuple[int, int], str] = {
    (1, 2): "companyName",
    (5, 2): "siteLocation",
    (7, 2): "jobActivitiesText",
    (9, 2): "plantEquipmentText",
    (7, 17): "swmsIssueDate",
}


def resolve_template_path(template_path: str | None = None) -> Path:
    """Resolves the SWMS Word template path from argument or environment."""
    if template_path:
        path = Path(template_path)
        if not path.is_file():
            raise DocumentRenderError(f"SWMS template not found: {path}")
        return path

    env_path = Path(__file__).resolve().parents[2] / "templates" / "swms-hrcw-risk-columns.docx"
    if env_path.is_file():
        return env_path

    if DEFAULT_TEMPLATE_PATH.is_file():
        return DEFAULT_TEMPLATE_PATH

    raise DocumentRenderError("SWMS Word template is not configured")


def _set_vertical_merge(cell, *, restart: bool) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in tc_pr.findall(qn("w:vMerge")):
        tc_pr.remove(child)

    v_merge = OxmlElement("w:vMerge")
    if restart:
        v_merge.set(qn("w:val"), "restart")
    tc_pr.append(v_merge)


def _set_cell_text_top(cell, text: str) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell.text = text
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _clear_cell(cell) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell.text = ""


def _apply_bullet_paragraph_format(paragraph, text: str) -> None:
    """Formats a paragraph as an indented bullet so wrapped lines align under the text."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.left_indent = BULLET_LEFT_INDENT
    paragraph.paragraph_format.first_line_indent = BULLET_HANGING_INDENT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = BULLET_SPACE_AFTER
    paragraph.text = f"• {text}"


def _set_risk_control_cell(cell, controls: list[str]) -> None:
    """Writes risk controls as hanging-indent bullets only (no hierarchy annotations)."""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell.text = ""

    cleaned_controls = [control.strip() for control in controls if control.strip()]
    paragraphs = cell.paragraphs
    next_paragraph_index = 0

    def next_paragraph():
        nonlocal next_paragraph_index
        if next_paragraph_index == 0 and paragraphs:
            paragraph = paragraphs[0]
        else:
            paragraph = cell.add_paragraph()
        next_paragraph_index += 1
        return paragraph

    for control in cleaned_controls:
        _apply_bullet_paragraph_format(next_paragraph(), control)


def _merge_column_vertically(table, column_index: int, start_row: int, end_row: int) -> None:
    if start_row >= end_row:
        return

    top_cell = table.rows[start_row].cells[column_index]
    _set_vertical_merge(top_cell, restart=True)
    top_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    for row_index in range(start_row + 1, end_row + 1):
        continuation_cell = table.rows[row_index].cells[column_index]
        _set_vertical_merge(continuation_cell, restart=False)
        _clear_cell(continuation_cell)


def _merge_cells_horizontally(row, start_column: int, end_column: int) -> None:
    if start_column >= end_column:
        return

    start_cell = row.cells[start_column]
    for column_index in range(start_column + 1, end_column + 1):
        start_cell.merge(row.cells[column_index])


def _prepare_steps_table_header(table) -> None:
    """Merges risk-control and hierarchy header cells into one informative column."""
    header_row = table.rows[2]
    _merge_cells_horizontally(header_row, RISK_CONTROL_COLUMN, HIERARCHY_COLUMN)
    _set_cell_text_top(header_row.cells[RISK_CONTROL_COLUMN], COMBINED_RISK_CONTROL_HEADER)


def _format_comma_separated_list(items: list[str]) -> str:
    """Formats a list of strings as a comma-separated line for template header cells."""
    return ", ".join(item.strip() for item in items if item.strip())


def _build_header_context(request: RenderSwmsRequest) -> dict[str, str]:
    job_text = _format_comma_separated_list(request.jobActivities)
    plant_text = _format_comma_separated_list(request.plantEquipment)
    issue_date = request.swmsIssueDate or date.today().strftime("%d %B %Y")

    return {
        "companyName": (request.companyName or "").strip(),
        "siteLocation": request.siteLocation.strip(),
        "jobActivitiesText": job_text,
        "plantEquipmentText": plant_text,
        "swmsIssueDate": issue_date,
    }


def _fill_header_table(document: Document, request: RenderSwmsRequest) -> None:
    if len(document.tables) <= HEADER_TABLE_INDEX:
        raise DocumentRenderError("Header table not found in SWMS template")

    header_table = document.tables[HEADER_TABLE_INDEX]
    context = _build_header_context(request)

    for (row_index, column_index), field_name in HEADER_CELL_MAP.items():
        if row_index >= len(header_table.rows):
            continue
        row = header_table.rows[row_index]
        if column_index >= len(row.cells):
            continue
        _set_cell_text_top(row.cells[column_index], context[field_name])

    if document.paragraphs:
        title_text = request.title.strip()
        if title_text:
            document.paragraphs[0].text = f"SAFE WORK METHOD STATEMENT (SWMS)\n{title_text}"


def _generate_pyramid_image_bytes() -> bytes:
    width = 900
    height = 420
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    blue_color = (43, 114, 186)
    text_color = (20, 60, 110)

    font_large = None
    font_small = None
    font_paths = [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for p in font_paths:
        if os.path.isfile(p):
            try:
                font_large = ImageFont.truetype(p, 26)
                font_small = ImageFont.truetype(p, 21)
                break
            except Exception:
                pass
    if not font_large:
        font_large = ImageFont.load_default()
        font_small = font_large

    tiers = [
        "ACT",
        "REGULATIONS",
        "CODES OF PRACTICE",
        "GUIDANCE MATERIALS",
        "AUSTRALIAN STANDARDS",
    ]

    top_y = 15
    bottom_y = 405
    total_h = bottom_y - top_y
    apex_x = width // 2
    base_half_w = 340
    left_base_x = apex_x - base_half_w
    right_base_x = apex_x + base_half_w

    num_tiers = len(tiers)
    tier_height = total_h / num_tiers

    draw.polygon(
        [(apex_x, top_y), (right_base_x, bottom_y), (left_base_x, bottom_y)],
        outline=blue_color,
        width=3,
    )

    for i in range(1, num_tiers):
        y = top_y + i * tier_height
        fraction = (y - top_y) / total_h
        half_w = base_half_w * fraction
        x_left = apex_x - half_w
        x_right = apex_x + half_w
        draw.line([(x_left, y), (x_right, y)], fill=blue_color, width=3)

    for i, text in enumerate(tiers):
        y_center = top_y + (i + 0.5) * tier_height
        font = font_large if len(text) <= 12 else font_small
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        x_pos = apex_x - tw / 2
        y_pos = y_center - th / 2 - 2
        draw.text((x_pos, y_pos), text, fill=text_color, font=font)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _set_cell_background(cell, fill_hex: str) -> None:
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_margins(cell, *, top: int = 40, bottom: int = 40, left: int = 80, right: int = 80) -> None:
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right}" w:type="dxa"/>\n'
        f'</w:tcMar>'
    )
    cell._tc.get_or_add_tcPr().append(tcMar)


def _set_table_borders(table) -> None:
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def _format_run(
    run,
    *,
    font_name: str = "Arial",
    font_size_pt: float = 8.5,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor | None = None,
) -> None:
    """Applies font family, size, weight and OpenXML script bindings for cross-platform visual consistency."""
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    # Ensure complex script properties match so text with punctuation (like '(Qld)')
    # does not drop bold or fall back to an unexpected font family in Word.
    for child in list(rPr):
        if child.tag in (qn("w:bCs"), qn("w:iCs")):
            rPr.remove(child)

    if bold:
        bCs = parse_xml(f'<w:bCs {nsdecls("w")}/>')
        rPr.append(bCs)
    if italic:
        iCs = parse_xml(f'<w:iCs {nsdecls("w")}/>')
        rPr.append(iCs)


def _clean_legislation_items(items: list[str]) -> list[str]:
    """Cleans and deduplicates legislation names (e.g. merging 'Act 2002' and 'Act 2002 (Qld)')."""
    seen_keys: set[str] = set()
    cleaned: list[str] = []
    for item in items:
        text = item.strip()
        if not text:
            continue
        norm_key = re.sub(
            r"\s*\((?:Qld|NSW|Vic|WA|SA|Tas|ACT|NT)\)\s*$",
            "",
            text,
            flags=re.IGNORECASE,
        ).strip().lower()
        if norm_key not in seen_keys:
            seen_keys.add(norm_key)
            cleaned.append(text)
    return cleaned


NSW_DEFAULT_ACTS = [
    "Work Health and Safety Act 2011",
]

NSW_DEFAULT_REGULATIONS = [
    "Work Health and Safety (General) Regulations 2017",
    "Work Health and Safety Amendment (Crystalline silica Substances) Amendment Regulation 2024",
]

NSW_DEFAULT_STANDARDS = [
    "AS 3610-1995 Formwork for concrete",
    "AS-NZS 4576-1995 Guidelines for scaffolding",
    "AS-1269.3 2005 Occupational Hearing Protection",
    "AS1336:2014 Occupational Eye Protection",
]

NSW_DEFAULT_GUIDELINES = [
    "Scaffolding near overhead electric lines information sheet (SWA)",
    "Working with crystalline silica substances guide 2024 (SWA)",
    "Working in the vicinity of overground and underground electric lines (SWA)",
    "Operating cranes and mobile plant near overhead electric lines (SWA)",
]

NSW_DEFAULT_COPS_LEFT = [
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
]

NSW_DEFAULT_COPS_RIGHT = [
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
]


def _is_nsw_state(state_name: str) -> bool:
    norm = state_name.strip().lower()
    return norm in (
        "new south wales",
        "nsw",
        "new south wales (nsw)",
        "new south wales state",
    )


def _fill_nsw_legislation_section(
    document: Document,
    request: RenderSwmsRequest,
    legislation: RenderLegislation,
    insert_before_target,
) -> None:
    # 1. Document Title
    title_text = request.title.strip()
    if title_text:
        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(4)
        run_title = p_title.add_run(title_text)
        _format_run(run_title, font_name="Arial", font_size_pt=18, bold=True, color=RGBColor(0, 0, 0))
        insert_before_target(p_title._p)

    # 2. Heading in Title Case: "New South Wales Legislation"
    p_heading = document.add_paragraph()
    p_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_heading.paragraph_format.space_before = Pt(6)
    p_heading.paragraph_format.space_after = Pt(6)
    run_heading = p_heading.add_run("New South Wales Legislation")
    _format_run(run_heading, font_name="Arial", font_size_pt=16, bold=True, color=RGBColor(0, 0, 0))
    insert_before_target(p_heading._p)

    # 3. Intro Paragraph
    p_desc = document.add_paragraph()
    p_desc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_desc.paragraph_format.space_before = Pt(2)
    p_desc.paragraph_format.space_after = Pt(8)
    p_desc.paragraph_format.line_spacing = 1.15
    desc_text = (
        "New South Wales is part of Australian harmonised Work Health and Safety Laws. "
        "The following legislation shows the various levels of legislation starting with the WHS Act, "
        "WHS Regulations, NSW Codes of Practice, Safe Work Australia Codes of Practice and Guidance materials, "
        "and Australian Standards. A person conducting a business or undertaking (PCBU) in New South Wales "
        "must ensure the highest standards of health and safety are implemented to protect their workers, "
        "the environment and relevant property."
    )
    run_desc = p_desc.add_run(desc_text)
    _format_run(run_desc, font_name="Arial", font_size_pt=8.5, bold=False, color=RGBColor(30, 30, 30))
    insert_before_target(p_desc._p)

    # 4. Table 1 (4 Columns: NSW Acts | NSW Regulations | Australian Standards, National COPs | Guidelines, Technical Notes)
    col_widths_t1 = [Inches(1.85), Inches(1.85), Inches(1.85), Inches(1.85)]
    table1 = document.add_table(rows=0, cols=4)
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    table1.autofit = False
    _set_table_borders(table1)

    def set_t1_row_widths(row):
        for idx, w in enumerate(col_widths_t1):
            row.cells[idx].width = w

    # Header Row
    hdr_row1 = table1.add_row()
    set_t1_row_widths(hdr_row1)
    hdr_titles = [
        "NSW Acts\n(Please Tick)",
        "NSW Regulations",
        "Australian Standards,\nNational COPs",
        "Guidelines, Technical Notes",
    ]
    for idx, title in enumerate(hdr_titles):
        c = hdr_row1.cells[idx]
        _set_cell_background(c, "0B2545")
        _set_cell_margins(c, top=60, bottom=60, left=40, right=40)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(title)
        _format_run(r, font_name="Arial", font_size_pt=8.5, bold=True, color=RGBColor(255, 255, 255))

    # Resolve Data for Table 1
    raw_acts = _clean_legislation_items(legislation.acts)
    raw_regs = _clean_legislation_items(legislation.regulations)
    raw_stds = _clean_legislation_items(legislation.australianStandards)
    raw_guides = _clean_legislation_items(legislation.guidelines)

    acts = raw_acts if raw_acts else list(NSW_DEFAULT_ACTS)
    regs = raw_regs if raw_regs else list(NSW_DEFAULT_REGULATIONS)
    standards = raw_stds if raw_stds else list(NSW_DEFAULT_STANDARDS)
    guidelines = raw_guides if raw_guides else list(NSW_DEFAULT_GUIDELINES)

    num_rows_t1 = max(len(acts), len(regs), len(standards), len(guidelines), 4)
    for r_idx in range(num_rows_t1):
        row = table1.add_row()
        set_t1_row_widths(row)
        cell_vals = [
            acts[r_idx] if r_idx < len(acts) else "",
            regs[r_idx] if r_idx < len(regs) else "",
            standards[r_idx] if r_idx < len(standards) else "",
            guidelines[r_idx] if r_idx < len(guidelines) else "",
        ]
        for c_idx, val in enumerate(cell_vals):
            c = row.cells[c_idx]
            _set_cell_margins(c, top=35, bottom=35, left=45, right=45)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            if val:
                r = p.add_run(val)
                _format_run(r, font_name="Arial", font_size_pt=7.5, bold=False, color=RGBColor(0, 0, 0))

    insert_before_target(table1._tbl)

    # Small spacer between Table 1 and Table 2
    p_spacer = document.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(3)
    p_spacer.paragraph_format.space_after = Pt(3)
    p_spacer.paragraph_format.line_spacing = 1.0
    insert_before_target(p_spacer._p)

    # 5. Table 2 (Codes of Practice: Checkbox | Left COPs | Right COPs)
    col_widths_t2 = [Inches(0.35), Inches(3.35), Inches(3.70)]
    table2 = document.add_table(rows=0, cols=3)
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    table2.autofit = False
    _set_table_borders(table2)

    def set_t2_row_widths(row):
        for idx, w in enumerate(col_widths_t2):
            row.cells[idx].width = w

    # Header Row
    hdr_row2 = table2.add_row()
    set_t2_row_widths(hdr_row2)
    hdr_cell2 = hdr_row2.cells[0]
    hdr_cell2.merge(hdr_row2.cells[1])
    hdr_cell2.merge(hdr_row2.cells[2])
    _set_cell_background(hdr_cell2, "0B2545")
    _set_cell_margins(hdr_cell2, top=60, bottom=60, left=80, right=80)
    p2 = hdr_cell2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run("Codes of Practice (Please Tick)")
    _format_run(r2, font_name="Arial", font_size_pt=9.5, bold=True, color=RGBColor(255, 255, 255))

    # Resolve Data for Table 2
    raw_cops = _clean_legislation_items(legislation.codesOfPractice)
    if not raw_cops or len(raw_cops) < 20:
        left_items = list(NSW_DEFAULT_COPS_LEFT)
        right_items = list(NSW_DEFAULT_COPS_RIGHT)
    else:
        if any("Mentally Healthy Workplaces" in item for item in raw_cops):
            mid = 14 if len(raw_cops) >= 28 else (len(raw_cops) + 1) // 2
            left_items = raw_cops[:mid]
            right_items = raw_cops[mid:]
        else:
            mid = (len(raw_cops) + 1) // 2
            left_items = raw_cops[:mid]
            right_items = raw_cops[mid:]

    num_rows_t2 = max(len(left_items), len(right_items), 14)
    for r_idx in range(num_rows_t2):
        row = table2.add_row()
        set_t2_row_widths(row)

        chk_cell = row.cells[0]
        left_cell = row.cells[1]
        right_cell = row.cells[2]

        _set_cell_margins(chk_cell, top=25, bottom=25, left=20, right=20)
        _set_cell_margins(left_cell, top=25, bottom=25, left=40, right=40)
        _set_cell_margins(right_cell, top=25, bottom=25, left=40, right=40)

        chk_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        left_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        right_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        p_chk = chk_cell.paragraphs[0]

        p_left = left_cell.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_left.paragraph_format.space_before = Pt(0)
        p_left.paragraph_format.space_after = Pt(0)
        p_left.paragraph_format.line_spacing = 1.0

        if r_idx < len(left_items) and left_items[r_idx]:
            r_left = p_left.add_run(left_items[r_idx])
            _format_run(r_left, font_name="Arial", font_size_pt=7.5, bold=False, color=RGBColor(0, 0, 0))

        p_right = right_cell.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_right.paragraph_format.space_before = Pt(0)
        p_right.paragraph_format.space_after = Pt(0)
        p_right.paragraph_format.line_spacing = 1.0

        if r_idx < len(right_items) and right_items[r_idx]:
            right_val = right_items[r_idx]
            is_bold = "Mentally Healthy Workplaces" in right_val
            r_right = p_right.add_run(right_val)
            _format_run(
                r_right,
                font_name="Arial",
                font_size_pt=8.0 if is_bold else 7.5,
                bold=is_bold,
                color=RGBColor(0, 0, 0),
            )

    insert_before_target(table2._tbl)

    # 6. Page Break
    p_pb = document.add_paragraph()
    p_pb.paragraph_format.space_before = Pt(0)
    p_pb.paragraph_format.space_after = Pt(0)
    r_pb = p_pb.add_run()
    r_pb.add_break(WD_BREAK.PAGE)
    insert_before_target(p_pb._p)


def _fill_standard_legislation_section(
    document: Document,
    request: RenderSwmsRequest,
    legislation: RenderLegislation,
    state_name: str,
    target_table_el,
    insert_before_target,
) -> None:
    # 1. Document Title
    title_text = request.title.strip()
    if title_text:
        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(4)
        run_title = p_title.add_run(title_text)
        _format_run(run_title, font_name="Arial", font_size_pt=16, bold=True, color=RGBColor(0, 0, 0))
        insert_before_target(p_title._p)

    # 2. State Legislation Heading
    p_heading = document.add_paragraph()
    p_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_heading.paragraph_format.space_before = Pt(4)
    p_heading.paragraph_format.space_after = Pt(6)
    run_heading = p_heading.add_run(f"{state_name.upper()} LEGISLATION")
    _format_run(run_heading, font_name="Arial", font_size_pt=14, bold=True, color=RGBColor(0, 0, 0))
    insert_before_target(p_heading._p)

    # 3. Hierarchy Description Paragraph
    p_desc = document.add_paragraph()
    p_desc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_desc.paragraph_format.space_before = Pt(2)
    p_desc.paragraph_format.space_after = Pt(6)
    p_desc.paragraph_format.line_spacing = 1.15
    desc_text = (
        f"{state_name} Work Health and Safety legislation aligns with the Australian harmonised "
        f"Work Health and Safety Laws. The following Hierarchy of legislation shows the various "
        f"levels of legislation starting with the WHS Act, WHS Regulations, Codes of Practice, "
        f"WHS guidance materials, and the Australian Standards. A person conducting business or "
        f"undertaking (PCBU) must ensure that the relevant parts of the Act, Regulations, Codes of "
        f"Practice, Guidance materials, and Australian Standards are implemented to secure the "
        f"health and safety of workers and to create a safe work environment."
    )
    run_desc = p_desc.add_run(desc_text)
    _format_run(run_desc, font_name="Arial", font_size_pt=8.5, bold=False, color=RGBColor(30, 30, 30))
    insert_before_target(p_desc._p)

    # 4. Hierarchy Pyramid Image
    try:
        pyramid_bytes = _generate_pyramid_image_bytes()
        p_img = document.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(8)
        run_img = p_img.add_run()
        run_img.add_picture(io.BytesIO(pyramid_bytes), width=Inches(3.3))
        insert_before_target(p_img._p)
    except Exception:
        pass

    # 5. Dynamic Legislation Table
    col_widths = [
        Inches(0.35), Inches(2.33), Inches(0.35), Inches(2.33),
        Inches(0.35), Inches(2.33), Inches(0.35), Inches(2.33),
    ]
    table = document.add_table(rows=0, cols=8)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table)

    def set_row_widths(row):
        for idx, w in enumerate(col_widths):
            row.cells[idx].width = w

    # Header Row: Dark Blue
    hdr_row = table.add_row()
    set_row_widths(hdr_row)
    hdr_cell = hdr_row.cells[0]
    for c in hdr_row.cells[1:]:
        hdr_cell.merge(c)
    _set_cell_background(hdr_cell, "0B2545")
    _set_cell_margins(hdr_cell, top=80, bottom=80, left=100, right=100)
    p_hdr = hdr_cell.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_before = Pt(2)
    p_hdr.paragraph_format.space_after = Pt(2)
    r_hdr = p_hdr.add_run(f"{state_name} Work Health and Safety Legislation")
    _format_run(r_hdr, font_name="Arial", font_size_pt=10.5, bold=True, color=RGBColor(255, 255, 255))

    has_any_records = False

    # Acts Rows (Uniform Arial 8.5pt Bold)
    acts = _clean_legislation_items(legislation.acts)
    if acts:
        has_any_records = True
        for act in acts:
            act_row = table.add_row()
            set_row_widths(act_row)
            cell = act_row.cells[0]
            for c in act_row.cells[1:]:
                cell.merge(c)
            _set_cell_margins(cell, top=45, bottom=45, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(act)
            _format_run(r, font_name="Arial", font_size_pt=8.5, bold=True, color=RGBColor(0, 0, 0))

    # Regulations Rows (Uniform Arial 8.5pt Bold)
    regs = _clean_legislation_items(legislation.regulations)
    if regs:
        has_any_records = True
        for reg in regs:
            reg_row = table.add_row()
            set_row_widths(reg_row)
            cell = reg_row.cells[0]
            for c in reg_row.cells[1:]:
                cell.merge(c)
            _set_cell_margins(cell, top=45, bottom=45, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(reg)
            _format_run(r, font_name="Arial", font_size_pt=8.5, bold=True, color=RGBColor(0, 0, 0))

    # Codes of Practice Section
    cops = _clean_legislation_items(legislation.codesOfPractice)
    if cops:
        has_any_records = True
        cop_hdr_row = table.add_row()
        set_row_widths(cop_hdr_row)
        cop_cell = cop_hdr_row.cells[0]
        for c in cop_hdr_row.cells[1:]:
            cop_cell.merge(c)
        _set_cell_background(cop_cell, "EAECEE")
        _set_cell_margins(cop_cell, top=60, bottom=60, left=100, right=100)
        p = cop_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        cop_title = (
            f"Worksafe {state_name} Codes of Practice: "
            "Select \u2611 the Code of Practice for the work tasks you are performing at the construction site."
        )
        r = p.add_run(cop_title)
        _format_run(r, font_name="Arial", font_size_pt=8.5, bold=True, color=RGBColor(0, 0, 0))

        for chunk_start in range(0, len(cops), 4):
            chunk = cops[chunk_start:chunk_start + 4]
            row = table.add_row()
            set_row_widths(row)
            for i in range(4):
                chk_cell = row.cells[i * 2]
                txt_cell = row.cells[i * 2 + 1]
                _set_cell_margins(chk_cell, top=25, bottom=25, left=30, right=30)
                _set_cell_margins(txt_cell, top=25, bottom=25, left=50, right=50)
                chk_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                txt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                p_chk = chk_cell.paragraphs[0]
                p_chk.text = ""
                p_txt = txt_cell.paragraphs[0]
                p_txt.text = ""
                p_txt.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_txt.paragraph_format.space_before = Pt(0)
                p_txt.paragraph_format.space_after = Pt(0)
                p_txt.paragraph_format.line_spacing = 1.0

                if i < len(chunk):
                    r_txt = p_txt.add_run(chunk[i])
                    _format_run(r_txt, font_name="Arial", font_size_pt=7.5, bold=False, color=RGBColor(0, 0, 0))

    # Australian Standards Section
    standards = _clean_legislation_items(legislation.australianStandards)
    if standards:
        has_any_records = True
        std_hdr_row = table.add_row()
        set_row_widths(std_hdr_row)
        std_cell = std_hdr_row.cells[0]
        for c in std_hdr_row.cells[1:]:
            std_cell.merge(c)
        _set_cell_background(std_cell, "EAECEE")
        _set_cell_margins(std_cell, top=60, bottom=60, left=100, right=100)
        p = std_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(
            "Australian Standards: Select \u2611 the Australian Standards for the work tasks you are performing at the construction site."
        )
        _format_run(r, font_name="Arial", font_size_pt=8.5, bold=True, color=RGBColor(0, 0, 0))

        for chunk_start in range(0, len(standards), 4):
            chunk = standards[chunk_start:chunk_start + 4]
            row = table.add_row()
            set_row_widths(row)
            for i in range(4):
                chk_cell = row.cells[i * 2]
                txt_cell = row.cells[i * 2 + 1]
                _set_cell_margins(chk_cell, top=25, bottom=25, left=30, right=30)
                _set_cell_margins(txt_cell, top=25, bottom=25, left=50, right=50)
                chk_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                txt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                p_chk = chk_cell.paragraphs[0]
                p_chk.text = ""
                p_txt = txt_cell.paragraphs[0]
                p_txt.text = ""
                p_txt.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_txt.paragraph_format.space_before = Pt(0)
                p_txt.paragraph_format.space_after = Pt(0)
                p_txt.paragraph_format.line_spacing = 1.0

                if i < len(chunk):
                    r_txt = p_txt.add_run(chunk[i])
                    _format_run(r_txt, font_name="Arial", font_size_pt=7.5, bold=False, color=RGBColor(0, 0, 0))

    # Graceful handling when no legislation records exist for the selected state
    if not has_any_records:
        empty_row = table.add_row()
        set_row_widths(empty_row)
        empty_cell = empty_row.cells[0]
        for c in empty_row.cells[1:]:
            empty_cell.merge(c)
        _set_cell_margins(empty_cell, top=50, bottom=50, left=100, right=100)
        p = empty_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(
            f"No specific legislation records currently catalogued for {state_name}. "
            "Applicable Commonwealth and model Work Health and Safety legislation, regulations, "
            "and national codes of practice apply."
        )
        _format_run(r, font_name="Arial", font_size_pt=8.5, bold=False, italic=True, color=RGBColor(80, 80, 80))

    insert_before_target(table._tbl)

    # 6. Page break after table
    p_pb = document.add_paragraph()
    p_pb.paragraph_format.space_before = Pt(0)
    p_pb.paragraph_format.space_after = Pt(0)
    r_pb = p_pb.add_run()
    r_pb.add_break(WD_BREAK.PAGE)
    insert_before_target(p_pb._p)


def _fill_legislation_section(document: Document, request: RenderSwmsRequest) -> None:
    legislation = request.legislation
    if legislation is None and not request.state:
        return

    if isinstance(legislation, dict):
        legislation = RenderLegislation.model_validate(legislation)
    elif legislation is None:
        legislation = RenderLegislation(state=request.state)

    state_name = (legislation.state or request.state or "Australian").strip()
    if not state_name:
        state_name = "Australian"

    if len(document.tables) <= 1:
        return

    # In the template, Table 0 is General Information and Table 1 is Qualifications.
    # Inserting before Table 1's XML element places this section right after Table 0 (Page 2).
    target_table_el = document.tables[1]._tbl

    def insert_before_target(elem):
        target_table_el.addprevious(elem)

    if _is_nsw_state(state_name):
        _fill_nsw_legislation_section(document, request, legislation, insert_before_target)
    else:
        _fill_standard_legislation_section(
            document, request, legislation, state_name, target_table_el, insert_before_target
        )


def _remove_steps_data_rows(table) -> None:
    while len(table.rows) > STEPS_DATA_START_ROW:
        table._tbl.remove(table.rows[-1]._tr)


def _find_steps_table(document: Document):
    for table in document.tables:
        if len(table.rows) >= 3:
            for row in table.rows[:3]:
                for cell in row.cells:
                    if "STEP No." in cell.text or "TASKS PERFORMED" in cell.text:
                        return table
    if len(document.tables) > STEPS_TABLE_INDEX:
        return document.tables[STEPS_TABLE_INDEX]
    raise DocumentRenderError("SWMS steps table not found in template")


def _fill_steps_table(document: Document, request: RenderSwmsRequest) -> None:
    table = _find_steps_table(document)
    _prepare_steps_table_header(table)
    _remove_steps_data_rows(table)

    for step in request.steps:
        hazard_count = len(step.hazards)
        start_row_index = len(table.rows)

        for hazard_index, hazard in enumerate(step.hazards):
            row = table.add_row()
            cells = row.cells

            if hazard_index == 0:
                _set_cell_text_top(cells[STEP_COLUMNS["step_no"]], step.stepNo)
                _set_cell_text_top(cells[STEP_COLUMNS["job_task"]], step.jobTaskElement)
            else:
                _clear_cell(cells[STEP_COLUMNS["step_no"]])
                _clear_cell(cells[STEP_COLUMNS["job_task"]])

            _set_cell_text_top(cells[STEP_COLUMNS["hazard"]], hazard.hazard)
            _set_cell_text_top(cells[STEP_COLUMNS["risk_level"]], hazard.riskLevel)
            _merge_cells_horizontally(row, RISK_CONTROL_COLUMN, HIERARCHY_COLUMN)
            _set_risk_control_cell(cells[STEP_COLUMNS["risk_control"]], hazard.controls)
            _set_cell_text_top(cells[STEP_COLUMNS["post_risk"]], hazard.postRiskLevel)
            # Merge the template's trailing micro-column into Responsible Person so content
            # rows match the header and do not leave a blank overlapping column.
            if len(cells) > RESPONSIBLE_TRAILING_COLUMN:
                _merge_cells_horizontally(
                    row,
                    STEP_COLUMNS["responsible"],
                    RESPONSIBLE_TRAILING_COLUMN,
                )
            _set_cell_text_top(cells[STEP_COLUMNS["responsible"]], hazard.responsiblePerson)

        end_row_index = start_row_index + hazard_count - 1
        if hazard_count > 1:
            _merge_column_vertically(table, STEP_COLUMNS["step_no"], start_row_index, end_row_index)
            _merge_column_vertically(table, STEP_COLUMNS["job_task"], start_row_index, end_row_index)


def render_swms_document_bytes(
    request: RenderSwmsRequest,
    template_path: str | None = None,
) -> bytes:
    """
    Renders a SWMS Word document from the template and structured step data.

    Args:
        request: Header metadata and generated steps/hazards.
        template_path: Optional override for the template file path.

    Returns:
        DOCX file bytes.
    """
    path = resolve_template_path(template_path)
    document = Document(str(path))

    _fill_header_table(document, request)
    _fill_legislation_section(document, request)
    _fill_steps_table(document, request)

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def render_swms_document_bytes_from_dict(
    payload: dict,
    template_path: str | None = None,
) -> bytes:
    """Validates payload and renders a SWMS Word document."""
    request = RenderSwmsRequest.model_validate(payload)
    return render_swms_document_bytes(request, template_path=template_path)
