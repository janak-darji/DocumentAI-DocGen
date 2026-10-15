"""DOCX hierarchical section and General Information field extraction using python-docx."""

from __future__ import annotations

import io
import re
from typing import BinaryIO

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from app.exceptions import CorruptFileError, DocumentParseError
from app.logger import logger
from app.schemas.document_sections import (
    ExtractedField,
    ExtractedSectionNode,
    ExtractedMainSection,
    ExtractedSubSection,
    ParsedDocumentSectionsResponse,
)

# Standard General Information field candidates commonly found in Construction SMPs
COMMON_SMP_FIELDS = [
    {
        "fieldKey": "principal_contractor",
        "fieldName": "Principal Contractor",
        "fieldType": "text",
        "placeholder": "e.g. ABC Constructions Pty Ltd",
        "helpText": "Name of the principal contractor responsible for site safety",
        "isRequired": True,
    },
    {
        "fieldKey": "project_name",
        "fieldName": "Project / Site Name",
        "fieldType": "text",
        "placeholder": "e.g. Commercial Office Refurbishment",
        "helpText": "Name of the construction project or facility",
        "isRequired": True,
    },
    {
        "fieldKey": "site_address",
        "fieldName": "Site Address",
        "fieldType": "text",
        "placeholder": "e.g. 123 Building Way, Sydney NSW 2000",
        "helpText": "Physical address where work is being conducted",
        "isRequired": True,
    },
    {
        "fieldKey": "site_supervisor",
        "fieldName": "Site Supervisor / Manager",
        "fieldType": "text",
        "placeholder": "e.g. John Doe",
        "helpText": "Name of appointed site supervisor",
        "isRequired": False,
    },
    {
        "fieldKey": "supervisor_phone",
        "fieldName": "Supervisor Phone",
        "fieldType": "phone",
        "placeholder": "e.g. 0412 345 678",
        "helpText": "Contact number for site supervisor",
        "isRequired": False,
    },
    {
        "fieldKey": "emergency_contact",
        "fieldName": "Emergency Contact Name",
        "fieldType": "text",
        "placeholder": "e.g. Jane Smith",
        "helpText": "Primary 24/7 site emergency contact person",
        "isRequired": True,
    },
    {
        "fieldKey": "emergency_phone",
        "fieldName": "Emergency Phone Number",
        "fieldType": "phone",
        "placeholder": "e.g. 0400 111 222",
        "helpText": "Emergency contact phone number",
        "isRequired": True,
    },
    {
        "fieldKey": "nearest_hospital",
        "fieldName": "Nearest Hospital",
        "fieldType": "text",
        "placeholder": "e.g. Royal Prince Alfred Hospital",
        "helpText": "Designated emergency medical facility",
        "isRequired": True,
    },
    {
        "fieldKey": "nearest_hospital_phone",
        "fieldName": "Hospital Phone",
        "fieldType": "phone",
        "placeholder": "e.g. (02) 9515 6111",
        "helpText": "Contact phone for the nearest hospital",
        "isRequired": False,
    },
    {
        "fieldKey": "fire_station",
        "fieldName": "Nearest Fire Station",
        "fieldType": "text",
        "placeholder": "e.g. Local Fire & Rescue NSW",
        "helpText": "Local fire service contact or address",
        "isRequired": False,
    },
    {
        "fieldKey": "police_station",
        "fieldName": "Nearest Police Station",
        "fieldType": "text",
        "placeholder": "e.g. Local Police Command",
        "helpText": "Local police station contact or address",
        "isRequired": False,
    },
    {
        "fieldKey": "commencement_date",
        "fieldName": "Work Commencement Date",
        "fieldType": "date",
        "placeholder": "YYYY-MM-DD",
        "helpText": "Scheduled start date of construction work",
        "isRequired": False,
    },
]

# Standard unnumbered or named SMP sections recognized with high confidence
KNOWN_SMP_SECTION_NAMES = {
    "introduction",
    "emergency numbers and general information",
    "emergency contacts",
    "whs policy and commitment",
    "work health and safety policy",
    "rehabilitation policy",
    "environmental policy",
    "work health and safety duties",
    "roles and responsibilities",
    "consultation cooperation coordination",
    "consultation, cooperation, coordination",
    "training and communication",
    "risk management",
    "high risk construction work",
    "hazard identification and risk assessment",
    "incident management and reporting",
    "site safety rules",
    "general site safety rules",
    "personal protective equipment",
    "plant and equipment",
    "hazardous chemicals",
    "first aid and emergency response",
    "monitoring and review",
    "document control and review",
    "definitions",
    "scope",
    "purpose",
    "appendices",
    "appendix",
}

UPPER_ROMAN_RE = re.compile(
    r"^(M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{1,3}))$"
)
LOWER_ROMAN_RE = re.compile(
    r"^(m{0,4}(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{1,3}))$"
)
ROMAN_VALS = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}

# Matches Table of Contents lines ending with dot leaders / spaces and a page number
TOC_LINE_RE = re.compile(r"^(.+?)(?:\s*[\.·\-_]{2,}|\t+|\s{2,})\s*(\d{1,4})$")


def _clean_text(text: str) -> str:
    """Strips leading/trailing whitespace and normalizes spaces."""
    return re.sub(r"[ \t]+", " ", text).strip()


def _slugify(text: str) -> str:
    """Creates a normalized snake_case key from text."""
    slug = re.sub(r"[^\w\s-]", "", text).strip().lower()
    slug = re.sub(r"[-\s]+", "_", slug)
    return slug[:100] or "section"


def _is_valid_roman(text: str) -> bool:
    s = text.strip()
    if not s:
        return False
    return bool(UPPER_ROMAN_RE.match(s) or LOWER_ROMAN_RE.match(s))


def _roman_to_int(s: str) -> int:
    s = s.lower().strip()
    total = 0
    prev_val = 0
    for char in reversed(s):
        val = ROMAN_VALS.get(char, 0)
        if val >= prev_val:
            total += val
        else:
            total -= val
        prev_val = val
    return total or 1


def parse_heading_prefix(text: str) -> tuple[str | None, str, int, str]:
    """
    Parses a heading candidate to extract its numbering prefix, clean title,
    numbering depth, and pattern type.

    Returns:
        (number, clean_title, depth, pattern_type)
    """
    raw = _clean_text(text)
    if not raw:
        return None, "", 1, "empty"

    # 1. Dotted numeric (e.g. "1", "1.1", "1.1.1", "6.1.1.1")
    # Matches: "1 WHS Policy", "1. Introduction", "1.1 General Policy", "6.1.1 Pre-Start Meeting"
    m_num = re.match(r"^(\d+(?:\.\d+)*)\s*[\.\-\:\)]?\s+(.+)$", raw)
    if m_num:
        num = m_num.group(1)
        title = _clean_text(m_num.group(2))
        depth = len(num.split("."))
        return num, title, depth, "numeric" if depth == 1 else "nested_numeric"

    # 2. Mixed numeric-alphabetic (e.g. "1.A", "1.1.A", "1.A.1", "A.1", "A.1.1")
    m_mixed = re.match(r"^(\d+(?:\.[A-Za-z0-9]+)+)\s*[\.\-\:\)]?\s+(.+)$", raw)
    if m_mixed:
        num = m_mixed.group(1)
        title = _clean_text(m_mixed.group(2))
        depth = len(num.split("."))
        return num, title, depth, "mixed"

    m_alpha_num = re.match(r"^([A-Za-z]\.\d+(?:\.[A-Za-z0-9]+)*)\s*[\.\-\:\)]?\s+(.+)$", raw)
    if m_alpha_num:
        num = m_alpha_num.group(1)
        title = _clean_text(m_alpha_num.group(2))
        depth = len(num.split("."))
        return num, title, depth, "mixed"

    # 3. Explicit prefix keyword (e.g. "Section 1: WHS", "Part A - Scope", "Clause 2.1")
    m_prefix = re.match(
        r"^(?:Section|Part|Chapter|Clause|Item)\s+([A-Za-z0-9\.]+)\s*[\.\-\:\)]?\s+(.+)$",
        raw,
        re.IGNORECASE,
    )
    if m_prefix:
        num = m_prefix.group(1)
        title = _clean_text(m_prefix.group(2))
        depth = len(num.split(".")) if "." in num else 1
        return num, title, depth, "prefixed"

    # 4. Roman numerals
    # Uppercase Roman (e.g. "I. Introduction", "II Responsibilities", "IV - Emergency")
    m_u_roman = re.match(
        r"^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX)\s*[\.\-\:\)]\s+(.+)$",
        raw,
    )
    if not m_u_roman:
        # Also allow space separation if followed by capitalized word and not English pronoun phrase
        m_u_roman_sp = re.match(
            r"^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX)\s+([A-Z][a-zA-Z\s]{2,})$",
            raw,
        )
        if m_u_roman_sp:
            first_word = m_u_roman_sp.group(2).split()[0].lower()
            if m_u_roman_sp.group(1) != "I" or first_word not in {
                "am",
                "have",
                "will",
                "would",
                "do",
                "can",
                "shall",
                "state",
                "declare",
                "certify",
                "agree",
            }:
                m_u_roman = m_u_roman_sp

    if m_u_roman:
        num = m_u_roman.group(1)
        title = _clean_text(m_u_roman.group(2))
        return num, title, 1, "roman"

    # Lowercase Roman with paren or dot (e.g. "i) General", "(ii) Safety", "iii. Requirements")
    m_l_roman = re.match(r"^(?:\(?([ivx]+)\)|\b([ivx]+)\))\s+(.+)$", raw)
    if not m_l_roman:
        m_l_roman = re.match(r"^([ivx]+)\.\s+(.+)$", raw)
    if m_l_roman and _is_valid_roman(m_l_roman.group(1) or m_l_roman.group(2)):
        num = (m_l_roman.group(1) or m_l_roman.group(2)).lower()
        title = _clean_text(
            m_l_roman.group(3) if len(m_l_roman.groups()) >= 3 and m_l_roman.group(3) else m_l_roman.group(2)
        )
        return num, title, 2, "roman_sub"

    # 5. Alphabetic sections
    # Uppercase single letter (e.g. "A Scope", "A. Scope", "B - Responsibilities", "A) Purpose")
    m_u_alpha = re.match(r"^([A-Z])\s*[\.\-\:\)]\s+(.+)$", raw)
    if not m_u_alpha:
        m_u_alpha_sp = re.match(r"^([A-Z])\s+([A-Z][a-zA-Z\s]{2,})$", raw)
        if m_u_alpha_sp:
            rest = m_u_alpha_sp.group(2).strip()
            # If letter is "A", check that it's a short title and not a common sentence
            if m_u_alpha_sp.group(1) != "A" or not rest.endswith("."):
                m_u_alpha = m_u_alpha_sp
    if m_u_alpha:
        num = m_u_alpha.group(1)
        title = _clean_text(m_u_alpha.group(2))
        return num, title, 1, "alphabetic"

    # Lowercase letter with paren or dot (e.g. "a) General Requirements", "(b) Safety Procedures")
    m_l_alpha = re.match(r"^(?:\(?([a-z])\)|\b([a-z])\))\s+(.+)$", raw)
    if not m_l_alpha:
        m_l_alpha = re.match(r"^([a-z])\.\s+(.+)$", raw)
    if m_l_alpha:
        num = (m_l_alpha.group(1) or m_l_alpha.group(2)).lower()
        title = _clean_text(m_l_alpha.group(len(m_l_alpha.groups())))
        return num, title, 2, "alphabetic_sub"

    # 6. Unnumbered heading (title only)
    return None, raw, 1, "unnumbered"


def _extract_placeholders(text: str) -> list[str]:
    """Finds {{placeholder}} and [Placeholder] tokens in text."""
    found: list[str] = []
    curly = re.findall(r"\{\{([a-zA-Z0-9_\-\s]+)\}\}", text)
    for c in curly:
        key = _slugify(c)
        if key and key not in found:
            found.append(key)

    square = re.findall(r"\[([A-Z][a-zA-Z0-9_\-\s]{2,40})\]", text)
    for s in square:
        key = _slugify(s)
        if key and key not in found:
            found.append(key)

    return found


def _parse_docx_elements(document: Document) -> list[dict]:
    """
    Traverses document body elements in order, extracting paragraph texts,
    styles, bold properties, font size, indentation, and Word numbering definitions.
    """
    elements: list[dict] = []

    for element in document.element.body:
        if element.tag == qn("w:p"):
            p = Paragraph(element, document)
            p_text = p.text.strip()
            if not p_text:
                continue

            # Check style name
            style_name = ""
            p_style = element.xpath("./w:pPr/w:pStyle/@w:val")
            if p_style:
                style_name = str(p_style[0]).lower()

            # Check bold status (if all or main runs are bold)
            bold_nodes = element.xpath(".//w:r/w:rPr/w:b")
            is_bold = len(bold_nodes) > 0

            # Check font size (half-points, e.g. 24 = 12pt, 28 = 14pt)
            sz_nodes = element.xpath(".//w:r/w:rPr/w:sz/@w:val")
            font_size = max([float(s) / 2 for s in sz_nodes], default=11.0)

            # Check Word numbering (w:numPr)
            ilvl_nodes = element.xpath("./w:pPr/w:numPr/w:ilvl/@w:val")
            num_id_nodes = element.xpath("./w:pPr/w:numPr/w:numId/@w:val")
            ilvl = int(ilvl_nodes[0]) if ilvl_nodes else None
            num_id = str(num_id_nodes[0]) if num_id_nodes else None

            elements.append({
                "type": "paragraph",
                "text": p_text,
                "style": style_name,
                "is_bold": is_bold,
                "font_size": font_size,
                "ilvl": ilvl,
                "num_id": num_id,
            })

        elif element.tag == qn("w:tbl"):
            tbl = Table(element, document)
            rows_data: list[list[str]] = []
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                if any(cells):
                    rows_data.append(cells)

            if rows_data:
                table_lines = [" | ".join(r) for r in rows_data]
                elements.append({
                    "type": "table",
                    "text": "\n".join(table_lines),
                    "rows": rows_data,
                })

    return elements


def extract_table_of_contents_entries(
    raw_elements: list[dict],
) -> tuple[list[dict], set[int]]:
    """
    Detects Table of Contents entries in raw elements.
    Returns:
        (toc_entries, set_of_element_indices_belonging_to_toc)
    """
    toc_entries: list[dict] = []
    toc_indices: set[int] = set()

    in_toc_section = False

    for idx, el in enumerate(raw_elements):
        if el["type"] != "paragraph":
            continue

        text = el["text"].strip()
        style = el.get("style", "")

        # Detect TOC heading start
        if text.lower() in {"contents", "table of contents", "table of content"}:
            in_toc_section = True
            toc_indices.add(idx)
            continue

        # Check for TOC style or dot leader line
        is_toc_style = style.startswith("toc") or "toc" in style
        m_toc = TOC_LINE_RE.match(text)
        if not m_toc and in_toc_section:
            m_toc = re.match(r"^(.+?)(?:\s*[\.·\-_]{2,}|\t+|\s+)(\d{1,4})$", text)

        if m_toc or (in_toc_section and is_toc_style):
            if m_toc:
                line_content = m_toc.group(1).strip()
                page_num = int(m_toc.group(2))
            else:
                line_content = text
                page_num = None

            num, title, depth, pat_type = parse_heading_prefix(line_content)

            # Determine TOC level from style or prefix depth
            level = depth
            if "toc 1" in style or style == "1":
                level = 1
            elif "toc 2" in style or style == "2":
                level = 2
            elif "toc 3" in style or style == "3":
                level = 3
            elif "toc 4" in style or style == "4":
                level = 4

            # If unnumbered entry (e.g. "Introduction", "Emergency Numbers")
            if num is None:
                level = 1

            toc_entries.append({
                "raw": text,
                "number": num,
                "title": title,
                "level": level,
                "page_number": page_num,
                "pattern_type": pat_type,
            })
            toc_indices.add(idx)
            in_toc_section = True

        elif in_toc_section:
            # End of TOC section when a non-TOC paragraph is encountered
            if not is_toc_style and not m_toc:
                in_toc_section = False

    return toc_entries, toc_indices


def _is_likely_heading(
    el: dict,
    toc_titles: set[str],
) -> tuple[bool, str | None, str, int, float, bool]:
    """
    Evaluates multi-signal heading detection on a paragraph element.

    Returns:
        (is_heading, number, title, level, confidence, is_ambiguous)
    """
    text = el["text"].strip()
    style = el.get("style", "")
    is_bold = el.get("is_bold", False)
    font_size = el.get("font_size", 11.0)
    ilvl = el.get("ilvl")

    num, title, depth, pat_type = parse_heading_prefix(text)
    norm_title = re.sub(r"[^\w\s]", "", title).strip().lower()
    norm_style = re.sub(r"[\s\-_]", "", style.lower())
    is_heading_style = norm_style.startswith("heading") or norm_style in {"1", "2", "3", "4", "title"}

    # 1. Negative filters: definitely not a heading
    if len(text) > 180 and not is_heading_style:
        return False, None, text, 1, 0.0, False
    if text.startswith(("- ", "• ", "* ")) and not is_heading_style:
        return False, None, text, 1, 0.0, False
    # Multi-sentence paragraphs ending in periods (unless explicitly a heading style)
    if not is_heading_style:
        if num is None and ". " in text and not text.endswith("."):
            return False, None, text, 1, 0.0, False
        if num is not None and ". " in title and not title.endswith("."):
            return False, None, text, 1, 0.0, False

    # 2. Signal A: Direct match with verified Table of Contents entry
    if norm_title in toc_titles or (num and norm_title in toc_titles):
        return True, num, title, depth, 1.0, False

    # 3. Signal B: Word Heading Style (support "Heading1", "Heading 1", "Heading_1", etc.)
    if "heading1" in norm_style or norm_style == "1":
        return True, num, title, 1, 0.98, False
    if "heading2" in norm_style or norm_style == "2":
        return True, num, title, 2, 0.98, False
    if "heading3" in norm_style or norm_style == "3":
        return True, num, title, 3, 0.98, False
    if "heading4" in norm_style or norm_style == "4":
        return True, num, title, 4, 0.98, False
    if "title" in norm_style:
        return True, num, title, 1, 0.95, False

    # 4. Signal C: Word list numbering level (w:numPr)
    if ilvl is not None:
        target_level = ilvl + 1
        return True, num, title, target_level, 0.92, False

    # 5. Signal D: Explicit numbering pattern
    if num is not None:
        # Check title formatting: title must start with uppercase letter/symbol and not be full sentence
        if title and (title[0].isupper() or title[0].isdigit() or title[0] in {"[", "("}):
            return True, num, title, depth, 0.90, False

    # 6. Signal E: Unnumbered standard heading (e.g. "Introduction", "Emergency Numbers and General Information")
    if norm_title in KNOWN_SMP_SECTION_NAMES:
        return True, None, title, 1, 0.88, False

    # 7. Signal F: Strong formatting signals (bold + larger font or uppercase + short line)
    if is_bold and len(text) < 70 and not text.endswith("."):
        if text.isupper() or title.istitle():
            return True, None, title, 1, 0.75, True

    return False, None, text, 1, 0.0, False


def build_hierarchy_tree(classified_items: list[dict]) -> list[ExtractedSectionNode]:
    """
    Constructs a strict hierarchical section tree using a stack-based algorithm.
    Preserves document order, duplicate numbering (e.g. multiple '5.1' headings),
    numbering restarts, and nests subsections cleanly under their parents.
    """
    root_nodes: list[ExtractedSectionNode] = []
    # Stack stores tuples of: (level, node, number)
    stack: list[tuple[int, ExtractedSectionNode, str | None]] = []
    active_node: ExtractedSectionNode | None = None
    order_counter = 1

    for item in classified_items:
        if item["is_heading"]:
            title = (item["title"] or "Untitled Section").strip()
            num = item["number"]
            level = item["level"]
            sec_key = _slugify(f"sec_{num or ''}_{title}")

            node = ExtractedSectionNode(
                number=num,
                sectionNumber=num,
                level=level,
                sectionKey=sec_key,
                title=title,
                content="",
                orderIndex=order_counter,
                pageNumber=item.get("page_number"),
                confidence=item.get("confidence", 1.0),
                isAmbiguous=item.get("is_ambiguous", False),
                children=[],
                subSections=[],
            )
            order_counter += 1

            # Stack resolution: pop while top of stack is at same or deeper level
            while stack and stack[-1][0] >= level:
                stack.pop()

            if stack:
                parent_level, parent_node, _ = stack[-1]
                parent_node.children.append(node)
                parent_node.subSections.append(node)
            else:
                root_nodes.append(node)

            stack.append((level, node, num))
            active_node = node

        else:
            # Body content (paragraph or table)
            content_text = item["text"].strip()
            if not content_text:
                continue

            if active_node is not None:
                if active_node.content:
                    active_node.content += "\n\n" + content_text
                else:
                    active_node.content = content_text

    return root_nodes


def parse_docx_document_sections(file_obj: BinaryIO) -> ParsedDocumentSectionsResponse:
    """
    Extracts document sections, subsections hierarchy, and General Information fields
    from an uploaded DOCX file.
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

    raw_elements = _parse_docx_elements(document)
    if not raw_elements:
        raise DocumentParseError("The uploaded DOCX document contains no readable text or content")

    # 1. Determine Document Title
    document_title = "Safety Management Plan"
    first_p = next((el["text"] for el in raw_elements if el["type"] == "paragraph"), None)
    if first_p and len(first_p) < 120 and not re.match(r"^\d+\.", first_p):
        document_title = first_p

    # 2. Extract Table of Contents entries (if present)
    toc_entries, toc_indices = extract_table_of_contents_entries(raw_elements)
    toc_titles = {
        re.sub(r"[^\w\s]", "", e["title"]).strip().lower()
        for e in toc_entries
        if e.get("title")
    }

    # 3. Classify elements as headings or body content
    classified_items: list[dict] = []
    body_elements = [el for idx, el in enumerate(raw_elements) if idx not in toc_indices]

    for el in body_elements:
        if el["type"] == "paragraph":
            is_head, num, title, level, conf, is_ambig = _is_likely_heading(el, toc_titles)

            # Retrieve page number from matching TOC entry if available
            page_num = None
            if is_head and toc_entries:
                norm_t = re.sub(r"[^\w\s]", "", title).strip().lower()
                matching_toc = next(
                    (t for t in toc_entries if re.sub(r"[^\w\s]", "", t["title"]).strip().lower() == norm_t),
                    None,
                )
                if matching_toc:
                    page_num = matching_toc.get("page_number")
                    # If TOC specified a level, confirm or adjust
                    if matching_toc.get("level"):
                        level = matching_toc["level"]

            classified_items.append({
                "is_heading": is_head,
                "number": num,
                "title": title,
                "level": level,
                "text": el["text"],
                "page_number": page_num,
                "confidence": conf,
                "is_ambiguous": is_ambig,
            })
        else:
            classified_items.append({
                "is_heading": False,
                "text": el["text"],
            })

    # 4. Build hierarchical tree from classified items
    sections = build_hierarchy_tree(classified_items)

    # 5. Fallback 1: If no sections detected from body but TOC exists, build tree from TOC entries
    if not sections and toc_entries:
        logger.info("Using Table of Contents outline to build section hierarchy")
        toc_classified = [
            {
                "is_heading": True,
                "number": e["number"],
                "title": e["title"],
                "level": e["level"],
                "text": "",
                "page_number": e.get("page_number"),
                "confidence": 0.95,
                "is_ambiguous": False,
            }
            for e in toc_entries
        ]
        sections = build_hierarchy_tree(toc_classified)



    # 7. Extract General Information Field Definitions
    all_document_text = [el["text"] for el in raw_elements]
    extracted_fields: list[ExtractedField] = []
    seen_keys: set[str] = set()

    combined_text = "\n".join(all_document_text)
    detected_placeholders = _extract_placeholders(combined_text)

    order_counter = 1
    for key in detected_placeholders:
        if key not in seen_keys:
            seen_keys.add(key)
            name = key.replace("_", " ").title()
            field_type = "date" if "date" in key else "phone" if "phone" in key or "contact" in key else "text"
            extracted_fields.append(
                ExtractedField(
                    fieldKey=key,
                    fieldName=name,
                    fieldType=field_type,
                    placeholder=f"Enter {name}",
                    helpText=f"Dynamically populated value for {name}",
                    isRequired=True,
                    orderIndex=order_counter,
                )
            )
            order_counter += 1

    for default_field in COMMON_SMP_FIELDS:
        f_key = default_field["fieldKey"]
        if f_key not in seen_keys:
            seen_keys.add(f_key)
            extracted_fields.append(
                ExtractedField(
                    fieldKey=f_key,
                    fieldName=default_field["fieldName"],
                    fieldType=default_field["fieldType"],
                    placeholder=default_field.get("placeholder"),
                    helpText=default_field.get("helpText"),
                    isRequired=default_field.get("isRequired", False),
                    orderIndex=order_counter,
                )
            )
            order_counter += 1

    return ParsedDocumentSectionsResponse(
        documentTitle=document_title,
        generalInformationFields=extracted_fields,
        sections=sections,
        tableOfContents=toc_entries,
    )


def parse_docx_sections_bytes(content: bytes) -> ParsedDocumentSectionsResponse:
    """Entry point for parsing raw bytes into document sections."""
    return parse_docx_document_sections(io.BytesIO(content))
