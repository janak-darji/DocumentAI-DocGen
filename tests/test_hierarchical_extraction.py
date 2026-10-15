"""Unit tests for hierarchical section and subsection extraction from DOCX."""

import io
from docx import Document

from app.services.document_section_parser import (
    parse_heading_prefix,
    build_hierarchy_tree,
    parse_docx_document_sections,
)


def test_parse_heading_prefixes():
    # 1. Numeric & Nested Numeric
    num, title, depth, ptype = parse_heading_prefix("1 WHS Policy and Commitment")
    assert num == "1"
    assert title == "WHS Policy and Commitment"
    assert depth == 1
    assert ptype == "numeric"

    num, title, depth, ptype = parse_heading_prefix("1.1 General Policy")
    assert num == "1.1"
    assert title == "General Policy"
    assert depth == 2
    assert ptype == "nested_numeric"

    num, title, depth, ptype = parse_heading_prefix("6.1.1 Pre-Start Meeting")
    assert num == "6.1.1"
    assert title == "Pre-Start Meeting"
    assert depth == 3
    assert ptype == "nested_numeric"

    # 2. Alphabetic & Nested Alphabetic
    num, title, depth, ptype = parse_heading_prefix("A Scope of Works")
    assert num == "A"
    assert title == "Scope of Works"
    assert depth == 1
    assert ptype == "alphabetic"

    num, title, depth, ptype = parse_heading_prefix("a) General Requirements")
    assert num == "a"
    assert title == "General Requirements"
    assert depth == 2
    assert ptype == "alphabetic_sub"

    num, title, depth, ptype = parse_heading_prefix("A.1 General Policy")
    assert num == "A.1"
    assert title == "General Policy"
    assert depth == 2
    assert ptype == "mixed"

    # 3. Roman Numerals
    num, title, depth, ptype = parse_heading_prefix("I. Introduction")
    assert num == "I"
    assert title == "Introduction"
    assert depth == 1
    assert ptype == "roman"

    num, title, depth, ptype = parse_heading_prefix("ii) Safety Procedures")
    assert num == "ii"
    assert title == "Safety Procedures"
    assert depth == 2
    assert ptype == "roman_sub"

    # 4. Mixed Numbering
    num, title, depth, ptype = parse_heading_prefix("1.A Site Establishment")
    assert num == "1.A"
    assert title == "Site Establishment"
    assert depth == 2
    assert ptype == "mixed"

    num, title, depth, ptype = parse_heading_prefix("1.1.A First Aid")
    assert num == "1.1.A"
    assert title == "First Aid"
    assert depth == 3
    assert ptype == "mixed"

    # 5. Prefixed
    num, title, depth, ptype = parse_heading_prefix("Section 1: Responsibilities")
    assert num == "1"
    assert title == "Responsibilities"
    assert depth == 1
    assert ptype == "prefixed"

    # 6. Unnumbered
    num, title, depth, ptype = parse_heading_prefix("Emergency Numbers and General Information")
    assert num is None
    assert title == "Emergency Numbers and General Information"
    assert depth == 1
    assert ptype == "unnumbered"


def test_alphabetic_and_nested_alphabetic_hierarchy():
    """Tests alphabetic sections (A -> A.1 -> a))."""
    classified = [
        {"is_heading": True, "number": "A", "title": "Scope", "level": 1, "text": "A Scope"},
        {"is_heading": False, "text": "Scope of work description."},
        {"is_heading": True, "number": "A.1", "title": "Inclusions", "level": 2, "text": "A.1 Inclusions"},
        {"is_heading": False, "text": "All internal fitouts."},
        {"is_heading": True, "number": "a", "title": "Demolition", "level": 3, "text": "a) Demolition"},
        {"is_heading": True, "number": "b", "title": "Carpentry", "level": 3, "text": "b) Carpentry"},
        {"is_heading": True, "number": "B", "title": "Responsibilities", "level": 1, "text": "B Responsibilities"},
    ]
    tree = build_hierarchy_tree(classified)
    assert len(tree) == 2
    assert tree[0].number == "A"
    assert tree[0].title == "Scope"
    assert len(tree[0].children) == 1
    assert tree[0].children[0].number == "A.1"
    assert len(tree[0].children[0].children) == 2
    assert tree[0].children[0].children[0].number == "a"
    assert tree[0].children[0].children[1].number == "b"
    assert tree[1].number == "B"
    assert tree[1].title == "Responsibilities"


def test_roman_numeral_hierarchy():
    """Tests Roman numeral sections (I -> i), ii))."""
    classified = [
        {"is_heading": True, "number": "I", "title": "Overview", "level": 1, "text": "I Overview"},
        {"is_heading": True, "number": "II", "title": "Risk Controls", "level": 1, "text": "II Risk Controls"},
        {"is_heading": True, "number": "i", "title": "Elimination", "level": 2, "text": "i) Elimination"},
        {"is_heading": True, "number": "ii", "title": "Substitution", "level": 2, "text": "ii) Substitution"},
    ]
    tree = build_hierarchy_tree(classified)
    assert len(tree) == 2
    assert tree[0].number == "I"
    assert tree[1].number == "II"
    assert len(tree[1].children) == 2
    assert tree[1].children[0].number == "i"
    assert tree[1].children[1].number == "ii"


def test_mixed_numbering_hierarchy():
    """Tests mixed numbering combinations such as 1.A, 1.1.A, A.1.1."""
    classified = [
        {"is_heading": True, "number": "1", "title": "Site Setup", "level": 1, "text": "1 Site Setup"},
        {"is_heading": True, "number": "1.A", "title": "Preliminaries", "level": 2, "text": "1.A Preliminaries"},
        {"is_heading": True, "number": "1.B", "title": "Execution", "level": 2, "text": "1.B Execution"},
        {"is_heading": True, "number": "1.B.1", "title": "Excavation Works", "level": 3, "text": "1.B.1 Excavation"},
    ]
    tree = build_hierarchy_tree(classified)
    assert len(tree) == 1
    assert tree[0].number == "1"
    assert len(tree[0].children) == 2
    assert tree[0].children[0].number == "1.A"
    assert tree[0].children[1].number == "1.B"
    assert len(tree[0].children[1].children) == 1
    assert tree[0].children[1].children[0].number == "1.B.1"


def test_numbering_restarts_within_new_parent():
    """Tests that numbering sequences can restart within a new parent section."""
    classified = [
        {"is_heading": True, "number": "1", "title": "Setup", "level": 1, "text": "1 Setup"},
        {"is_heading": True, "number": "1.1", "title": "Access", "level": 2, "text": "1.1 Access"},
        {"is_heading": True, "number": "2", "title": "Hazards", "level": 1, "text": "2 Hazards"},
        {"is_heading": True, "number": "1.1", "title": "Manual Handling", "level": 2, "text": "1.1 Manual Handling"},
    ]
    tree = build_hierarchy_tree(classified)
    assert len(tree) == 2
    assert tree[0].number == "1"
    assert tree[0].children[0].number == "1.1"
    assert tree[0].children[0].title == "Access"

    assert tree[1].number == "2"
    assert tree[1].children[0].number == "1.1"
    assert tree[1].children[0].title == "Manual Handling"


def test_malformed_numbering_graceful_handling():
    """Tests that orphan/malformed subsections (e.g. 2.1 when no 2 exists) are preserved."""
    classified = [
        {"is_heading": True, "number": "2.1", "title": "Orphan Subsection", "level": 2, "text": "2.1 Orphan Subsection"},
        {"is_heading": False, "text": "Important safety notes that must not be deleted."},
    ]
    tree = build_hierarchy_tree(classified)
    assert len(tree) == 1
    assert tree[0].title == "Orphan Subsection"
    assert "Important safety notes" in tree[0].content


def test_hierarchy_tree_construction_with_duplicate_and_nested():
    """
    Tests stack-based hierarchy builder with:
    - Unnumbered top-level sections
    - Numeric 1 -> 1.1, 1.2
    - Duplicate numbering (two 5.1s)
    - 3-level nesting (6 -> 6.1 -> 6.1.1, 6.1.2)
    - Content association
    """
    classified = [
        {"is_heading": True, "number": None, "title": "Introduction", "level": 1, "text": "Introduction"},
        {"is_heading": False, "text": "This is introductory text."},
        {"is_heading": True, "number": None, "title": "Emergency Numbers and General Information", "level": 1, "text": "Emergency Numbers"},
        {"is_heading": True, "number": "1", "title": "WHS Policy and Commitment", "level": 1, "text": "1 WHS Policy"},
        {"is_heading": False, "text": "Policy preamble content."},
        {"is_heading": True, "number": "1.1", "title": "General Policy", "level": 2, "text": "1.1 General Policy"},
        {"is_heading": False, "text": "General policy content details."},
        {"is_heading": True, "number": "1.2", "title": "Commitment to policy objective", "level": 2, "text": "1.2 Commitment"},
        {"is_heading": True, "number": "2", "title": "Rehabilitation Policy", "level": 1, "text": "2 Rehabilitation Policy"},
        {"is_heading": True, "number": "5", "title": "Roles and Responsibilities", "level": 1, "text": "5 Roles and Responsibilities"},
        {"is_heading": True, "number": "5.1", "title": "Workers and Sub-Contractors Roles and Responsibilities", "level": 2, "text": "5.1 Workers"},
        {"is_heading": True, "number": "5.1", "title": "Register of all workers with specific WHS responsibilities", "level": 2, "text": "5.1 Register"},
        {"is_heading": True, "number": "6", "title": "Consultation, Cooperation, Coordination", "level": 1, "text": "6 Consultation"},
        {"is_heading": True, "number": "6.1", "title": "Consultation Procedures", "level": 2, "text": "6.1 Consultation Procedures"},
        {"is_heading": True, "number": "6.1.1", "title": "Pre-Start Meeting", "level": 3, "text": "6.1.1 Pre-Start Meeting"},
        {"is_heading": False, "text": "Pre-start meeting happens daily at 7 AM."},
        {"is_heading": True, "number": "6.1.2", "title": "Toolbox Meeting (TBM)", "level": 3, "text": "6.1.2 Toolbox Meeting"},
        {"is_heading": True, "number": "7", "title": "Training and Communication", "level": 1, "text": "7 Training and Communication"},
    ]

    tree = build_hierarchy_tree(classified)

    # 1. Root level checks
    root_titles = [node.title for node in tree]
    assert root_titles == [
        "Introduction",
        "Emergency Numbers and General Information",
        "WHS Policy and Commitment",
        "Rehabilitation Policy",
        "Roles and Responsibilities",
        "Consultation, Cooperation, Coordination",
        "Training and Communication",
    ]

    # 2. Check Introduction has content and no children
    assert tree[0].number is None
    assert tree[0].level == 1
    assert "introductory text" in tree[0].content
    assert len(tree[0].children) == 0

    # 3. Check Section 1 has subsections 1.1 and 1.2
    sec1 = tree[2]
    assert sec1.number == "1"
    assert "Policy preamble" in sec1.content
    assert len(sec1.children) == 2
    assert sec1.children[0].number == "1.1"
    assert "General policy content details" in sec1.children[0].content
    # Content must NOT be duplicated in parent
    assert "General policy content details" not in sec1.content

    # 4. Check Section 5 preserves DUPLICATE 5.1s separately
    sec5 = tree[4]
    assert sec5.number == "5"
    assert len(sec5.children) == 2
    assert sec5.children[0].number == "5.1"
    assert sec5.children[0].title == "Workers and Sub-Contractors Roles and Responsibilities"
    assert sec5.children[1].number == "5.1"
    assert sec5.children[1].title == "Register of all workers with specific WHS responsibilities"

    # 5. Check Section 6 3-level hierarchy: 6 -> 6.1 -> 6.1.1, 6.1.2
    sec6 = tree[5]
    assert sec6.number == "6"
    assert len(sec6.children) == 1
    sub6_1 = sec6.children[0]
    assert sub6_1.number == "6.1"
    assert sub6_1.title == "Consultation Procedures"
    assert sub6_1.level == 2
    assert len(sub6_1.children) == 2
    assert sub6_1.children[0].number == "6.1.1"
    assert sub6_1.children[0].title == "Pre-Start Meeting"
    assert sub6_1.children[0].level == 3
    assert "Pre-start meeting happens daily" in sub6_1.children[0].content
    assert sub6_1.children[1].number == "6.1.2"
    assert sub6_1.children[1].title == "Toolbox Meeting (TBM)"
    assert sub6_1.children[1].level == 3


def test_docx_full_parse_with_toc_and_body():
    """
    Creates an in-memory DOCX file matching the Table of Contents in the screenshot
    and tests that parse_docx_document_sections extracts the exact hierarchy.
    """
    doc = Document()
    doc.add_paragraph("Commercial Fitout Safety Management Plan")

    # Add Table of Contents section
    doc.add_paragraph("Contents")
    doc.add_paragraph("Introduction .................................................... 6")
    doc.add_paragraph("Emergency Numbers and General Information ....................... 6")
    doc.add_paragraph("1   WHS Policy and Commitment ................................... 9")
    doc.add_paragraph("    1.1   General Policy ........................................ 9")
    doc.add_paragraph("    1.2   Commitment to policy objective......................... 9")
    doc.add_paragraph("    1.3   Scope of Policy – who it applies to.................... 9")
    doc.add_paragraph("    1.4   Site Safety Plan....................................... 9")
    doc.add_paragraph("2   Rehabilitation Policy ....................................... 10")
    doc.add_paragraph("3   Environmental Policy......................................... 10")
    doc.add_paragraph("4   Work Health and Safety Duties ............................... 11")
    doc.add_paragraph("    4.1   Duty of a PCBU/Business Owner.......................... 11")
    doc.add_paragraph("    4.2   Duty of Workers including Site Supervisors ............ 11")
    doc.add_paragraph("    4.3   Duty of other persons at the workplace (visitors) ..... 11")
    doc.add_paragraph("5   Roles and Responsibilities .................................. 12")
    doc.add_paragraph("    5.1   Workers and Sub-Contractors Roles and Responsibilities  .......... 12")
    doc.add_paragraph("    5.1   Register of all workers with specific WHS responsibilities ....... 12")
    doc.add_paragraph("6   Consultation, Cooperation, Coordination ..................... 13")
    doc.add_paragraph("    6.1   Consultation Procedures................................ 13")
    doc.add_paragraph("        6.1.1   Pre-Start Meeting ............................... 13")
    doc.add_paragraph("        6.1.2   Toolbox Meeting (TBM) ........................... 13")
    doc.add_paragraph("        6.1.3   Direct Consultation with Workers ................ 13")
    doc.add_paragraph("7   Training and Communication .................................. 14")

    # Body sections with paragraphs
    doc.add_paragraph("Introduction", style="Heading 1")
    doc.add_paragraph("This document outlines the safety management procedures for {{project_name}}.")

    doc.add_paragraph("Emergency Numbers and General Information", style="Heading 1")
    doc.add_paragraph("Emergency phone: {{emergency_phone}}.")

    doc.add_paragraph("1 WHS Policy and Commitment", style="Heading 1")
    doc.add_paragraph("Management commitment statement.")

    doc.add_paragraph("1.1 General Policy", style="Heading 2")
    doc.add_paragraph("General policy commitments and safety protocols.")

    doc.add_paragraph("1.2 Commitment to policy objective", style="Heading 2")
    doc.add_paragraph("Commitment details.")

    doc.add_paragraph("2 Rehabilitation Policy", style="Heading 1")
    doc.add_paragraph("Return to work program procedures.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)

    res = parse_docx_document_sections(bio)
    assert len(res.sections) >= 4
    assert res.sections[0].title == "Introduction"
    assert res.sections[0].number is None
    assert res.sections[0].level == 1
    assert "safety management procedures" in res.sections[0].content

    assert res.sections[1].title == "Emergency Numbers and General Information"
    assert res.sections[1].number is None

    assert res.sections[2].number == "1"
    assert res.sections[2].title == "WHS Policy and Commitment"
    assert len(res.sections[2].children) >= 2
    assert res.sections[2].children[0].number == "1.1"
    assert res.sections[2].children[0].title == "General Policy"
    assert res.sections[2].children[1].number == "1.2"

    assert res.sections[3].number == "2"
    assert res.sections[3].title == "Rehabilitation Policy"

    # Verify placeholder extraction
    field_keys = {f.fieldKey for f in res.generalInformationFields}
    assert "project_name" in field_keys
    assert "emergency_phone" in field_keys


def test_document_with_different_heading_structure():
    """
    Tests a document with an alternative heading structure:
    'Part A', 'Part B', Roman numerals, and numbered clauses without a TOC.
    """
    doc = Document()
    doc.add_paragraph("Alternative Safety Policy Document")

    doc.add_paragraph("Part A - General Governance", style="Heading 1")
    doc.add_paragraph("Policy provisions for general governance on site.")

    doc.add_paragraph("Clause 1.1 Scope and Objective", style="Heading 2")
    doc.add_paragraph("Applies to all construction trades.")

    doc.add_paragraph("Clause 1.2 Responsibilities", style="Heading 2")
    doc.add_paragraph("Principal contractor and subcontractor obligations.")

    doc.add_paragraph("Part B - Operational Risk Controls", style="Heading 1")
    doc.add_paragraph("Operational rules for high risk activities.")

    doc.add_paragraph("I. Hot Works", style="Heading 2")
    doc.add_paragraph("Fire watch and permit required.")

    doc.add_paragraph("II. Confined Space", style="Heading 2")
    doc.add_paragraph("Atmospheric testing mandatory.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)

    res = parse_docx_document_sections(bio)
    assert len(res.sections) == 2

    # Part A
    part_a = res.sections[0]
    assert part_a.title == "General Governance"
    assert part_a.level == 1
    assert len(part_a.children) == 2
    assert part_a.children[0].number == "1.1"
    assert part_a.children[0].title == "Scope and Objective"
    assert part_a.children[1].number == "1.2"
    assert part_a.children[1].title == "Responsibilities"

    # Part B
    part_b = res.sections[1]
    assert part_b.title == "Operational Risk Controls"
    assert part_b.level == 1
    assert len(part_b.children) == 2
    assert part_b.children[0].number == "I"
    assert part_b.children[0].title == "Hot Works"
    assert part_b.children[1].number == "II"
    assert part_b.children[1].title == "Confined Space"
