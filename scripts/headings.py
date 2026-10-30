from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
try:
    from .core_xml import insert_ordered_p_pr
except ImportError:
    from core_xml import insert_ordered_p_pr


def add_academic_heading(
    document,
    text: str,
    level: int = 1,
    font_name: str = "Times New Roman",
    center_if_level_1: bool = True,
) -> Paragraph:
    heading_style = f"Heading {level}"
    try:
        p = document.add_paragraph(text, style=heading_style)
    except KeyError:
        p = document.add_paragraph(text)

    fmt = p.paragraph_format
    fmt.keep_with_next = True
    fmt.space_before = Pt(12 if level == 1 else 8)
    fmt.space_after = Pt(6)

    if level == 1 and center_if_level_1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    for run in p.runs:
        run.font.name = font_name
        run.bold = True
        run.font.size = Pt(14 if level == 1 else (13 if level == 2 else 12))

    set_outline_level(p, level - 1)
    return p


def set_outline_level(paragraph: Paragraph, level: int) -> Paragraph:
    pPr = paragraph._p.get_or_add_pPr()
    outlineLvl = pPr.find(qn("w:outlineLvl"))
    if outlineLvl is None:
        outlineLvl = OxmlElement("w:outlineLvl")
        insert_ordered_p_pr(pPr, outlineLvl)
    outlineLvl.set(qn("w:val"), str(level))
    return paragraph
