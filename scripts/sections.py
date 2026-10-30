from __future__ import annotations

from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.section import Section
from docx.shared import Inches


def set_a4_page_size(section: Section) -> Section:
    section.page_width = Inches(8.267)
    section.page_height = Inches(11.692)
    return section


def set_letter_page_size(section: Section) -> Section:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    return section


def set_academic_margins(
    section: Section,
    top_cm: float = 3.0,
    bottom_cm: float = 3.0,
    left_cm: float = 4.0,
    right_cm: float = 3.0,
) -> Section:
    section.top_margin = Inches(top_cm / 2.54)
    section.bottom_margin = Inches(bottom_cm / 2.54)
    section.left_margin = Inches(left_cm / 2.54)
    section.right_margin = Inches(right_cm / 2.54)
    return section


def set_two_column_section(section: Section, col_space_inches: float = 0.5) -> Section:
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), str(int(col_space_inches * 1440)))
    cols.set(qn("w:equalWidth"), "1")
    return section


def add_section_break(document, start_type: str = "new_page") -> Section:
    type_map = {
        "new_page": WD_SECTION_START.NEW_PAGE,
        "continuous": WD_SECTION_START.CONTINUOUS,
        "even_page": WD_SECTION_START.EVEN_PAGE,
        "odd_page": WD_SECTION_START.ODD_PAGE,
    }
    st = type_map.get(start_type.lower(), WD_SECTION_START.NEW_PAGE)
    new_section = document.add_section(st)
    return new_section
