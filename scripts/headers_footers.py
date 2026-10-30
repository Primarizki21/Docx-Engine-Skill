from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.section import Section
from docx.shared import Pt


def add_page_number_to_footer(section: Section, align: str = "center") -> None:
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p.alignment = align_map.get(align.lower(), WD_ALIGN_PARAGRAPH.CENTER)

    fld_simple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
    p._p.append(fld_simple)


def add_page_number_with_total(section: Section, align: str = "center") -> None:
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p.alignment = align_map.get(align.lower(), WD_ALIGN_PARAGRAPH.CENTER)

    run_page = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
    run_of = parse_xml(f'<w:r {nsdecls("w")}><w:t xml:space="preserve"> dari </w:t></w:r>')
    run_total = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="NUMPAGES"/>')

    p._p.append(run_page)
    p._p.append(run_of)
    p._p.append(run_total)


def set_running_header(
    section: Section,
    text: str,
    align: str = "right",
    font_name: str = "Times New Roman",
    size_pt: float = 10.0,
) -> None:
    header = section.header
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p.alignment = align_map.get(align.lower(), WD_ALIGN_PARAGRAPH.RIGHT)
    p.text = text
    for run in p.runs:
        run.font.name = font_name
        run.font.size = Pt(size_pt)
