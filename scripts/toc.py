from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt
from docx.text.paragraph import Paragraph


def add_table_of_contents(
    document,
    title: str = "DAFTAR ISI",
    heading_levels: str = "1-3",
) -> Paragraph:
    if title:
        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(12)

    p_toc = document.add_paragraph()
    fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\o &quot;{heading_levels}&quot; \\h \\z \\u"/>')
    p_toc._p.append(fld)
    return p_toc


def add_list_of_figures(
    document,
    title: str = "DAFTAR GAMBAR",
    seq_name: str = "Gambar",
) -> Paragraph:
    if title:
        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(12)

    p_lof = document.add_paragraph()
    fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\c &quot;{seq_name}&quot; \\h \\z"/>')
    p_lof._p.append(fld)
    return p_lof


def add_list_of_tables(
    document,
    title: str = "DAFTAR TABEL",
    seq_name: str = "Tabel",
) -> Paragraph:
    if title:
        p_title = document.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(12)

    p_lot = document.add_paragraph()
    fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\c &quot;{seq_name}&quot; \\h \\z"/>')
    p_lot._p.append(fld)
    return p_lot
