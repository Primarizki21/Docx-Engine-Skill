from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
try:
    from .core_xml import insert_ordered_p_pr
    from .sections import add_section_break
except ImportError:
    from core_xml import insert_ordered_p_pr
    from sections import add_section_break

def add_academic_heading(
    document,
    text: str,
    level: int = 1,
    font_name: str = "Times New Roman",
    center_if_level_1: bool = True,
    section_break: bool = False,
    strict_length: bool = False,
) -> Paragraph:
    if level >= 2 and len(text) > 45:
        msg = f"Judul heading level {level} melebihi batas 45 karakter ({len(text)} karakter): '{text}'"
        if strict_length:
            raise ValueError(msg)
        import warnings
        warnings.warn(msg, UserWarning, stacklevel=2)
    if section_break and len(document.paragraphs) > 0:
        add_section_break(document, start_type="new_page")
    heading_style = f"Heading {level}"
    try:
        p = document.add_paragraph(text, style=heading_style)
    except KeyError:
        p = document.add_paragraph(text)

    fmt = p.paragraph_format
    fmt.keep_with_next = True
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)

    if level == 1 and center_if_level_1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    is_bold = level <= 2
    for run in p.runs:
        run.font.name = font_name
        run.bold = is_bold
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


def add_chapter_heading(
    document,
    text: str,
    font_name: str = "Times New Roman",
) -> Paragraph:
    """Menambahkan Heading 1 BAB baru dengan pemisah Section Break (Next Page).

    Jika dokumen sudah memiliki paragraf sebelumnya, otomatis menyisipkan
    Section Break (Next Page) agar setiap bab terisolasi dalam seksi terpisah.
    """
    return add_academic_heading(
        document,
        text,
        level=1,
        font_name=font_name,
        center_if_level_1=True,
        section_break=True,
    )
