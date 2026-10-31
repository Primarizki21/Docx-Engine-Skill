from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
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
    bold: bool | None = None,
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

    is_bold = (level <= 2) if bold is None else bold
    for run in p.runs:
        run.font.name = font_name
        run.bold = is_bold
        run.font.size = Pt(14 if level == 1 else (13 if level == 2 else 12))
        run.font.color.rgb = RGBColor(0, 0, 0)
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

def is_caption_paragraph(paragraph: Paragraph) -> bool:
    """Mengecek apakah paragraf adalah caption (berisi field SEQ atau diawali Gambar/Tabel)."""
    for fld in paragraph._p.iter(qn("w:fldSimple")):
        instr = fld.get(qn("w:instr"), "")
        if "SEQ Gambar" in instr or "SEQ Tabel" in instr:
            return True
    text = paragraph.text.strip().lower()
    return bool(re.match(r"^(gambar|tabel)\s+(\d+|[ivxlcdm]+|[a-z]\b|:|\.)", text))


def get_max_body_heading_level(document) -> int:
    """Mendeteksi level heading tertinggi pada body text (di luar caption)."""
    max_level = 0
    for p in document.paragraphs:
        if is_caption_paragraph(p):
            continue
        level = 0
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None:
            oLvl = pPr.find(qn("w:outlineLvl"))
            if oLvl is not None and oLvl.get(qn("w:val"), "").isdigit():
                level = max(level, int(oLvl.get(qn("w:val"))) + 1)
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                m = re.match(r"^Heading\s*(\d+)$", pStyle.get(qn("w:val"), ""), re.IGNORECASE)
                level = max(level, int(m.group(1))) if m else level
        style_name = getattr(getattr(p, "style", None), "name", "")
        m_style = re.match(r"^Heading\s*(\d+)$", style_name, re.IGNORECASE)
        level = max(level, int(m_style.group(1))) if m_style else level
        max_level = max(max_level, level)
    return max_level
