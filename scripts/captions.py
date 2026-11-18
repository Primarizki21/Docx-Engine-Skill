from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
try:
    from .core_xml import insert_ordered_p_pr
    from .headings import set_outline_level
    from .lexicon import italic_segments
    from .references import add_bookmark
except ImportError:
    from core_xml import insert_ordered_p_pr
    from headings import set_outline_level
    from lexicon import italic_segments
    from references import add_bookmark


def apply_caption_heading(paragraph: Paragraph, level: int) -> Paragraph:
    """Menerapkan style Heading dan outlineLvl pada caption tanpa mengubah run formatting."""
    lvl = max(1, min(9, int(level)))
    heading_style = f"Heading {lvl}"
    try:
        paragraph.style = heading_style
    except (KeyError, ValueError):
        pPr = paragraph._p.get_or_add_pPr()
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            pStyle = OxmlElement("w:pStyle")
            insert_ordered_p_pr(pPr, pStyle)
        pStyle.set(qn("w:val"), f"Heading{lvl}")
    set_outline_level(paragraph, lvl - 1)
    return paragraph


def detect_caption_type(paragraph: Paragraph) -> str | None:
    """Mendeteksi tipe caption ('Gambar' atau 'Tabel') dari awalan teks."""
    text = (paragraph.text or "").strip().lower()
    if text.startswith("gambar"):
        return "Gambar"
    if text.startswith("tabel"):
        return "Tabel"
    return None


def _is_chapter_heading(paragraph: Paragraph) -> bool:
    """Heading 1 diawali 'BAB ' yang menentukan nomor bab caption."""
    pPr = paragraph._p.pPr
    pStyle = None if pPr is None else pPr.pStyle
    outline = None if pPr is None else pPr.find(qn("w:outlineLvl"))
    is_h1 = (
        (getattr(paragraph.style, "name", "") or "") == "Heading 1"
        or (pStyle is not None and (pStyle.val or "") == "Heading1")
        or (outline is not None and outline.get(qn("w:val")) == "0")
    )
    return is_h1 and (paragraph.text or "").strip().startswith("BAB ")


def _caption_number(document, kind: str) -> str:
    """Nomor X.Y: X urutan BAB, Y urutan caption sejenis di dalam bab itu."""
    chapter = 0
    seq = 0
    for paragraph in document.paragraphs:
        if _is_chapter_heading(paragraph):
            chapter += 1
            seq = 0
        elif detect_caption_type(paragraph) == kind:
            seq += 1
    chapter = max(1, chapter)
    return f"{chapter}.{seq + 1}"


def add_seq_caption(
    document,
    label: str,
    description: str,
    bookmark_name: str | None = None, bookmark_id: int = 1,
    font_name: str = "Times New Roman", size_pt: float = 10.0, align: str | None = None,
    strict_length: bool = True, bold: bool = False, italic: bool = False, separator: str = ".",
    space_before_pt: float = 0.0, space_after_pt: float = 0.0,
    heading_level: int | str | None = "auto",
) -> Paragraph:
    clean_desc = " ".join((description or "").split())
    if len(clean_desc) > 45:
        raise ValueError(
            f"Deskripsi caption {label} melebihi batas 45 karakter "
            f"({len(clean_desc)} karakter): '{clean_desc}'"
        )
    is_table = (label or "").strip().lower().startswith("tabel")
    kind = "Tabel" if is_table else "Gambar"
    number = _caption_number(document, kind)
    p = document.add_paragraph()
    apply_caption_heading(p, 5 if is_table else 4)
    if align is None:
        target_align = "left" if is_table else "center"
    else:
        target_align = align.lower()
    align_map = {"left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT}
    p.alignment = align_map.get(target_align, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)

    def style_run(run, is_italic: bool = False) -> None:
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = False
        run.italic = is_italic
        run.font.color.rgb = RGBColor(0, 0, 0)

    style_run(p.add_run(f"{kind} {number}"))
    if clean_desc:
        sep = separator.strip()
        style_run(p.add_run(f"{sep} " if sep else " "))
        for segment, is_italic in italic_segments(clean_desc):
            style_run(p.add_run(segment), is_italic)
    if bookmark_name:
        add_bookmark(p, bookmark_id=bookmark_id, bookmark_name=bookmark_name)
    return p
