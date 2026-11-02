from __future__ import annotations

import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
try:
    from .core_xml import insert_ordered_p_pr
    from .headings import get_max_body_heading_level, set_outline_level
    from .references import add_bookmark
except ImportError:
    from core_xml import insert_ordered_p_pr
    from headings import get_max_body_heading_level, set_outline_level
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
    """Mendeteksi tipe caption ('Gambar' atau 'Tabel') dari field SEQ atau awalan teks."""
    for fld in paragraph._p.iter(qn("w:fldSimple")):
        instr = fld.get(qn("w:instr"), "")
        if "SEQ Gambar" in instr:
            return "Gambar"
        if "SEQ Tabel" in instr:
            return "Tabel"
    m = re.match(r"^(gambar|tabel)\b", paragraph.text.strip(), re.IGNORECASE)
    return m.group(1).capitalize() if m else None


def resolve_caption_headings(document) -> dict[str, int]:
    """Menyelaraskan level heading caption secara dinamis terhadap max body heading level."""
    max_body = get_max_body_heading_level(document)
    target_img = max(4, max_body + 1) if max_body >= 4 else 4
    target_tbl = min(9, target_img + 1)
    target_img = min(9, target_img)
    resolved = {"Gambar": target_img, "Tabel": target_tbl, "max_body": max_body}

    for p in document.paragraphs:
        cap_type = detect_caption_type(p)
        if cap_type in resolved:
            apply_caption_heading(p, resolved[cap_type])
    return resolved

def add_seq_caption(
    document,
    label: str,
    description: str,
    bookmark_name: str | None = None, bookmark_id: int = 1,
    font_name: str = "Times New Roman", size_pt: float = 10.0, align: str | None = None,
    strict_length: bool = False, bold: bool = False, italic: bool = False, separator: str = ".",
    space_before_pt: float = 0.0, space_after_pt: float = 0.0,
    heading_level: int | str | None = "auto",
) -> Paragraph:
    clean_desc = re.sub(r"[\r\n]+", " ", description or "").strip()
    if clean_desc and len(clean_desc) > 45:
        msg = f"Deskripsi caption {label} melebihi batas 45 karakter ({len(clean_desc)} karakter): '{clean_desc}'"
        if strict_length:
            raise ValueError(msg)
        import warnings
        warnings.warn(msg, UserWarning, stacklevel=2)
    p = document.add_paragraph()
    if heading_level == "auto":
        max_body = get_max_body_heading_level(document)
        target_img = max(4, max_body + 1) if max_body >= 4 else 4
        lvl = target_img if label.lower().startswith("gambar") else min(9, target_img + 1)
        apply_caption_heading(p, lvl)
    elif isinstance(heading_level, int):
        apply_caption_heading(p, heading_level)
    if align is None:
        target_align = "left" if label.lower().startswith("tabel") else "center"
    else:
        target_align = align.lower()
    align_map = {"left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT}
    p.alignment = align_map.get(target_align, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)

    run_label = p.add_run(f"{label} ")
    run_label.font.name = font_name
    run_label.font.size = Pt(size_pt)
    run_label.bold = bold
    run_label.italic = italic
    run_label.font.color.rgb = RGBColor(0, 0, 0)

    # Field SEQ (misal SEQ Gambar \* ARABIC)
    seq_xml = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="SEQ {label} \\* ARABIC"/>')
    p._p.append(seq_xml)

    if clean_desc:
        sep_prefix = f"{separator.strip()} " if separator.strip() else " "
        run_desc = p.add_run(f"{sep_prefix}{clean_desc}")
        run_desc.font.name = font_name
        run_desc.font.size = Pt(size_pt)
        run_desc.bold = bold
        run_desc.italic = italic
        run_desc.font.color.rgb = RGBColor(0, 0, 0)
    if bookmark_name:
        add_bookmark(p, bookmark_id=bookmark_id, bookmark_name=bookmark_name)

    return p
