from __future__ import annotations

import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor
try:
    from .references import add_bookmark
except ImportError:
    from references import add_bookmark


def add_seq_caption(
    document,
    label: str,
    description: str,
    bookmark_name: str | None = None,
    bookmark_id: int = 1,
    font_name: str = "Times New Roman",
    size_pt: float = 10.0,
    align: str = "center",
    strict_length: bool = False,
    bold: bool = False,
    separator: str = ".",
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
) -> Paragraph:
    clean_desc = re.sub(r"[\r\n]+", " ", description or "").strip()
    if clean_desc and len(clean_desc) > 45:
        msg = f"Deskripsi caption {label} melebihi batas 45 karakter ({len(clean_desc)} karakter): '{clean_desc}'"
        if strict_length:
            raise ValueError(msg)
        import warnings
        warnings.warn(msg, UserWarning, stacklevel=2)
    p = document.add_paragraph()
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p.alignment = align_map.get(align.lower(), WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)

    # Prefix label (e.g. "Gambar ") - murni italic, no bold
    run_label = p.add_run(f"{label} ")
    run_label.font.name = font_name
    run_label.font.size = Pt(size_pt)
    run_label.bold = bold
    run_label.italic = True
    run_label.font.color.rgb = RGBColor(0, 0, 0)

    # Field SEQ (misal SEQ Gambar \* ARABIC)
    seq_xml = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="SEQ {label} \\* ARABIC"/>')
    p._p.append(seq_xml)

    # Separator & Description - murni italic, no bold
    if clean_desc:
        sep_prefix = f"{separator.strip()} " if separator.strip() else " "
        run_desc = p.add_run(f"{sep_prefix}{clean_desc}")
        run_desc.font.name = font_name
        run_desc.font.size = Pt(size_pt)
        run_desc.bold = bold
        run_desc.italic = True
        run_desc.font.color.rgb = RGBColor(0, 0, 0)
    if bookmark_name:
        add_bookmark(p, bookmark_id=bookmark_id, bookmark_name=bookmark_name)

    return p
