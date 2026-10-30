from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt
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
) -> Paragraph:
    if description and len(description) > 45:
        msg = f"Deskripsi caption {label} melebihi batas 45 karakter ({len(description)} karakter): '{description}'"
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
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)

    # Prefix label (e.g. "Gambar ")
    run_label = p.add_run(f"{label} ")
    run_label.font.name = font_name
    run_label.font.size = Pt(size_pt)
    run_label.bold = True

    # Field SEQ (misal SEQ Gambar \* ARABIC)
    seq_xml = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="SEQ {label} \\* ARABIC"/>')
    p._p.append(seq_xml)

    # Separator & Description
    run_desc = p.add_run(f": {description}")
    run_desc.font.name = font_name
    run_desc.font.size = Pt(size_pt)
    run_desc.italic = True

    if bookmark_name:
        add_bookmark(p, bookmark_id=bookmark_id, bookmark_name=bookmark_name)

    return p
