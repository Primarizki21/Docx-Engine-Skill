from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.text.paragraph import Paragraph


def set_paragraph_spacing(
    p: Paragraph,
    before_pt: float = 0.0,
    after_pt: float = 6.0,
    line_spacing: float = 1.15,
) -> Paragraph:
    fmt = p.paragraph_format
    fmt.space_before = Pt(before_pt)
    fmt.space_after = Pt(after_pt)
    fmt.line_spacing = line_spacing
    return p


def set_paragraph_indent(
    p: Paragraph,
    left_inches: float = 0.0,
    right_inches: float = 0.0,
    first_line_inches: float = 0.0,
) -> Paragraph:
    fmt = p.paragraph_format
    fmt.left_indent = Inches(left_inches) if left_inches else None
    fmt.right_indent = Inches(right_inches) if right_inches else None
    fmt.first_line_indent = Inches(first_line_inches) if first_line_inches else None
    return p


def set_paragraph_alignment(p: Paragraph, align: str = "left") -> Paragraph:
    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    p.alignment = mapping.get(align.lower(), WD_ALIGN_PARAGRAPH.LEFT)
    return p


def format_academic_paragraph(
    p: Paragraph,
    first_line_indent_cm: float = 1.0,
    line_spacing: float = 1.5,
) -> Paragraph:
    fmt = p.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.first_line_indent = Inches(first_line_indent_cm / 2.54)
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def format_table_cell_paragraph(
    p: Paragraph,
    line_spacing: float = 1.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    align: str = "left",
) -> Paragraph:
    """Memformat paragraf sel tabel: spasi single 1.0 (line=240), space before/after 0 pt."""
    fmt = p.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.space_before = Pt(space_before_pt)
    fmt.space_after = Pt(space_after_pt)
    fmt.first_line_indent = None
    fmt.left_indent = None
    set_paragraph_alignment(p, align)
    return p


def format_toc_entry_paragraph(
    p: Paragraph,
    line_spacing: float = 1.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 2.0,
) -> Paragraph:
    """Memformat entri daftar isi: spasi single 1.0 (line=240) dan jarak antar baris rapat."""
    fmt = p.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.space_before = Pt(space_before_pt)
    fmt.space_after = Pt(space_after_pt)
    return p


def format_preface_paragraph(
    p: Paragraph,
    first_line_indent_cm: float = 1.0,
    line_spacing: float = 1.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 4.0,
) -> Paragraph:
    """Memformat paragraf Kata Pengantar: spasi single 1.0 (line=240), indent 1 cm, dan rata justified."""
    fmt = p.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.first_line_indent = Inches(first_line_indent_cm / 2.54)
    fmt.space_before = Pt(space_before_pt)
    fmt.space_after = Pt(space_after_pt)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p
