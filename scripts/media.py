from __future__ import annotations

from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph

try:
    from .captions import apply_caption_heading
    from .lexicon import italic_segments
except ImportError:
    from captions import apply_caption_heading
    from lexicon import italic_segments


def _clean_caption_text(text: str | None) -> str:
    """Gabungkan break baris menjadi satu spasi, tanpa regex."""
    out: list[str] = []
    pending_break = False
    for ch in text or "":
        if ch in "\r\n":
            pending_break = True
            continue
        if pending_break:
            if out and out[-1] != " ":
                out.append(" ")
            pending_break = False
        out.append(ch)
    return "".join(out).strip()


def add_image_with_caption(
    document,
    image_path: str | Path,
    caption: str,
    width_inches: float = 5.5,
    caption_font_name: str = "Times New Roman",
    caption_size_pt: float = 10.0,
    heading_level: int = 4,
) -> tuple[Paragraph, Paragraph]:
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")

    # 1 inci = 914400 EMU
    p_img = document.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(0)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(str(path), width=Inches(width_inches))

    p_cap = document.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(0)
    p_cap.paragraph_format.keep_with_next = False
    apply_caption_heading(p_cap, heading_level)

    clean_cap = _clean_caption_text(caption)
    for text, is_italic in italic_segments(clean_cap):
        run = p_cap.add_run(text)
        run.font.name = caption_font_name
        run.font.size = Pt(caption_size_pt)
        run.italic = is_italic
        run.bold = False
        run.font.color.rgb = RGBColor(0, 0, 0)
    return p_img, p_cap
