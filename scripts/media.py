from __future__ import annotations

import re
from pathlib import Path
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
try:
    from .captions import apply_caption_heading
    from .headings import get_max_body_heading_level
except ImportError:
    from captions import apply_caption_heading
    from headings import get_max_body_heading_level

def add_image_with_caption(
    document,
    image_path: str | Path,
    caption: str,
    width_inches: float = 5.5,
    caption_font_name: str = "Times New Roman",
    caption_size_pt: float = 10.0,
    heading_level: int | str | None = "auto",
) -> tuple[Paragraph, Paragraph]:
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")

    # 1 inci = 914400 EMU
    p_img = document.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(4)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(str(path), width=Inches(width_inches))

    p_cap = document.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(10)
    p_cap.paragraph_format.keep_with_next = False
    if heading_level == "auto":
        max_body = get_max_body_heading_level(document)
        target_img = max(4, max_body + 1) if max_body >= 4 else 4
        apply_caption_heading(p_cap, target_img)
    elif isinstance(heading_level, int):
        apply_caption_heading(p_cap, heading_level)

    clean_cap = re.sub(r"[\r\n]+", " ", caption or "").strip()
    run_cap = p_cap.add_run(clean_cap)
    run_cap.font.name = caption_font_name
    run_cap.font.size = Pt(caption_size_pt)
    run_cap.bold = False
    run_cap.italic = True
    run_cap.font.color.rgb = RGBColor(0, 0, 0)
    return p_img, p_cap
