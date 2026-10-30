from __future__ import annotations

import re
from docx.shared import Pt, RGBColor
from docx.text.run import Run


def format_run(
    run: Run,
    font_name: str = "Times New Roman",
    size_pt: float = 12.0,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color_rgb: tuple[int, int, int] | None = None,
) -> Run:
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    return run


def add_formatted_run(
    paragraph,
    text: str,
    font_name: str = "Times New Roman",
    size_pt: float = 12.0,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color_rgb: tuple[int, int, int] | None = None,
    use_smart_quotes: bool = False,
) -> Run:
    final_text = convert_to_smart_quotes(text) if use_smart_quotes else text
    run = paragraph.add_run(final_text)
    return format_run(
        run=run,
        font_name=font_name,
        size_pt=size_pt,
        bold=bold,
        italic=italic,
        underline=underline,
        color_rgb=color_rgb,
    )


def convert_to_smart_quotes(text: str) -> str:
    # Mengubah straight quotes (" dan ') menjadi smart quotes (‘ ’, “ ”)
    text = re.sub(r'(^|[\s(\[{])"', r'\1“', text)
    text = re.sub(r'"', r'”', text)
    text = re.sub(r"(^|[\s(\[{])'", r'\1‘', text)
    text = re.sub(r"'", r'’', text)
    return text


def add_markdown_runs(
    paragraph,
    text: str,
    font_name: str = "Times New Roman",
    size_pt: float = 12.0,
    base_bold: bool = False,
    base_italic: bool = False,
    color_rgb: tuple[int, int, int] | None = (0, 0, 0),
    use_smart_quotes: bool = False,
) -> list[Run]:
    """Menambahkan teks dengan parsing format markdown inline (**tebal** dan *miring*)."""
    processed_text = convert_to_smart_quotes(text) if use_smart_quotes else text
    tokens = re.split(r"(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*)", processed_text)
    runs = []
    for token in tokens:
        if not token:
            continue
        is_bold = base_bold
        is_italic = base_italic
        run_text = token
        if token.startswith("***") and token.endswith("***") and len(token) >= 6:
            is_bold = True
            is_italic = True
            run_text = token[3:-3]
        elif token.startswith("**") and token.endswith("**") and len(token) >= 4:
            is_bold = True
            run_text = token[2:-2]
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            is_italic = True
            run_text = token[1:-1]

        r = paragraph.add_run(run_text)
        format_run(
            run=r,
            font_name=font_name,
            size_pt=size_pt,
            bold=is_bold,
            italic=is_italic,
            color_rgb=color_rgb,
        )
        runs.append(r)
    return runs
