from __future__ import annotations

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
    chars = []
    for index, char in enumerate(text):
        if char not in "\"'":
            chars.append(char)
            continue
        previous = text[index - 1] if index else ""
        opening = index == 0 or previous.isspace() or previous in "([{"
        if char == '"':
            chars.append("“" if opening else "”")
        else:
            chars.append("‘" if opening else "’")
    return "".join(chars)


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
    runs = []

    def emit(value: str, bold: bool, italic: bool) -> None:
        if not value:
            return
        run = paragraph.add_run(value)
        format_run(run=run, font_name=font_name, size_pt=size_pt, bold=bold,
                   italic=italic, color_rgb=color_rgb)
        runs.append(run)

    plain_start = index = 0
    while index < len(processed_text):
        line_end = processed_text.find("\n", index)
        if line_end == -1:
            line_end = len(processed_text)
        end = None
        for width, delimiter in ((3, "***"), (2, "**"), (1, "*")):
            if not processed_text.startswith(delimiter, index):
                continue
            closing = processed_text.find(delimiter, index + width, line_end)
            if closing != -1:
                end = closing + width
                break
        if end is None:
            index += 1
            continue
        token = processed_text[index:end]
        is_bold, is_italic, run_text = base_bold, base_italic, token
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
        emit(processed_text[plain_start:index], base_bold, base_italic)
        emit(run_text, is_bold, is_italic)
        index = plain_start = end
    emit(processed_text[plain_start:], base_bold, base_italic)
    return runs
