from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
try:
    from .core_xml import insert_ordered_p_pr
except ImportError:
    from core_xml import insert_ordered_p_pr


def add_bullet_item(document, text: str, level: int = 0) -> Paragraph:
    style_name = "List Bullet" if level == 0 else f"List Bullet {min(level + 1, 3)}"
    try:
        p = document.add_paragraph(text, style=style_name)
    except KeyError:
        p = document.add_paragraph(text, style="List Bullet")
    return p


def add_numbered_item(document, text: str, level: int = 0) -> Paragraph:
    style_name = "List Number" if level == 0 else f"List Number {min(level + 1, 3)}"
    try:
        p = document.add_paragraph(text, style=style_name)
    except KeyError:
        p = document.add_paragraph(text, style="List Number")
    return p


def attach_num_pr(paragraph: Paragraph, num_id: int, ilvl: int = 0) -> Paragraph:
    pPr = paragraph._p.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = OxmlElement("w:numPr")
        insert_ordered_p_pr(pPr, numPr)

    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), str(ilvl))
    numPr.append(ilvl_elem)

    num_id_elem = OxmlElement("w:numId")
    num_id_elem.set(qn("w:val"), str(num_id))
    numPr.append(num_id_elem)
    return paragraph

import re
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor


def strip_manual_numbering_prefix(text: str) -> str:
    """Menghapus awalan penomoran manual seperti '1. ', 'a. ', '• ' agar tidak dobel nomor."""
    return re.sub(r"^(\d+[\.\)]\s*|[A-Za-z][\.\)]\s*|[•\-\*]\s*)", "", text).strip()


def add_numbered_paragraph(
    document,
    text: str,
    num_id: int = 1,
    ilvl: int = 0,
    style: str = "List Paragraph",
    font_name: str = "Times New Roman",
    size_pt: float = 12.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.15,
    ind_left_dxa: int | None = None,
    ind_first_line_dxa: int | None = None,
    keep_next: bool = False,
    strip_prefix: bool = True,
) -> Paragraph:
    """Menambahkan paragraf bernomor Word asli (ListParagraph + w:numPr) dengan perataan justified."""
    clean_text = strip_manual_numbering_prefix(text) if strip_prefix else text

    try:
        p = document.add_paragraph(style=style)
    except KeyError:
        try:
            p = document.add_paragraph(style="ListParagraph")
        except KeyError:
            p = document.add_paragraph()
            pPr_init = p._p.get_or_add_pPr()
            pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="ListParagraph"/>')
            insert_ordered_p_pr(pPr_init, pStyle)

    pPr = p._p.get_or_add_pPr()

    if keep_next:
        keepNext = parse_xml(f'<w:keepNext {nsdecls("w")}/>')
        insert_ordered_p_pr(pPr, keepNext)

    numPr = parse_xml(
        f'<w:numPr {nsdecls("w")}>\n'
        f'  <w:ilvl w:val="{ilvl}"/>\n'
        f'  <w:numId w:val="{num_id}"/>\n'
        f'</w:numPr>'
    )
    insert_ordered_p_pr(pPr, numPr)

    spacing_dxa = int(line_spacing * 240) if line_spacing > 0 else 240
    spacing = parse_xml(
        f'<w:spacing {nsdecls("w")} w:before="{int(space_before_pt * 20)}" w:after="{int(space_after_pt * 20)}" w:line="{spacing_dxa}" w:lineRule="auto"/>'
    )
    insert_ordered_p_pr(pPr, spacing)

    if ind_left_dxa is not None:
        first_line_attr = f' w:firstLine="{ind_first_line_dxa}"' if ind_first_line_dxa is not None else ""
        ind = parse_xml(f'<w:ind {nsdecls("w")} w:left="{ind_left_dxa}"{first_line_attr}/>')
        insert_ordered_p_pr(pPr, ind)

    jc = parse_xml(f'<w:jc {nsdecls("w")} w:val="both"/>')
    insert_ordered_p_pr(pPr, jc)

    # Parsing markdown inline (**tebal** dan *miring*)
    tokens = re.split(r"(\*\*.*?\*\*|\*.*?\*)", clean_text)
    for token in tokens:
        if not token:
            continue
        run_bold = False
        run_italic = False
        run_text = token
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            run_bold = True
            run_text = token[2:-2]
        elif token.startswith("*") and token.endswith("*") and len(token) >= 2:
            run_italic = True
            run_text = token[1:-1]

        run = p.add_run(run_text)
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = run_bold
        run.italic = run_italic
        run.font.color.rgb = RGBColor(0, 0, 0)

    return p


def add_numbered_list(
    document,
    items: list[str],
    num_id: int = 1,
    ilvl: int = 0,
    space_before_pt: float = 3.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.15,
) -> list[Paragraph]:
    """Menambahkan serangkaian butir daftar bernomor secara berurutan."""
    result = []
    for item in items:
        p = add_numbered_paragraph(
            document,
            text=item,
            num_id=num_id,
            ilvl=ilvl,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_spacing=line_spacing,
        )
        result.append(p)
    return result
