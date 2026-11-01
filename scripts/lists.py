from __future__ import annotations

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.text.paragraph import Paragraph

try:
    from .core_xml import insert_ordered_p_pr
    from .typography import add_markdown_runs
except ImportError:
    from core_xml import insert_ordered_p_pr
    from typography import add_markdown_runs

LOWER_ROMAN_SET = {"i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "xi", "xii"}
UPPER_ROMAN_SET = {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"}


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


def parse_list_prefix(text: str) -> tuple[str, str | None, str | None]:
    """Mendeteksi prefiks penomoran/list deterministik tanpa regex.
    
    Mengembalikan (clean_text, num_fmt, suffix) atau (text, None, None) jika bukan list.
    """
    stripped = text.strip()
    if not stripped:
        return text, None, None

    parts = stripped.split(None, 1)
    if len(parts) < 2:
        return text, None, None

    token, rest = parts[0], parts[1].strip()

    # 1. Simbol bullet (•, -, *, –)
    if token in {"•", "-", "*", "–"}:
        return rest, "bullet", None

    # 2. Format kurung lengkap: (1), (a), (i), (A), (I)
    if token.startswith("(") and token.endswith(")"):
        core = token[1:-1]
        suffix = "()"
        if core.isdigit():
            return rest, "decimal", suffix
        if core in LOWER_ROMAN_SET:
            return rest, "lowerRoman", suffix
        if core in UPPER_ROMAN_SET:
            return rest, "upperRoman", suffix
        if len(core) == 1 and core.isalpha():
            return rest, "lowerLetter" if core.islower() else "upperLetter", suffix

    # 3. Format berakhiran titik atau kurung tutup: 1., 1), a., a), i., i), A., A)
    if token.endswith(".") or token.endswith(")"):
        suffix = token[-1]
        core = token[:-1]

        # Desimal murni (1., 2)) - bukan subbab seperti 1.1 atau 2.3.1
        if core.isdigit():
            return rest, "decimal", suffix

        if core in LOWER_ROMAN_SET:
            return rest, "lowerRoman", suffix
        if core in UPPER_ROMAN_SET:
            return rest, "upperRoman", suffix

        if len(core) == 1 and core.isalpha():
            return rest, "lowerLetter" if core.islower() else "upperLetter", suffix

    return text, None, None


def strip_manual_numbering_prefix(text: str) -> str:
    """Menghapus awalan penomoran manual secara deterministik tanpa regex."""
    clean_text, _, _ = parse_list_prefix(text)
    return clean_text


def ensure_numbering_scheme(
    document,
    num_fmt: str = "decimal",
    suffix: str = ".",
    restart: bool = False,
    ind_left_dxa: int = 720,
    ind_hanging_dxa: int = 360,
) -> int:
    """Mendaftarkan atau mencari skema penomoran di word/numbering.xml (ECMA-376).
    
    Mendukung format: 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'bullet'.
    Mengembalikan numId yang valid untuk disematkan ke w:numPr.
    """
    numbering_part = document.part.numbering_part
    numbering_el = numbering_part._element

    if num_fmt == "bullet":
        lvl_text = ""
        r_pr_xml = '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
    elif suffix == "()":
        lvl_text = "(%1)"
        r_pr_xml = ""
    elif suffix == ")":
        lvl_text = "%1)"
        r_pr_xml = ""
    else:
        lvl_text = "%1."
        r_pr_xml = ""

    target_abstract_id = None
    for abstract in numbering_el.findall(qn("w:abstractNum")):
        lvl = abstract.find(qn("w:lvl"))
        if lvl is not None:
            fmt_elem = lvl.find(qn("w:numFmt"))
            txt_elem = lvl.find(qn("w:lvlText"))
            if (
                fmt_elem is not None
                and fmt_elem.get(qn("w:val")) == num_fmt
                and txt_elem is not None
                and txt_elem.get(qn("w:val")) == lvl_text
            ):
                target_abstract_id = int(abstract.get(qn("w:abstractNumId")))
                break

    abstract_ids = [int(el.get(qn("w:abstractNumId"))) for el in numbering_el.findall(qn("w:abstractNum"))]
    num_ids = [int(el.get(qn("w:numId"))) for el in numbering_el.findall(qn("w:num"))]

    if target_abstract_id is None:
        target_abstract_id = max(abstract_ids, default=0) + 1
        abstract_xml = (
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{target_abstract_id}">\n'
            f'  <w:multiLevelType w:val="singleLevel"/>\n'
            f'  <w:lvl w:ilvl="0">\n'
            f'    <w:start w:val="1"/>\n'
            f'    <w:numFmt w:val="{num_fmt}"/>\n'
            f'    <w:lvlText w:val="{lvl_text}"/>\n'
            f'    <w:lvlJc w:val="left"/>\n'
            f'    <w:pPr>\n'
            f'      <w:ind w:left="{ind_left_dxa}" w:hanging="{ind_hanging_dxa}"/>\n'
            f'    </w:pPr>\n'
            f'    {r_pr_xml}\n'
            f'  </w:lvl>\n'
            f'</w:abstractNum>'
        )
        abstract_el = parse_xml(abstract_xml)
        first_num = numbering_el.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract_el)
        else:
            numbering_el.append(abstract_el)

    if not restart:
        for num in numbering_el.findall(qn("w:num")):
            ref = num.find(qn("w:abstractNumId"))
            if ref is not None and int(ref.get(qn("w:val"))) == target_abstract_id:
                return int(num.get(qn("w:numId")))

    new_num_id = max(num_ids, default=0) + 1
    num_xml = (
        f'<w:num {nsdecls("w")} w:numId="{new_num_id}">\n'
        f'  <w:abstractNumId w:val="{target_abstract_id}"/>\n'
        f'</w:num>'
    )
    num_el = parse_xml(num_xml)
    numbering_el.append(num_el)
    return new_num_id


def add_numbered_paragraph(
    document,
    text: str,
    num_fmt: str = "auto",
    num_id: int | None = None,
    ilvl: int = 0,
    style: str = "List Paragraph",
    font_name: str = "Times New Roman",
    size_pt: float = 12.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.15,
    ind_left_dxa: int | None = None,
    ind_hanging_dxa: int = 360,
    keep_next: bool = False,
    strip_prefix: bool = True,
    restart: bool = False,
) -> Paragraph:
    """Menambahkan paragraf bernomor Word asli (ListParagraph + w:numPr) dengan perataan justified."""
    clean_text = text
    detected_fmt = None
    suffix = "."

    if strip_prefix:
        parsed_text, detected_fmt, detected_suffix = parse_list_prefix(text)
        if detected_fmt is not None:
            clean_text = parsed_text
            suffix = detected_suffix or "."

    target_fmt = detected_fmt if (num_fmt == "auto" and detected_fmt) else (num_fmt if num_fmt != "auto" else "decimal")

    if num_id is None:
        num_id = ensure_numbering_scheme(
            document,
            num_fmt=target_fmt,
            suffix=suffix,
            restart=restart,
            ind_left_dxa=ind_left_dxa or 720,
            ind_hanging_dxa=ind_hanging_dxa,
        )

    try:
        p = document.add_paragraph(style=style)
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

    left_val = ind_left_dxa or 720
    ind = parse_xml(f'<w:ind {nsdecls("w")} w:left="{left_val}" w:hanging="{ind_hanging_dxa}"/>')
    insert_ordered_p_pr(pPr, ind)

    jc = parse_xml(f'<w:jc {nsdecls("w")} w:val="both"/>')
    insert_ordered_p_pr(pPr, jc)

    add_markdown_runs(
        p,
        clean_text,
        font_name=font_name,
        size_pt=size_pt,
        color_rgb=(0, 0, 0),
    )
    return p


def add_numbered_list(
    document,
    items: list[str],
    num_fmt: str = "auto",
    num_id: int | None = None,
    ilvl: int = 0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.15,
    restart: bool = False,
) -> list[Paragraph]:
    """Menambahkan serangkaian butir daftar bernomor secara berurutan."""
    if not items:
        return []

    resolved_num_id = num_id
    if resolved_num_id is None:
        target_fmt = num_fmt
        suffix = "."
        if num_fmt == "auto":
            _, detected_fmt, detected_suffix = parse_list_prefix(items[0])
            target_fmt = detected_fmt or "decimal"
            suffix = detected_suffix or "."

        resolved_num_id = ensure_numbering_scheme(
            document,
            num_fmt=target_fmt,
            suffix=suffix,
            restart=restart,
        )

    result = []
    for item in items:
        p = add_numbered_paragraph(
            document,
            text=item,
            num_fmt=num_fmt,
            num_id=resolved_num_id,
            ilvl=ilvl,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_spacing=line_spacing,
        )
        result.append(p)
    return result
