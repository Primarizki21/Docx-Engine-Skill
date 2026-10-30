from __future__ import annotations

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.text.paragraph import Paragraph


def add_bookmark(paragraph: Paragraph, bookmark_id: int, bookmark_name: str) -> None:
    p = paragraph._p
    tag_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bookmark_id}" w:name="{bookmark_name}"/>')
    tag_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bookmark_id}"/>')
    p.insert(0, tag_start)
    p.append(tag_end)


def add_cross_reference_run(
    paragraph: Paragraph,
    display_text: str,
    bookmark_name: str,
    color_hex: str = "0000FF",
) -> None:
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} w:anchor="{bookmark_name}"/>')
    run_xml = parse_xml(
        f'<w:r {nsdecls("w")}>'
        f"  <w:rPr>"
        f'    <w:color w:val="{color_hex}"/>'
        f'    <w:u w:val="single"/>'
        f"  </w:rPr>"
        f"  <w:t>{display_text}</w:t>"
        f"</w:r>"
    )
    hyperlink.append(run_xml)
    paragraph._p.append(hyperlink)


def add_citation_text(paragraph: Paragraph, citation_label: str) -> None:
    run = paragraph.add_run(f" [{citation_label}]")
    run.font.name = "Times New Roman"
