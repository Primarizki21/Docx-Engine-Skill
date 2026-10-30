from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt
from docx.text.paragraph import Paragraph

M_DECL = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def add_omml_block(paragraph: Paragraph, omml_inner_xml: str) -> Paragraph:
    xml_str = f"<m:oMathPara {M_DECL}><m:oMath>{omml_inner_xml}</m:oMath></m:oMathPara>"
    elem = parse_xml(xml_str)
    paragraph._p.append(elem)
    return paragraph


def add_inline_math(paragraph: Paragraph, omml_inner_xml: str) -> Paragraph:
    xml_str = f"<m:oMath {M_DECL}>{omml_inner_xml}</m:oMath>"
    elem = parse_xml(xml_str)
    paragraph._p.append(elem)
    return paragraph


def make_omml_text(text: str) -> str:
    return f"<m:r><m:t>{text}</m:t></m:r>"


def make_omml_fraction(num: str, den: str) -> str:
    return f"<m:f><m:num>{make_omml_text(num)}</m:num><m:den>{make_omml_text(den)}</m:den></m:f>"


def make_omml_subscript(base: str, sub: str) -> str:
    return f"<m:sSub><m:e>{make_omml_text(base)}</m:e><m:sub>{make_omml_text(sub)}</m:sub></m:sSub>"


def make_omml_superscript(base: str, sup: str) -> str:
    return f"<m:sSup><m:e>{make_omml_text(base)}</m:e><m:sup>{make_omml_text(sup)}</m:sup></m:sSup>"


def add_centered_equation(document, omml_inner_xml: str) -> Paragraph:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    add_omml_block(p, omml_inner_xml)
    return p
