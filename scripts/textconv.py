from __future__ import annotations

import difflib
import zipfile
from pathlib import Path

from lxml import etree

try:
    from .core_xml import qn
except ImportError:
    from core_xml import qn


_PARSER = etree.XMLParser(resolve_entities=False)


def _para_text(p: etree._Element) -> str:
    """Gabungkan seluruh w:t dalam satu paragraf (stabil terhadap perpecahan run)."""
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _cell_text(tc: etree._Element) -> str:
    """Gabungkan paragraf di dalam satu sel tabel."""
    return "".join(_para_text(p) for p in tc.iter(qn("w:p")))


def _level(p: etree._Element) -> int | None:
    """Level 1-9 dari outlineLvl (0-8) atau gaya Heading N."""
    p_pr = p.find(qn("w:pPr"))
    if p_pr is None:
        return None
    outline = p_pr.find(qn("w:outlineLvl"))
    if outline is not None:
        value = outline.get(qn("w:val"), "")
        if value.isdigit():
            return int(value) + 1
    style = p_pr.find(qn("w:pStyle"))
    if style is not None:
        value = style.get(qn("w:val"), "")
        if value.startswith("Heading"):
            digits = value[len("Heading"):].strip()
            if digits.isdigit():
                return int(digits)
    return None


def _classify(p: etree._Element, text: str) -> str:
    level = _level(p)
    if level in (4, 5) and text.startswith(("Gambar ", "Tabel ")):
        return f"C{level}"
    if level in (1, 2, 3):
        return f"H{level}"
    return "P"


def _has_sect_pr(p: etree._Element) -> bool:
    p_pr = p.find(qn("w:pPr"))
    return p_pr is not None and p_pr.find(qn("w:sectPr")) is not None


def canonical_lines(docx_path) -> list[str]:
    """Hasilkan baris `JENIS|teks` dari word/document.xml secara deterministik."""
    with zipfile.ZipFile(docx_path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"), _PARSER)
    body = root.find(qn("w:body"))
    if body is None:
        return []

    lines: list[str] = []
    for child in body:
        tag = etree.QName(child).localname
        if tag == "p":
            text = _para_text(child)
            lines.append(f"{_classify(child, text)}|{text}")
            if _has_sect_pr(child):
                lines.append("SECT|")
        elif tag == "tbl":
            for tr in child.findall(qn("w:tr")):
                cells = [_cell_text(tc) for tc in tr.findall(qn("w:tc"))]
                lines.append("T|" + "\t".join(cells))
        elif tag == "sectPr":
            lines.append("SECT|")
    return lines

def _as_lines(value) -> list[str]:
    if isinstance(value, Path) or (isinstance(value, str) and Path(value).is_file()):
        return canonical_lines(value)
    if isinstance(value, str):
        return value.splitlines()
    return list(value)


def canonical_diff(a, b) -> list[str]:
    """Diff unified dari dua daftar baris kanonik (atau dua path docx)."""
    return list(difflib.unified_diff(_as_lines(a), _as_lines(b), fromfile="a", tofile="b", lineterm=""))

def write_canonical(docx_path, out) -> Path:
    """Tulis baris kanonik ke berkas teks dan kembalikan Path tujuannya."""
    lines = canonical_lines(docx_path)
    target = Path(out)
    target.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
    return target
