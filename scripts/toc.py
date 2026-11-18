from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
_SPECS = (("DAFTAR ISI", "Static TOC Title", "Static TOC Entry"),
          ("DAFTAR GAMBAR", "Static LOF Title", "Static LOF Entry"),
          ("DAFTAR TABEL", "Static LOT Title", "Static LOT Entry"))
_ENTRY = tuple(spec[2] for spec in _SPECS)

def _level(paragraph) -> int | None:
    name = paragraph.style.name if paragraph.style is not None else ""
    return int(name[8:]) if name.startswith("Heading ") and name[8:].isdigit() else None

def collect_toc_entries(document) -> list[tuple[int, str]]:
    return [(level, p.text.strip()) for p in document.paragraphs
            if (level := _level(p)) is not None and p.text.strip()]

def collect_caption_entries(document, kind: str = "Gambar") -> list[str]:
    prefix = f"{kind} "
    found = []
    for p in document.paragraphs:
        text = p.text.strip()
        if p.style.name in _ENTRY or "\t" in text or not text.lower().startswith(prefix.lower()):
            continue
        if text[len(prefix):].strip().partition(" ")[0][:1].isdigit():
            found.append(text)
    return found

def _style(document, name):
    return document.styles[name] if name in [s.name for s in document.styles] else (
        document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH))

def _entry(paragraph, text: str, page=None) -> None:
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    for value in [text] + (["\t", str(page)] if page is not None else []):
        run = paragraph.add_run(value)
        run.font.name, run.font.size, run.font.bold = "Times New Roman", Pt(10), False
        run.font.color.rgb = RGBColor(0, 0, 0)
    fmt = paragraph.paragraph_format
    fmt.space_before, fmt.space_after, fmt.line_spacing = Pt(0), Pt(0), 1.0
    fmt.first_line_indent = Pt(0)
    fmt.tab_stops.clear_all(); fmt.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)

def _add(document, spec, title, entries, pages):
    last = None
    if title:
        last = document.add_paragraph()
        last.style = _style(document, spec[1])
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = last.add_run(title)
        run.bold, run.font.name, run.font.size = True, "Times New Roman", Pt(14)
        last.paragraph_format.space_before, last.paragraph_format.space_after = Pt(12), Pt(12)
    for index, text in enumerate(entries):
        last = document.add_paragraph()
        last.style = _style(document, spec[2])
        page = pages.get(text) if isinstance(pages, dict) else (pages[index] if pages and index < len(pages) else None)
        _entry(last, text, page)
    return last

def add_table_of_contents(document, title="DAFTAR ISI", heading_levels="1-3", page_numbers=None):
    first, sep, last = str(heading_levels).partition("-")
    levels = range(int(first), int(last) + 1) if sep and first.isdigit() and last.isdigit() else {int(heading_levels)}
    return _add(document, _SPECS[0], title, [t for n, t in collect_toc_entries(document) if n in levels], page_numbers)

def add_list_of_figures(document, title="DAFTAR GAMBAR", seq_name="Gambar", page_numbers=None):
    return _add(document, _SPECS[1], title, collect_caption_entries(document, seq_name), page_numbers)

def add_list_of_tables(document, title="DAFTAR TABEL", seq_name="Tabel", page_numbers=None):
    return _add(document, _SPECS[2], title, collect_caption_entries(document, seq_name), page_numbers)

def render_page_numbers(docx_path) -> dict[str, int]:
    soffice, pdftotext = shutil.which("soffice"), shutil.which("pdftotext")
    missing = [name for name, found in (("soffice", soffice), ("pdftotext", pdftotext)) if not found]
    if missing:
        raise RuntimeError("Executable tidak tersedia: " + ", ".join(missing))
    source = Path(docx_path)
    with tempfile.TemporaryDirectory(prefix="docx-toc-") as temp:
        work = Path(temp)
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(work), str(source)], check=True, capture_output=True)
        pdfs = sorted(work.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError(f"LibreOffice tidak menghasilkan PDF untuk {source}")
        out = work / "pages.txt"
        subprocess.run([pdftotext, "-layout", str(pdfs[0]), str(out)], check=True, capture_output=True)
        pages = [p for p in out.read_text(encoding="utf-8", errors="replace").split("\f") if p.strip()]
    document = docx.Document(str(source))
    texts = [text for _, text in collect_toc_entries(document)]
    texts += collect_caption_entries(document, "Gambar") + collect_caption_entries(document, "Tabel")
    flat = [" ".join(page.split()) for page in pages]
    skip = {i for i, page in enumerate(flat) if any(spec[0] in page for spec in _SPECS)}
    mapping: dict[str, int] = {}
    for text in dict.fromkeys(texts):
        needle = " ".join(text.split())
        for index, page in enumerate(flat):
            if index not in skip and needle in page:
                mapping[text] = index + 1
                break
    return mapping

def build_static_lists(docx_path, output_docx) -> dict[str, int]:
    before = render_page_numbers(docx_path)
    document = docx.Document(str(docx_path))
    for paragraph in document.paragraphs:
        if paragraph.style.name in _ENTRY:
            text = paragraph.text.partition("\t")[0].strip()
            _entry(paragraph, text, before.get(text))
    document.save(str(output_docx))
    after = render_page_numbers(output_docx)
    if after != before:
        raise RuntimeError("Nomor halaman bergeser setelah daftar statis diisi")
    return after
