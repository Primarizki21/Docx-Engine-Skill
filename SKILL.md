---
name: docx-engine
description: "Use this skill whenever you need to create, read, edit, format, or scaffold Microsoft Word (.docx) documents using Python and OOXML. Covers paragraph and character formatting, APA academic tables, numbered and bulleted lists, heading outline levels, automatic TOC/LOT/LOF, SEQ captions, OMML math equations, headers and footers with page numbers, A4 and margin setup, XML inspection, auto-repair, and project scaffolding."
license: MIT
compatibility: claude-code opencode omp
---

# DOCX Engine: Modular Python & OOXML Toolchain

A modular, zero-black-box toolchain for generating and manipulating professional academic and technical Word (.docx) documents in Python.

## Core Philosophy

1. **Python-Native**: Uses `python-docx` and `lxml` under `uv`. No Node.js dependencies.
2. **Modular Micro-Scripts**: Code is split into focused files (< 120 lines each). Agents read only the module needed for the active task.
3. **Read-Modify-Write**: Fully supports opening existing documents and templates, modifying them surgically, and saving them back.
4. **Self-Scaffolding**: Can be ejected into any target project using `scaffold.py`.

---

## Quick Reference Table

| Module | Location | Primary Functions | Typical Use Case |
|---|---|---|---|
| **Sections** | `scripts/sections.py` | `set_a4_page_size`, `set_academic_margins`, `set_two_column_section`, `add_section_break` | Setting up A4 paper, standard 4-3-3-3 cm margins, or multi-column sections |
| **Headings** | `scripts/headings.py` | `add_academic_heading`, `add_chapter_heading`, `set_outline_level` | Adding Heading 1 to 4 with proper outline levels, and creating new chapters with section breaks |
| **Paragraphs** | `scripts/paragraphs.py` | `set_paragraph_spacing`, `set_paragraph_indent`, `set_paragraph_alignment`, `format_academic_paragraph`, `format_preface_paragraph` | Setting 1.5 line spacing for body, single 1.0 spacing for Kata Pengantar, and 1 cm first-line indent |
| **Typography** | `scripts/typography.py` | `format_run`, `add_formatted_run`, `convert_to_smart_quotes` | Font family (Times New Roman), sizes, bold, italic, and typographic quotes |
| **Lists** | `scripts/lists.py` | `add_bullet_item`, `add_numbered_item`, `add_numbered_paragraph`, `add_numbered_list`, `attach_num_pr`, `strip_manual_numbering_prefix` | Adding multi-level bullet points and numbered lists with native Word numPr and markdown runs |
| **Tables** | `scripts/tables.py` | `create_table`, `apply_academic_table_borders`, `apply_grid_table_borders`, `set_table_col_widths_dxa`, `set_table_compact_margins`, `remove_cell_shading`, `format_table_cells` | Building APA or Full-Grid tables with exact DXA widths, zero shading (no color), compact padding, and single line spacing |
| **Media** | `scripts/media.py` | `add_image_with_caption` | Centering images, converting dimensions to EMU (1 in = 914400 EMU), anti-orphan `keep_with_next`, and italic captions |
| **Captions** | `scripts/captions.py` | `add_seq_caption`, `apply_caption_heading`, `detect_caption_type`, `resolve_caption_headings` | Adding SEQ field captions with dynamic heading levels (Gambar H4, Tabel H5, dynamic shifting) |
| **Headers/Footers** | `scripts/headers_footers.py` | `add_page_number_to_footer`, `add_page_number_with_total`, `set_running_header` | Adding PAGE and NUMPAGES field codes and running header text |
| **Diff & Inspection** | `scripts/diff_docx.py` | `run_diff_inspection` | Automated semantic and normalized XML diffing of user Word edits against git baseline or another DOCX |
| **References** | `scripts/references.py` | `add_bookmark`, `add_cross_reference_run`, `add_citation_text` | Bookmarking figures/tables and creating clickable internal hyperlinks |
| **Table of Contents** | `scripts/toc.py` | `add_table_of_contents`, `add_list_of_figures`, `add_list_of_tables` | Generating dynamic field codes for TOC, LOT, and LOF |
| **Equations (Math)** | `scripts/equations.py` | `add_omml_block`, `add_inline_math`, `add_centered_equation`, `make_omml_fraction` | Generating Office Math (OMML) equations for formulas and metrics |
| **Settings** | `scripts/settings.py` | `enable_update_fields_in_unpacked`, `enable_update_fields_in_doc` | Injecting updateFields flag into settings.xml so Word auto-refreshes TOC and SEQ |
| **Package** | `scripts/package.py` | `unpack_docx`, `pack_docx`, `parse_xml_part`, `save_xml_part`, `validate_xml_parts` | Unpacking zip, editing raw XML, auto-repairing whitespace, and repacking |
| **Scaffold** | `scripts/scaffold.py` | `scaffold_engine` | Copying the entire engine into any target project directory |
| **Replace & Templates** | `scripts/replace.py` | `replace_exact_text`, `replace_placeholders`, `replace_exact_text_in_paragraph` | Safe, run-preserving exact match string and template placeholder replacement |

---

## How to Use in Any Target Project

### Option 1: Eject / Scaffold into Project (Recommended for Complete Independence)
To copy the entire engine directly into a target repository so the project owns its code:
```bash
uv run python ~/.omp/agent/skills/docx-engine/scripts/scaffold.py --target ./src/engine
```
Then install dependencies in the target project:
```bash
uv add python-docx lxml
```
And import directly:
```python
from engine import tables, headings, sections, toc
```

### Option 2: CLI Microservice Execution
To run operations directly from the skill folder without touching target source files:
```bash
# Unpack a document for inspection
uv run python ~/.omp/agent/skills/docx-engine/scripts/package.py unpack ./laporan.docx ./unpacked

# Validate XML parts
uv run python ~/.omp/agent/skills/docx-engine/scripts/package.py validate ./unpacked

# Repack with auto-repair
uv run python ~/.omp/agent/skills/docx-engine/scripts/package.py pack ./unpacked ./laporan_fixed.docx

# Inspect user manual edits against git HEAD
uv run python ~/.omp/agent/skills/docx-engine/scripts/diff_docx.py ./laporan.docx --git

# Inspect user manual edits between two DOCX files with clean XML diff
uv run python ~/.omp/agent/skills/docx-engine/scripts/diff_docx.py ./laporan_base.docx ./laporan_edited.docx --mode xml
```

## Inviolable Robustness Rules

### Precedence Hierarchy: Sensible Defaults vs. Explicit User Override
- **User Explicit Request (Highest Precedence)**: If the user explicitly specifies a custom style or format (e.g., APA table style instead of grid, bold heading 3, specific font or size, custom margins), agents MUST prioritize the user's explicit instructions and adapt function parameters accordingly without forcing default values.
- **Sensible Defaults (Automatic Fallback)**: When no specific formatting instructions are provided by the user, agents MUST automatically apply the engine's built-in defaults (full grid borders, 10 pt table font, 0 cm cell padding, bold Heading 1 & 2, regular non-bold Heading 3, zero space before/after, <= 45 char limit).
- **Technical Invariants (Always Enforced)**: Core OOXML file integrity rules (DXA dual-widths, `<w:pPr>` element order, whitespace preservation, dynamic updateFields, and safe exact-match replacement) must never be violated regardless of user styling requests.

1. **Dual-Width DXA for Tables**:
   Always set table width AND individual cell widths in DXA units (1 inch = 1440 DXA). Never use percentages. Cell widths must sum to the table width.
2. **Auto-Update Fields via Settings**:
   When using dynamic fields (TOC, SEQ captions, PAGE numbers), always call `enable_update_fields_in_unpacked` or `enable_update_fields_in_doc`. This ensures Word refreshes all fields when the user opens the document.
3. **Element Ordering in `<w:pPr>`**:
   Never append child elements randomly into paragraph properties. Use `insert_ordered_p_pr()` from `core_xml.py` to respect the ECMA-376 schema sequence (`pStyle`, `numPr`, `spacing`, `ind`, `jc`, `rPr`).
4. **Preserve Whitespace**:
   Any `<w:t>` containing leading or trailing spaces must include `xml:space="preserve"`. The `package.py` packer automatically repairs this.
5. **EMU Scaling for Media**:
   Always convert image dimensions to EMUs (English Metric Units: 914,400 EMUs = 1 inch, 360,000 EMUs = 1 cm).
6. **Strict Transparent Table Headers & Default Full Grid Config**:
   Never inject `<w:shd w:fill="FFFFFF"/>` or `clear` unless a solid background is explicitly needed. Omit `<w:shd>` entirely so that background letterheads, kop frames, and watermark images are not masked. Enforce table defaults: solid 0.5 pt grid borders on all sides (`sz=4` via `create_table()` default), font size 10 pt, explicit black font (`RGBColor(0, 0, 0)`), bold header row, 0 cm margin/padding on all sides, and 0 pt paragraph spacing/indent via `format_table_cells()`.
7. **Context-Aware Line Spacing**:
   Use 1.5 line spacing (`line="360" lineRule="auto"`) for academic chapter body text, but enforce single 1.0 line spacing (`line="240" lineRule="auto"`) for Kata Pengantar (`format_preface_paragraph`), table cells, and compact TOC entries.
8. **Anti-Orphan Pagination Guards**:
   Always enforce `keep_with_next = True` on headings and image paragraphs so they never separate from their succeeding body paragraph or caption.
9. **New Chapters (BAB) MUST Use Section Break (Next Page)**:
   Always separate preliminary pages (Kata Pengantar, Daftar Isi) from BAB I, and each subsequent chapter (BAB II, BAB III, dst.) using Section Break (Next Page) via `add_chapter_heading()` or `add_section_break(doc, start_type="new_page")`. NEVER use simple page breaks (`doc.add_page_break()` or Ctrl+Enter / `<w:br w:type="page"/>`) to divide chapters.
10. **Batas Karakter Subbab, Subsubbab, dan Caption (<= 45 Karakter)**:
    Subbab (Heading 2), subsubbab (Heading 3), serta caption gambar, tabel, dan lampiran TIDAK BOLEH melebihi 45 karakter agar judul padat dan tidak membungkus berlebihan pada TOC/LOT/LOF.
11. **Hierarki Tebal (Bold) Heading**:
    Heading 1 (BAB) dan Heading 2 (Subbab) wajib dicetak TEBAL (`run.bold = True`). Heading 3 (Subsubbab) dan seterusnya DILARANG BOLD (`run.bold = False`, `<w:b w:val="0"/>`), 12 pt Regular, warna teks hitam pekat.
12. **Nol Spasi Sebelum dan Sesudah Seluruh Paragraf dan Heading (Zero Space Before & After)**:
    Teks, paragraf akademik, preface, serta seluruh tingkatan Heading (Heading 1 / BAB, Heading 2 / Subbab, Heading 3 / Subsubbab, dst.) wajib 0 pt space before dan 0 pt space after (`space_before=Pt(0)`, `space_after=Pt(0)` / 0 cm). Jarak antarbaris dikontrol penuh oleh line spacing (1.5 atau 1.0).
13. **Exact-Match Template Replacement (Larangan Penggunaan Regex)**:
    Saat memodifikasi dokumen atau mengisi template, gunakan selalu `replace_exact_text()` atau `replace_placeholders()` dari `scripts/replace.py`. DILARANG menggunakan regex `re.sub` pada berkas .docx karena memecah run formatting dan rentan korupsi XML.
14. **Wajib Menggunakan Native Numbered List (Larangan Penomoran Manual)**:
    Dilarang keras mengetik manual penomoran pada paragraf biasa (`"1. Teks"`). Seluruh butir daftar berurutan wajib memanggil `add_numbered_paragraph()` atau `add_numbered_list()` dari `scripts/lists.py` untuk mengaktifkan native Word `<w:numPr>` dan menjamin *hanging indent* lurus rapi.
15. **Standar Format Caption: Murni Italic 10 pt Tanpa Bold & Tanpa Narasi Penjelas**:
    Seluruh caption (Gambar, Tabel, Lampiran) wajib satu baris ringkas (<= 45 karakter), murni Italic (`run.italic = True`), 10 pt, dan TANPA BOLD (`run.bold = False`). DILARANG menyisipkan break line (`\n`) atau menambahkan paragraf narasi penjelas di bawah caption.
16. **Hierarki Heading Caption Dinamis (Gambar = Heading 4, Tabel = Heading 5, Dynamic Shifting)**:
    Caption Gambar default `Heading 4` (outlineLvl 3) dan caption Tabel default `Heading 5` (outlineLvl 4). Jika dokumen memuat body heading level 4 atau lebih tinggi ($N \ge 4$), level heading caption bergeser secara dinamis ($\text{Gambar} = \max(4, N + 1)$, $\text{Tabel} = \text{Gambar} + 1$). Pemanggilan `add_seq_caption()` otomatis menerapkan heading level dinamis ini, dan `resolve_caption_headings(doc)` menyinkronkan seluruh dokumen.
---

## Standard Recipes

### Recipe 1: Standard Academic Report Document Setup (Indonesia Format)
```python
import docx
from sections import set_a4_page_size, set_academic_margins, add_section_break
from headers_footers import add_page_number_to_footer, set_running_header
from headings import add_academic_heading, add_chapter_heading
from paragraphs import format_academic_paragraph, format_preface_paragraph
from typography import add_formatted_run

doc = docx.Document()
sec = doc.sections[0]

# A4 paper with standard 4-3-3-3 cm margins
set_a4_page_size(sec)
set_academic_margins(sec, top_cm=3.0, bottom_cm=3.0, left_cm=4.0, right_cm=3.0)
set_running_header(sec, "Laporan Riset Akademik")
add_page_number_to_footer(sec, align="center")

# Preliminaries: Kata Pengantar with single line spacing 1.0
add_academic_heading(doc, "KATA PENGANTAR", level=1)
p_preface = doc.add_paragraph()
format_preface_paragraph(p_preface)
add_formatted_run(p_preface, "Puji syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa atas selesainya laporan riset ini.")

# Chapter 1: separated by Section Break (Next Page) automatically
add_chapter_heading(doc, "BAB I PENDAHULUAN")
p = doc.add_paragraph()
format_academic_paragraph(p)
add_formatted_run(p, "Chatbot akademik FTMM COMPASS mengimplementasikan arsitektur hybrid RAG-CAG.")
doc.save("laporan.docx")
```

### Recipe 2: Adding an APA Scientific Table with Shading
```python
from tables import create_table, apply_academic_table_borders, set_cell_shading, set_cell_margins

table = create_table(doc, rows=3, cols=2, col_widths_inches=[3.5, 2.5])
table.rows[0].cells[0].paragraphs[0].text = "Metrik Evaluasi"
table.rows[0].cells[1].paragraphs[0].text = "Nilai Pengujian"

table.rows[1].cells[0].paragraphs[0].text = "P50 Latency"
table.rows[1].cells[1].paragraphs[0].text = "340 ms"

table.rows[2].cells[0].paragraphs[0].text = "Akurasi Klasifikasi"
table.rows[2].cells[1].paragraphs[0].text = "94.2%"

apply_academic_table_borders(table)
set_cell_shading(table.rows[0].cells[0], "F0F0F0")
set_cell_shading(table.rows[0].cells[1], "F0F0F0")
```

### Recipe 3: Sequential Captions and Cross-References
```python
from captions import add_seq_caption
from references import add_cross_reference_run
# Add caption with dynamic SEQ field and bookmark (pure italic, 10 pt, no bold, single-line)
add_seq_caption(doc, label="Tabel", description="Komparasi Metrik Latensi", bookmark_name="tbl_latensi")

# Refer to it in body text
p_body = doc.add_paragraph("Hasil evaluasi komparasi dapat dilihat pada ")
add_cross_reference_run(p_body, "Tabel 1", bookmark_name="tbl_latensi")
```

### Recipe 4: Full Grid Table with Exact DXA Widths & Transparent Header
```python
from tables import create_table, apply_grid_table_borders, set_table_col_widths_dxa, set_table_compact_margins, format_table_cells

headers = ["Tipologi Klaster", "Cakupan Wilayah", "Rata-rata TPT", "Informalitas", "Rata-rata Upah"]
rows = [
    ["Klaster 1: Urban-Industrial Trap", "8 Provinsi", "6,58%", "43,2%", "Rp3,85 Juta"],
    ["Klaster 2: Agraris-Informal Shelter", "18 Provinsi", "2,82%", "75,4%", "Rp2,62 Juta"],
]

table = create_table(doc, rows=len(rows) + 1, cols=len(headers))
set_table_col_widths_dxa(table, [1746, 2360, 1418, 1275, 1418])
apply_grid_table_borders(table, sz=4)  # 0.5 pt solid grid borders
set_table_compact_margins(table, left_dxa=0, right_dxa=0)
format_table_cells(table, font_size_pt=10.0, line_spacing=1.0, ensure_no_color=True)
```

### Recipe 5: Native Word Numbered Paragraphs with Inline Markdown
```python
from lists import add_numbered_paragraph

# Adds a justified paragraph with ListParagraph style, w:numPr, and parsed italic/bold runs
add_numbered_paragraph(
    doc,
    text="Krisis Ketenagakerjaan: Berisiko meninggalkan luka permanen (*scarring effect*) pada kapasitas kerja.",
    num_id=3,
    ilvl=0,
    line_spacing=1.15,
)
```

### Recipe 6: Inspecting & Reverse-Syncing User Word Edits
```bash
# User edits output/laporan.docx in Word, then inspect differences against git:
uv run python ~/.omp/agent/skills/docx-engine/scripts/diff_docx.py output/laporan.docx --git
```

### Recipe 7: Exact-Match Template Replacement & Placeholder Filling
```python
import docx
from replace import replace_exact_text, replace_placeholders

doc = docx.Document("template.docx")

# Replace specific text across the entire document
replace_exact_text(doc, "[JUDUL_LAPORAN]", "Sistem Klasifikasi Citra")

# Bulk placeholder replacement across paragraphs, tables, and headers/footers
mapping = {
    "{{NAMA}}": "Prima Rizki",
    "{{NIM}}": "164221000",
    "{{TANGGAL}}": "30 Oktober 2026",
}
replace_placeholders(doc, mapping)
doc.save("laporan_filled.docx")
```
