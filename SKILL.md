---
name: docx-engine
description: "Use this skill whenever you need to create, read, edit, format, or scaffold Microsoft Word (.docx) documents using Python and OOXML. Covers paragraph and character formatting, APA academic tables, numbered and bulleted lists, heading outline levels, static TOC/LOT/LOF with two-pass PDF page numbers, hierarchical Gambar/Tabel captions, OMML math equations, headers and footers with page numbers, A4 and margin setup, deterministic packaging, canonical text diff, lexicon and acronym consistency, fail-fast QA verification, XML inspection, auto-repair, and project scaffolding."
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
| **Sections** | `scripts/sections.py` | `set_a4_page_size`, `set_academic_margins`, `set_two_column_section`, `add_section_break` | Setting up A4 paper with the default 4-4-3-3 cm margins, or multi-column sections |
| **Headings** | `scripts/headings.py` | `add_academic_heading`, `add_chapter_heading`, `set_outline_level` | Adding Heading 1 to 5 with proper outline levels; Heading 1 `BAB ...` sequence drives hierarchical caption numbers |
| **Paragraphs** | `scripts/paragraphs.py` | `set_paragraph_spacing`, `set_paragraph_indent`, `set_paragraph_alignment`, `format_academic_paragraph`, `format_preface_paragraph` | 1.5 line spacing and 720 twips first-line indent (0.5 in / 1.27 cm) for body; single 1.0 spacing and 1 cm indent for Kata Pengantar |
| **Typography** | `scripts/typography.py` | `format_run`, `add_formatted_run`, `convert_to_smart_quotes` | Font family (Times New Roman), sizes, bold, italic, and typographic quotes |
| **Lists** | `scripts/lists.py` | `add_bullet_item`, `add_numbered_item`, `add_numbered_paragraph`, `add_numbered_list`, `ensure_numbering_scheme`, `parse_list_prefix`, `strip_manual_numbering_prefix` | Multi-format lists (decimal, lowerLetter a., upperLetter A., lowerRoman i., bullet) with native Word numPr, deterministic parsing without regex, and markdown runs |
| **Tables** | `scripts/tables.py` | `create_table(vertical_center=False)`, `format_table_cells(vertical_center=False)`, `apply_academic_table_borders`, `apply_grid_table_borders`, `set_table_col_widths_dxa`, `set_table_compact_margins`, `remove_cell_shading`, `set_cell_vertical_alignment`, `set_table_vertical_alignment` | APA or Full-Grid tables with exact DXA widths, zero shading, compact padding, single line spacing; body cells `vAlign=top`, header row `vAlign=center` + `jc=center`; `vertical_center=True` only on explicit request |
| **Media** | `scripts/media.py` | `add_image_with_caption` | Centering images, converting dimensions to EMU (1 in = 914400 EMU), anti-orphan `keep_with_next`, one image one caption, compact monochrome layout with print font >= 9 pt |
| **Captions** | `scripts/captions.py` | `add_seq_caption`, `apply_caption_heading`, `detect_caption_type` | Hierarchical `Gambar X.Y` / `Tabel X.Y` numbers from BAB order; Gambar Heading 4, Tabel Heading 5; italic only for lexicon terms; <=45 char fail-fast; no SEQ field |
| **Headers/Footers** | `scripts/headers_footers.py` | `add_page_number_to_footer`, `add_page_number_with_total`, `set_running_header` | Adding PAGE and NUMPAGES field codes and running header text |
| **Diff & Inspection** | `scripts/diff_docx.py` | `run_diff_inspection`, `inspect_annotations_diff` | Semantic and annotation diffing against git baseline or standalone DOCX, fed by canonical `JENIS|teks` lines from `textconv.py` |
| **Annotations** | `scripts/annotations.py` | `extract_document_annotations`, `format_annotations_report` | Extracting highlighted text, custom font colors, Word comments, and track changes (ins/del) with breadcrumb context |
| **References** | `scripts/references.py` | `add_bookmark`, `add_cross_reference_run`, `add_citation_text` | Bookmarking figures/tables and creating clickable internal hyperlinks |
| **Table of Contents** | `scripts/toc.py` | `add_table_of_contents`, `add_list_of_figures`, `add_list_of_tables` | Static TOC/LOT/LOF from headings and captions; page numbers from two LibreOffice PDF renders; explicit error when `soffice`/`pdftotext` is absent |
| **Equations (Math)** | `scripts/equations.py` | `add_omml_block`, `add_inline_math`, `add_centered_equation`, `make_omml_fraction` | Generating Office Math (OMML) equations for formulas and metrics |
| **Settings** | `scripts/settings.py` | `enable_update_fields_in_unpacked`, `enable_update_fields_in_doc` | Injecting updateFields flag into settings.xml so Word refreshes PAGE/NUMPAGES only; TOC is static |
| **Text Conversion** | `scripts/textconv.py` | `canonical_lines` | Producing stable `JENIS|teks` canonical lines for regression diff and manual-drift inspection |
| **Mutation Guard** | `scripts/mutation_guard.py` | scoped replace assertions | Enforcing `expected=1`, before/after assertions, and idempotency for every scoped text mutation |
| **Lexicon** | `scripts/lexicon.py`, `data/lexicon.json` | `italic_segments`, banned forms, acronym expansions | Single source of truth for italic terms, forbidden variants, and approved acronyms |
| **Acronyms** | `scripts/acronyms.py` | acronym expansion checks | Ensuring acronyms appear only after their full form in the body text |
| **Sentence Metrics** | `scripts/sentence_metrics.py` | narrative sentence word counts | Enforcing <=20 content words per narrative sentence, excluding parenthetical citations |
| **Citation QA** | `scripts/citation_qa.py` | citation placement checks | Ensuring citations attach to claim sentences |
| **Reference QA** | `scripts/reference_qa.py` | reference consistency checks | Validating reference entries against in-text citations |
| **Verification Runner** | `scripts/verify.py` | fail-fast QA runner | Running all QA gates with a 35-second budget via `docx-engine-verify` |
| **Regression Diff** | `scripts/regression_diff.py` | canonical regression comparison | Flagging unintended changes against the approved baseline via `docx-engine-regression` |
| **Manual Drift** | `scripts/manual_drift.py` | manual edit drift report | Detecting user edits outside allowed scope via `docx-engine-drift` |
| **Data** | `data/` | `lexicon.json` | Shared JSON store for lexicon terms, banned forms, and acronym expansions |
| **Package** | `scripts/package.py` | `unpack_docx`, `pack_docx`, `parse_xml_part`, `save_xml_part`, `validate_xml_parts` | Unpacking zip, editing raw XML, auto-repairing whitespace, and deterministic repack (sorted entries, 1980-01-01 timestamp, `external_attr=0`, skip write when bytes are identical) |
| **Scaffold** | `scripts/scaffold.py` | `scaffold_engine` | Copying the entire engine into any target project directory |
| **Replace & Templates** | `scripts/replace.py` | `replace_exact_text`, `replace_placeholders`, `replace_exact_text_in_paragraph` | Scoped, run-preserving exact-match replacement with `expected=1`, before/after assertions, idempotency, and no regex |

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

# Run all QA gates (fail-fast, 35 s budget)
docx-engine-verify ./laporan.docx

# Emit stable canonical JENIS|teks lines
docx-engine-canon ./laporan.docx > laporan.canon

# Compare against the approved canonical baseline
docx-engine-regression laporan.canon baseline.canon

# Report manual edits outside allowed scope
docx-engine-drift ./laporan.docx
```

## Inviolable Robustness Rules

### Precedence Hierarchy: Sensible Defaults vs. Explicit User Override
- **User Explicit Request (Highest Precedence)**: If the user explicitly specifies a custom style or format (e.g., APA table style instead of grid, bold heading 3, specific font or size, custom margins), agents MUST prioritize the user's explicit instructions and adapt function parameters accordingly without forcing default values.
- **Sensible Defaults (Automatic Fallback)**: When no specific formatting instructions are provided by the user, agents MUST automatically apply the engine's built-in defaults (full grid borders, 10 pt table font, 0 cm cell padding, bold Heading 1 & 2, regular non-bold Heading 3, zero space before/after, <= 45 char limit).
- **Technical Invariants (Always Enforced)**: Core OOXML file integrity rules (DXA dual-widths, `<w:pPr>` element order, whitespace preservation, updateFields for page-number fields only, deterministic packaging, and safe scoped exact-match replacement) must never be violated regardless of user styling requests.

1. **Dual-Width DXA for Tables**:
   Always set table width AND individual cell widths in DXA units (1 inch = 1440 DXA). Never use percentages. Cell widths must sum to the table width.
2. **Auto-Update Fields via Settings**:
   When using dynamic fields (PAGE and NUMPAGES only), always call `enable_update_fields_in_unpacked` or `enable_update_fields_in_doc`. The TOC/LOT/LOF are static lists and captions no longer use SEQ fields, so only page-number fields need the updateFields flag.
3. **Element Ordering in `<w:pPr>`**:
   Never append child elements randomly into paragraph properties. Use `insert_ordered_p_pr()` from `core_xml.py` to respect the ECMA-376 schema sequence (`pStyle`, `numPr`, `spacing`, `ind`, `jc`, `rPr`).
4. **Preserve Whitespace**:
   Any `<w:t>` containing leading or trailing spaces must include `xml:space="preserve"`. The `package.py` packer automatically repairs this.
5. **EMU Scaling for Media**:
   Always convert image dimensions to EMUs (English Metric Units: 914,400 EMUs = 1 inch, 360,000 EMUs = 1 cm).
6. **Strict Transparent Table Headers & Default Full Grid Config**:
   Never inject `<w:shd w:fill="FFFFFF"/>` or `clear` unless a solid background is explicitly requested. Omit `<w:shd>` entirely so that background letterheads, kop frames, and watermark images are not masked. Enforce table defaults: solid 0.5 pt grid borders on all sides (`sz=4` via `create_table()` default), font size 10 pt, explicit black font (`RGBColor(0, 0, 0)`), bold header row, 0 cm margin/padding on all sides, 0 pt paragraph spacing/indent. Vertical alignment is selective, not uniform: body cells use `vAlign=top`, header row uses `vAlign=center` plus `jc=center`. `create_table(vertical_center=False)` and `format_table_cells(vertical_center=False)` are the defaults; pass `vertical_center=True` only when the user explicitly asks for all cells centered. Gray cell shading is an explicit override, never a default. Use `col0_left=True` if the first column must be left-aligned.
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
13. **Prinsip Anti-Regex & Mutasi Exact-Match Berscope**:
    Saat memodifikasi dokumen, mengisi template, atau parsing prefiks list, dilarang keras regex (`re.sub`, `re.match`, dsb.) karena rapuh, rentan salah potong, dan berisiko merusak struktur run XML. Gunakan `replace_exact_text()` / `replace_placeholders()` dari `scripts/replace.py` dengan mutasi berscope, `expected=1`, asersi sebelum/sesudah, dan idempotensi; `scripts/mutation_guard.py` menjaga kontrak ini. Parsing prefiks list memakai tokenizing deterministik via `parse_list_prefix()`.
14. **Wajib Menggunakan Native Numbered List (Larangan Penomoran Manual)**:
    Dilarang keras mengetik manual penomoran pada paragraf biasa (`"1. Teks"`, `"a. Teks"`, `"• Teks"`). Seluruh butir daftar berurutan wajib memanggil `add_numbered_paragraph()` atau `add_numbered_list()` dari `scripts/lists.py`. Engine secara otomatis mendeteksi varian format (`decimal` 1., `lowerLetter` a., `upperLetter` A., `lowerRoman` i., `bullet`), mendaftarkan skema ke `word/numbering.xml`, dan menyematkan native Word `<w:numPr>` dengan *hanging indent* lurus rapi.
15. **Standar Format Caption: Nomor Hierarkis X.Y, Italic Terdaftar, 10 pt Times New Roman, Tanpa Bold, Perataan Kontekstual**:
    Caption memakai nomor hierarkis `Gambar X.Y` / `Tabel X.Y`, bukan field SEQ. Label dan nomor tegak; italic HANYA untuk istilah yang terdaftar di `data/lexicon.json`, diterapkan lewat `lexicon.italic_segments(text)`. Format satu baris ringkas (<= 45 karakter, fail-fast), Times New Roman 10 pt, TANPA BOLD. Perataan: `center` untuk Gambar, `left` untuk Tabel. Dilarang break line (`\n`), paragraf narasi tambahan di bawah caption, nomor subbab, kode eksperimen, atau nama variabel program di caption. Satu gambar satu caption.
16. **Hierarki Heading Caption Tetap (Gambar = Heading 4, Tabel = Heading 5, Tanpa Pergeseran Dinamis)**:
    Caption Gambar memakai Heading 4 (outlineLvl 3) dan caption Tabel memakai Heading 5 (outlineLvl 4), tetap. `add_seq_caption()` menerapkan level ini secara langsung; `resolve_caption_headings()` tidak lagi dipakai dan level caption tidak pernah bergeser mengikuti heading lain.

17. **Margin Akademik dan Indent Paragraf**:
    `sections.set_academic_margins()` memakai default 4-4-3-3 cm. `paragraphs.format_academic_paragraph()` memakai indent baris pertama 720 twips (1,27 cm / 0,5 inci) dengan spasi 1,5; `format_preface_paragraph()` tetap 1 cm dengan spasi 1,0. Jangan menimpa nilai ini tanpa permintaan eksplisit.

18. **Daftar Statis dan Dua Render PDF**:
    `toc.py` membangun Daftar Isi, Daftar Gambar, dan Daftar Tabel sebagai daftar statis dari heading dan caption, bukan field dinamis. Nomor halaman diambil dari render PDF LibreOffice dua kali (render awal, lalu render ulang setelah nomor disisipkan) agar stabil. Jika `soffice` atau `pdftotext` tidak tersedia, engine wajib error eksplisit, bukan fallback diam.

19. **Mutasi Berscope dengan Jaminan**:
    Setiap mutasi teks wajib berscope pada node yang dituju, memakai `expected=1`, asersi sebelum/sesudah, dan idempotensi. `scripts/mutation_guard.py` memastikan tidak ada perubahan di luar scope dan pengulangan operasi tidak menggandakan hasil.

20. **Paket DOCX Deterministik**:
    `package.pack_docx()` menghasilkan zip deterministik: entry diurutkan, timestamp 1980-01-01, dan `external_attr=0`. Jika byte hasil identik dengan file existing, tulis dilewati (skip) agar diff tetap bersih.

21. **Teks Kanonik untuk Diff Stabil**:
    `textconv.canonical_lines()` mengubah dokumen menjadi baris kanonik `JENIS|teks` sehingga diff tidak dipengaruhi urutan atribut XML atau format run. Pakai ini untuk regression diff dan inspeksi perubahan manual.

22. **SSOT Istilah dan Akronim**:
    `data/lexicon.json` lewat `scripts/lexicon.py` adalah satu-satunya sumber istilah italic, bentuk terlarang, dan ekspansi akronim; `scripts/acronyms.py` menerapkannya. Satu istilah per konsep. Akronim hanya dipakai setelah kepanjangan ditulis di badan naskah.

23. **Kalimat Narasi dan Sitasi**:
    `scripts/sentence_metrics.py` membatasi isi narasi <= 20 kata per kalimat (di luar sitasi kurung); `scripts/citation_qa.py` dan `scripts/reference_qa.py` memastikan sitasi menempel pada kalimat klaim dan referensi konsisten.

24. **Konvensi Flowchart dan Diagram**:
    Word OOXML tidak menjalankan sintaks diagram (Mermaid, Draw.io). Render ke `.png` lebih dulu, lalu sisipkan via `add_image_with_caption()`. Simpul diagram <= 4 kata, satu gambar satu caption, tanpa nomor subbab/kode eksperimen/nama variabel program di caption maupun simpul. Layout kompak monokrom dengan font cetak >= 9 pt.

25. **Post-Mortem Jalan Buntu**:
    Berhenti setelah dua hipotesis gugur atau dua siklus probe. Catat gejala, akar, obat, dan waktu deteksi (< 5 menit) sebelum mencoba arah baru. `scripts/verify.py` adalah runner fail-fast untuk seluruh QA dengan anggaran 35 detik; jalankan lewat `docx-engine-verify`.
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

# A4 paper with the default 4-4-3-3 cm margins
set_a4_page_size(sec)
set_academic_margins(sec)  # default 4-4-3-3 cm
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

### Recipe 2: Adding an APA Scientific Table (No Shading by Default)
```python
from tables import create_table, apply_academic_table_borders, set_cell_shading

table = create_table(doc, rows=3, cols=2, col_widths_inches=[3.5, 2.5])
table.rows[0].cells[0].paragraphs[0].text = "Metrik Evaluasi"
table.rows[0].cells[1].paragraphs[0].text = "Nilai Pengujian"

table.rows[1].cells[0].paragraphs[0].text = "P50 Latency"
table.rows[1].cells[1].paragraphs[0].text = "340 ms"

table.rows[2].cells[0].paragraphs[0].text = "Akurasi Klasifikasi"
table.rows[2].cells[1].paragraphs[0].text = "94.2%"

apply_academic_table_borders(table)
# Default: no shading, body cells vAlign=top, header row vAlign=center + jc=center.
# Gray shading is an explicit override only, never a default:
# set_cell_shading(table.rows[0].cells[0], "F0F0F0")
```

### Recipe 3: Hierarchical Captions and Cross-References
```python
from captions import add_seq_caption
from references import add_cross_reference_run

# Number derived from BAB order plus per-chapter caption sequence: Tabel 1.1.
# Label and number are upright; italic applies only to data/lexicon.json terms.
add_seq_caption(
    doc,
    label="Tabel",
    description="Komparasi Metrik Latensi",
    bookmark_name="tbl_latensi",
)

# Refer to it in body text
p_body = doc.add_paragraph("Hasil evaluasi komparasi dapat dilihat pada ")
add_cross_reference_run(p_body, "Tabel 1.1", bookmark_name="tbl_latensi")
```

### Recipe 4: Full Grid Table with Exact DXA Widths and Top-Aligned Body Cells
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
# vertical_center=False is the default: body cells vAlign=top,
# header row vAlign=center + jc=center. Set True only on explicit request.
format_table_cells(
    table,
    font_size_pt=10.0,
    line_spacing=1.0,
    ensure_no_color=True,
    vertical_center=False,
)
```

### Recipe 5: Native Multi-Format Word Lists (Auto-Detect or Explicit)
```python
from lists import add_numbered_paragraph, add_numbered_list

# 1. Auto-detection: awalan 'a.', 'b.', '1.', dsb. otomatis menjadi list asli Word
add_numbered_paragraph(doc, "a. Tahap **analisis** kebutuhan sistem")
add_numbered_paragraph(doc, "b. Tahap *perancangan* arsitektur")

# 2. Menambahkan daftar berurutan sekaligus (angka, huruf, romawi)
add_numbered_list(
    doc,
    items=["Tahap evaluasi awal", "Tahap optimasi bobot model"],
    num_fmt="lowerLetter",  # 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'bullet'
    line_spacing=1.15,
)
```

### Recipe 6: Canonical Diff of User Word Edits
```bash
# Canonical JENIS|teks lines keep the diff stable across XML formatting noise:
docx-engine-canon output/laporan.docx > laporan.canon
docx-engine-regression laporan.canon baseline.canon

# Scoped manual drift against git HEAD:
uv run python ~/.omp/agent/skills/docx-engine/scripts/diff_docx.py output/laporan.docx --git
```

### Recipe 7: Scoped Exact-Match Template Replacement
```python
import docx
from replace import replace_exact_text, replace_placeholders

doc = docx.Document("template.docx")

# Scoped replacement: one expected match, guarded before/after, idempotent, no regex
replace_exact_text(doc, "[JUDUL_LAPORAN]", "Sistem Klasifikasi Citra", expected=1)

# Bulk placeholder replacement across paragraphs, tables, and headers/footers
mapping = {
    "{{NAMA}}": "Prima Rizki",
    "{{NIM}}": "164221000",
    "{{TANGGAL}}": "30 Oktober 2026",
}
replace_placeholders(doc, mapping, expected=1)
doc.save("laporan_filled.docx")
```
