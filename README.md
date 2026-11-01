# Docx Engine

A zero-black-box Python & OOXML toolchain for generating, formatting, and inspecting professional academic, technical, and corporate Word (`.docx`) documents.

Designed as both a **universal AI Agent Skill** (compatible with Claude Code, OpenCode, OMP, Cursor, Windsurf) and an **ejectable standalone Python library**.

---

## Why Docx Engine?

Vanilla `python-docx` works for basic text, but complex academic theses, journals, and formal reports often run into frustrating OOXML quirks:

| Problem in Word Generation | How Docx Engine Solves It |
|---|---|
| **Corrupted Documents**: Word rejects out-of-order paragraph properties | Enforces ECMA-376 schema order inside `<w:pPr>` (`pStyle` -> `numPr` -> `spacing` -> `ind` -> `jc` -> `rPr`) |
| **Broken Chapter Pagination**: Ctrl+Enter page breaks trap chapters in one section | Automates Section Break (Next Page) per BAB so headers, footers, and page numbers isolate cleanly |
| **Outdated TOC & Page Numbers**: Word opens files with blank or stale TOC fields | Automatically injects `w:updateFields` into `settings.xml` to force field refresh on file open |
| **Deformed Tables**: Columns distort across pages or mask letterhead watermarks | Enforces dual-width DXA columns, transparent (no-fill) headers, and compact 0 cm padding |
| **Inconsistent Spacing**: Manual spacing mixes up 1.5 body vs single-spaced front matter | Provides context-aware presets: 1.5 for chapter body, single 1.0 for Kata Pengantar and tables |
| **Ragged / Fragile Numbering**: Manual string numbers ("1. Text") lack hanging indents and cannot sync | Enforces native Word `<w:numPr>` (`numId`, `ilvl`) with automatic prefix stripping so lines align cleanly |
| **Inconsistent Captions & AI Slop**: Captions mix bold styles or inject multi-line narrative explanations | Enforces pure 10 pt italic non-bold captions, dynamic `SEQ` fields, and single-line titles (<= 45 chars) |
| **Accidental Bold Subsubbabs**: Word's default template renders Heading 3 in bold and accent blue | Locks Heading 1-2 to bold and Heading 3-4 to regular non-bold black text with 0 pt spacing |

---

## Quickstart (Python API)

Create a complete standardized document with a few lines of code:

```python
import docx
from scripts import (
    set_a4_page_size,
    set_academic_margins,
    add_academic_heading,
    add_chapter_heading,
    format_academic_paragraph,
    format_preface_paragraph,
    add_numbered_paragraph,
    add_seq_caption,
    create_table,
    apply_grid_table_borders,
    format_table_cells,
    add_page_number_to_footer,
    enable_update_fields_in_doc,
)

doc = docx.Document()
sec = doc.sections[0]

# 1. Setup A4 Paper & Standard Margins (Top/Bottom/Right: 3 cm, Left: 4 cm)
set_a4_page_size(sec)
set_academic_margins(sec, top_cm=3.0, bottom_cm=3.0, left_cm=4.0, right_cm=3.0)
add_page_number_to_footer(sec, align="center")

# 2. Preliminary Front Matter: Kata Pengantar (Single Line Spacing 1.0)
add_academic_heading(doc, "KATA PENGANTAR", level=1)
p_preface = doc.add_paragraph("Segala puji dan syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa...")
format_preface_paragraph(p_preface)

# 3. Chapter 1: Automatically starts with a Section Break (Next Page)
add_chapter_heading(doc, "BAB I PENDAHULUAN")
p_body = doc.add_paragraph("Latar belakang penelitian mengkaji arsitektur hybrid model...")
format_academic_paragraph(p_body)

# 4. Subbab (Heading 2 = Bold) and Subsubbab (Heading 3 = Regular Non-Bold)
add_academic_heading(doc, "1.1 Latar Belakang Masalah", level=2)
add_academic_heading(doc, "1.1.1 Identifikasi Gap Riset", level=3)

# 5. Native Word Lists (Multi-Format: Numbers, Letters a., Roman i., Bullet)
add_numbered_paragraph(doc, "a. Evaluasi latensi awal dan throughput pipeline.")
add_numbered_paragraph(doc, "b. Optimasi bobot attention dan pemangkasan model.")

# 6. Sequential Caption (Pure Italic, 10 pt, No Bold, Single-Line)
add_seq_caption(doc, label="Tabel", description="Komparasi Metrik Latensi")

# 7. Standard Table: 10 pt black font, bold header, 0 cm margins, full grid borders
table = create_table(doc, rows=3, cols=2, col_widths_inches=[3.5, 2.5])
table.rows[0].cells[0].paragraphs[0].text = "Parameter"
table.rows[0].cells[1].paragraphs[0].text = "Nilai"
table.rows[1].cells[0].paragraphs[0].text = "Latency"
table.rows[1].cells[1].paragraphs[0].text = "340 ms"
table.rows[2].cells[0].paragraphs[0].text = "Accuracy"
table.rows[2].cells[1].paragraphs[0].text = "94.2%"

apply_grid_table_borders(table)
format_table_cells(table)

# 8. Enable Word Dynamic Field Auto-Update on File Open
enable_update_fields_in_doc(doc)
doc.save("laporan_akademik.docx")
```

---

## Installation & Setup

You can run this project using [Astral uv](https://github.com/astral-sh/uv) (recommended) or standard `pip`.

### Option A: Using Astral uv (Fast & Cross-Platform)

```bash
# Clone the repository
git clone git@github.com:Primarizki21/Docx-Engine-Skill.git docx-engine
cd docx-engine

# Sync dependencies in a virtual environment
uv sync
```

### Option B: Using Standard Python Virtualenv & pip

```bash
# Linux / macOS
python3 -m venv .venv
source .venv/bin/activate
pip install -e .

# Windows (Command Prompt)
python -m venv .venv
.venv\Scripts\activate.bat
pip install -e .

# Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
```

---

## Setting Up as an AI Agent Skill

This repository includes a comprehensive `SKILL.md` specification and can be plugged directly into various AI coding agents.

### 1. Linking into Your Agent's Skills Folder

Link this repository into whichever skills directory your AI tool monitors:

**On Linux / macOS:**
```bash
ln -s "/absolute/path/to/docx-engine" "<YOUR_SKILLS_DIRECTORY>/docx-engine"

# Examples:
# For Claude Code: ln -s "/path/to/docx-engine" "$HOME/.claude/skills/docx-engine"
# For OMP / OpenCode: ln -s "/path/to/docx-engine" "$HOME/.omp/agent/skills/docx-engine"
```

**On Windows (PowerShell as Administrator or Developer Mode):**
```powershell
New-Item -ItemType SymbolicLink -Path "<YOUR_SKILLS_DIRECTORY>\docx-engine" -Target "C:\path\to\docx-engine"

# Examples:
# For Claude Code: New-Item -ItemType SymbolicLink -Path "$HOME\.claude\skills\docx-engine" -Target "C:\path\to\docx-engine"
# For OMP / OpenCode: New-Item -ItemType SymbolicLink -Path "$HOME\.omp\agent\skills\docx-engine" -Target "C:\path\to\docx-engine"
```

### 2. Using Inside Project Workspaces (Cursor, Windsurf, Copilot, Roo)

If your coding assistant operates within a specific project workspace:
1. Copy or scaffold the `scripts/` folder into your project using the scaffold command:
   ```bash
   uv run docx-engine-scaffold --target ./src/engine
   ```
2. Reference `SKILL.md` in your project's agent instructions (such as `CLAUDE.md`, `.cursorrules`, or workspace system prompt).

---

## CLI Toolchain

Built-in command-line tools for scaffolding, XML surgery, and document diffing:

```bash
# 1. Eject / scaffold the engine directly into a target project
uv run docx-engine-scaffold --target /path/to/project/src/engine

# 2. Unpack docx to raw XML, auto-repair whitespace, and repack cleanly
uv run docx-engine-package unpack ./document.docx ./unpacked
uv run docx-engine-package validate ./unpacked
uv run docx-engine-package pack ./unpacked ./repaired.docx

# 3. Diff Word document changes against git HEAD or between versions
uv run docx-engine-diff ./document.docx --git
uv run docx-engine-diff ./base.docx ./edited.docx --mode xml
```

---

## Modules Overview

Every module under `scripts/` is tightly scoped (< 120 lines) and can be used independently:

| Module | File | Key Capabilities |
|---|---|---|
| **Sections** | `scripts/sections.py` | A4/Letter sizing, 4-3-3-3 cm margins, two-column layout, section breaks |
| **Headings** | `scripts/headings.py` | Heading 1-2 bold, Heading 3-4 regular non-bold, zero space before/after |
| **Paragraphs** | `scripts/paragraphs.py` | 1.5 academic body spacing, 1.0 Kata Pengantar spacing, alignments, indents |
| **Tables** | `scripts/tables.py` | Dual DXA widths, Full-Grid and APA borders, 0 cm padding, black 10 pt font |
| **Typography** | `scripts/typography.py` | Character runs, Times New Roman presets, markdown parsing, smart quotes |
| **Lists** | `scripts/lists.py` | Native Word `w:numPr` numbered lists with hanging indent, prefix stripping |
| **TOC & Fields** | `scripts/toc.py` | Dynamic Table of Contents, List of Figures (LOF), List of Tables (LOT) |
| **Captions** | `scripts/captions.py` | Pure italic 10 pt captions, dynamic headings (Gambar H4, Tabel H5, dynamic shifting), bookmarks |
| **Equations** | `scripts/equations.py` | Native Office Math (OMML) blocks, fractions, centered formulas |
| **Headers/Footers**| `scripts/headers_footers.py` | Running headers, `PAGE` and `NUMPAGES` footer fields |
| **Settings** | `scripts/settings.py` | Injects `w:updateFields` to guarantee dynamic fields update on file open |
| **Package** | `scripts/package.py` | Unpack zip, inspect raw XML, auto-repair missing whitespace, repack |
| **Diff** | `scripts/diff_docx.py` | Semantic and XML diff inspection comparing Word edits against git baseline |
| **Scaffold** | `scripts/scaffold.py` | Copies engine modules into target projects for complete ownership |
| **Replace** | `scripts/replace.py` | Run-preserving exact match string & placeholder replacement for templates |

---

## License

MIT License.
