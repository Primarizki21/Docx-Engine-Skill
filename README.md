# docx-engine

Modular, zero-black-box Python & OOXML toolchain for generating and manipulating professional academic and technical Word (`.docx`) documents.

## Core Features

- **Astral `uv` Native**: Pure Python powered by `python-docx` and `lxml` via `uv`.
- **Modular Micro-Scripts**: Focused modules under `scripts/` (each `< 120` lines) for surgical document operations.
- **Read-Modify-Write**: Open existing templates or documents, modify targeted XML/nodes, and save back cleanly.
- **Academic & Publication Standards**:
  - Full-grid and APA scientific tables (dual DXA widths, compact margins, transparent headers).
  - Multi-level bullet and numbered lists with native Word `w:numPr`.
  - Academic heading hierarchy (Heading 1–4) with outline levels.
  - Automatic dynamic TOC, List of Figures (LOF), List of Tables (LOT).
  - Sequential captions (`SEQ`) and clickable cross-references.
  - OMML math equations (fractions, matrices, formulas).
  - Page numbers (`PAGE` / `NUMPAGES`) in headers/footers with A4 margin presets.
  - Word diff & inspection tool (`docx-engine-diff`) for tracking user edits against git baselines.

---

## Quickstart with Astral `uv`

### Installation & Sync

```bash
# Clone the repository
git clone <repo-url> docx-engine
cd docx-engine

# Sync dependencies and virtual environment via Astral uv
uv sync --all-extras
```

### Run Tests

```bash
uv run pytest
```

### CLI Commands

Entrypoints provided out-of-the-box:

```bash
# Eject / scaffold the engine into another project
uv run docx-engine-scaffold --target /path/to/project/src/engine

# Unpack and repack docx packages with XML auto-repair
uv run docx-engine-package unpack ./document.docx ./unpacked
uv run docx-engine-package validate ./unpacked
uv run docx-engine-package pack ./unpacked ./repaired.docx

# Diff Word documents against git HEAD or between versions
uv run docx-engine-diff ./document.docx --git
uv run docx-engine-diff ./base.docx ./edited.docx --mode xml
```

---

## Modules Overview

| Module | Location | Primary Functions |
|---|---|---|
| **Sections** | `scripts/sections.py` | `set_a4_page_size`, `set_academic_margins`, `set_two_column_section`, `add_section_break` |
| **Headings** | `scripts/headings.py` | `add_academic_heading`, `set_outline_level` |
| **Paragraphs** | `scripts/paragraphs.py` | `set_paragraph_spacing`, `set_paragraph_indent`, `set_paragraph_alignment`, `format_academic_paragraph` |
| **Typography** | `scripts/typography.py` | `format_run`, `add_formatted_run`, `add_markdown_runs`, `convert_to_smart_quotes` |
| **Lists** | `scripts/lists.py` | `add_bullet_item`, `add_numbered_item`, `add_numbered_paragraph`, `add_numbered_list` |
| **Tables** | `scripts/tables.py` | `create_table`, `apply_academic_table_borders`, `apply_grid_table_borders`, `set_table_col_widths_dxa`, `set_table_compact_margins`, `format_table_cells` |
| **Media** | `scripts/media.py` | `add_image_with_caption` |
| **Headers/Footers** | `scripts/headers_footers.py` | `add_page_number_to_footer`, `add_page_number_with_total`, `set_running_header` |
| **References** | `scripts/references.py` | `add_bookmark`, `add_cross_reference_run`, `add_citation_text` |
| **Table of Contents** | `scripts/toc.py` | `add_table_of_contents`, `add_list_of_figures`, `add_list_of_tables` |
| **Equations** | `scripts/equations.py` | `add_omml_block`, `add_inline_math`, `add_centered_equation`, `make_omml_fraction` |
| **Settings** | `scripts/settings.py` | `enable_update_fields_in_unpacked`, `enable_update_fields_in_doc` |
| **Package** | `scripts/package.py` | `unpack_docx`, `pack_docx`, `parse_xml_part`, `save_xml_part`, `validate_xml_parts` |
| **Scaffold** | `scripts/scaffold.py` | `scaffold_engine` |
| **Diff** | `scripts/diff_docx.py` | `run_diff_inspection` |

---

## License

MIT
