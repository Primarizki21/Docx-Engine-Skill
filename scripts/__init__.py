from __future__ import annotations

from .captions import add_seq_caption, apply_caption_heading, detect_caption_type, resolve_caption_headings
from .core_xml import NAMESPACES, W_NS, ensure_xml_space, insert_ordered_p_pr, qn
from .headers_footers import add_page_number_to_footer, add_page_number_with_total, set_running_header
from .headings import add_academic_heading, add_chapter_heading, get_max_body_heading_level, set_outline_level
from .lists import (
    add_bullet_item,
    add_numbered_item,
    add_numbered_list,
    add_numbered_paragraph,
    attach_num_pr,
    ensure_numbering_scheme,
    parse_list_prefix,
    strip_manual_numbering_prefix,
)
from .equations import add_centered_equation, add_inline_math, add_omml_block, make_omml_fraction
from .media import add_image_with_caption
from .package import pack_docx, parse_xml_part, save_xml_part, unpack_docx, validate_xml_parts
from .paragraphs import (
    format_academic_paragraph,
    format_preface_paragraph,
    format_table_cell_paragraph,
    format_toc_entry_paragraph,
    set_paragraph_alignment,
    set_paragraph_indent,
    set_paragraph_spacing,
)
from .references import add_bookmark, add_citation_text, add_cross_reference_run
from .replace import replace_exact_text, replace_exact_text_in_paragraph, replace_placeholders
from .sections import (
    add_section_break,
    set_a4_page_size,
    set_academic_margins,
    set_letter_page_size,
    set_two_column_section,
)
from .settings import enable_update_fields_in_doc, enable_update_fields_in_unpacked
from .tables import (
    apply_academic_table_borders,
    apply_grid_table_borders,
    create_table,
    format_table_cells,
    remove_cell_shading,
    set_cell_margins,
    set_cell_shading,
    set_table_col_widths_dxa,
    set_table_col_widths_inches,
    set_table_compact_margins,
)
from .toc import add_list_of_figures, add_list_of_tables, add_table_of_contents
from .typography import add_formatted_run, add_markdown_runs, convert_to_smart_quotes, format_run
from .diff_docx import run_diff_inspection

__all__ = [
    "add_academic_heading",
    "add_chapter_heading",
    "set_outline_level",
    "add_page_number_to_footer",
    "add_page_number_with_total",
    "set_running_header",
    "add_bullet_item",
    "add_numbered_item",
    "add_numbered_paragraph",
    "add_numbered_list",
    "ensure_numbering_scheme",
    "parse_list_prefix",
    "strip_manual_numbering_prefix",
    "attach_num_pr",
    "add_image_with_caption",
    "unpack_docx",
    "pack_docx",
    "parse_xml_part",
    "save_xml_part",
    "validate_xml_parts",
    "format_academic_paragraph",
    "format_preface_paragraph",
    "format_table_cell_paragraph",
    "format_toc_entry_paragraph",
    "set_paragraph_alignment",
    "set_paragraph_indent",
    "set_paragraph_spacing",
    "add_bookmark",
    "add_citation_text",
    "add_cross_reference_run",
    "set_a4_page_size",
    "set_letter_page_size",
    "set_academic_margins",
    "set_two_column_section",
    "add_section_break",
    "create_table",
    "set_table_col_widths_inches",
    "set_cell_shading",
    "set_cell_margins",
    "apply_academic_table_borders",
    "apply_grid_table_borders",
    "set_table_col_widths_dxa",
    "set_table_compact_margins",
    "remove_cell_shading",
    "format_table_cells",
    "format_run",
    "add_formatted_run",
    "add_markdown_runs",
    "run_diff_inspection",
    "replace_exact_text",
    "replace_exact_text_in_paragraph",
    "replace_placeholders",
    "add_table_of_contents",
    "add_list_of_figures",
    "add_list_of_tables",
    "add_seq_caption",
    "apply_caption_heading",
    "detect_caption_type",
    "resolve_caption_headings",
    "get_max_body_heading_level",
    "add_omml_block",
    "add_inline_math",
    "add_centered_equation",
    "make_omml_fraction",
    "enable_update_fields_in_doc",
    "enable_update_fields_in_unpacked",
    "insert_ordered_p_pr",
    "ensure_xml_space",
    "qn",
    "NAMESPACES",
    "W_NS",
]
