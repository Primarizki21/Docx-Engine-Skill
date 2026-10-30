from __future__ import annotations

import sys
from pathlib import Path
import docx

# Add scripts directory to path for testing
SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from captions import add_seq_caption
from core_xml import NAMESPACES, qn
from headers_footers import add_page_number_to_footer, add_page_number_with_total, set_running_header
from headings import add_academic_heading
from lists import add_bullet_item, add_numbered_item
from equations import add_centered_equation, make_omml_fraction
from package import pack_docx, parse_xml_part, unpack_docx, validate_xml_parts
from paragraphs import format_academic_paragraph
from references import add_bookmark, add_citation_text, add_cross_reference_run
from scaffold import scaffold_engine
from sections import add_section_break, set_a4_page_size, set_academic_margins, set_two_column_section
from settings import enable_update_fields_in_unpacked
from tables import apply_academic_table_borders, create_table, set_cell_margins, set_cell_shading
from toc import add_list_of_figures, add_list_of_tables, add_table_of_contents
from typography import add_formatted_run, convert_to_smart_quotes
from tables import (
    apply_grid_table_borders,
    format_table_cells,
    set_table_col_widths_dxa,
    set_table_compact_margins,
)
from lists import (
    add_numbered_list,
    add_numbered_paragraph,
    strip_manual_numbering_prefix,
)
from typography import add_markdown_runs
from diff_docx import run_diff_inspection


def test_full_academic_document_generation(tmp_path: Path):
    doc = docx.Document()

    # 1. Section 1: Preliminaries (Cover / TOC / LOT / LOF)
    sec1 = doc.sections[0]
    set_a4_page_size(sec1)
    set_academic_margins(sec1, top_cm=3.0, bottom_cm=3.0, left_cm=4.0, right_cm=3.0)
    set_running_header(sec1, "Laporan Evaluasi Chatbot IRIS")
    add_page_number_to_footer(sec1, align="center")

    add_table_of_contents(doc, title="DAFTAR ISI")
    add_list_of_figures(doc, title="DAFTAR GAMBAR")
    add_list_of_tables(doc, title="DAFTAR TABEL")

    # 2. Section 2: Main Content
    sec2 = add_section_break(doc, start_type="new_page")
    add_page_number_with_total(sec2, align="right")

    h1 = add_academic_heading(doc, "BAB I PENDAHULUAN", level=1)
    assert h1.text == "BAB I PENDAHULUAN"

    p_intro = doc.add_paragraph()
    format_academic_paragraph(p_intro)
    add_formatted_run(p_intro, "Sistem chatbot akademik mengintegrasikan model bahasa lokal.", bold=False)
    add_citation_text(p_intro, "FTMM-2026")

    # Typography smart quotes
    quote_text = convert_to_smart_quotes('"Hybrid RAG-CAG" adalah arsitektur yang diuji.')
    assert "“" in quote_text and "”" in quote_text

    # Lists
    b1 = add_bullet_item(doc, "Tinjauan kurikulum prodi")
    n1 = add_numbered_item(doc, "Tahap 1 evaluasi latensi")
    assert b1.text == "Tinjauan kurikulum prodi"
    assert n1.text == "Tahap 1 evaluasi latensi"

    # Math equation
    eq_xml = f"{make_omml_fraction('S_{CAG} + S_{RAG}', '2')}"
    add_centered_equation(doc, eq_xml)

    # Tables with APA border and dual width
    table = create_table(doc, rows=3, cols=2, col_widths_inches=[3.0, 3.0])
    table.rows[0].cells[0].paragraphs[0].text = "Varian"
    table.rows[0].cells[1].paragraphs[0].text = "Akurasi"
    table.rows[1].cells[0].paragraphs[0].text = "Hybrid RAG-CAG"
    table.rows[1].cells[1].paragraphs[0].text = "94.2%"
    table.rows[2].cells[0].paragraphs[0].text = "Pure RAG"
    table.rows[2].cells[1].paragraphs[0].text = "81.5%"
    apply_academic_table_borders(table)
    set_cell_shading(table.rows[0].cells[0], "EAEAEA")
    set_cell_margins(table.rows[0].cells[0], top_pt=4, bottom_pt=4, left_pt=6, right_pt=6)

    # Captions and Cross-references
    p_cap = add_seq_caption(doc, label="Tabel", description="Komparasi Akurasi", bookmark_name="tbl_komparasi")
    assert "Tabel " in p_cap.text

    p_ref = doc.add_paragraph()
    add_cross_reference_run(p_ref, "Lihat Tabel 1", bookmark_name="tbl_komparasi")

    # 3. Save, Unpack, Enable UpdateFields, Validate, Pack
    target_docx = tmp_path / "complete_report.docx"
    doc.save(str(target_docx))
    assert target_docx.exists()

    unpacked_dir = tmp_path / "unpacked"
    unpack_docx(target_docx, unpacked_dir)
    assert (unpacked_dir / "word" / "document.xml").exists()

    # Inject updateFields into settings.xml
    enable_update_fields_in_unpacked(unpacked_dir)
    settings_root, _ = parse_xml_part(unpacked_dir, "word/settings.xml")
    assert settings_root.find(qn("w:updateFields")) is not None

    # Validate XML integrity
    errors = validate_xml_parts(unpacked_dir)
    assert len(errors) == 0

    repacked_docx = tmp_path / "final_report.docx"
    pack_docx(unpacked_dir, repacked_docx)
    assert repacked_docx.exists()

    # Reload document to guarantee Word-parsable
    reloaded_doc = docx.Document(str(repacked_docx))
    assert len(reloaded_doc.paragraphs) > 5


def test_scaffolding_functionality(tmp_path: Path):
    target_engine_dir = tmp_path / "scaffolded_engine"
    dest = scaffold_engine(target_engine_dir)
    assert dest.exists()
    assert (dest / "__init__.py").exists()
    assert (dest / "tables.py").exists()
    assert (dest / "toc.py").exists()
    assert (dest / "core_xml.py").exists()


def test_grid_table_and_transparent_header(tmp_path: Path):
    doc = docx.Document()
    table = create_table(doc, rows=3, cols=3)
    set_table_col_widths_dxa(table, [2000, 3000, 2000])
    apply_grid_table_borders(table, sz=4)
    set_table_compact_margins(table, left_dxa=0, right_dxa=0)
    format_table_cells(table, font_size_pt=10.0, line_spacing=1.0, ensure_no_color=True)

    out_file = tmp_path / "test_grid.docx"
    doc.save(str(out_file))

    # Reload and verify
    reloaded = docx.Document(str(out_file))
    t = reloaded.tables[0]
    tblPr = t._tbl.tblPr

    # Check tblW
    tblW = tblPr.find(qn("w:tblW"))
    assert tblW.attrib.get(qn("w:w")) == "7000"

    # Check borders
    borders = tblPr.find(qn("w:tblBorders"))
    assert borders is not None
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = borders.find(qn(f"w:{side}"))
        assert b is not None
        assert b.attrib.get(qn("w:sz")) == "4"

    # Check no shading
    for row in t.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            assert tcPr.find(qn("w:shd")) is None


def test_numbered_paragraphs_and_markdown_runs(tmp_path: Path):
    doc = docx.Document()

    # Test strip prefix
    assert strip_manual_numbering_prefix("1. Krisis Ekonomi") == "Krisis Ekonomi"
    assert strip_manual_numbering_prefix("• Butir Pertama") == "Butir Pertama"

    # Add numbered paragraph with markdown
    p = add_numbered_paragraph(
        doc,
        text="1. Krisis Pemuda: Mengalami **dampak permanen** (*scarring effect*) pada karir.",
        num_id=3,
        ilvl=0,
        line_spacing=1.15,
    )

    # Verify numPr
    numPr = p._p.pPr.find(qn("w:numPr"))
    assert numPr is not None
    assert numPr.find(qn("w:numId")).attrib.get(qn("w:val")) == "3"
    assert numPr.find(qn("w:ilvl")).attrib.get(qn("w:val")) == "0"

    # Verify bold and italic runs
    runs = p.runs
    bold_found = any(r.bold and "dampak permanen" in r.text for r in runs)
    italic_found = any(r.italic and "scarring effect" in r.text for r in runs)
    assert bold_found
    assert italic_found

    # Add numbered list
    list_items = ["Item Satu", "Item Dua", "Item Tiga"]
    added_paras = add_numbered_list(doc, list_items, num_id=4, ilvl=0)
    assert len(added_paras) == 3

    out_file = tmp_path / "test_numbered.docx"
    doc.save(str(out_file))
    assert out_file.exists()


def test_diff_docx_inspection(tmp_path: Path):
    # Doc 1: Baseline with APA table
    doc1 = docx.Document()
    t1 = create_table(doc1, rows=2, cols=2)
    apply_academic_table_borders(t1)
    p1 = doc1.add_paragraph("Teks awal")
    file1 = tmp_path / "base.docx"
    doc1.save(str(file1))

    # Doc 2: Modified with Grid table and numbered paragraph
    doc2 = docx.Document()
    t2 = create_table(doc2, rows=2, cols=2)
    apply_grid_table_borders(t2, sz=4)
    format_table_cells(t2, ensure_no_color=True)
    p2 = add_numbered_paragraph(doc2, "Teks baru bernomor", num_id=2, ilvl=0)
    file2 = tmp_path / "mod.docx"
    doc2.save(str(file2))

    # Run diff inspection
    report = run_diff_inspection(file1, file2, mode="summary")
    assert "TABLES" in report
    assert "NUMBERING" in report
    assert "REKOMENDASI KODE GENERATOR PYTHON" in report

    # Test XML diff mode
    xml_diff = run_diff_inspection(file1, file2, mode="xml")
    assert "---" in xml_diff
    assert "+++" in xml_diff
