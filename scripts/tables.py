from __future__ import annotations

import argparse
import json
from pathlib import Path
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table, _Cell


def create_table(
    document,
    rows: int,
    cols: int,
    col_widths_inches: list[float] | None = None,
    align: str = "center",
    compact_margins: bool = True,
) -> Table:
    table = document.add_table(rows=rows, cols=cols)
    align_map = {
        "center": WD_TABLE_ALIGNMENT.CENTER,
        "left": WD_TABLE_ALIGNMENT.LEFT,
        "right": WD_TABLE_ALIGNMENT.RIGHT,
    }
    table.alignment = align_map.get(align.lower(), WD_TABLE_ALIGNMENT.CENTER)
    if compact_margins:
        set_table_compact_margins(table, top_dxa=0, bottom_dxa=0, left_dxa=0, right_dxa=0)


    if col_widths_inches and len(col_widths_inches) == cols:
        set_table_col_widths_inches(table, col_widths_inches)
    return table


def set_table_col_widths_inches(table: Table, widths_inches: list[float]) -> None:
    # Standar Dual-Width DXA (1 inci = 1440 DXA)
    total_dxa = sum(int(w * 1440) for w in widths_inches)
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")

    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for idx, cell in enumerate(row.cells):
            dxa_val = int(widths_inches[idx] * 1440)
            cell.width = Inches(widths_inches[idx])
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(dxa_val))
            tcW.set(qn("w:type"), "dxa")


def set_cell_shading(cell: _Cell, color_hex: str) -> None:
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))


def set_cell_margins(cell: _Cell, top_pt: int = 4, bottom_pt: int = 4, left_pt: int = 6, right_pt: int = 6) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, pt in [("top", top_pt), ("bottom", bottom_pt), ("left", left_pt), ("right", right_pt)]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(int(pt * 20)))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def apply_academic_table_borders(table: Table) -> Table:
    # Standar format tabel ilmiah / APA
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'  <w:insideH w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)

    if len(table.rows) > 0:
        header_tr = table.rows[0]._tr.get_or_add_trPr()
        header_tr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f"</w:tcBorders>"
            )
            tcPr.append(tcBorders)
    return table


def set_table_col_widths_dxa(table: Table, widths_dxa: list[int]) -> None:
    """Mengatur lebar kolom tabel secara presisi dalam satuan DXA (1 inci = 1440 DXA)."""
    total_dxa = sum(widths_dxa)
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total_dxa))
    tblW.set(qn("w:type"), "dxa")

    # Perbarui elemen tblGrid
    tblGrid = table._tbl.tblGrid
    if tblGrid is not None:
        tblGrid.clear()
        for w in widths_dxa:
            col = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w}"/>')
            tblGrid.append(col)

    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for idx, cell in enumerate(row.cells):
            if idx < len(widths_dxa):
                dxa_val = widths_dxa[idx]
                cell.width = Inches(dxa_val / 1440.0)
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = OxmlElement("w:tcW")
                    tcPr.append(tcW)
                tcW.set(qn("w:w"), str(dxa_val))
                tcW.set(qn("w:type"), "dxa")


def apply_grid_table_borders(table: Table, sz: int = 4, color: str = "auto") -> Table:
    """Menerapkan border grid kotak penuh (solid 0.5 pt / sz=4) pada seluruh batas tabel."""
    tblPr = table._tbl.tblPr
    # Hapus border lama jika ada
    existing_borders = tblPr.find(qn("w:tblBorders"))
    if existing_borders is not None:
        tblPr.remove(existing_borders)

    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    # Pastikan cantSplit di semua baris dan tblHeader di baris 0
    for idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if idx == 0 and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    return table


def set_table_compact_margins(
    table: Table,
    top_dxa: int = 0,
    bottom_dxa: int = 0,
    left_dxa: int = 0,
    right_dxa: int = 0,
) -> None:
    """Mengatur margin sel rapat (tblCellMar) pada level tabel induk."""
    tblPr = table._tbl.tblPr
    existing_mar = tblPr.find(qn("w:tblCellMar"))
    if existing_mar is not None:
        tblPr.remove(existing_mar)

    tblCellMar = parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top_dxa}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left_dxa}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom_dxa}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right_dxa}" w:type="dxa"/>\n'
        f'</w:tblCellMar>'
    )
    tblPr.append(tblCellMar)


def remove_cell_shading(cell: _Cell) -> None:
    """Menghapus background/shading sel sehingga 100% transparan / no color."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing_shd = tcPr.find(qn("w:shd"))
    if existing_shd is not None:
        tcPr.remove(existing_shd)


def format_table_cells(
    table: Table,
    font_name: str = "Times New Roman",
    font_size_pt: float = 10.0,
    line_spacing: float = 1.0,
    header_bold: bool = True,
    header_center: bool = True,
    col0_left: bool = True,
    ensure_no_color: bool = True,
    compact_margins: bool = True,
    font_color_black: bool = True,
) -> None:
    """Memformat seluruh sel tabel: font 10 pt hitam, single spasi (line=240), space after/before=0,
    margin sel rapat 0 cm, serta perataan yang proporsional."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    if compact_margins:
        set_table_compact_margins(table, top_dxa=0, bottom_dxa=0, left_dxa=0, right_dxa=0)

    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx == 0)
        for col_idx, cell in enumerate(row.cells):
            if ensure_no_color:
                remove_cell_shading(cell)

            for p in cell.paragraphs:
                # Spasi baris 1.0 (single), space before/after 0 pt, dan indent 0
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.left_indent = Pt(0)
                p.paragraph_format.right_indent = Pt(0)
                p.paragraph_format.line_spacing = line_spacing
                # Alignment
                if is_header and header_center:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_idx == 0 and col0_left:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                for run in p.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size_pt)
                    if font_color_black:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    if is_header and header_bold:
                        run.bold = True
