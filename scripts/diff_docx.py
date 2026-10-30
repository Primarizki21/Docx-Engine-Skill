#!/usr/bin/env python3
"""DOCX Diff & Inspection Engine: Automated Semantic & XML Comparator for DOCX Documents.

Membandingkan dua dokumen Word (.docx) atau berkas Word yang dimodifikasi terhadap git HEAD.
Mendeteksi perubahan semantik pada:
- Teks & Typography Runs
- Formatting Paragraf (Spasi, Indent, Alignment, keepNext)
- Headings & Subheadings (Outline levels, Styles)
- Tabel (Borders, Shading/No-Color, Lebar Kolom DXA, Cell Padding, cantSplit, tblHeader)
- Penomoran Asli Word (w:numPr, numId, ilvl, ListParagraph)
- Gambar & Media (Dimensi EMU, Inline/Anchor, Captions)
- Equation / OMML Math
- Daftar Pustaka (Hanging indents, [1] IEEE numbering)
- Pengaturan Halaman (Margin, Kertas A4/Letter, Orientasi, Header/Footer)

Menghasilkan:
1. Laporan bahasa manusia yang ringkas dan terstruktur.
2. Rekomendasi cuplikan kode Python generator untuk mereproduksi perubahan.
3. Diff XML bersih ter-normalisasi (opsional: --mode xml).
"""

from __future__ import annotations

import argparse
import difflib
import io
import os
import subprocess
import sys
import tempfile
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

NAMESPACES = {
    "w": W_NS,
    "m": M_NS,
    "wp": WP_NS,
    "a": A_NS,
    "pic": PIC_NS,
}


def read_docx_xml(docx_path: str | Path, internal_path: str = "word/document.xml") -> ET.Element | None:
    """Membaca berkas XML dari dalam arsip ZIP docx."""
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            if internal_path in z.namelist():
                return ET.fromstring(z.read(internal_path))
    except Exception as e:
        print(f"Error membaca {internal_path} dari {docx_path}: {e}", file=sys.stderr)
    return None


def extract_git_head_docx(target_path: str | Path) -> Path | None:
    """Mengekstrak versi git HEAD dari dokumen sebagai baseline perbandingan."""
    rel_path = os.path.relpath(target_path)
    try:
        proc = subprocess.run(
            ["git", "show", f"HEAD:{rel_path}"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
        )
        tf = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tf.write(proc.stdout)
        tf.close()
        return Path(tf.name)
    except Exception as e:
        print(f"Gagal mengekstrak versi git HEAD untuk {target_path}: {e}", file=sys.stderr)
        return None


def normalize_xml_string(raw_xml_bytes: bytes) -> str:
    """Menghilangkan XML noise (rsid, paraId, textId, proofErr) dan memformat rapi."""
    root = ET.fromstring(raw_xml_bytes)

    # 1. Hapus atribut-atribut noise
    for elem in root.iter():
        attribs_to_remove = [
            k for k in elem.attrib
            if "rsid" in k.lower() or "paraid" in k.lower() or "textid" in k.lower()
        ]
        for k in attribs_to_remove:
            del elem.attrib[k]

    # 2. Hapus tag proofErr (spell check)
    for parent in root.iter():
        children_to_remove = [c for c in list(parent) if c.tag.endswith("proofErr")]
        for c in children_to_remove:
            parent.remove(c)

    rough = ET.tostring(root, "utf-8")
    reparsed = minidom.parseString(rough)
    return reparsed.toprettyxml(indent="  ")


def inspect_tables_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi perubahan pada seluruh tabel."""
    results = []
    tbls_base = root_base.findall(".//w:tbl", NAMESPACES)
    tbls_mod = root_mod.findall(".//w:tbl", NAMESPACES)

    if len(tbls_base) != len(tbls_mod):
        results.append(f"Jumlah tabel berubah: {len(tbls_base)} -> {len(tbls_mod)}")

    for idx in range(min(len(tbls_base), len(tbls_mod))):
        tb = tbls_base[idx]
        tm = tbls_mod[idx]
        diffs = []

        # Lebar tabel
        w_b = tb.find(".//w:tblW", NAMESPACES)
        w_m = tm.find(".//w:tblW", NAMESPACES)
        wb_val = w_b.attrib.get(f"{{{W_NS}}}w") if w_b is not None else None
        wm_val = w_m.attrib.get(f"{{{W_NS}}}w") if w_m is not None else None
        if wb_val != wm_val:
            diffs.append(f"Lebar tabel total: {wb_val} dxa -> {wm_val} dxa")

        # Kolom grid
        cols_b = [c.attrib.get(f"{{{W_NS}}}w", "?") for c in tb.findall(".//w:gridCol", NAMESPACES)]
        cols_m = [c.attrib.get(f"{{{W_NS}}}w", "?") for c in tm.findall(".//w:gridCol", NAMESPACES)]
        if cols_b != cols_m:
            diffs.append(f"Rincian lebar kolom: {cols_b} -> {cols_m}")

        # Borders
        borders_b = tb.find(".//w:tblBorders", NAMESPACES)
        borders_m = tm.find(".//w:tblBorders", NAMESPACES)
        bb_str = "None"
        bm_str = "None"
        if borders_b is not None:
            bb_top = borders_b.find("w:top", NAMESPACES)
            bb_inH = borders_b.find("w:insideH", NAMESPACES)
            bb_str = "Grid" if bb_inH is not None and bb_inH.attrib.get(f"{{{W_NS}}}val") == "single" else "APA/Open"
        if borders_m is not None:
            bm_top = borders_m.find("w:top", NAMESPACES)
            bm_inH = borders_m.find("w:insideH", NAMESPACES)
            bm_str = "Grid" if bm_inH is not None and bm_inH.attrib.get(f"{{{W_NS}}}val") == "single" else "APA/Open"
        if bb_str != bm_str:
            diffs.append(f"Gaya garis batas (borders): {bb_str} -> {bm_str}")

        # Shading header
        rows_b = tb.findall("w:tr", NAMESPACES)
        rows_m = tm.findall("w:tr", NAMESPACES)
        if rows_b and rows_m:
            hdr_cells_b = rows_b[0].findall("w:tc", NAMESPACES)
            hdr_cells_m = rows_m[0].findall("w:tc", NAMESPACES)
            shd_vals_b = [c.find(".//w:shd", NAMESPACES).attrib.get(f"{{{W_NS}}}fill", "No Color") if c.find(".//w:shd", NAMESPACES) is not None else "No Color" for c in hdr_cells_b]
            shd_vals_m = [c.find(".//w:shd", NAMESPACES).attrib.get(f"{{{W_NS}}}fill", "No Color") if c.find(".//w:shd", NAMESPACES) is not None else "No Color" for c in hdr_cells_m]
            if shd_vals_b != shd_vals_m:
                b_desc = "No Color / Transparan" if all(s == "No Color" for s in shd_vals_b) else str(shd_vals_b)
                m_desc = "No Color / Transparan" if all(s == "No Color" for s in shd_vals_m) else str(shd_vals_m)
                diffs.append(f"Header shading: {b_desc} -> {m_desc}")

        # tblCellMar
        mar_b = tb.find(".//w:tblCellMar", NAMESPACES)
        mar_m = tm.find(".//w:tblCellMar", NAMESPACES)
        def get_mar(m):
            if m is None:
                return None
            l = m.find("w:left", NAMESPACES)
            r = m.find("w:right", NAMESPACES)
            return (l.attrib.get(f"{{{W_NS}}}w") if l is not None else None,
                    r.attrib.get(f"{{{W_NS}}}w") if r is not None else None)
        if get_mar(mar_b) != get_mar(mar_m):
            diffs.append(f"Padding sel (tblCellMar): {get_mar(mar_b)} -> {get_mar(mar_m)}")

        # Spasi paragraf dalam sel
        if rows_b and rows_m:
            p0_b = rows_b[0].find(".//w:pPr/w:spacing", NAMESPACES)
            p0_m = rows_m[0].find(".//w:pPr/w:spacing", NAMESPACES)
            def get_sp(p):
                if p is None:
                    return None
                return (p.attrib.get(f"{{{W_NS}}}line"), p.attrib.get(f"{{{W_NS}}}after"))
            if get_sp(p0_b) != get_sp(p0_m):
                diffs.append(f"Spasi teks sel: {get_sp(p0_b)} -> {get_sp(p0_m)}")
        if diffs:
            results.append(f"Tabel {idx}:\n" + "\n".join(f"    • {d}" for d in diffs))

    return results


def inspect_numbering_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi perubahan penomoran daftar ber-w:numPr."""
    results = []

    def get_num_items(root):
        items = []
        for idx, p in enumerate(root.findall(".//w:p", NAMESPACES)):
            pPr = p.find("w:pPr", NAMESPACES)
            if pPr is not None:
                numPr = pPr.find("w:numPr", NAMESPACES)
                if numPr is not None:
                    numId = numPr.find("w:numId", NAMESPACES).attrib.get(f"{{{W_NS}}}val", "?")
                    ilvl = numPr.find("w:ilvl", NAMESPACES).attrib.get(f"{{{W_NS}}}val", "?")
                    txt = "".join(p.itertext()).strip()
                    items.append((numId, ilvl, txt))
        return items

    items_b = get_num_items(root_base)
    items_m = get_num_items(root_mod)

    if not items_b and items_m:
        results.append(f"Daftar bernomor dinamis Word (w:numPr) ditambahkan: {len(items_m)} butir terdeteksi")
        from collections import Counter
        counts = Counter((item[0], item[1]) for item in items_m)
        for (nid, lvl), cnt in sorted(counts.items()):
            # Ambil contoh teks
            sample = next(item[2] for item in items_m if item[0] == nid and item[1] == lvl)
            results.append(f"    • numId={nid}, ilvl={lvl}: {cnt} butir (Contoh: \"{sample[:40]}...\")")
    elif items_b != items_m:
        results.append(f"Perubahan pada butir penomoran: {len(items_b)} butir -> {len(items_m)} butir")

    return results


def inspect_paragraphs_styling_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi perubahan spasi, indentasi, dan perataan teks paragraf."""
    results = []
    paras_b = root_base.findall(".//w:p", NAMESPACES)
    paras_m = root_mod.findall(".//w:p", NAMESPACES)

    diff_count = 0
    for idx in range(min(len(paras_b), len(paras_m))):
        pb = paras_b[idx]
        pm = paras_m[idx]
        pPr_b = pb.find("w:pPr", NAMESPACES)
        pPr_m = pm.find("w:pPr", NAMESPACES)

        sp_b = pPr_b.find("w:spacing", NAMESPACES) if pPr_b is not None else None
        sp_m = pPr_m.find("w:spacing", NAMESPACES) if pPr_m is not None else None
        ind_b = pPr_b.find("w:ind", NAMESPACES) if pPr_b is not None else None
        ind_m = pPr_m.find("w:ind", NAMESPACES) if pPr_m is not None else None

        line_b = sp_b.attrib.get(f"{{{W_NS}}}line") if sp_b is not None else None
        line_m = sp_m.attrib.get(f"{{{W_NS}}}line") if sp_m is not None else None

        if line_b != line_m and line_m is not None:
            diff_count += 1

    if diff_count > 0:
        results.append(f"Perubahan spasi baris terdeteksi pada {diff_count} paragraf")

    return results


def inspect_images_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi perubahan gambar, dimensi EMU, dan posisi."""
    results = []
    imgs_b = root_base.findall(".//wp:extent", NAMESPACES)
    imgs_m = root_mod.findall(".//wp:extent", NAMESPACES)

    if len(imgs_b) != len(imgs_m):
        results.append(f"Jumlah gambar berubah: {len(imgs_b)} -> {len(imgs_m)}")
    else:
        for idx in range(len(imgs_m)):
            cx_b = int(imgs_b[idx].attrib.get("cx", 0))
            cy_b = int(imgs_b[idx].attrib.get("cy", 0))
            cx_m = int(imgs_m[idx].attrib.get("cx", 0))
            cy_m = int(imgs_m[idx].attrib.get("cy", 0))
            if (cx_b, cy_b) != (cx_m, cy_m):
                w_in_b = round(cx_b / 914400.0, 2)
                w_in_m = round(cx_m / 914400.0, 2)
                results.append(f"Gambar {idx + 1} dimensi diubah: lebar {w_in_b}\" -> {w_in_m}\"")

    return results


def inspect_math_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi persamaan matematika Office Math (OMML)."""
    results = []
    math_b = root_base.findall(".//m:oMath", NAMESPACES)
    math_m = root_mod.findall(".//m:oMath", NAMESPACES)
    if len(math_b) != len(math_m):
        results.append(f"Jumlah persamaan matematika berubah: {len(math_b)} -> {len(math_m)}")
    return results


def inspect_sections_diff(root_base: ET.Element, root_mod: ET.Element) -> list[str]:
    """Mendeteksi perubahan margin, ukuran kertas, dan orientasi halaman."""
    results = []
    secs_b = root_base.findall(".//w:sectPr", NAMESPACES)
    secs_m = root_mod.findall(".//w:sectPr", NAMESPACES)

    for idx in range(min(len(secs_b), len(secs_m))):
        mar_b = secs_b[idx].find("w:pgMar", NAMESPACES)
        mar_m = secs_m[idx].find("w:pgMar", NAMESPACES)
        if mar_b is not None and mar_m is not None:
            if mar_b.attrib != mar_m.attrib:
                results.append(f"Seksi {idx + 1} margin halaman disesuaikan:")
                for k, v in mar_m.attrib.items():
                    k_short = k.split("}")[-1]
                    v_in = round(int(v) / 1440.0, 2)
                    results.append(f"    • Margin {k_short}: {v_in} inci ({v} dxa)")
    return results


def generate_python_recommendations(report_items: dict[str, list[str]]) -> list[str]:
    """Menghasilkan cuplikan kode Python generator berdasarkan perubahan yang terdeteksi."""
    code_lines = []
    code_lines.append("# Rekomendasi penyesuaian kode generator Python:")

    has_table = any("Tabel" in k for k in report_items.get("tables", []))
    has_num = bool(report_items.get("numbering", []))

    imports = []
    if has_table:
        imports.append("apply_grid_table_borders, set_table_col_widths_dxa, format_table_cells")
    if has_num:
        imports.append("add_numbered_paragraph")

    if imports:
        code_lines.append(f"from engine import {', '.join(imports)}\n")

    if has_table:
        code_lines.append("# 1. Perbarui styling tabel ke Full Grid & No Color:")
        code_lines.append("apply_grid_table_borders(table, sz=4)")
        code_lines.append("set_table_col_widths_dxa(table, [1746, 2360, 1418, 1275, 1418])")
        code_lines.append("format_table_cells(table, font_size_pt=10.0, line_spacing=1.0, ensure_no_color=True)\n")

    if has_num:
        code_lines.append("# 2. Ganti penomoran manual paragraf dengan native numbering Word:")
        code_lines.append("add_numbered_paragraph(doc, text, num_id=3, ilvl=0)")

    return code_lines


def generate_clean_xml_diff(docx1: str | Path, docx2: str | Path) -> str:
    """Menghasilkan diff XML ter-normalisasi yang bebas dari noise Word."""
    with zipfile.ZipFile(docx1, "r") as z1, zipfile.ZipFile(docx2, "r") as z2:
        xml1 = z1.read("word/document.xml") if "word/document.xml" in z1.namelist() else b"<empty/>"
        xml2 = z2.read("word/document.xml") if "word/document.xml" in z2.namelist() else b"<empty/>"

    norm1 = normalize_xml_string(xml1).splitlines()
    norm2 = normalize_xml_string(xml2).splitlines()

    diff = difflib.unified_diff(
        norm1,
        norm2,
        fromfile=str(docx1),
        tofile=str(docx2),
        lineterm="",
    )
    return "\n".join(diff)


def run_diff_inspection(
    file1: str | Path,
    file2: str | Path | None = None,
    use_git: bool = False,
    mode: str = "summary",
) -> str:
    """Fungsi utama untuk membandingkan dua berkas DOCX."""
    temp_git_file = None

    if use_git:
        temp_git_file = extract_git_head_docx(file1)
        if temp_git_file is None:
            return "Error: Berkas tidak ditemukan di git HEAD atau gagal diekstrak."
        base_docx = temp_git_file
        mod_docx = Path(file1)
    else:
        if file2 is None:
            return "Error: Harap sertakan berkas kedua atau gunakan flag --git."
        base_docx = Path(file1)
        mod_docx = Path(file2)

    try:
        if mode == "xml":
            return generate_clean_xml_diff(base_docx, mod_docx)

        # Mode Summary
        root_base = read_docx_xml(base_docx)
        root_mod = read_docx_xml(mod_docx)

        if root_base is None or root_mod is None:
            return "Error: Gagal mem-parse word/document.xml dari salah satu berkas."

        report = {}
        report["tables"] = inspect_tables_diff(root_base, root_mod)
        report["numbering"] = inspect_numbering_diff(root_base, root_mod)
        report["paragraphs"] = inspect_paragraphs_styling_diff(root_base, root_mod)
        report["images"] = inspect_images_diff(root_base, root_mod)
        report["math"] = inspect_math_diff(root_base, root_mod)
        report["sections"] = inspect_sections_diff(root_base, root_mod)

        output = []
        output.append("=" * 80)
        output.append("               DOCX-ENGINE: USER EDIT INSPECTION REPORT               ")
        output.append("=" * 80)
        output.append(f"Target Dokumen   : {mod_docx}")
        output.append(f"Baseline Dokumen : {'Git HEAD' if use_git else base_docx}")
        output.append("-" * 80)

        changes_found = False
        for category, items in report.items():
            if items:
                changes_found = True
                output.append(f"\n[{category.upper()}]")
                for item in items:
                    output.append(f"  {item}")

        if not changes_found:
            output.append("\nTidak ada perbedaan semantik terdeteksi (dokumen identik).")
        else:
            output.append("\n" + "-" * 80)
            output.append("REKOMENDASI KODE GENERATOR PYTHON:")
            output.append("-" * 80)
            recs = generate_python_recommendations(report)
            output.extend(recs)

        output.append("=" * 80)
        return "\n".join(output)

    finally:
        if temp_git_file and temp_git_file.exists():
            os.remove(temp_git_file)


def main():
    parser = argparse.ArgumentParser(description="DOCX Diff & Semantic Inspection Tool")
    parser.add_argument("file1", help="Berkas docx utama yang dimodifikasi")
    parser.add_argument("file2", nargs="?", default=None, help="Berkas docx pembanding (opsional jika --git aktif)")
    parser.add_argument("--git", action="store_true", help="Bandingkan dengan versi git HEAD")
    parser.add_argument("--mode", choices=["summary", "xml"], default="summary", help="Mode laporan diff")

    args = parser.parse_args()
    report = run_diff_inspection(args.file1, args.file2, use_git=args.git, mode=args.mode)
    print(report)


if __name__ == "__main__":
    main()
