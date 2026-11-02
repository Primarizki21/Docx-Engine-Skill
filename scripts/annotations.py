from __future__ import annotations

import zipfile
from pathlib import Path
import lxml.etree as etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _extract_runs_annotations(container, location_prefix: str, context_text: str) -> list[dict]:
    items = []
    # 1. Highlights & custom font colors in runs
    for r in container.findall(f".//{{{W_NS}}}r"):
        rPr = r.find(f"{{{W_NS}}}rPr")
        if rPr is None:
            continue
        txt = "".join(r.itertext()).strip()
        if not txt:
            continue
        hl = rPr.find(f"{{{W_NS}}}highlight")
        if hl is not None:
            val = hl.get(f"{{{W_NS}}}val") or "highlight"
            items.append({
                "kind": "highlight", "location": location_prefix, "color": val,
                "text": txt, "context": context_text, "author": "User"
            })
        col = rPr.find(f"{{{W_NS}}}color")
        if col is not None:
            val = (col.get(f"{{{W_NS}}}val") or "").upper()
            if val and val not in ("000000", "AUTO"):
                items.append({
                    "kind": "color", "location": location_prefix, "color": val,
                    "text": txt, "context": context_text, "author": "User"
                })

    # 2. Track changes (w:ins & w:del)
    for tag_name, kind, col in [("ins", "insertion", "green"), ("del", "deletion", "red")]:
        for el in container.findall(f".//{{{W_NS}}}{tag_name}"):
            txt = "".join(el.itertext()).strip()
            if txt:
                author = el.get(f"{{{W_NS}}}author") or "User"
                items.append({
                    "kind": kind, "location": location_prefix, "color": col,
                    "text": txt, "context": context_text, "author": author
                })
    return items


def extract_document_annotations(docx_path: str | Path) -> list[dict]:
    """Mengekstrak seluruh teks ber-highlight, warna revisi, komentar, dan track changes."""
    path = Path(docx_path)
    if not path.exists():
        return []

    annotations: list[dict] = []
    active_heading = "Awal Naskah"

    try:
        with zipfile.ZipFile(path, "r") as z:
            # 1. Periksa word/document.xml
            if "word/document.xml" in z.namelist():
                doc_xml = z.read("word/document.xml")
                root = etree.fromstring(doc_xml)
                body = root.find(f"{{{W_NS}}}body")
                if body is not None:
                    p_idx = 0
                    t_idx = 0
                    for child in body:
                        tag = child.tag.split("}")[-1]
                        if tag == "p":
                            p_idx += 1
                            p_text = "".join(child.itertext()).strip()
                            pPr = child.find(f"{{{W_NS}}}pPr")
                            pStyle = pPr.find(f"{{{W_NS}}}pStyle") if pPr is not None else None
                            style_val = (pStyle.get(f"{{{W_NS}}}val") if pStyle is not None else "").lower()
                            if "heading" in style_val or p_text.upper().startswith("BAB "):
                                active_heading = p_text[:40] if p_text else style_val
                            loc = f"[{active_heading}] Paragraf #{p_idx}"
                            annotations.extend(_extract_runs_annotations(child, loc, p_text))
                        elif tag == "tbl":
                            t_idx += 1
                            for r_i, row in enumerate(child.findall(f".//{{{W_NS}}}tr")):
                                for c_i, cell in enumerate(row.findall(f".//{{{W_NS}}}tc")):
                                    c_text = "".join(cell.itertext()).strip()
                                    loc = f"[{active_heading}] Tabel #{t_idx} (R{r_i} C{c_i})"
                                    annotations.extend(_extract_runs_annotations(cell, loc, c_text))

            # 2. Periksa word/comments.xml
            if "word/comments.xml" in z.namelist():
                c_xml = z.read("word/comments.xml")
                c_root = etree.fromstring(c_xml)
                for c in c_root.findall(f".//{{{W_NS}}}comment"):
                    author = c.get(f"{{{W_NS}}}author") or "User"
                    c_txt = "".join(c.itertext()).strip()
                    if c_txt:
                        annotations.append({
                            "kind": "comment", "location": "Word Comments Part",
                            "color": "comment", "text": c_txt, "context": c_txt, "author": author
                        })
    except Exception:
        return []

    return annotations


def format_annotations_report(annotations: list[dict]) -> list[str]:
    """Memformat daftar anotasi pengguna menjadi baris laporan CLI terstruktur."""
    if not annotations:
        return ["Tidak ada sorotan highlight, komentar, atau revisi terdeteksi (dokumen bersih)."]
    lines = [f"Ditemukan {len(annotations)} butir sorotan reviu pengguna:"]
    for a in annotations:
        tag = a["color"].upper() if a["kind"] in ("highlight", "color") else a["kind"].upper()
        author_str = f" oleh {a['author']}" if a["author"] and a["author"] != "User" else ""
        lines.append(f"  • [{tag}]{author_str} di {a['location']}:")
        lines.append(f"    Teks: \"{a['text']}\"")
        if a["context"] and a["context"] != a["text"]:
            ctx = a["context"][:100] + "..." if len(a["context"]) > 100 else a["context"]
            lines.append(f"    Konteks: \"{ctx}\"")
    return lines
