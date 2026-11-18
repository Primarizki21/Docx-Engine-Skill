from __future__ import annotations

import argparse
import hashlib
import subprocess
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from lxml import etree
try:
    from .core_xml import qn
    from .textconv import canonical_diff
except ImportError:
    from core_xml import qn
    from textconv import canonical_diff

_PARSER = etree.XMLParser(resolve_entities=False)

def _git_baseline(docx_path, workdir):
    source = Path(docx_path).resolve()
    root = subprocess.run(["git", "rev-parse", "--show-toplevel"], cwd=source.parent, capture_output=True, text=True)
    if root.returncode:
        raise RuntimeError(f"Git tidak tersedia untuk {source}: {root.stderr.strip()}")
    repo = Path(root.stdout.strip())
    rel = source.relative_to(repo).as_posix()
    show = subprocess.run(["git", "show", f"HEAD:{rel}"], cwd=repo, capture_output=True)
    if show.returncode or not show.stdout:
        raise RuntimeError(f"Baseline HEAD tidak tersedia untuk {rel}: {show.stderr.decode(errors='replace').strip()}")
    target = Path(workdir) / "baseline.docx"
    target.write_bytes(show.stdout)
    return target

def compare_with_baseline(docx_path, baseline_path=None) -> list[str]:
    """Delta baris kanonik terhadap baseline eksplisit atau versi HEAD di git."""
    with TemporaryDirectory(prefix="docx-regression-") as tmp:
        baseline = Path(baseline_path) if baseline_path is not None else _git_baseline(docx_path, tmp)
        return list(canonical_diff(baseline, docx_path))

def _root(docx_path):
    with zipfile.ZipFile(docx_path) as archive:
        return etree.fromstring(archive.read("word/document.xml"), _PARSER)

def _attr(parent, tag, default=None):
    el = parent.find(qn(tag)) if parent is not None else None
    return el.get(qn("w:val"), default) if el is not None else None

def _run_signature(root):
    out = []
    for p in root.iter(qn("w:p")):
        runs = [(_attr(r.find(qn("w:rPr")), "w:rStyle"),
                 (_attr(r.find(qn("w:rPr")), "w:b", "1"), _attr(r.find(qn("w:rPr")), "w:i", "1")),
                 (_attr(r.find(qn("w:rPr")), "w:sz"), _attr(r.find(qn("w:rPr")), "w:szCs")))
                for r in p.findall(qn("w:r"))]
        out.append((_attr(p.find(qn("w:pPr")), "w:pStyle"), tuple(runs)))
    return out

def _image_anchors(docx_path):
    with zipfile.ZipFile(docx_path) as archive:
        names = set(archive.namelist())
        root = etree.fromstring(archive.read("word/document.xml"), _PARSER)
        rels_path = "word/_rels/document.xml.rels"
        rels = etree.fromstring(archive.read(rels_path), _PARSER) if rels_path in names else []
        targets = {rel.get("Id"): rel.get("Target", "") for rel in rels}
        anchors = []
        for box in list(root.iter(qn("wp:inline"))) + list(root.iter(qn("wp:anchor"))):
            extent, blip = box.find(qn("wp:extent")), box.find(".//" + qn("a:blip"))
            digest = None
            if blip is not None:
                rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
                name = targets.get(rid, "").lstrip("/")
                name = name if name.startswith("word/") else "word/" + name
                if name in names:
                    digest = hashlib.sha256(archive.read(name)).hexdigest()
            anchors.append((digest, extent.get("cx") if extent is not None else None, extent.get("cy") if extent is not None else None))
        return anchors

def _table_geometry(root):
    out = []
    for tbl in root.iter(qn("w:tbl")):
        width = tbl.find(qn("w:tblPr")).find(qn("w:tblW")) if tbl.find(qn("w:tblPr")) is not None else None
        grid = tbl.find(qn("w:tblGrid"))
        cols = tuple(c.get(qn("w:w")) for c in grid.findall(qn("w:gridCol"))) if grid is not None else ()
        cells = []
        for tc in (tc for tr in tbl.findall(qn("w:tr")) for tc in tr.findall(qn("w:tc"))):
            pr = tc.find(qn("w:tcPr"))
            cw = pr.find(qn("w:tcW")) if pr is not None else None
            cells.append((cw.get(qn("w:w")), cw.get(qn("w:type"))) if cw is not None else None)
        tbl_w = (width.get(qn("w:w")), width.get(qn("w:type"))) if width is not None else None
        out.append((tbl_w, cols, tuple(cells)))
    return out

def _probe(label, before, after):
    out = [f"Jumlah {label.lower()} berubah: {len(before)} -> {len(after)}"] if len(before) != len(after) else []
    out += [f"{label} {i} berubah: {a!r} -> {b!r}" for i, (a, b) in enumerate(zip(before, after), 1) if a != b]
    return out

def audit_regression(docx_path, baseline_path=None) -> list[str]:
    """Gabungkan delta kanonik dengan tiga probe struktur dokumen."""
    with TemporaryDirectory(prefix="docx-regression-") as tmp:
        baseline = Path(baseline_path) if baseline_path is not None else _git_baseline(docx_path, tmp)
        findings = list(canonical_diff(baseline, docx_path))
        old, new = _root(baseline), _root(docx_path)
        findings += _probe("Tanda tangan run paragraf (run signature)", _run_signature(old), _run_signature(new))
        findings += _probe("Jangkar gambar (image anchor)", _image_anchors(baseline), _image_anchors(docx_path))
        findings += _probe("Geometri tabel (table geometry)", _table_geometry(old), _table_geometry(new))
        return findings

def main() -> None:
    parser = argparse.ArgumentParser(description="Audit regresi DOCX terhadap baseline")
    parser.add_argument("docx")
    parser.add_argument("--baseline")
    args = parser.parse_args()
    findings = audit_regression(args.docx, args.baseline)
    if findings:
        print("\n".join(findings))
    raise SystemExit(1 if findings else 0)

if __name__ == "__main__":
    main()
