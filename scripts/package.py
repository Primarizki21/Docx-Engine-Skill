from __future__ import annotations

import argparse
import os
import sys
import zipfile
from pathlib import Path
from lxml import etree
try:
    from .core_xml import ensure_xml_space
except ImportError:
    from core_xml import ensure_xml_space


def unpack_docx(docx_path: str | Path, output_dir: str | Path) -> Path:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as z:
        z.extractall(out)
    return out


def pack_docx(source_dir: str | Path, output_docx: str | Path, auto_repair: bool = True) -> Path:
    src = Path(source_dir)
    target = Path(output_docx)
    target.parent.mkdir(parents=True, exist_ok=True)

    if auto_repair:
        doc_xml = src / "word" / "document.xml"
        if doc_xml.exists():
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(str(doc_xml), parser)
            ensure_xml_space(tree.getroot())
            doc_xml.write_bytes(etree.tostring(tree, encoding="utf-8", xml_declaration=True, standalone="yes"))

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src):
            for file in files:
                full_path = Path(root) / file
                arcname = full_path.relative_to(src)
                z.write(full_path, arcname)
    return target


def parse_xml_part(unpacked_dir: str | Path, relative_path: str) -> tuple[etree._Element, Path]:
    target_path = Path(unpacked_dir) / relative_path
    if not target_path.exists():
        raise FileNotFoundError(f"Part not found: {relative_path}")
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(target_path), parser)
    return tree.getroot(), target_path


def save_xml_part(root: etree._Element, target_path: str | Path) -> None:
    data = etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone="yes")
    Path(target_path).write_bytes(data)


def validate_xml_parts(unpacked_dir: str | Path) -> list[str]:
    errors = []
    base = Path(unpacked_dir)
    for p in base.rglob("*.xml"):
        try:
            etree.parse(str(p))
        except Exception as e:
            errors.append(f"{p.relative_to(base)}: {e}")
    return errors


def main() -> None:
    parser = argparse.ArgumentParser(description="DOCX Package Utilities")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_unpack = sub.add_parser("unpack", help="Unpack docx to folder")
    p_unpack.add_argument("docx", help="Input docx path")
    p_unpack.add_argument("out_dir", help="Output directory")

    p_pack = sub.add_parser("pack", help="Pack folder to docx")
    p_pack.add_argument("in_dir", help="Input directory")
    p_pack.add_argument("docx", help="Output docx path")

    p_val = sub.add_parser("validate", help="Validate unpacked XML")
    p_val.add_argument("in_dir", help="Directory to validate")

    args = parser.parse_args()
    if args.cmd == "unpack":
        out = unpack_docx(args.docx, args.out_dir)
        print(f"Unpacked to {out}")
    elif args.cmd == "pack":
        target = pack_docx(args.in_dir, args.docx)
        print(f"Packed to {target}")
    elif args.cmd == "validate":
        errs = validate_xml_parts(args.in_dir)
        if errs:
            print("Validation errors found:")
            for err in errs:
                print(f"  - {err}")
            sys.exit(1)
        print("All XML parts are valid.")


if __name__ == "__main__":
    main()
