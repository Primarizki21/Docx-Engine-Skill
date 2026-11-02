from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

MODULE_FILES = [
    "__init__.py",
    "core_xml.py",
    "package.py",
    "settings.py",
    "sections.py",
    "headings.py",
    "paragraphs.py",
    "typography.py",
    "lists.py",
    "tables.py",
    "media.py",
    "headers_footers.py",
    "references.py",
    "toc.py",
    "captions.py",
    "equations.py",
    "diff_docx.py",
    "replace.py",
    "annotations.py",
]


def scaffold_engine(target_dir: str | Path, overwrite: bool = False) -> Path:
    dest = Path(target_dir).resolve()
    dest.mkdir(parents=True, exist_ok=True)
    src_dir = Path(__file__).parent.resolve()

    copied = 0
    for file_name in MODULE_FILES:
        src_file = src_dir / file_name
        dest_file = dest / file_name

        if not src_file.exists():
            continue

        if dest_file.exists() and not overwrite:
            print(f"Skipping existing: {dest_file.name}")
            continue

        shutil.copy2(src_file, dest_file)
        copied += 1

    print(f"Successfully scaffolded {copied} docx-engine modules to: {dest}")
    print("\nNext steps in your target project:")
    print("  1. Add dependencies: uv add python-docx lxml")
    print("  2. Import modules: from <target_path> import tables, headings, toc, etc.")
    return dest


def main() -> None:
    parser = argparse.ArgumentParser(description="Scaffold docx-engine modules into any project.")
    parser.add_argument("--target", "-t", required=True, help="Target destination directory")
    parser.add_argument("--force", "-f", action="store_true", help="Overwrite existing files in target")
    args = parser.parse_args()

    scaffold_engine(target_dir=args.target, overwrite=args.force)


if __name__ == "__main__":
    main()
