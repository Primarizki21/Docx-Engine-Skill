from __future__ import annotations

import argparse
import ast
import sys
from pathlib import Path

EXEMPT_MARKER = "ponytail: mutation-guard-exempt"
_REGEX_CALLS = frozenset({"sub", "match", "search", "split"})


def _python_files(path: Path):
    if path.is_file():
        if path.suffix == ".py":
            yield path
        return
    for child in sorted(path.rglob("*.py")):
        yield child


def _flagged_call(func: ast.expr) -> str | None:
    if not isinstance(func, ast.Attribute):
        return None
    if func.attr == "replace":
        return ".replace"
    if (
        func.attr in _REGEX_CALLS
        and isinstance(func.value, ast.Name)
        and func.value.id == "re"
    ):
        return "re." + func.attr
    return None


def _exempt_lines(lines: list[str]) -> set[int]:
    return {number for number, line in enumerate(lines, 1) if EXEMPT_MARKER in line}


def scan_source(path: str | Path) -> list[str]:
    target = Path(path)
    findings: list[tuple[str, int, str]] = []

    for source in _python_files(target):
        text = source.read_text(encoding="utf-8")
        exempt = _exempt_lines(text.splitlines())
        try:
            tree = ast.parse(text, filename=str(source))
        except (SyntaxError, ValueError):
            continue

        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            label = _flagged_call(node.func)
            if label is None:
                continue
            end = node.end_lineno or node.lineno
            if any(line in exempt for line in range(node.lineno, end + 1)):
                continue
            findings.append((str(source), node.lineno, label))

    return [f"{source}:{line}: {label}" for source, line, label in sorted(findings)]


def main() -> None:
    parser = argparse.ArgumentParser(description="Pemindai AST untuk mutasi teks berisiko")
    parser.add_argument(
        "path",
        nargs="?",
        default=str(Path(__file__).resolve().parent),
        help="Berkas .py atau direktori yang dipindai (default: direktori scripts)",
    )
    args = parser.parse_args()

    findings = scan_source(args.path)
    for finding in findings:
        print(finding)
    raise SystemExit(1 if findings else 0)


if __name__ == "__main__":
    main()
