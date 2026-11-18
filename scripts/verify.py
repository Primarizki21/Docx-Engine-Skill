from __future__ import annotations

import argparse
import sys
import time
from tempfile import TemporaryDirectory

import docx

try:
    from . import acronyms, citation_qa, lexicon, mutation_guard, package, reference_qa, sentence_metrics, textconv
except ImportError:
    import acronyms
    import citation_qa
    import lexicon
    import mutation_guard
    import package
    import reference_qa
    import sentence_metrics
    import textconv

BUDGET_SECONDS = 35


def _text_lines(docx_path: str) -> list[str]:
    """Ambil teks baris kanonik tanpa prefiks JENIS|."""
    lines = textconv.canonical_lines(docx_path)
    return [line.split("|", 1)[1] if "|" in line else line for line in lines]


def _forbidden_findings(lines: list[str]) -> list[str]:
    findings = []
    for term, approved in lexicon.forbidden_map().items():
        needle = term.lower()
        for line in lines:
            if needle in line.lower():
                findings.append(f"bentuk terlarang '{term}' -> '{approved}': {line}")
                break
    return sorted(findings)


def _acronym_findings(lines: list[str]) -> list[str]:
    used = acronyms.expansion_required("\n".join(lines))
    return [f"akronim tanpa kepanjangan: {name}" for name in sorted(set(used))]


def _xml_findings(docx_path: str) -> list[str]:
    with TemporaryDirectory(prefix="docx-verify-") as unpacked:
        package.unpack_docx(docx_path, unpacked)
        return package.validate_xml_parts(unpacked)


def run_checks(docx_path=None, source_dir="scripts", strict=True) -> list[str]:
    """Jalankan gerbang QA berurutan; berhenti pada gerbang pertama yang menemukan masalah."""
    findings = mutation_guard.scan_source(source_dir)
    if findings:
        return findings
    if docx_path is None:
        return []

    findings = _xml_findings(docx_path)
    if findings:
        return findings

    lines = _text_lines(docx_path)
    findings = _forbidden_findings(lines)
    if findings:
        return findings

    findings = _acronym_findings(lines)
    if findings:
        return findings

    findings = sentence_metrics.audit(lines, strict=strict)
    if findings:
        return findings

    document = docx.Document(docx_path)
    findings = citation_qa.audit_claim_citations(document, strict=strict)
    if findings:
        return findings

    return reference_qa.audit_references(document)


def main() -> None:
    parser = argparse.ArgumentParser(description="Runner QA fail-fast untuk dokumen DOCX")
    parser.add_argument("--docx", help="Path dokumen DOCX yang diperiksa")
    parser.add_argument("--source", default="scripts", help="Direktori sumber untuk mutation_guard")
    parser.add_argument(
        "--no-strict",
        action="store_true",
        help="Kembalikan temuan audit sebagai daftar, tanpa ValueError",
    )
    args = parser.parse_args()

    started = time.perf_counter()
    try:
        findings = run_checks(args.docx, args.source, strict=not args.no_strict)
    except ValueError as exc:
        findings = [str(exc)]
    elapsed = time.perf_counter() - started

    for finding in findings:
        print(finding)
    if elapsed > BUDGET_SECONDS:
        print(f"Peringatan: verifikasi melewati anggaran {BUDGET_SECONDS} detik ({elapsed:.1f} detik).")

    raise SystemExit(1 if findings else 0)


if __name__ == "__main__":
    main()
