from __future__ import annotations

from typing import Iterable
from docx.text.paragraph import Paragraph


def _replace_range_in_runs(runs, start_idx: int, end_idx: int, new_text: str) -> None:
    cur = 0
    intervals = []
    for r in runs:
        rlen = len(r.text)
        intervals.append((r, cur, cur + rlen))
        cur += rlen

    first_idx = None
    last_idx = None
    for i, (r, r_start, r_end) in enumerate(intervals):
        if r_start <= start_idx < r_end or (start_idx == r_start and start_idx == end_idx):
            if first_idx is None:
                first_idx = i
        if r_start < end_idx <= r_end:
            last_idx = i

    if first_idx is None or last_idx is None:
        return

    r_first, r_fstart, _ = intervals[first_idx]
    r_last, r_lstart, _ = intervals[last_idx]

    offset_in_first = start_idx - r_fstart
    offset_in_last = end_idx - r_lstart

    if first_idx == last_idx:
        r_first.text = r_first.text[:offset_in_first] + new_text + r_first.text[offset_in_last:]
    else:
        r_first.text = r_first.text[:offset_in_first] + new_text
        for k in range(first_idx + 1, last_idx):
            intervals[k][0].text = ""
        r_last.text = r_last.text[offset_in_last:]


def replace_exact_text_in_paragraph(p: Paragraph, old_text: str, new_text: str) -> int:
    """Mengganti teks persis (exact match) di dalam satu paragraf dengan menjaga formatting runs."""
    if not old_text or old_text == new_text or old_text not in p.text:
        return 0

    if not p.runs:
        p.text = p.text.replace(old_text, new_text)
        return 1

    start_indices = []
    idx = p.text.find(old_text)
    while idx != -1:
        start_indices.append(idx)
        idx = p.text.find(old_text, idx + len(old_text))

    for start_idx in reversed(start_indices):
        end_idx = start_idx + len(old_text)
        _replace_range_in_runs(p.runs, start_idx, end_idx, new_text)

    return len(start_indices)


def _iter_all_paragraphs(doc) -> Iterable[Paragraph]:
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            yield p
        for t in sec.header.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
        for p in sec.footer.paragraphs:
            yield p
        for t in sec.footer.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


def replace_exact_text(doc, old_text: str, new_text: str) -> int:
    """Mengganti teks exact match di seluruh bagian dokumen (paragraf, tabel, header/footer)."""
    if not old_text or old_text == new_text:
        return 0
    total_replaced = 0
    for p in _iter_all_paragraphs(doc):
        total_replaced += replace_exact_text_in_paragraph(p, old_text, new_text)
    return total_replaced


def replace_placeholders(doc, mapping: dict[str, str]) -> dict[str, int]:
    """Mengganti sekumpulan placeholder template exact match secara aman tanpa regex."""
    results = {}
    for placeholder, val in mapping.items():
        results[placeholder] = replace_exact_text(doc, placeholder, val)
    return results
