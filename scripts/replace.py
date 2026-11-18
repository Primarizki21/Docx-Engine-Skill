from __future__ import annotations

from typing import Iterable

from docx.text.paragraph import Paragraph


def _is_heading(p: Paragraph) -> bool:
    name = str(getattr(p.style, "name", "") or "").lower()
    return name.startswith("heading") or name == "title"


def _replace_range_in_runs(runs, start_idx: int, end_idx: int, new_text: str) -> None:
    """Replace [start_idx, end_idx) across runs, preserving run formatting."""
    bounds = []
    cur = 0
    for r in runs:
        bounds.append((r, cur, cur + len(r.text)))
        cur += len(r.text)
    hits = [i for i, (_, a, b) in enumerate(bounds) if a < end_idx and b > start_idx]
    if not hits:
        return
    first, last = hits[0], hits[-1]
    r_first, a_first, _ = bounds[first]
    r_last, a_last, _ = bounds[last]
    off_first, off_last = max(start_idx - a_first, 0), end_idx - a_last
    if first == last:
        r_first.text = r_first.text[:off_first] + new_text + r_first.text[off_last:]
    else:
        r_first.text = r_first.text[:off_first] + new_text
        for i in range(first + 1, last):
            bounds[i][0].text = ""
        r_last.text = r_last.text[off_last:]


def replace_exact_text_in_paragraph(p: Paragraph, old: str, new: str, expected: int = 1) -> int:
    """Guarded exact-match replacement in one paragraph; run-preserving, idempotent."""
    if not old or old == new or old not in p.text:
        return 0
    found = p.text.count(old)
    if found > expected:
        raise ValueError(f"expected <= {expected} occurrence(s) of {old!r}, found {found}")
    if not p.runs:
        p.text = p.text.replace(old, new)  # ponytail: mutation-guard-exempt
    else:
        starts = []
        idx = p.text.find(old)
        while idx != -1:
            starts.append(idx)
            idx = p.text.find(old, idx + len(old))
        for start in reversed(starts):
            _replace_range_in_runs(p.runs, start, start + len(old), new)
    if old not in new:
        assert old not in p.text
    assert new in p.text
    return found


def _table_paragraphs(tables) -> Iterable[Paragraph]:
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def _iter_all_paragraphs(doc) -> Iterable[Paragraph]:
    yield from doc.paragraphs
    yield from _table_paragraphs(doc.tables)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            yield from part.paragraphs
            yield from _table_paragraphs(part.tables)


def replace_exact_text(doc, old: str, new: str, expected: int = 1) -> int:
    """Guarded exact-match replacement across paragraphs, tables and headers/footers."""
    if not old or old == new:
        return 0
    paragraphs = list(_iter_all_paragraphs(doc))
    total = sum(p.text.count(old) for p in paragraphs)
    if total == 0:
        return 0
    if total > expected:
        raise ValueError(f"expected <= {expected} occurrence(s) of {old!r}, found {total}")
    changed = sum(replace_exact_text_in_paragraph(p, old, new, expected=expected) for p in paragraphs)
    assert changed == total
    if old not in new:
        assert all(old not in p.text for p in paragraphs)
    assert any(new in p.text for p in paragraphs)
    return changed


def replace_placeholders(doc, mapping: dict[str, str], expected: int = 1) -> dict[str, int]:
    """Replace each placeholder via the guarded helper; missing keys are idempotent."""
    results = {}
    for placeholder, val in mapping.items():
        results[placeholder] = replace_exact_text(doc, placeholder, val, expected=expected)
    return results


def replace_in_scope(document, start_heading: str, old: str, new: str, expected: int = 1) -> int:
    """Replace only in paragraphs below the unique start heading, up to the next heading."""
    paragraphs = list(document.paragraphs)
    starts = [i for i, p in enumerate(paragraphs) if p.text == start_heading and _is_heading(p)]
    if len(starts) != 1:
        raise ValueError(f"expected exactly one heading {start_heading!r}, found {len(starts)}")
    begin = starts[0] + 1
    end = begin
    while end < len(paragraphs) and not _is_heading(paragraphs[end]):
        end += 1
    scope = paragraphs[begin:end]
    total = sum(p.text.count(old) for p in scope)
    if total == 0:
        return 0
    if total > expected:
        raise ValueError(f"expected <= {expected} occurrence(s) of {old!r}, found {total} in scope")
    changed = sum(replace_exact_text_in_paragraph(p, old, new, expected=expected) for p in scope)
    assert changed == total
    return changed
