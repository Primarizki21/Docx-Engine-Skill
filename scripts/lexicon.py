from __future__ import annotations

import json
from pathlib import Path

LEXICON_PATH = Path(__file__).resolve().parent.parent / "data" / "lexicon.json"


def load_lexicon() -> dict:
    """Baca SSOT leksikon dari data/lexicon.json."""
    with LEXICON_PATH.open(encoding="utf-8") as handle:
        return json.load(handle)


def forbidden_map() -> dict[str, str]:
    """Petakan bentuk terlarang ke istilah baku yang disetujui."""
    return dict(load_lexicon().get("forbidden", {}))


def acronym_map() -> dict[str, str]:
    """Petakan akronim yang disetujui ke kepanjangannya."""
    return dict(load_lexicon().get("acronyms", {}))


def _matching_terms() -> list[tuple[str, str]]:
    """Pasangan (huruf kecil, teks asli) diurutkan dari pencocokan terpanjang."""
    terms = []
    for term in load_lexicon().get("italic_terms", []):
        if term:
            terms.append((term.lower(), term))
    return sorted(terms, key=lambda item: (-len(item[0]), item[0]))


def italic_segments(text: str) -> list[tuple[str, bool]]:
    """Pecah teks menjadi (potongan asli, italic) untuk istilah terdaftar.

    Pencocokan memakai huruf kecil tanpa syarat batas kata, dengan prioritas
    istilah terpanjang lebih dulu agar "buffer gap" menang atas "gap".
    """
    if not text:
        return []

    terms = _matching_terms()
    length = len(text)
    segments: list[tuple[str, bool]] = []
    plain_start = 0
    index = 0

    while index < length:
        matched = 0
        for needle, original in terms:
            size = len(original)
            if text[index:index + size].lower() == needle:
                matched = size
                break
        if matched == 0:
            index += 1
            continue

        if index > plain_start:
            segments.append((text[plain_start:index], False))
        end = index + matched
        segments.append((text[index:end], True))
        index = end
        plain_start = index

    if plain_start < length:
        segments.append((text[plain_start:], False))
    return segments
