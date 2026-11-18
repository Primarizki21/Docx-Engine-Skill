from __future__ import annotations

try:
    from .lexicon import acronym_map
except ImportError:
    from lexicon import acronym_map


def _contains_word(text: str, word: str) -> bool:
    """Cari kemunculan `word` sebagai kata utuh memakai string methods."""
    if not word:
        return False
    size = len(word)
    start = 0
    while True:
        index = text.find(word, start)
        if index < 0:
            return False
        left_ok = index == 0 or not text[index - 1].isalnum()
        right = index + size
        right_ok = right == len(text) or not text[right].isalnum()
        if left_ok and right_ok:
            return True
        start = index + 1


def _has_expansion(text: str, expansion: str) -> bool:
    """Cek keberadaan kepanjangan secara case-insensitive."""
    if not expansion:
        return False
    return expansion.lower() in text.lower()


def _pairs(acronyms):
    if acronyms is None:
        acronyms = acronym_map()
    return list(acronyms.items())


def expansion_required(text: str, acronyms=None) -> list[str]:
    """Akronim yang dipakai tanpa kepanjangannya di teks yang sama."""
    needed: list[str] = []
    for acronym, expansion in _pairs(acronyms):
        if not acronym or not expansion:
            continue
        if _contains_word(text, acronym) and not _has_expansion(text, expansion):
            needed.append(acronym)
    return needed


def find_dropped_expansions(before: str, after: str, acronyms=None) -> list[str]:
    """Kepanjangan yang ada sebelum rewrite namun hilang sesudahnya."""
    dropped: list[str] = []
    for acronym, expansion in _pairs(acronyms):
        if not acronym or not expansion:
            continue
        if (
            _has_expansion(before, expansion)
            and _contains_word(after, acronym)
            and not _has_expansion(after, expansion)
        ):
            dropped.append(acronym)
    return dropped


def assert_no_dropped_expansions(before: str, after: str, acronyms=None) -> None:
    """Naikkan ValueError bila rewrite menghilangkan kepanjangan akronim."""
    dropped = find_dropped_expansions(before, after, acronyms)
    if dropped:
        raise ValueError(
            "Kepanjangan akronim hilang setelah rewrite: " + ", ".join(dropped)
        )
