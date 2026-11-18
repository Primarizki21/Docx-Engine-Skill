from __future__ import annotations

DEFAULT_LIMIT = 20
_MIN_RUN = 3
_ABBREVIATIONS = ("dkk.", "dll.", "dsb.", "s.d.")
_TERMINATORS = ".!?"

def _ends_with_abbreviation(text: str, end: int) -> bool:
    tail = text[:end].lower()
    for abbr in _ABBREVIATIONS:
        if tail.endswith(abbr):
            start = end - len(abbr)
            if start == 0 or not text[start - 1].isalnum():
                return True
    return False

def _is_sentence_end(text: str, index: int) -> bool:
    if text[index] != ".":
        return True
    prev = text[index - 1] if index else ""
    nxt = text[index + 1] if index + 1 < len(text) else ""
    if (prev.isdigit() and nxt.isdigit()) or _ends_with_abbreviation(text, index + 1):
        return False
    return not nxt.islower()

def split_sentences(text: str) -> list[str]:
    """Pecah kalimat tanpa memotong singkatan, desimal, atau isi kurung."""
    if not text or not text.strip():
        return []
    sentences: list[str] = []
    start = depth = 0
    for index, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")":
            depth = max(0, depth - 1)
        elif char in _TERMINATORS and depth == 0 and _is_sentence_end(text, index):
            end = index + 1
            sentence = text[start:end].strip()
            if sentence:
                sentences.append(sentence)
            start = end
    remainder = text[start:].strip()
    if remainder:
        sentences.append(remainder)
    return sentences

def _without_year_citations(sentence: str) -> str:
    """Buang sitasi kurung yang memuat tahun empat digit."""
    kept: list[str] = []
    group: list[str] = []
    depth = 0
    for char in sentence:
        if char == "(":
            if depth == 0:
                group = ["("]
            else:
                group.append(char)
            depth += 1
        elif char == ")":
            group.append(char)
            depth = max(0, depth - 1)
            if depth == 0:
                inside = "".join(group)
                if not any(inside[i:i + 4].isdigit() for i in range(len(inside) - 3)):
                    kept.append(inside)
                group = []
        elif depth:
            group.append(char)
        else:
            kept.append(char)
    if group:
        kept.append("".join(group))
    return "".join(kept)

def content_word_count(sentence: str) -> int:
    """Hitung kata isi, abaikan sitasi kurung bertahun empat digit."""
    cleaned = _without_year_citations(sentence)
    return sum(1 for token in cleaned.split() if any(c.isalnum() for c in token))

def paragraph_metrics(text: str) -> dict:
    """Ukur kalimat paragraf memakai batas panjang bawaan 20 kata isi."""
    sentences = split_sentences(text)
    counts = [content_word_count(s) for s in sentences]
    best = current = 0
    for count in counts:
        current = current + 1 if count > DEFAULT_LIMIT else 0
        best = max(best, current)
    return {
        "sentences": sentences, "word_counts": counts, "max_words": max(counts, default=0),
        "long_run": best if best >= _MIN_RUN else 0,
    }

def audit(texts, limit: int = DEFAULT_LIMIT, strict: bool = False) -> list[str]:
    """Laporkan kalimat melebihi batas dan rentetan tiga kalimat panjang."""
    if isinstance(texts, str):
        texts = [texts]
    findings: list[str] = []
    sequence: list[tuple[str, int]] = []
    for paragraph, text in enumerate(texts, 1):
        for sentence, count in enumerate(paragraph_metrics(text)["word_counts"], 1):
            label = f"paragraf {paragraph} kalimat {sentence}"
            sequence.append((label, count))
            if count > limit:
                findings.append(f"{label}: {count} kata isi melebihi batas {limit}")
    run_start: int | None = None
    for index in range(len(sequence) + 1):
        count = sequence[index][1] if index < len(sequence) else 0
        if count > limit:
            if run_start is None:
                run_start = index
        elif run_start is not None:
            if index - run_start >= _MIN_RUN:
                labels = ", ".join(item[0] for item in sequence[run_start:index])
                findings.append(f"{index - run_start} kalimat berurutan melebihi {limit} kata isi: {labels}")
            run_start = None
    if findings and strict:
        raise ValueError(f"{len(findings)} temuan metrik kalimat: " + "; ".join(findings))
    return findings
