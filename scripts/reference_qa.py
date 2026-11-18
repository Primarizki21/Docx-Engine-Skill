from __future__ import annotations

_PREFIXES = ("Gambar", "Tabel")


def _number(text: str) -> str | None:
    if not text or not text[0].isdigit():
        return None
    index = 0
    while index < len(text) and text[index].isdigit():
        index += 1
    if index == len(text) or text[index] != ".":
        return None
    index += 1
    start = index
    while index < len(text) and text[index].isdigit():
        index += 1
    return text[:index] if index > start else None


def _leading_key(text: str, prefix: str) -> str | None:
    needle = prefix.lower()
    if text.lower().find(needle) != 0:
        return None
    number = _number(text[len(needle):].lstrip())
    return f"{prefix} {number}" if number is not None else None


def _mentions(text: str, prefix: str) -> list[str]:
    lowered = text.lower()
    needle = prefix.lower()
    keys: list[str] = []
    start = 0
    while True:
        position = lowered.find(needle, start)
        if position < 0:
            return keys
        number = _number(text[position + len(needle):].lstrip())
        if number is not None:
            key = f"{prefix} {number}"
            if key not in keys:
                keys.append(key)
        start = position + len(needle)


def _equation_numbers(text: str) -> list[str]:
    numbers: list[str] = []
    start = 0
    while True:
        opening = text.find("(", start)
        if opening < 0:
            return numbers
        closing = text.find(")", opening + 1)
        if closing < 0:
            return numbers
        inner = text[opening + 1:closing]
        number = _number(inner)
        if number is not None and len(number) == len(inner) and number not in numbers:
            numbers.append(number)
        start = closing + 1


def _entries(document):
    paragraphs = getattr(document, "paragraphs", document)
    for item in paragraphs:
        text = item.text if hasattr(item, "text") else item
        yield str(text or "").strip()


def _collect(document):
    captions: list[str] = []
    equations: list[str] = []
    references: list[str] = []
    for text in _entries(document):
        if not text:
            continue
        caption = None
        for prefix in _PREFIXES:
            caption = _leading_key(text, prefix)
            if caption is not None:
                break
        if caption is not None:
            if caption not in captions:
                captions.append(caption)
            continue
        numbers = _equation_numbers(text)
        if len(numbers) == 1 and text == f"({numbers[0]})":
            if numbers[0] not in equations:
                equations.append(numbers[0])
            continue
        for prefix in _PREFIXES:
            for key in _mentions(text, prefix):
                if key not in references:
                    references.append(key)
        for number in numbers:
            key = f"({number})"
            if key not in references:
                references.append(key)
    return captions, equations, references


def audit_references(document) -> list[str]:
    """Temukan rujukan tanpa objek serta caption tanpa rujukan narasi."""
    captions, equations, references = _collect(document)
    findings: list[str] = []
    for key in references:
        exists = key[1:-1] in equations if key.startswith("(") else key in captions
        if not exists:
            findings.append(f"Rujukan tanpa objek: {key}")
    for key in captions:
        if key not in references:
            findings.append(f"Caption tanpa rujukan narasi: {key}")
    return findings
