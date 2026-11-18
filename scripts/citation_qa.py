from __future__ import annotations

from pathlib import Path

try:
    from .sentence_metrics import content_word_count, split_sentences
except ImportError:
    from sentence_metrics import content_word_count, split_sentences


_CLAIM = "klaim"
_DESIGN = "rancangan penelitian ini"
_DERIVATION = "turunan/simpulan"
_TRANSITION = "transisi/rujukan internal"
_DEFINITION = "definisi metode lazim"

_TRANSITION_MARKERS = (
    "gambar ", "tabel ", "persamaan ", "bagian ini", "bab ini",
    "lihat gambar", "lihat tabel", "seperti ditunjukkan", "pada bagian",
    "selanjutnya", "berikutnya", "pertama,", "kedua,", "terakhir,",
)
_EXTERNAL_MARKERS = (
    "penelitian sebelumnya", "studi sebelumnya", "penelitian terdahulu",
    "studi terdahulu", "literatur",
)
_DESIGN_MARKERS = (
    "penelitian ini", "studi ini", "kajian ini", "tujuan penelitian",
    "kami mengusulkan", "kami menggunakan", "kami menerapkan",
    "kami mengembangkan", "kami melakukan", "kami mengevaluasi",
    "kami menguji", "metode kami", "pendekatan kami", "ruang lingkup",
)
_DERIVATION_MARKERS = (
    "dengan demikian", "sehingga", "oleh karena itu", "dari hasil",
    "kami menyimpulkan", "kesimpulannya", "dapat disimpulkan",
    "hal ini menunjukkan",
)
_DEFINITION_MARKERS = (
    "adalah", "merupakan", "didefinisikan sebagai", "dihitung sebagai",
    "diartikan sebagai", "dirumuskan sebagai", "dikenal sebagai",
    "disebut juga", "yaitu",
)


def _classify(sentence: str) -> str:
    """Golongkan kalimat ke salah satu dari lima kategori QA."""
    lower = sentence.lower()
    for markers, kind in (
        (_TRANSITION_MARKERS, _TRANSITION),
        (_EXTERNAL_MARKERS, _CLAIM),
        (_DESIGN_MARKERS, _DESIGN),
        (_DERIVATION_MARKERS, _DERIVATION),
        (_DEFINITION_MARKERS, _DEFINITION),
    ):
        if any(marker in lower for marker in markers):
            return kind
    return _CLAIM


def _has_year_citation(text: str) -> bool:
    """Benar bila ada kurung berisi tahun empat digit."""
    start = 0
    while start < len(text):
        opening, closing = text.find("(", start), text.find(")", start + 1)
        if opening < 0 or closing < opening:
            return False
        inner, index = text[opening + 1:closing], 0
        while index < len(inner):
            if inner[index].isdigit():
                end = index
                while end < len(inner) and inner[end].isdigit():
                    end += 1
                if end - index == 4 and 1000 <= int(inner[index:end]) <= 2099:
                    return True
                index = end
            else:
                index += 1
        start = closing + 1
    return False


def _extract_texts(document) -> list[str]:
    """Terima path docx, objek Document, daftar teks, atau satu string."""
    if isinstance(document, (str, Path)):
        path = Path(document)
        if not path.is_file():
            return [str(document)]
        from docx import Document
        return [p.text for p in Document(str(path)).paragraphs]
    if isinstance(document, (list, tuple)):
        return [str(item) for item in document]
    paragraphs = getattr(document, "paragraphs", None)
    return [p.text for p in paragraphs] if paragraphs is not None else [str(document)]


def audit_claim_citations(document, strict: bool = False) -> list[str]:
    """Laporkan klaim tanpa sitasi tahun dan sitasi yang tak melekat pada klaim."""
    findings: list[str] = []
    for text in _extract_texts(document):
        for line in str(text).splitlines():
            for sentence in split_sentences(line):
                sentence = sentence.strip()
                if (
                    not sentence
                    or not any(ch.islower() for ch in sentence)
                    or sentence[-1] not in ".!?)"
                ):
                    continue
                cited = _has_year_citation(sentence)
                if cited and content_word_count(sentence) < 2:
                    findings.append(f"Sitasi tanpa klaim: {sentence}")
                    continue
                kind = _classify(sentence)
                if kind == _CLAIM and not cited:
                    findings.append(f"Klaim tanpa sitasi: {sentence}")
                elif kind != _CLAIM and cited:
                    findings.append(f"Sitasi pada {kind}: {sentence}")
    if strict and findings:
        raise ValueError("; ".join(findings))
    return findings
