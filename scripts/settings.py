from __future__ import annotations

from pathlib import Path
from lxml import etree
try:
    from .core_xml import NAMESPACES, W_NS, qn
except ImportError:
    from core_xml import NAMESPACES, W_NS, qn


def enable_update_fields_in_unpacked(unpacked_dir: str | Path) -> bool:
    settings_path = Path(unpacked_dir) / "word" / "settings.xml"
    if not settings_path.exists():
        # Buat settings.xml minimal jika belum ada
        root = etree.Element(f"{{{W_NS}}}settings", nsmap={"w": W_NS})
    else:
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(settings_path), parser)
        root = tree.getroot()

    existing = root.find(qn("w:updateFields"))
    if existing is None:
        elem = etree.SubElement(root, qn("w:updateFields"))
        elem.set(qn("w:val"), "true")
    else:
        existing.set(qn("w:val"), "true")

    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings_path.write_bytes(etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone="yes"))
    return True


def enable_update_fields_in_doc(document) -> None:
    # Memasang flag updateFields langsung ke objek Document python-docx jika memungkinkan
    try:
        settings_elem = document.settings.element
        update_fields = settings_elem.find(qn("w:updateFields"))
        if update_fields is None:
            update_fields = etree.SubElement(settings_elem, qn("w:updateFields"))
        update_fields.set(qn("w:val"), "true")
    except Exception:
        pass
