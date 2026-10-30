from __future__ import annotations

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NAMESPACES = {
    "w": W_NS,
    "r": R_NS,
    "wp": WP_NS,
    "a": A_NS,
    "pic": PIC_NS,
    "m": M_NS,
    "xml": XML_NS,
}

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing",
    "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
    "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]


def qn(tag: str) -> str:
    prefix, local = tag.split(":") if ":" in tag else ("", tag)
    if not prefix:
        return local
    uri = NAMESPACES.get(prefix)
    if not uri:
        raise ValueError(f"Unknown namespace prefix: {prefix}")
    return f"{{{uri}}}{local}"


def insert_ordered_p_pr(p_pr_elem: etree._Element, new_elem: etree._Element) -> etree._Element:
    new_local = etree.QName(new_elem).localname
    new_idx = PPR_ORDER.index(new_local) if new_local in PPR_ORDER else 999

    for child in p_pr_elem:
        child_local = etree.QName(child).localname
        child_idx = PPR_ORDER.index(child_local) if child_local in PPR_ORDER else 999
        if child_idx > new_idx:
            child.addprevious(new_elem)
            return new_elem
    p_pr_elem.append(new_elem)
    return new_elem


def ensure_xml_space(root_elem: etree._Element) -> None:
    for t_node in root_elem.xpath(".//w:t", namespaces=NAMESPACES):
        text = t_node.text or ""
        if text.startswith(" ") or text.endswith(" "):
            t_node.set(f"{{{XML_NS}}}space", "preserve")


def get_next_rel_id(rels_root: etree._Element) -> str:
    max_id = 0
    for rel in rels_root.findall(f"{{{R_NS}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            max_id = max(max_id, int(rid[3:]))
    return f"rId{max_id + 1}"
