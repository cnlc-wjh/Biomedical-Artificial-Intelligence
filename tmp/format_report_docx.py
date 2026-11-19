from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
ET.register_namespace("w", W)


def q(name: str) -> str:
    return f"{{{W}}}{name}"


def set_attr(el: ET.Element, name: str, value: str) -> None:
    el.set(q(name), value)


def get_or_make(parent: ET.Element, tag: str) -> ET.Element:
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = ET.SubElement(parent, q(tag))
    return node


def remove_italics(style: ET.Element, root: ET.Element | None = None) -> None:
    for rpr in style.findall(".//w:rPr", NS):
        for tag in ("i", "iCs"):
            for node in rpr.findall(f"w:{tag}", NS):
                rpr.remove(node)


def border(parent: ET.Element, tag: str, val: str, size: str = "8") -> None:
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = ET.SubElement(parent, q(tag))
    set_attr(node, "val", val)
    if val != "nil":
        set_attr(node, "sz", size)
        set_attr(node, "space", "0")
        set_attr(node, "color", "000000")


def set_three_line_table(table: ET.Element) -> None:
    tbl_pr = get_or_make(table, "tblPr")
    borders = get_or_make(tbl_pr, "tblBorders")
    for tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border(borders, tag, "nil")
    border(borders, "top", "single", "8")
    border(borders, "bottom", "single", "8")

    rows = table.findall("w:tr", NS)
    for row_index, row in enumerate(rows):
        cells = row.findall("w:tc", NS)
        for cell in cells:
            tc_pr = get_or_make(cell, "tcPr")
            tc_borders = get_or_make(tc_pr, "tcBorders")
            for tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
                border(tc_borders, tag, "nil")
            if row_index == 0:
                border(tc_borders, "top", "single", "8")
                border(tc_borders, "bottom", "single", "8")
            if row_index == len(rows) - 1:
                border(tc_borders, "bottom", "single", "8")


def remove_headers_and_footers(parts: dict[str, bytes]) -> None:
    document = ET.fromstring(parts["word/document.xml"])
    for sect_pr in document.findall(".//w:sectPr", NS):
        for tag in ("headerReference", "footerReference"):
            for node in sect_pr.findall(f"w:{tag}", NS):
                sect_pr.remove(node)
    parts["word/document.xml"] = ET.tostring(document, encoding="utf-8", xml_declaration=True)

    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", rel_ns)
    if "word/_rels/document.xml.rels" in parts:
        rels = ET.fromstring(parts["word/_rels/document.xml.rels"])
        for rel in list(rels):
            target = rel.get("Target", "")
            if "header" in target.lower() or "footer" in target.lower():
                rels.remove(rel)
        parts["word/_rels/document.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", ct_ns)
    if "[Content_Types].xml" in parts:
        content_types = ET.fromstring(parts["[Content_Types].xml"])
        for override in list(content_types):
            part_name = override.get("PartName", "").lower()
            if "/word/header" in part_name or "/word/footer" in part_name:
                content_types.remove(override)
        parts["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)

    for name in list(parts):
        if re.match(r"word/(header|footer)\d+\.xml$", name):
            del parts[name]


def patch_docx(source: Path, target: Path) -> None:
    with zipfile.ZipFile(source, "r") as zin:
        parts = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    document = ET.fromstring(parts["word/document.xml"])
    for table in document.findall(".//w:tbl", NS):
        set_three_line_table(table)

    for paragraph in document.findall(".//w:body/w:p", NS):
        p_style = paragraph.find("./w:pPr/w:pStyle", NS)
        style_id = p_style.get(q("val"), "") if p_style is not None else ""
        if style_id in {"FirstParagraph", "BodyText"}:
            ppr = get_or_make(paragraph, "pPr")
            ind = get_or_make(ppr, "ind")
            set_attr(ind, "firstLine", "480")
            if q("hanging") in ind.attrib:
                del ind.attrib[q("hanging")]
        texts = paragraph.findall(".//w:t", NS)
        full_text = "".join(t.text or "" for t in texts)
        if full_text.startswith("图 "):
            for rpr in paragraph.findall(".//w:rPr", NS):
                for tag in ("i", "iCs"):
                    for node in rpr.findall(f"w:{tag}", NS):
                        rpr.remove(node)

    parts["word/document.xml"] = ET.tostring(document, encoding="utf-8", xml_declaration=True)
    remove_headers_and_footers(parts)

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: format_report_docx.py INPUT.docx OUTPUT.docx")
    patch_docx(Path(sys.argv[1]), Path(sys.argv[2]))
