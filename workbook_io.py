"""Small OOXML edits that preserve cached formula results and workbook formatting.

Loading and saving an entire file with openpyxl discards formula caches. Portfolio
exports instead replace only edited text cells and application-owned hidden sheets.
"""
from __future__ import annotations

import datetime as dt
import io
import math
import numbers
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl.utils import get_column_letter


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
XML = "http://www.w3.org/XML/1998/namespace"
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def tag(name):
    return f"{{{MAIN}}}{name}"


def parse_xml(data):
    # Keep namespace prefixes referenced by mc:Ignorable and extension attributes.
    declarations = []
    for _, pair in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
        prefix, uri = pair
        declarations.append((prefix, uri))
        if not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)
    root = ET.fromstring(data)
    used = set()
    for node in root.iter():
        for qualified in [node.tag, *node.attrib]:
            if qualified.startswith("{"):
                used.add(qualified[1:].split("}", 1)[0])
    for prefix, uri in declarations:
        if prefix and uri not in used:
            root.set(f"xmlns:{prefix}", uri)
    return root


def xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sheet_paths(archive):
    book = parse_xml(archive.read("xl/workbook.xml"))
    relationships = parse_xml(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in relationships if r.get("TargetMode") != "External"}
    result = {}
    for sheet in book.find(tag("sheets")):
        target = targets.get(sheet.get(f"{{{REL}}}id"))
        if target:
            result[sheet.get("name")] = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target))
    return result


def clean_text(value):
    return ILLEGAL.sub(" ", "" if value is None else str(value)).replace("\r\n", "\n").replace("\r", "\n")[:32767]


def write_text(cell, value):
    for child in list(cell):
        cell.remove(child)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, tag("is"))
    node = ET.SubElement(inline, tag("t"), {f"{{{XML}}}space": "preserve"})
    node.text = clean_text(value)


def _rewrite(archive, replacements, removed=()):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as dest:
        for item in archive.infolist():
            if item.filename not in removed:
                dest.writestr(item, replacements.pop(item.filename, archive.read(item.filename)))
        for name, data in replacements.items():
            dest.writestr(name, data)
    return output.getvalue()


def patch_text_cells(file_bytes: bytes, edits: dict[str, dict[str, str]]) -> bytes:
    if not any(edits.values()):
        return file_bytes
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as archive:
        paths = sheet_paths(archive)
        replacements = {}
        for sheet, changes in edits.items():
            if not changes:
                continue
            if sheet not in paths:
                raise ValueError("An edited sheet is missing from the workbook.")
            root = parse_xml(archive.read(paths[sheet]))
            data = root.find(tag("sheetData"))
            if data is None:
                raise ValueError("An edited sheet contains no worksheet data.")
            rows = {int(row.get("r")): row for row in data}
            for coordinate, value in changes.items():
                match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", coordinate)
                if not match:
                    raise ValueError("Invalid export cell coordinate.")
                row_number = int(match[2])
                row = rows.get(row_number)
                if row is None:
                    row = ET.Element(tag("row"), {"r": str(row_number)})
                    data.append(row)
                    rows[row_number] = row
                cell = next((c for c in row if c.get("r") == coordinate), None)
                if cell is None:
                    cell = ET.SubElement(row, tag("c"), {"r": coordinate})
                write_text(cell, value)
                row[:] = sorted(row, key=lambda c: (len(re.sub(r"\d", "", c.get("r", ""))), re.sub(r"\d", "", c.get("r", ""))))
            data[:] = sorted(data, key=lambda row: int(row.get("r", 0)))
            replacements[paths[sheet]] = xml_bytes(root)
        return _rewrite(archive, replacements)


def _sheet_xml(frame: pd.DataFrame) -> bytes:
    root = ET.Element(tag("worksheet"))
    data = ET.SubElement(root, tag("sheetData"))
    for row_number, values in enumerate([list(frame.columns), *frame.itertuples(index=False, name=None)], 1):
        row = ET.SubElement(data, tag("row"), {"r": str(row_number)})
        for col, value in enumerate(values, 1):
            cell = ET.SubElement(row, tag("c"), {"r": f"{get_column_letter(col)}{row_number}"})
            missing = value is None
            if not missing:
                try:
                    missing = bool(pd.isna(value))
                except (ValueError, TypeError):
                    pass
            if missing:
                write_text(cell, "")
            elif isinstance(value, bool):
                cell.set("t", "b")
                ET.SubElement(cell, tag("v")).text = "1" if value else "0"
            elif isinstance(value, numbers.Real) and math.isfinite(float(value)):
                ET.SubElement(cell, tag("v")).text = str(value)
            else:
                if isinstance(value, (dt.date, dt.datetime, pd.Timestamp)):
                    value = value.isoformat()
                write_text(cell, value)
    return xml_bytes(root)


def replace_hidden_sheets(file_bytes: bytes, frames: dict[str, pd.DataFrame]) -> bytes:
    if not frames:
        return file_bytes
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as archive:
        paths = sheet_paths(archive)
        book = parse_xml(archive.read("xl/workbook.xml"))
        rels = parse_xml(archive.read("xl/_rels/workbook.xml.rels"))
        content_types = parse_xml(archive.read("[Content_Types].xml"))
        sheets = book.find(tag("sheets"))
        replacements = {}
        used_ids = {r.get("Id") for r in rels}
        used_paths = set(archive.namelist())
        next_sheet_id = max((int(s.get("sheetId", 0)) for s in sheets), default=0) + 1
        for name, frame in frames.items():
            if not name.startswith("_") or len(name) > 31:
                raise ValueError("Only application-owned hidden sheets can be replaced.")
            existing = next((s for s in sheets if " ".join(s.get("name", "").casefold().split()) == " ".join(name.casefold().split())), None)
            if existing is not None:
                path = paths[existing.get("name")]
                existing.set("name", name)
                existing.set("state", "hidden")
            else:
                seq = 1
                while f"xl/worksheets/wpo{seq}.xml" in used_paths:
                    seq += 1
                path = f"xl/worksheets/wpo{seq}.xml"
                used_paths.add(path)
                rid_no = 1
                while f"rIdWpo{rid_no}" in used_ids:
                    rid_no += 1
                rid = f"rIdWpo{rid_no}"
                used_ids.add(rid)
                ET.SubElement(rels, f"{{{PKG}}}Relationship", {"Id": rid, "Type": f"{REL}/worksheet", "Target": path.removeprefix("xl/")})
                ET.SubElement(sheets, tag("sheet"), {"name": name, "sheetId": str(next_sheet_id), "state": "hidden", f"{{{REL}}}id": rid})
                next_sheet_id += 1
            if not any(o.get("PartName") == f"/{path}" for o in content_types):
                ET.SubElement(content_types, f"{{{TYPES}}}Override", {"PartName": f"/{path}", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
            replacements[path] = _sheet_xml(frame)
        replacements["xl/workbook.xml"] = xml_bytes(book)
        replacements["xl/_rels/workbook.xml.rels"] = xml_bytes(rels)
        replacements["[Content_Types].xml"] = xml_bytes(content_types)
        return _rewrite(archive, replacements)
