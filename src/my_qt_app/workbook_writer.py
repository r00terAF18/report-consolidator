"""Update report values without reserializing the original Excel drawings/styles."""
import posixpath
import re
from math import isclose
from datetime import date, datetime, time
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import to_excel


def check_saved_values(path: Path, assignments: dict) -> dict:
    """Read the saved values before reporting a successful import."""
    saved = load_workbook(path, data_only=False)
    checked = 0
    nonempty = 0
    try:
        for sheet, cells in assignments.items():
            for coordinate, expected in cells.items():
                actual = saved[sheet][coordinate].value
                equal = actual == expected
                if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
                    equal = isclose(actual, expected, rel_tol=1e-12, abs_tol=1e-9)
                if not equal:
                    raise ValueError(f'Saved value differs at {sheet}!{coordinate}: expected {expected!r}, found {actual!r}')
                checked += 1
                nonempty += expected is not None
        return {'status': 'matched', 'assigned_cells': checked, 'nonempty_values': nonempty}
    finally:
        saved.close()


def save_report_package(target: Path, output: Path, wb, changes: dict) -> None:
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    docrel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    def xml_text(value):
        return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(value))
    with ZipFile(target) as original:
        workbook_xml = original.read("xl/workbook.xml").decode("utf-8")
        relationships_xml = original.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        workbook = ET.fromstring(workbook_xml)
        relationships = ET.fromstring(relationships_xml)
        targets = {r.attrib["Id"]: posixpath.normpath("xl/" + r.attrib["Target"])
                   if not r.attrib["Target"].startswith("/") else r.attrib["Target"].lstrip("/")
                   for r in relationships}
        updated = {}
        for sheet in workbook.find(f"{{{ns}}}sheets"):
            name = sheet.attrib["name"]
            if name not in changes:
                continue
            path = targets[sheet.attrib[f"{{{docrel}}}id"]]
            source = original.read(path).decode("utf-8")
            for coord, value in changes[name].items():
                pattern = rf'<c\b(?=[^>]*\br="{coord}")[^>]*?(?:/>|>.*?</c>)'
                match = re.search(pattern, source, re.S)
                attributes = dict(ET.fromstring(match[0]).attrib) if match else {"r": coord}
                attributes.pop("t", None)
                cell = ET.Element("c", attributes)
                if value is not None:
                    if isinstance(value, str):
                        cell.set("t", "inlineStr")
                        text = ET.SubElement(ET.SubElement(cell, "is"), "t")
                        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        text.text = xml_text(value)
                    else:
                        cell.set("t", "n")
                        ET.SubElement(cell, "v").text = str(to_excel(value, wb.epoch) if isinstance(value, (time, date, datetime)) else value)
                replacement = ET.tostring(cell, encoding="unicode")
                if match:
                    source = source[:match.start()] + replacement + source[match.end():]
                else:
                    row_number, column = coordinate_to_tuple(coord)
                    row_match = re.search(rf'<row\b(?=[^>]*\br="{row_number}")[^>]*?(?:/>|>.*?</row>)', source, re.S)
                    if not row_match:
                        raise ValueError(f"Refusing to create a new target row for {name}: {coord}")
                    row_xml = row_match[0]
                    if row_xml.endswith("/>"):
                        row_xml = row_xml[:-2] + "></row>"
                    next_cell = next((m for m in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"', row_xml)
                                      if coordinate_to_tuple(m[1])[1] > column), None)
                    at = next_cell.start() if next_cell else row_xml.rfind("</row>")
                    row_xml = row_xml[:at] + replacement + row_xml[at:]
                    source = source[:row_match.start()] + row_xml + source[row_match.end():]
            updated[path] = source.encode("utf-8")
        with ZipFile(output, "w") as result:
            for entry in original.infolist():
                result.writestr(entry, updated.pop(entry.filename, original.read(entry.filename)))
            for name, content in updated.items():
                result.writestr(name, content)
