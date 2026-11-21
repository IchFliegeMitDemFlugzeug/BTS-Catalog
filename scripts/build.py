#!/usr/bin/env python3
# Build BTS.xlsm while retaining the nine designed worksheets of BTS.xlsx.

from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "assets" / "BTS.xlsx"
OUTPUT = ROOT / "BTS.xlsm"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("x", MAIN)
ET.register_namespace("r", OFFICE_REL)
ET.register_namespace("", CONTENT_TYPES)

def append_override(root, name, mime):
    # Add or replace an OOXML content-type declaration.
    for item in root.findall("{" + CONTENT_TYPES + "}Override"):
        if item.get("PartName") == name:
            item.set("ContentType", mime)
            return
    ET.SubElement(root, "{" + CONTENT_TYPES + "}Override", PartName=name, ContentType=mime)

def modify_xml(filename, data):
    # Modify only package metadata needed to embed VBA. Leave worksheet content as-is.
    root = ET.fromstring(data)
    if filename == "[Content_Types].xml":
        append_override(root, "/xl/workbook.xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        append_override(root, "/xl/vbaProject.bin", "application/vnd.ms-office.vbaProject")
    elif filename == "xl/_rels/workbook.xml.rels":
        ET.SubElement(
            root, "{" + RELATIONS + "}Relationship",
            Id="rIdBTSVBA",
            Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            Target="vbaProject.bin",
        )
    elif filename == "xl/workbook.xml":
        pr = root.find("{" + MAIN + "}workbookPr")
        if pr is None:
            pr = ET.Element("{" + MAIN + "}workbookPr")
            root.insert(0, pr)
        pr.set("codeName", "ThisWorkbook")
    elif filename == "xl/worksheets/sheet1.xml":
        pr = root.find("{" + MAIN + "}sheetPr")
        if pr is None:
            pr = ET.Element("{" + MAIN + "}sheetPr")
            root.insert(0, pr)
        pr.set("codeName", "Sheet1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

# A short-lived donor workbook supplies a real, compiled-structure VBA project.
with TemporaryDirectory() as temp:
    donor = Path(temp) / "donor.xlsm"
    with ExcelFile.create_new(str(donor)) as book:
        book.vba_project().add_module("BTS", (ROOT / "src" / "BTS.bas").read_text(encoding="utf-8"))
        book.vba_project().add_module("Localization", (ROOT / "src" / "Localization.bas").read_text(encoding="utf-8"))
        book.set_module("ThisWorkbook", (ROOT / "src" / "ThisWorkbook.cls").read_text(encoding="utf-8"))
        book.save()

    # Copy the VBA binary into the nine-sheet original workbook without recreating sheets.
    with ZipFile(donor) as macro_zip:
        vba_binary = macro_zip.read("xl/vbaProject.bin")
    with ZipFile(BASE) as base, ZipFile(OUTPUT, "w", compression=ZIP_DEFLATED) as out:
        assert base.testzip() is None
        for item in base.infolist():
            data = base.read(item.filename)
            if item.filename in {
                "[Content_Types].xml",
                "xl/_rels/workbook.xml.rels",
                "xl/workbook.xml",
                "xl/worksheets/sheet1.xml",
            }:
                data = modify_xml(item.filename, data)
            out.writestr(item, data)
        out.writestr("xl/vbaProject.bin", vba_binary, compress_type=ZIP_DEFLATED)

# Verify OOXML packaging and embedded VBA source survive the round trip.
with ZipFile(OUTPUT) as zipfile:
    assert zipfile.testzip() is None
    names = zipfile.namelist()
    assert "xl/vbaProject.bin" in names
    assert len([x for x in names if x.startswith("xl/worksheets/sheet") and x.endswith(".xml")]) == 9
    assert zipfile.read("xl/vbaProject.bin") == vba_binary
    assert b"rIdBTSVBA" in zipfile.read("xl/_rels/workbook.xml.rels")
    assert b"macroEnabled" in zipfile.read("[Content_Types].xml")

with ExcelFile(str(OUTPUT)) as result:
    names = set(result.module_names())
    assert {"BTS", "Localization", "ThisWorkbook"} <= names, names
    assert "Public Sub BTS_AddItem()" in result.get_module("BTS")
    assert "Public Sub BTS_ExportOrder()" in result.get_module("BTS")
    assert "Private Sub Workbook_Open()" in result.get_module("ThisWorkbook")

print(f"Built {OUTPUT.name}: {OUTPUT.stat().st_size} bytes, nine sheets, real VBA.")
