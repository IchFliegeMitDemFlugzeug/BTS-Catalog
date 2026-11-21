#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a macro-enabled BTS workbook from versioned VBA modules."""

from __future__ import annotations

# Path is used to locate the checked-in source files and generated workbook.
from pathlib import Path

# ZipFile checks that Excel sees the expected VBA project package part.
from zipfile import ZipFile

# ExcelFile writes a real VBA project into an Excel workbook.
from pyopenvba import ExcelFile

# The repository root is one directory above scripts/.
ROOT = Path(__file__).resolve().parents[1]

# This is the workbook delivered to users; do not use a fake .xlsm extension.
OUTPUT = ROOT / "BTS.xlsm"

# Read the main VBA module without changing its source encoding.
main = (ROOT / "src" / "BTS.bas").read_text(encoding="utf-8")

# Read the Unicode localization module, generated as code-point expressions.
labels = (ROOT / "src" / "Localization.bas").read_text(encoding="utf-8")

# Read the ThisWorkbook event handler for initialization and dependent lists.
workbook_events = (ROOT / "src" / "ThisWorkbook.cls").read_text(encoding="utf-8")

# Create a fresh, valid Excel macro workbook with a real VBA project.
with ExcelFile.create_new(str(OUTPUT)) as wb:
    # Keep the business logic as a standard, independently editable module.
    wb.vba_project().add_module("BTS", main)

    # Keep Russian interface strings in their own standard VBA module.
    wb.vba_project().add_module("Localization", labels)

    # Attach event handlers to Excel's existing ThisWorkbook document module.
    wb.set_module("ThisWorkbook", workbook_events)

    # Write the VBA-enabled workbook to disk.
    wb.save()

# Reopen via the VBA parser, not by simply trusting the output extension.
with ExcelFile(str(OUTPUT)) as result:
    # Check that all three modules survived round-tripping.
    names = set(result.module_names())
    assert {"BTS", "Localization", "ThisWorkbook"} <= names, names

    # Make sure the principal entry points are still part of VBA code.
    assert "Public Sub BTS_AddItem()" in result.get_module("BTS")
    assert "Public Sub BTS_ExportOrder()" in result.get_module("BTS")
    assert "Private Sub Workbook_Open()" in result.get_module("ThisWorkbook")

# Check the Microsoft Excel container, including its genuine embedded VBA part.
with ZipFile(OUTPUT) as archive:
    # A macro workbook must include this binary project object.
    assert "xl/vbaProject.bin" in archive.namelist()

    # The VBA project cannot be an empty placeholder.
    assert len(archive.read("xl/vbaProject.bin")) > 1000

# Report the exact produced file and its byte size for CI logs.
print(f"Built {OUTPUT.name}: {OUTPUT.stat().st_size} bytes; modules: {', '.join(sorted(names))}")
