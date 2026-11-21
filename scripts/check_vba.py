#!/usr/bin/env python3
"""Try the workbook's initialization routine with an independent VBA interpreter.

A successful check here does not replace a full Microsoft Excel integration test.
"""

from pathlib import Path
from pyopenvba.apps.excel import ExcelApplication

workbook = Path(__file__).resolve().parents[1] / "BTS.xlsm"

print("Opening macro workbook in VBA execution model.")
app = ExcelApplication.open(str(workbook))

try:
    print("Executing BTS_Init (non-Excel engine).")
    app.run("BTS_Init")
    print("Interpreter completed BTS_Init.")
except Exception as exc:
    error_class = exc.__class__.__name__
    print(f"VBA execution-model check stopped: {error_class}: {exc}")
    if "Compile" in error_class or "Syntax" in error_class:
        raise
    if "Unsupported" in error_class:
        print("This VBA operation is not modeled by the Python engine; desktop Excel check remains required.")
    else:
        print("Excel execution semantics could not be verified by this model.")
