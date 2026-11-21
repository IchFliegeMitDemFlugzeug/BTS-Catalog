#!/usr/bin/env python3
"""Check VBA loading and report what the headless interpreter can actually verify."""

from pathlib import Path

from pyopenvba.apps.excel import ExcelApplication

# Build output is located at the repository root.
workbook = Path(__file__).resolve().parents[1] / "BTS.xlsm"

try:
    # Opening triggers Workbook_Open and BTS_Init, just like desktop Excel.
    print("Checking workbook open and initialization.")
    app = ExcelApplication.open(str(workbook))

    # If opening succeeds, test the public initializer once more.
    app.run("BTS_Init")
    print("VBA interpreter completed BTS_Init.")
except Exception as exc:
    # Distinguish invalid VBA syntax from a missing method in this test engine.
    kind = type(exc).__name__
    print(f"Execution-model result: {kind}: {exc}")

    # Compilation and syntax errors are genuine blockers: fail the workflow.
    if "Compile" in kind or "Syntax" in kind:
        raise

    # A missing Excel UI operation does not establish a failure in real Excel.
    if "Unsupported" in kind:
        print("A desktop Excel test is still required; UI support is incomplete here.")
    else:
        # Any other failure must not be silently represented as a passing test.
        raise
