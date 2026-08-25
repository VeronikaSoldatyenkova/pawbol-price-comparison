"""Production entrypoint with robust Excel-column discovery.

Some supplier workbooks have a title only in A1 while B1/C1/... are blank.
Pandas with nrows=0 then reports only the first column even though later rows
contain data.  The UI preview reads the real sheet and sees those columns as
Column1, Column2, ... .  Patch the FastAPI module so metadata and validation
use a representative/authoritative column set as well.
"""

from pathlib import Path

import pandas as pd

import main as _main


METADATA_SCAN_ROWS = 250


def _robust_workbook_metadata(path: str) -> dict:
    try:
        xls = pd.ExcelFile(path)
    except Exception as exc:
        raise ValueError(f"Could not open workbook: {exc}") from exc

    sheets = []
    for sheet in xls.sheet_names:
        try:
            # nrows=0 can lose columns whose first-row header is blank.
            # A small data scan reveals the actual used worksheet width while
            # remaining much cheaper than loading an entire large pricelist.
            sample = pd.read_excel(
                path,
                sheet_name=sheet,
                nrows=METADATA_SCAN_ROWS,
                dtype=object,
            )
            sample = _main._normalise_excel_columns(sample)
            columns = [str(c) for c in sample.columns]
        except Exception:
            columns = []
        sheets.append({"name": str(sheet), "columns": columns})
    return {"sheets": sheets}


def _authoritative_sheet_columns(file_meta, sheet_name):
    """Return the same normalized columns that comparison will actually use."""
    try:
        df = _main._read_sheet(file_meta, sheet_name)
        return [str(c) for c in df.columns]
    except Exception:
        # Keep the metadata fallback so validation can still produce a useful
        # message if the workbook itself becomes unreadable.
        for sheet in file_meta.get("metadata", {}).get("sheets", []):
            if sheet.get("name") == sheet_name:
                return list(sheet.get("columns") or [])
        return []


_main._workbook_metadata = _robust_workbook_metadata
_main._sheet_columns = _authoritative_sheet_columns

# Uvicorn imports this object.
app = _main.app
