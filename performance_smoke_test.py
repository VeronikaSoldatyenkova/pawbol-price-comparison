"""Smoke checks for the large-file comparison performance layer."""
import tempfile
from pathlib import Path

import pandas as pd

from performance_optimization import _read_selected_columns
from production_app import app


# The production app must expose exactly one optimized handler on the URLs used
# by the existing templates.
compare_routes = [
    r for r in app.router.routes
    if getattr(r, "path", None) == "/compare/{workspace_id}"
    and "POST" in (getattr(r, "methods", set()) or set())
]
assert len(compare_routes) == 1, len(compare_routes)

download_routes = [
    r for r in app.router.routes
    if getattr(r, "path", None) == "/download/{workspace_id}"
    and "GET" in (getattr(r, "methods", set()) or set())
]
assert len(download_routes) == 1, len(download_routes)

assert any(getattr(r, "path", None) == "/compare-progress/{workspace_id}" for r in app.router.routes)
assert any(getattr(r, "path", None) == "/api/compare-progress/{workspace_id}" for r in app.router.routes)

# Selected-column reading must preserve normalized mapping names and avoid
# materializing unrelated supplier columns.
with tempfile.TemporaryDirectory() as temp_dir:
    path = Path(temp_dir) / "large_supplier.xlsx"
    pd.DataFrame(
        {
            "SKU": ["A", "B"],
            "Description": ["One", "Two"],
            "Price": [10.0, 20.0],
            "Unused": [1, 2],
        }
    ).to_excel(path, index=False)

    file_meta = {
        "name": path.name,
        "path": str(path),
        "metadata": {
            "sheets": [
                {"name": "Sheet1", "columns": ["SKU", "Description", "Price", "Unused"]}
            ]
        },
    }
    selected = _read_selected_columns(file_meta, "Sheet1", ["SKU", "Price"])
    assert list(selected.columns) == ["SKU", "Price"]
    assert len(selected) == 2

print("Performance smoke test: OK")
