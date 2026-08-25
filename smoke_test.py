"""Build-time smoke tests for the production price-comparison image."""
import tempfile
from pathlib import Path

import pandas as pd

from core import (
    build_code_lookups,
    compare_all,
    create_excel,
    dataframe_to_table_model,
    filter_result_by_codes,
    parse_requested_codes,
)

powerbi = pd.DataFrame({
    "EAN": ["111", "222"],
    "SKU": ["A", "B"],
    "Latest Pricelist Value": [10.0, 20.0],
    "Realisation Summ": [1.0, 1.0],
    "Net Available Qty": [0, 0],
    "Days Since Last Sale": [3, 5],
    "MinStock": [0, 0],
    "MaxStock": [2, 2],
})
supplier = pd.DataFrame({"EAN": ["111", "999"], "SKU": ["A", "B"], "Price": [8.0, 18.0]})
config = {
    "supplier_name": "Test Supplier",
    "file_name": "supplier.xlsx",
    "sheet_name": "Sheet1",
    "match_method": "EAN + SKU",
    "ean_column": "EAN",
    "sku_column": "SKU",
    "price_column": "Price",
}
result, supplier_cols, duplicate_info, _ = compare_all(
    powerbi, {"type": "PowerBI Pricelist"}, [(supplier, config)]
)
assert not result.columns.duplicated().any()
assert result["Matched Suppliers"].eq(1).all()
assert set(result["Cheapest Supplier"]) == {"Test Supplier"}

lookup_ean, lookup_sku = build_code_lookups(result)
requests = parse_requested_codes("111\t9\nB\t19\nNOPE\t10")
quick = filter_result_by_codes(result, requests, lookup_ean, lookup_sku, supplier_cols)
assert list(quick["Requested Code"]) == ["111", "B", "NOPE"]
assert quick.iloc[2]["Lookup Status"] == "CODE NOT FOUND"
assert quick.iloc[0]["Below Target"] == "Test Supplier"
dataframe_to_table_model(quick, supplier_cols, target_mode=True)

excel = create_excel(
    result, supplier_cols, [config], duplicate_info, "current.xlsx", {"type": "PowerBI Pricelist"}
)
assert len(excel) > 1000

free_current = pd.DataFrame({
    "Barcode": ["111", "222"],
    "Article": ["A", "B"],
    "NetPrice": [10.0, 20.0],
    "Brand": ["X", "Y"],
})
free_result, _, _, _ = compare_all(
    free_current,
    {
        "type": "Free format pricelist",
        "ean_column": "Barcode",
        "sku_column": "Article",
        "price_column": "NetPrice",
        "extra_columns": ["Brand"],
    },
    [(supplier, config)],
)
assert "Brand" in free_result.columns
assert "Realisation Summ" not in free_result.columns

from main import _normalise_excel_columns, _preview_payload, health
from runtime_main import (
    _authoritative_sheet_columns,
    _displayed_excel_bytes,
    _robust_workbook_metadata,
    _sort_displayed_result,
)

blank_headers = pd.DataFrame(
    [["x", "y", "A"]],
    columns=["Unnamed: 0", "", "SKU"],
)
blank_headers = _normalise_excel_columns(blank_headers)
assert list(blank_headers.columns) == ["Column1", "Column2", "SKU"]
preview = _preview_payload(blank_headers)
assert preview["columns"] == ["Column1", "Column2", "SKU"]
assert preview["rows_shown"] == 1

# Regression: a supplier workbook may contain a title only in A1 while the
# real Code / Designation / Price data begins several rows later.
with tempfile.TemporaryDirectory() as temp_dir:
    path = Path(temp_dir) / "supplier_blank_headers.xlsx"
    raw = pd.DataFrame(
        [
            ["NEGOWATT - Price list", None, None],
            ["2026-08-03", None, None],
            [None, None, None],
            ["Code", "Designation", "Price"],
            ["A9A15215", "Transformer", 29.697],
        ]
    )
    raw.to_excel(path, index=False, header=False)

    metadata = _robust_workbook_metadata(str(path))
    cols = metadata["sheets"][0]["columns"]
    assert cols == ["NEGOWATT - Price list", "Column1", "Column2"], cols

    file_meta = {
        "name": path.name,
        "path": str(path),
        "metadata": metadata,
    }
    actual_cols = _authoritative_sheet_columns(file_meta, metadata["sheets"][0]["name"])
    assert actual_cols == ["NEGOWATT - Price list", "Column1", "Column2"], actual_cols

# Displayed-result download must export the current filtered table, including
# Quick Filter target columns, and must respect a selected browser sort.
sorted_quick = _sort_displayed_result(quick, "Target Price", "desc")
assert list(sorted_quick["Target Price"].dropna()) == sorted(
    list(sorted_quick["Target Price"].dropna()), reverse=True
)
shown_excel = _displayed_excel_bytes(sorted_quick, supplier_cols)
assert len(shown_excel) > 1000

assert health()["status"] == "ok"
print("Price Comparison smoke test: OK")
