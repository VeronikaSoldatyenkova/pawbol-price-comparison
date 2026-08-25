"""Build-time smoke tests for the production price-comparison image."""
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

from main import health
assert health()["status"] == "ok"
print("Price Comparison smoke test: OK")
