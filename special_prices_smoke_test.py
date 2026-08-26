"""Regression tests for Special Prices From Supplier quick-filter mode."""

import pandas as pd

import core
from special_prices import (
    BELOW_SPECIAL,
    SPECIAL_PRICE,
    _recalculate_legrand_price,
    _special_table_model,
    build_special_price_result,
)


assert abs(_recalculate_legrand_price(10.0, 20.0) - 13.75) < 1e-12

full = pd.DataFrame(
    {
        "EAN": ["111", "222"],
        "SKU": ["A", "B"],
        "Our Price": [20.0, 8.0],
        "Supplier A Price": [12.0, 7.0],
        "Supplier B Price": [15.0, 11.0],
        "Cheapest Price": [12.0, 7.0],
        "Cheapest Supplier": ["Supplier A", "Supplier A"],
        "Saving €": [8.0, 1.0],
        "Saving %": [0.4, 0.125],
        "Status": ["CHEAPER", "CHEAPER"],
        "Matched Suppliers": [2, 2],
        "_relevant_for_display": [True, True],
    }
)

ean_lookup = {"111": 0, "222": 1}
sku_lookup = {"A": 0, "B": 1}
supplier_cols = ["Supplier A Price", "Supplier B Price"]
requests = core.parse_requested_codes("111\t10\nB\t10\nNOPE\t9")

special, skipped = build_special_price_result(
    full,
    requests,
    ean_lookup,
    sku_lookup,
    supplier_cols,
    legrand=True,
)

assert skipped == 0
assert list(special["Requested Code"]) == ["111", "B", "NOPE"]
assert abs(float(special.iloc[0][SPECIAL_PRICE]) - 13.75) < 1e-12
assert special.iloc[0][BELOW_SPECIAL] == "Supplier A"

# Our=8 and raw special=10 => 10 + ((8-10)*.3*1.25) = 9.25.
assert abs(float(special.iloc[1][SPECIAL_PRICE]) - 9.25) < 1e-12
assert special.iloc[1][BELOW_SPECIAL] == "Our Price, Supplier A"
assert special.iloc[2]["Lookup Status"] == "CODE NOT FOUND"
assert float(special.iloc[2][SPECIAL_PRICE]) == 9.0
assert special.iloc[2][BELOW_SPECIAL] == ""

model = _special_table_model(special, supplier_cols)
assert SPECIAL_PRICE in model["columns"]
assert BELOW_SPECIAL in model["columns"]
assert "Target Price" not in model["columns"]
assert "Below Target" not in model["columns"]

# Production must actually install the endpoints used by the updated template.
from production_app import app

routes = {
    (getattr(route, "path", None), method)
    for route in app.router.routes
    for method in (getattr(route, "methods", set()) or set())
}
assert ("/results/{workspace_id}/quick-filter", "POST") in routes
assert ("/download/{workspace_id}/shown-smart", "POST") in routes

print("Special Prices smoke test: OK")
