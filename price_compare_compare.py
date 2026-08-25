import hashlib

import numpy as np
import pandas as pd

from price_compare_common import (
    file_digest,
    normalize_ean,
    normalize_sku,
    parse_number,
    prepare_our_prices,
    prepare_supplier_price_maps,
)


def compare_all_suppliers(our_df, our_config, supplier_configs):
    result, our_extra_columns = prepare_our_prices(our_df, our_config)
    supplier_price_columns = []
    duplicate_info = {}

    normalized_our_ean = result["_ean_key"]
    normalized_our_sku = result["_sku_key"]

    for config in supplier_configs:
        supplier_name = config["supplier_name"]
        price_col = f"{supplier_name} Price"
        price_maps, duplicate_count = prepare_supplier_price_maps(
            config["dataframe"], config
        )

        if config["match_method"] == "EAN":
            result[price_col] = normalized_our_ean.map(price_maps["EAN"])
        elif config["match_method"] == "SKU":
            result[price_col] = normalized_our_sku.map(price_maps["SKU"])
        else:
            ean_prices = normalized_our_ean.map(price_maps["EAN"])
            sku_prices = normalized_our_sku.map(price_maps["SKU"])
            result[price_col] = ean_prices.combine_first(sku_prices)

        supplier_price_columns.append(price_col)
        duplicate_info[supplier_name] = duplicate_count

    result = result.drop(columns=["_ean_key", "_sku_key"])
    our_price = pd.to_numeric(result["Our Price"], errors="coerce")

    if supplier_price_columns:
        supplier_prices = result[supplier_price_columns].apply(
            pd.to_numeric, errors="coerce"
        )
        cheapest_alternative = supplier_prices.min(axis=1, skipna=True)
        result["Matched Suppliers"] = supplier_prices.notna().sum(axis=1).astype(int)
        all_prices = pd.concat(
            [our_price.rename("Our Price"), supplier_prices], axis=1
        )
    else:
        cheapest_alternative = pd.Series(np.nan, index=result.index, dtype=float)
        result["Matched Suppliers"] = 0
        all_prices = our_price.to_frame(name="Our Price")

    has_any_price = all_prices.notna().any(axis=1)
    result["Cheapest Price"] = all_prices.min(axis=1, skipna=True)

    cheapest_cols = pd.Series(pd.NA, index=result.index, dtype="object")
    if has_any_price.any():
        cheapest_cols.loc[has_any_price] = all_prices.loc[has_any_price].idxmin(axis=1)

    cheapest_names = cheapest_cols.astype("string")
    supplier_mask = cheapest_names.notna() & cheapest_names.ne("Our Price")
    cheapest_names.loc[supplier_mask] = cheapest_names.loc[supplier_mask].str.replace(
        r" Price$", "", regex=True
    )
    result["Cheapest Supplier"] = cheapest_names.where(has_any_price, "").fillna("")

    result["Saving €"] = our_price - result["Cheapest Price"]
    result["Saving %"] = np.where(
        our_price.notna() & our_price.ne(0),
        result["Saving €"] / our_price,
        np.nan,
    )

    alt = cheapest_alternative
    result["Status"] = np.select(
        [
            alt.isna(),
            our_price.isna(),
            alt < (our_price - 0.0000001),
            alt > (our_price + 0.0000001),
        ],
        ["NOT FOUND", "OUR PRICE MISSING", "CHEAPER", "MORE EXPENSIVE"],
        default="SAME PRICE",
    )

    final_columns = (
        ["EAN", "SKU", "Our Price"]
        + our_extra_columns
        + supplier_price_columns
        + [
            "Cheapest Price",
            "Cheapest Supplier",
            "Saving €",
            "Saving %",
            "Status",
            "Matched Suppliers",
            "_relevant_for_display",
        ]
    )
    result = result[final_columns].copy()

    status_order = {
        "CHEAPER": 1,
        "SAME PRICE": 2,
        "MORE EXPENSIVE": 3,
        "OUR PRICE MISSING": 4,
        "NOT FOUND": 5,
    }
    result["_sort_order"] = result["Status"].map(status_order).fillna(99)
    result = (
        result.sort_values(
            ["_sort_order", "Saving €"],
            ascending=[True, False],
            na_position="last",
            kind="stable",
        )
        .drop(columns=["_sort_order"])
        .reset_index(drop=True)
    )

    return result, supplier_price_columns, duplicate_info, our_extra_columns


def parse_requested_codes(text):
    if not text:
        return []

    requests = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if "\t" in line:
            code_part, target_part = line.split("\t", 1)
            code = code_part.strip()
            target_raw = target_part.strip()
        else:
            code = line
            target_raw = ""

        if not code:
            continue

        target_price = parse_number(target_raw) if target_raw else np.nan
        target_valid = (
            bool(target_raw)
            and pd.notna(target_price)
            and float(target_price) > 0
        )

        requests.append(
            {
                "code": code,
                "target_price": float(target_price) if target_valid else np.nan,
                "target_raw": target_raw,
                "target_valid": target_valid,
            }
        )
    return requests


def build_code_lookups(full_result):
    ean_lookup = {}
    sku_lookup = {}

    for idx, key in zip(full_result.index, full_result["EAN"].map(normalize_ean)):
        if key and key not in ean_lookup:
            ean_lookup[key] = idx

    for idx, key in zip(full_result.index, full_result["SKU"].map(normalize_sku)):
        if key and key not in sku_lookup:
            sku_lookup[key] = idx

    return ean_lookup, sku_lookup


def filter_result_by_codes(
    full_result,
    requested_codes,
    ean_lookup,
    sku_lookup,
    supplier_price_columns,
):
    if not requested_codes:
        return pd.DataFrame()

    rows = []
    text_columns = {"EAN", "SKU", "Cheapest Supplier", "Status"}

    for request in requested_codes:
        requested_code = request if isinstance(request, str) else request.get("code", "")
        target_price = np.nan if isinstance(request, str) else request.get("target_price", np.nan)

        ean_key = normalize_ean(requested_code)
        sku_key = normalize_sku(requested_code)
        matched_index = None
        lookup_status = ""

        if ean_key and ean_key in ean_lookup:
            matched_index = ean_lookup[ean_key]
            lookup_status = "MATCHED BY EAN"
        elif sku_key and sku_key in sku_lookup:
            matched_index = sku_lookup[sku_key]
            lookup_status = "MATCHED BY SKU"

        if matched_index is not None:
            row_dict = full_result.loc[matched_index].to_dict()
        else:
            row_dict = {
                col: ("" if col in text_columns else np.nan)
                for col in full_result.columns
            }
            row_dict["_relevant_for_display"] = False
            lookup_status = "CODE NOT FOUND"

        below_target = []
        if matched_index is not None and pd.notna(target_price):
            our_price = row_dict.get("Our Price", np.nan)
            if pd.notna(our_price) and float(our_price) < float(target_price):
                below_target.append("Our Price")

            for price_col in supplier_price_columns:
                supplier_price = row_dict.get(price_col, np.nan)
                if pd.notna(supplier_price) and float(supplier_price) < float(target_price):
                    supplier_name = (
                        price_col[:-len(" Price")]
                        if price_col.endswith(" Price")
                        else price_col
                    )
                    below_target.append(supplier_name)

        row_dict["Requested Code"] = requested_code
        row_dict["Target Price"] = target_price
        row_dict["Below Target"] = ", ".join(below_target)
        row_dict["Lookup Status"] = lookup_status
        rows.append(row_dict)

    filtered = pd.DataFrame(rows)
    first_columns = ["Requested Code", "Target Price", "Below Target", "Lookup Status"]
    remaining_columns = [col for col in full_result.columns if col not in first_columns]
    return filtered[first_columns + remaining_columns]


def lightweight_supplier_configs(supplier_configs):
    return [
        {key: value for key, value in config.items() if key != "dataframe"}
        for config in supplier_configs
    ]


def comparison_signature(our_bytes, our_sheet, our_config, supplier_configs):
    parts = [
        file_digest(our_bytes),
        str(our_sheet),
        str(our_config.get("type")),
        str(our_config.get("ean_column")),
        str(our_config.get("sku_column")),
        str(our_config.get("price_column")),
        "||".join(map(str, our_config.get("extra_columns", []))),
    ]

    for config in supplier_configs:
        parts.extend(
            [
                config["file_digest"],
                config["supplier_name"].casefold(),
                str(config["sheet_name"]),
                str(config["match_method"]),
                str(config.get("ean_column")),
                str(config.get("sku_column")),
                str(config["price_column"]),
            ]
        )

    return hashlib.sha256("||".join(parts).encode("utf-8")).hexdigest()
