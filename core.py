import io
import math
import re
from typing import Any

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_col_to_name

POWERBI_REQUIRED_COLUMNS = [
    "EAN",
    "SKU",
    "Latest Pricelist Value",
    "Realisation Summ",
    "Net Available Qty",
    "Days Since Last Sale",
    "MinStock",
    "MaxStock",
]

POWERBI_NUMERIC_COLUMNS = [
    "Latest Pricelist Value",
    "Realisation Summ",
    "Net Available Qty",
    "Days Since Last Sale",
    "MinStock",
    "MaxStock",
]

POWERBI_OUTPUT_COLUMNS = [
    "Realisation Summ",
    "Net Available Qty",
    "Days Since Last Sale",
    "MinStock",
    "MaxStock",
]

RESERVED_RESULT_COLUMNS = {
    "EAN", "SKU", "Our Price", "Cheapest Price", "Cheapest Supplier",
    "Saving €", "Saving %", "Status", "Matched Suppliers", "_relevant_for_display",
}


def normalize_ean(value: Any):
    if pd.isna(value): return None
    if isinstance(value, (int, np.integer)): return str(int(value))
    if isinstance(value, (float, np.floating)):
        if math.isnan(value): return None
        if value.is_integer(): return str(int(value))
    text = str(value).strip().replace("\u00A0", "")
    text = re.sub(r"\s+", "", text)
    if not text: return None
    if re.fullmatch(r"\d+\.0+", text): text = text.split(".")[0]
    return text.upper()


def normalize_sku(value: Any):
    if pd.isna(value): return None
    if isinstance(value, (int, np.integer)): return str(int(value)).upper()
    if isinstance(value, (float, np.floating)):
        if math.isnan(value): return None
        if value.is_integer(): return str(int(value)).upper()
    text = str(value).strip().replace("\u00A0", "")
    text = re.sub(r"\s+", "", text)
    return text.upper() if text else None


def clean_display_value(value: Any):
    if pd.isna(value): return None
    return str(value).strip().replace("\u00A0", "")


def parse_number(value: Any):
    if pd.isna(value): return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        try: return float(value)
        except (TypeError, ValueError): return np.nan
    text = str(value).strip().replace("\u00A0", "").replace(" ", "")
    if not text: return np.nan
    text = re.sub(r"[€$£]", "", text)
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    try: return float(text)
    except (TypeError, ValueError): return np.nan


def parse_numeric_series(series: pd.Series) -> pd.Series:
    result = pd.to_numeric(series, errors="coerce")
    fallback_mask = result.isna() & series.notna()
    if fallback_mask.any(): result.loc[fallback_mask] = series.loc[fallback_mask].map(parse_number)
    return result.astype(float)


def guess_column(columns, kind):
    columns = list(columns)
    if not columns: return None
    lowered = {str(c).strip().lower(): c for c in columns}
    exact_names = {
        "EAN": ["ean", "gtin", "barcode", "ean13", "ean-13", "ean code"],
        "SKU": ["sku", "article", "article no", "article number", "item no", "item number", "product code", "material"],
    }
    if kind in exact_names:
        for name in exact_names[kind]:
            if name in lowered: return lowered[name]
        partials = ["ean", "gtin", "barcode"] if kind == "EAN" else ["sku", "article", "item no", "item number", "product code", "material"]
        for col in columns:
            if any(part in str(col).lower() for part in partials): return col
    if kind == "PRICE":
        for word in ["latest pricelist value", "salesprice", "sales price", "net price", "purchase price", "unit price", "supplier price", "price", "cost"]:
            for col in columns:
                if word in str(col).lower(): return col
    return columns[0]


def unique_extra_output_names(extra_columns):
    mapping, used = {}, set(RESERVED_RESULT_COLUMNS)
    for source_col in extra_columns:
        candidate = str(source_col)
        if candidate in used: candidate = f"Our - {candidate}"
        base, suffix = candidate, 2
        while candidate in used:
            candidate = f"{base} ({suffix})"; suffix += 1
        mapping[source_col] = candidate; used.add(candidate)
    return mapping


def prepare_current_pricelist(current_df: pd.DataFrame, config: dict):
    if config["type"] == "PowerBI Pricelist":
        missing = [c for c in POWERBI_REQUIRED_COLUMNS if c not in current_df.columns]
        if missing: raise ValueError("PowerBI Pricelist is missing required columns: " + ", ".join(missing))
        our = current_df.copy()
        for col in POWERBI_NUMERIC_COLUMNS: our[col] = parse_numeric_series(our[col])
        our["EAN"] = our["EAN"].map(normalize_ean)
        our["SKU"] = our["SKU"].map(clean_display_value)
        our["Our Price"] = our["Latest Pricelist Value"]
        our["_ean_key"] = our["EAN"].map(normalize_ean)
        our["_sku_key"] = our["SKU"].map(normalize_sku)
        our["_relevant_for_display"] = our["Realisation Summ"].fillna(0).gt(0) | our["Net Available Qty"].fillna(0).gt(0)
        keep = ["EAN", "SKU", "Our Price"] + POWERBI_OUTPUT_COLUMNS + ["_ean_key", "_sku_key", "_relevant_for_display"]
        return our[keep].copy(), list(POWERBI_OUTPUT_COLUMNS)

    ean_column, sku_column = config.get("ean_column"), config.get("sku_column")
    price_column, extra_columns = config["price_column"], list(config.get("extra_columns") or [])
    if not ean_column and not sku_column: raise ValueError("Free format pricelist requires at least one identifier: EAN or SKU.")
    if price_column not in current_df.columns: raise ValueError("Selected current price column does not exist in the selected sheet.")
    if ean_column and ean_column not in current_df.columns: raise ValueError("Selected current EAN column does not exist in the selected sheet.")
    if sku_column and sku_column not in current_df.columns: raise ValueError("Selected current SKU column does not exist in the selected sheet.")
    if price_column in {ean_column, sku_column}: raise ValueError("Current Price column cannot also be an identifier column.")

    result = pd.DataFrame(index=current_df.index)
    result["EAN"] = current_df[ean_column].map(normalize_ean) if ean_column else None
    result["SKU"] = current_df[sku_column].map(clean_display_value) if sku_column else None
    result["Our Price"] = parse_numeric_series(current_df[price_column])
    if result["Our Price"].notna().sum() == 0: raise ValueError("No valid numeric prices were found in the selected current Price column.")
    extra_name_map, extra_output_columns = unique_extra_output_names(extra_columns), []
    for source_col in extra_columns:
        if source_col not in current_df.columns or source_col in {ean_column, sku_column, price_column}: continue
        output_col = extra_name_map[source_col]
        result[output_col] = current_df[source_col]; extra_output_columns.append(output_col)
    result["_ean_key"] = result["EAN"].map(normalize_ean)
    result["_sku_key"] = result["SKU"].map(normalize_sku)
    result["_relevant_for_display"] = True
    return result, extra_output_columns


def _build_identifier_price_map(supplier, id_column, normalizer):
    temp = supplier[[id_column, "_price"]].copy()
    temp["_key"] = temp[id_column].map(normalizer)
    temp = temp.dropna(subset=["_key", "_price"])
    duplicate_count = int(temp.loc[temp["_key"].duplicated(keep=False), "_key"].nunique())
    return temp.groupby("_key", sort=False)["_price"].min().to_dict(), duplicate_count


def prepare_supplier_maps(supplier_df, config):
    price_column = config["price_column"]
    if price_column not in supplier_df.columns: raise ValueError(f"{config['supplier_name']}: selected Price column does not exist.")
    supplier = supplier_df.copy(); supplier["_price"] = parse_numeric_series(supplier[price_column])
    if supplier["_price"].notna().sum() == 0: raise ValueError(f"{config['supplier_name']}: selected Price column has no valid numeric prices.")
    maps, duplicate_count, method = {}, 0, config["match_method"]
    if method in ("EAN", "EAN + SKU"):
        col = config.get("ean_column")
        if not col or col not in supplier.columns: raise ValueError(f"{config['supplier_name']}: select a valid EAN column.")
        maps["EAN"], dup = _build_identifier_price_map(supplier, col, normalize_ean); duplicate_count += dup
    if method in ("SKU", "EAN + SKU"):
        col = config.get("sku_column")
        if not col or col not in supplier.columns: raise ValueError(f"{config['supplier_name']}: select a valid SKU column.")
        maps["SKU"], dup = _build_identifier_price_map(supplier, col, normalize_sku); duplicate_count += dup
    return maps, duplicate_count


def compare_all(current_df, current_config, supplier_items):
    result, current_extra_columns = prepare_current_pricelist(current_df, current_config)
    supplier_price_columns, duplicate_info = [], {}
    ean_keys, sku_keys = result["_ean_key"], result["_sku_key"]
    for supplier_df, config in supplier_items:
        supplier_name = config["supplier_name"].strip(); price_col = f"{supplier_name} Price"
        maps, dup = prepare_supplier_maps(supplier_df, config); method = config["match_method"]
        if method == "EAN": result[price_col] = ean_keys.map(maps["EAN"])
        elif method == "SKU": result[price_col] = sku_keys.map(maps["SKU"])
        else: result[price_col] = ean_keys.map(maps["EAN"]).combine_first(sku_keys.map(maps["SKU"]))
        supplier_price_columns.append(price_col); duplicate_info[supplier_name] = dup

    result = result.drop(columns=["_ean_key", "_sku_key"]); our_price = pd.to_numeric(result["Our Price"], errors="coerce")
    if supplier_price_columns:
        supplier_prices = result[supplier_price_columns].apply(pd.to_numeric, errors="coerce")
        cheapest_alternative = supplier_prices.min(axis=1, skipna=True)
        result["Matched Suppliers"] = supplier_prices.notna().sum(axis=1).astype(int)
        all_prices = pd.concat([our_price.rename("Our Price"), supplier_prices], axis=1)
    else:
        cheapest_alternative = pd.Series(np.nan, index=result.index, dtype=float); result["Matched Suppliers"] = 0
        all_prices = our_price.to_frame(name="Our Price")
    has_any = all_prices.notna().any(axis=1); result["Cheapest Price"] = all_prices.min(axis=1, skipna=True)
    cheapest_cols = pd.Series(pd.NA, index=result.index, dtype="object")
    if has_any.any(): cheapest_cols.loc[has_any] = all_prices.loc[has_any].idxmin(axis=1)
    cheapest_names = cheapest_cols.astype("string"); mask = cheapest_names.notna() & cheapest_names.ne("Our Price")
    cheapest_names.loc[mask] = cheapest_names.loc[mask].str.replace(r" Price$", "", regex=True)
    result["Cheapest Supplier"] = cheapest_names.where(has_any, "").fillna("")
    result["Saving €"] = our_price - result["Cheapest Price"]
    result["Saving %"] = np.where(our_price.notna() & our_price.ne(0), result["Saving €"] / our_price, np.nan)
    alt = cheapest_alternative
    result["Status"] = np.select([alt.isna(), our_price.isna(), alt < (our_price - 1e-7), alt > (our_price + 1e-7)], ["NOT FOUND", "OUR PRICE MISSING", "CHEAPER", "MORE EXPENSIVE"], default="SAME PRICE")
    final_columns = ["EAN", "SKU", "Our Price"] + current_extra_columns + supplier_price_columns + ["Cheapest Price", "Cheapest Supplier", "Saving €", "Saving %", "Status", "Matched Suppliers", "_relevant_for_display"]
    duplicates = [str(c) for c in pd.Index(final_columns)[pd.Index(final_columns).duplicated()].unique()]
    if duplicates: raise ValueError("Internal comparison produced duplicate output columns: " + ", ".join(duplicates))
    result = result[final_columns].copy()
    order = {"CHEAPER": 1, "SAME PRICE": 2, "MORE EXPENSIVE": 3, "OUR PRICE MISSING": 4, "NOT FOUND": 5}
    result["_sort_order"] = result["Status"].map(order).fillna(99)
    result = result.sort_values(["_sort_order", "Saving €"], ascending=[True, False], na_position="last", kind="stable").drop(columns=["_sort_order"]).reset_index(drop=True)
    return result, supplier_price_columns, duplicate_info, current_extra_columns


def build_code_lookups(full_result):
    ean_lookup, sku_lookup = {}, {}
    for idx, key in zip(full_result.index, full_result["EAN"].map(normalize_ean)):
        if key and key not in ean_lookup: ean_lookup[key] = idx
    for idx, key in zip(full_result.index, full_result["SKU"].map(normalize_sku)):
        if key and key not in sku_lookup: sku_lookup[key] = idx
    return ean_lookup, sku_lookup


def parse_requested_codes(text):
    requests = []
    for raw_line in (text or "").splitlines():
        line = raw_line.strip()
        if not line: continue
        if "\t" in line: code_part, target_part = line.split("\t", 1); code, target_raw = code_part.strip(), target_part.strip()
        else: code, target_raw = line, ""
        if not code: continue
        target = parse_number(target_raw) if target_raw else np.nan
        valid = bool(target_raw) and pd.notna(target) and float(target) > 0
        requests.append({"code": code, "target_price": float(target) if valid else np.nan, "target_raw": target_raw, "target_valid": valid})
    return requests


def filter_result_by_codes(full_result, requests, ean_lookup, sku_lookup, supplier_price_columns):
    rows, text_cols = [], {"EAN", "SKU", "Cheapest Supplier", "Status"}
    for req in requests:
        code, target = req["code"], req.get("target_price", np.nan)
        ean_key, sku_key = normalize_ean(code), normalize_sku(code); matched_index, lookup_status = None, ""
        if ean_key and ean_key in ean_lookup: matched_index, lookup_status = ean_lookup[ean_key], "MATCHED BY EAN"
        elif sku_key and sku_key in sku_lookup: matched_index, lookup_status = sku_lookup[sku_key], "MATCHED BY SKU"
        if matched_index is not None: row = full_result.loc[matched_index].to_dict()
        else:
            row = {c: ("" if c in text_cols else np.nan) for c in full_result.columns}; row["_relevant_for_display"] = False; lookup_status = "CODE NOT FOUND"
        below_target = []
        if matched_index is not None and pd.notna(target):
            our_price = row.get("Our Price", np.nan)
            if pd.notna(our_price) and float(our_price) < float(target): below_target.append("Our Price")
            for col in supplier_price_columns:
                value = row.get(col, np.nan)
                if pd.notna(value) and float(value) < float(target): below_target.append(col[:-6] if col.endswith(" Price") else col)
        row.update({"Requested Code": code, "Target Price": target, "Below Target": ", ".join(below_target), "Lookup Status": lookup_status}); rows.append(row)
    first = ["Requested Code", "Target Price", "Below Target", "Lookup Status"]
    remaining = [c for c in full_result.columns if c not in first]
    filtered = pd.DataFrame(rows)
    return filtered[first + remaining] if not filtered.empty else pd.DataFrame(columns=first + remaining)


def create_excel(full_result, supplier_price_columns, supplier_configs, duplicate_info, current_file_name, current_config):
    export_result = full_result.drop(columns=["_relevant_for_display"]).copy(); relevant_mask = full_result["_relevant_for_display"].fillna(False).astype(bool)
    summary_rows = [
        ("Current pricelist file", current_file_name), ("Current pricelist type", current_config["type"]), ("Suppliers uploaded", len(supplier_configs)),
        ("All current-pricelist products exported", len(full_result)), ("Products shown by default", int(relevant_mask.sum())),
        ("Products matched with at least one supplier", int(full_result["Matched Suppliers"].gt(0).sum())),
        ("Products where at least one supplier is cheaper", int((full_result["Status"] == "CHEAPER").sum())),
        ("Products not found at any supplier", int((full_result["Status"] == "NOT FOUND").sum())),
    ]
    for config in supplier_configs:
        supplier_name, price_col = config["supplier_name"], f"{config['supplier_name']} Price"
        summary_rows.extend([("", ""), (f"{supplier_name} - source file", config["file_name"]), (f"{supplier_name} - sheet", config["sheet_name"]), (f"{supplier_name} - match method", config["match_method"]), (f"{supplier_name} - price column", config["price_column"]), (f"{supplier_name} - matched products", int(full_result[price_col].notna().sum())), (f"{supplier_name} - duplicated identifiers", duplicate_info.get(supplier_name, 0))])
    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"]); buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export_result.to_excel(writer, sheet_name="Price Comparison", index=False); summary_df.to_excel(writer, sheet_name="Summary", index=False)
        workbook, ws, sws = writer.book, writer.sheets["Price Comparison"], writer.sheets["Summary"]
        header_fmt = workbook.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True, "bg_color": "#0F766E", "font_color": "#FFFFFF"})
        money_fmt = workbook.add_format({"num_format": '€#,##0.00;[Red]-€#,##0.00'}); pct_fmt = workbook.add_format({"num_format": "0.00%"})
        supplier_cheaper_fmt = workbook.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"}); cheapest_fmt = workbook.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E", "bold": True}); not_found_fmt = workbook.add_format({"bg_color": "#F1F5F9", "font_color": "#475569"})
        positions = {name: export_result.columns.get_loc(name) for name in export_result.columns}
        for i, name in enumerate(export_result.columns):
            ws.write(0, i, name, header_fmt); width = 17 if name == "EAN" else 22 if name in {"SKU", "Cheapest Price", "Cheapest Supplier"} else 18; ws.set_column(i, i, width)
        for name in ["Our Price", *supplier_price_columns, "Cheapest Price", "Saving €"]:
            if name in positions: ws.set_column(positions[name], positions[name], 18, money_fmt)
        if "Saving %" in positions: ws.set_column(positions["Saving %"], positions["Saving %"], 12, pct_fmt)
        ws.freeze_panes(1, 3)
        if len(export_result):
            ws.autofilter(0, 0, len(export_result), len(export_result.columns) - 1); our_letter = xl_col_to_name(positions["Our Price"])
            for supplier_col in supplier_price_columns:
                sidx, sletter = positions[supplier_col], xl_col_to_name(positions[supplier_col])
                ws.conditional_format(1, sidx, len(export_result), sidx, {"type": "formula", "criteria": f'=AND({sletter}2<>"",${our_letter}2<>"",{sletter}2<${our_letter}2)', "format": supplier_cheaper_fmt})
            for name in ["Cheapest Price", "Cheapest Supplier"]:
                idx = positions[name]; ws.conditional_format(1, idx, len(export_result), idx, {"type": "no_blanks", "format": cheapest_fmt})
            status_idx, status_letter = positions["Status"], xl_col_to_name(positions["Status"])
            ws.conditional_format(1, 0, len(export_result), len(export_result.columns) - 1, {"type": "formula", "criteria": f'=${status_letter}2="NOT FOUND"', "format": not_found_fmt})
        sws.set_column("A:A", 50); sws.set_column("B:B", 45)
        for i, name in enumerate(summary_df.columns): sws.write(0, i, name, header_fmt)
        sws.freeze_panes(1, 0)
    buffer.seek(0); return buffer.getvalue()


def dataframe_to_table_model(df, supplier_price_columns, target_mode=False):
    columns = [c for c in df.columns if c != "_relevant_for_display"]; rows = []
    money_cols = {"Our Price", "Target Price", "Cheapest Price", "Saving €", *supplier_price_columns}
    for _, row in df.iterrows():
        target = row.get("Target Price", np.nan); has_target = target_mode and pd.notna(target); our_price = row.get("Our Price", np.nan)
        row_cells, not_found_code = [], row.get("Lookup Status", "") == "CODE NOT FOUND"
        for col in columns:
            value, cls = row.get(col), ""
            if not_found_code: cls = "cell-not-found"
            elif has_target:
                if col == "Target Price": cls = "cell-target"
                elif col in ["Our Price", *supplier_price_columns, "Cheapest Price"] and pd.notna(value) and float(value) < float(target): cls = "cell-target"
                elif col == "Below Target" and str(value or "").strip(): cls = "cell-target"
            else:
                if col in supplier_price_columns and pd.notna(value) and pd.notna(our_price) and float(value) < float(our_price): cls = "cell-cheaper"
                if col in {"Cheapest Price", "Cheapest Supplier"} and pd.notna(value) and str(value) != "": cls = "cell-cheapest"
            is_missing = False
            if not isinstance(value, (list, dict, tuple, set)):
                try:
                    missing_value = pd.isna(value); is_missing = bool(missing_value) if np.isscalar(missing_value) else False
                except Exception: is_missing = False
            if is_missing: display = ""
            elif col in money_cols:
                try: display = f"€{float(value):,.2f}"
                except Exception: display = str(value)
            elif col == "Saving %":
                try: display = f"{float(value):.2%}"
                except Exception: display = str(value)
            elif isinstance(value, float) and value.is_integer(): display = str(int(value))
            else: display = str(value)
            row_cells.append({"value": display, "class": cls})
        rows.append(row_cells)
    return {"columns": columns, "rows": rows}
