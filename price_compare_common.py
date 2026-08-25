import hashlib
import io
import math
import re

import numpy as np
import pandas as pd


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
    "EAN",
    "SKU",
    "Our Price",
    "Cheapest Price",
    "Cheapest Supplier",
    "Saving €",
    "Saving %",
    "Status",
    "Matched Suppliers",
    "_relevant_for_display",
}


def safe_widget_key(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(text))


def file_digest(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def normalize_ean(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        if math.isnan(value):
            return None
        if value.is_integer():
            return str(int(value))

    text = str(value).strip().replace("\u00A0", "")
    text = re.sub(r"\s+", "", text)
    if not text:
        return None
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    return text.upper()


def normalize_sku(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, np.integer)):
        return str(int(value)).upper()
    if isinstance(value, (float, np.floating)):
        if math.isnan(value):
            return None
        if value.is_integer():
            return str(int(value)).upper()

    text = str(value).strip().replace("\u00A0", "")
    text = re.sub(r"\s+", "", text)
    return text.upper() if text else None


def clean_display_value(value):
    if pd.isna(value):
        return None
    return str(value).strip().replace("\u00A0", "")


def parse_number(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            return float(value)
        except (TypeError, ValueError):
            return np.nan

    text = str(value).strip().replace("\u00A0", "").replace(" ", "")
    if not text:
        return np.nan
    text = re.sub(r"[€$£]", "", text)

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        return float(text)
    except (TypeError, ValueError):
        return np.nan


def parse_numeric_series(series: pd.Series) -> pd.Series:
    result = pd.to_numeric(series, errors="coerce")
    fallback_mask = result.isna() & series.notna()
    if fallback_mask.any():
        result.loc[fallback_mask] = series.loc[fallback_mask].map(parse_number)
    return result.astype(float)


def guess_column(columns, kind):
    columns = list(columns)
    if not columns:
        return None

    lowered = {str(c).strip().lower(): c for c in columns}
    exact_names = {
        "EAN": ["ean", "gtin", "barcode", "ean13", "ean-13", "ean code"],
        "SKU": [
            "sku",
            "article",
            "article no",
            "article number",
            "item no",
            "item number",
            "product code",
            "material",
        ],
    }

    if kind in exact_names:
        for name in exact_names[kind]:
            if name in lowered:
                return lowered[name]

        partials = (
            ["ean", "gtin", "barcode"]
            if kind == "EAN"
            else ["sku", "article", "item no", "item number", "product code", "material"]
        )
        for col in columns:
            low = str(col).lower()
            if any(part in low for part in partials):
                return col

    if kind == "PRICE":
        price_words = [
            "salesprice",
            "sales price",
            "net price",
            "purchase price",
            "unit price",
            "supplier price",
            "latest pricelist value",
            "price",
            "cost",
        ]
        for word in price_words:
            for col in columns:
                if word in str(col).lower():
                    return col

    return columns[0]


def get_sheet_names(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def read_excel_sheet(file_bytes, sheet_name):
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        dtype=object,
    )
    # Browser forms send strings; normalizing column labels keeps mapping stable.
    df.columns = [str(col) for col in df.columns]
    return df


def unique_extra_output_names(extra_columns):
    mapping = {}
    used = set(RESERVED_RESULT_COLUMNS)

    for source_col in extra_columns:
        candidate = str(source_col)
        if candidate in used:
            candidate = f"Our - {candidate}"

        base = candidate
        suffix = 2
        while candidate in used:
            candidate = f"{base} ({suffix})"
            suffix += 1

        mapping[source_col] = candidate
        used.add(candidate)

    return mapping


def prepare_our_prices(our_df, our_config):
    """Prepare PowerBI or free-format current-pricelist data."""
    if our_config["type"] == "PowerBI Pricelist":
        our = our_df.copy()
        for col in POWERBI_NUMERIC_COLUMNS:
            our[col] = parse_numeric_series(our[col])

        our["EAN"] = our["EAN"].map(normalize_ean)
        our["SKU"] = our["SKU"].map(clean_display_value)
        our["Our Price"] = our["Latest Pricelist Value"]
        our["_ean_key"] = our["EAN"].map(normalize_ean)
        our["_sku_key"] = our["SKU"].map(normalize_sku)
        our["_relevant_for_display"] = (
            our["Realisation Summ"].fillna(0).gt(0)
            | our["Net Available Qty"].fillna(0).gt(0)
        )

        keep = (
            ["EAN", "SKU", "Our Price"]
            + POWERBI_OUTPUT_COLUMNS
            + ["_ean_key", "_sku_key", "_relevant_for_display"]
        )
        return our[keep].copy(), list(POWERBI_OUTPUT_COLUMNS)

    source = our_df.copy()
    result = pd.DataFrame(index=source.index)
    ean_column = our_config.get("ean_column")
    sku_column = our_config.get("sku_column")
    price_column = our_config["price_column"]
    extra_columns = list(our_config.get("extra_columns", []))
    extra_name_map = unique_extra_output_names(extra_columns)

    result["EAN"] = source[ean_column].map(normalize_ean) if ean_column else None
    result["SKU"] = source[sku_column].map(clean_display_value) if sku_column else None
    result["Our Price"] = parse_numeric_series(source[price_column])

    extra_output_columns = []
    for source_col in extra_columns:
        output_col = extra_name_map[source_col]
        result[output_col] = source[source_col]
        extra_output_columns.append(output_col)

    result["_ean_key"] = result["EAN"].map(normalize_ean)
    result["_sku_key"] = result["SKU"].map(normalize_sku)
    result["_relevant_for_display"] = True
    return result, extra_output_columns


def _build_identifier_price_map(supplier, id_column, normalizer):
    temp = supplier[[id_column, "_price"]].copy()
    temp["_key"] = temp[id_column].map(normalizer)
    temp = temp.dropna(subset=["_key", "_price"]).copy()

    duplicate_count = int(
        temp.loc[temp["_key"].duplicated(keep=False), "_key"].nunique()
    )
    best = temp.groupby("_key", sort=False)["_price"].min()
    return best.to_dict(), duplicate_count


def prepare_supplier_price_maps(supplier_df, config):
    supplier = supplier_df.copy()
    supplier["_price"] = parse_numeric_series(supplier[config["price_column"]])

    match_method = config["match_method"]
    price_maps = {}
    duplicate_count = 0

    if match_method in ("EAN", "EAN + SKU"):
        ean_map, ean_duplicates = _build_identifier_price_map(
            supplier, config["ean_column"], normalize_ean
        )
        price_maps["EAN"] = ean_map
        duplicate_count += ean_duplicates

    if match_method in ("SKU", "EAN + SKU"):
        sku_map, sku_duplicates = _build_identifier_price_map(
            supplier, config["sku_column"], normalize_sku
        )
        price_maps["SKU"] = sku_map
        duplicate_count += sku_duplicates

    return price_maps, duplicate_count
