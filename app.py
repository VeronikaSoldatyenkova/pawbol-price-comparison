import hashlib
import io
import math
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from xlsxwriter.utility import xl_col_to_name


# =========================================================
# PAGE CONFIGURATION + UI THEME
# =========================================================

st.set_page_config(
    page_title="Price List Comparison",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        :root {
            --pc-primary: #0f766e;
            --pc-primary-hover: #115e59;
            --pc-primary-soft: #ecfdf5;
            --pc-ink: #0f172a;
            --pc-muted: #64748b;
            --pc-border: #e2e8f0;
            --pc-surface: #ffffff;
            --pc-soft: #f8fafc;
            --pc-warning: #fff7ed;
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1550px;
        }

        h1, h2, h3, h4 {
            color: var(--pc-ink);
        }

        .pc-hero {
            padding: 1.25rem 1.4rem;
            border: 1px solid var(--pc-border);
            border-radius: 16px;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            margin-bottom: 1rem;
        }

        .pc-hero-title {
            margin: 0;
            font-size: 1.85rem;
            font-weight: 750;
            letter-spacing: -0.02em;
            color: var(--pc-ink);
        }

        .pc-hero-subtitle {
            margin: .35rem 0 0 0;
            color: var(--pc-muted);
            font-size: .98rem;
        }

        .pc-step {
            display: inline-flex;
            align-items: center;
            gap: .55rem;
            margin: .4rem 0 .65rem 0;
            font-size: 1.05rem;
            font-weight: 700;
            color: var(--pc-ink);
        }

        .pc-step-number {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            border-radius: 999px;
            background: var(--pc-primary-soft);
            color: var(--pc-primary);
            font-size: .86rem;
            font-weight: 800;
        }

        div[data-testid="stMetric"] {
            background: var(--pc-surface);
            border: 1px solid var(--pc-border);
            border-radius: 12px;
            padding: .85rem 1rem;
        }

        div[data-testid="stMetricLabel"] {
            color: var(--pc-muted);
        }

        div[data-testid="stExpander"] {
            border: 1px solid var(--pc-border);
            border-radius: 12px;
            overflow: hidden;
        }

        div[data-testid="stFileUploader"] {
            border-radius: 12px;
        }

        /* Primary actions: teal, never Streamlit red */
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"],
        button[kind="primary"] {
            background: var(--pc-primary) !important;
            border-color: var(--pc-primary) !important;
            color: white !important;
            box-shadow: none !important;
        }

        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover,
        button[kind="primary"]:hover {
            background: var(--pc-primary-hover) !important;
            border-color: var(--pc-primary-hover) !important;
            color: white !important;
        }

        /* Secondary buttons: clean white / slate */
        .stButton > button[kind="secondary"],
        .stDownloadButton > button[kind="secondary"],
        button[kind="secondary"] {
            background: #ffffff !important;
            border-color: #cbd5e1 !important;
            color: #334155 !important;
            box-shadow: none !important;
        }

        .stButton > button[kind="secondary"]:hover,
        .stDownloadButton > button[kind="secondary"]:hover,
        button[kind="secondary"]:hover {
            background: #f8fafc !important;
            border-color: #94a3b8 !important;
            color: #0f172a !important;
        }

        button:focus {
            box-shadow: 0 0 0 2px rgba(15, 118, 110, .18) !important;
        }

        .stTextInput input:focus,
        .stTextArea textarea:focus {
            border-color: var(--pc-primary) !important;
            box-shadow: 0 0 0 1px var(--pc-primary) !important;
        }

        div[data-baseweb="select"] > div:focus-within {
            border-color: var(--pc-primary) !important;
        }

        .pc-note {
            border-left: 3px solid var(--pc-primary);
            background: #f0fdfa;
            color: #334155;
            border-radius: 8px;
            padding: .75rem .9rem;
            margin: .35rem 0 .8rem 0;
            font-size: .92rem;
        }

        .pc-muted {
            color: var(--pc-muted);
            font-size: .9rem;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# OURPRICES CONFIGURATION
# =========================================================

REQUIRED_OUR_COLUMNS = [
    "EAN",
    "SKU",
    "Latest Pricelist Value",
    "Realisation Summ",
    "Net Available Qty",
    "Days Since Last Sale",
    "MinStock",
    "MaxStock",
]

OUR_NUMERIC_COLUMNS = [
    "Latest Pricelist Value",
    "Realisation Summ",
    "Net Available Qty",
    "Days Since Last Sale",
    "MinStock",
    "MaxStock",
]


# =========================================================
# HELPERS
# =========================================================

def section_title(number: int, title: str):
    st.markdown(
        f"""
        <div class="pc-step">
            <span class="pc-step-number">{number}</span>
            <span>{title}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def safe_widget_key(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(text))


@st.cache_data(show_spinner=False)
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


def parse_number(value):
    """Fallback parser for uncommon text price formats."""
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
    """
    Fast path for numeric Excel cells, with a fallback only for text cells
    that pandas could not convert directly.
    """
    result = pd.to_numeric(series, errors="coerce")
    fallback_mask = result.isna() & series.notna()

    if fallback_mask.any():
        result.loc[fallback_mask] = series.loc[fallback_mask].map(parse_number)

    return result.astype(float)


def guess_column(columns, kind):
    columns = list(columns)
    lowered = {str(c).strip().lower(): c for c in columns}

    exact_names = {
        "EAN": [
            "ean",
            "gtin",
            "barcode",
            "ean13",
            "ean-13",
            "ean code",
        ],
        "SKU": [
            "sku",
            "article",
            "article no",
            "article number",
            "item no",
            "item number",
            "product code",
        ],
    }

    if kind in exact_names:
        for name in exact_names[kind]:
            if name in lowered:
                return lowered[name]

        if kind == "EAN":
            partials = ["ean", "gtin", "barcode"]
        else:
            partials = [
                "sku",
                "article",
                "item no",
                "item number",
                "product code",
                "material",
            ]

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
            "price",
            "cost",
        ]

        for word in price_words:
            for col in columns:
                if word in str(col).lower():
                    return col

    return columns[0]


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def read_excel_sheet(file_bytes, sheet_name):
    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        dtype=object,
    )


def prepare_our_prices(our_df):
    our = our_df.copy()

    for col in OUR_NUMERIC_COLUMNS:
        our[col] = parse_numeric_series(our[col])

    our["EAN"] = our["EAN"].map(normalize_ean)
    our["SKU"] = our["SKU"].map(
        lambda x: None if pd.isna(x) else str(x).strip().replace("\u00A0", "")
    )

    # Precompute matching keys once. These are used for every supplier.
    our["_ean_key"] = our["EAN"].map(normalize_ean)
    our["_sku_key"] = our["SKU"].map(normalize_sku)

    # Browser-only filter. Excel export still receives all rows.
    our["_relevant_for_display"] = (
        our["Realisation Summ"].fillna(0).gt(0)
        | our["Net Available Qty"].fillna(0).gt(0)
    )

    return our


def _build_identifier_price_map(supplier, id_column, normalizer):
    temp = supplier[[id_column, "_price"]].copy()
    temp["_key"] = temp[id_column].map(normalizer)
    temp = temp.dropna(subset=["_key", "_price"]).copy()

    duplicate_count = int(
        temp.loc[
            temp["_key"].duplicated(keep=False),
            "_key",
        ].nunique()
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
            supplier,
            config["ean_column"],
            normalize_ean,
        )
        price_maps["EAN"] = ean_map
        duplicate_count += ean_duplicates

    if match_method in ("SKU", "EAN + SKU"):
        sku_map, sku_duplicates = _build_identifier_price_map(
            supplier,
            config["sku_column"],
            normalize_sku,
        )
        price_maps["SKU"] = sku_map
        duplicate_count += sku_duplicates

    return price_maps, duplicate_count


def compare_all_suppliers(our_df, supplier_configs):
    """
    Compare every OurPrices row against all suppliers.

    EAN + SKU rule:
      1. EAN match first.
      2. SKU is used only when the EAN lookup has no supplier price.

    Cheapest Price includes Our Price as well as every matched supplier price.
    Our Price has tie priority, so if Our Price equals the lowest supplier price,
    Cheapest Supplier is shown as "Our Price".
    """
    result = prepare_our_prices(our_df)

    result = result[
        [
            "EAN",
            "SKU",
            "Latest Pricelist Value",
            "Realisation Summ",
            "Net Available Qty",
            "Days Since Last Sale",
            "MinStock",
            "MaxStock",
            "_ean_key",
            "_sku_key",
            "_relevant_for_display",
        ]
    ].copy()

    result = result.rename(columns={"Latest Pricelist Value": "Our Price"})

    supplier_price_columns = []
    duplicate_info = {}

    normalized_our_ean = result["_ean_key"]
    normalized_our_sku = result["_sku_key"]

    for config in supplier_configs:
        supplier_name = config["supplier_name"]
        price_col = f"{supplier_name} Price"

        price_maps, duplicate_count = prepare_supplier_price_maps(
            config["dataframe"],
            config,
        )

        match_method = config["match_method"]

        if match_method == "EAN":
            result[price_col] = normalized_our_ean.map(price_maps["EAN"])

        elif match_method == "SKU":
            result[price_col] = normalized_our_sku.map(price_maps["SKU"])

        else:
            ean_prices = normalized_our_ean.map(price_maps["EAN"])
            sku_prices = normalized_our_sku.map(price_maps["SKU"])
            result[price_col] = ean_prices.combine_first(sku_prices)

        supplier_price_columns.append(price_col)
        duplicate_info[supplier_name] = duplicate_count

    # Remove internal matching keys before producing the result.
    result = result.drop(columns=["_ean_key", "_sku_key"])

    our_price = pd.to_numeric(result["Our Price"], errors="coerce")

    if supplier_price_columns:
        supplier_prices = result[supplier_price_columns].apply(
            pd.to_numeric,
            errors="coerce",
        )

        # Keep supplier-only information for the existing matched/status logic.
        # This remains separate from Cheapest Price because Cheapest Price now
        # also includes Our Price.
        cheapest_alternative = supplier_prices.min(axis=1, skipna=True)
        result["Matched Suppliers"] = supplier_prices.notna().sum(axis=1).astype(int)

        # Our Price is intentionally first so it wins an exact-price tie.
        all_prices = pd.concat(
            [our_price.rename("Our Price"), supplier_prices],
            axis=1,
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

    # Supplier price columns are named "<Supplier> Price". Keep "Our Price"
    # unchanged while stripping the suffix from supplier names.
    cheapest_names = cheapest_cols.astype("string")
    supplier_name_mask = cheapest_names.notna() & cheapest_names.ne("Our Price")
    cheapest_names.loc[supplier_name_mask] = cheapest_names.loc[
        supplier_name_mask
    ].str.replace(r" Price$", "", regex=True)

    result["Cheapest Supplier"] = cheapest_names.where(has_any_price, "").fillna("")

    # Savings are now against the cheapest price overall. If Our Price is
    # already the lowest price, Saving € and Saving % are zero.
    result["Saving €"] = our_price - result["Cheapest Price"]

    result["Saving %"] = np.where(
        our_price.notna() & our_price.ne(0),
        result["Saving €"] / our_price,
        np.nan,
    )

    # Status keeps its original supplier-comparison meaning so filters such as
    # "Only supplier cheaper" continue to work correctly.
    alt = cheapest_alternative
    our = our_price

    result["Status"] = np.select(
        [
            alt.isna(),
            our.isna(),
            alt < (our - 0.0000001),
            alt > (our + 0.0000001),
        ],
        [
            "NOT FOUND",
            "OUR PRICE MISSING",
            "CHEAPER",
            "MORE EXPENSIVE",
        ],
        default="SAME PRICE",
    )

    final_columns = (
        ["EAN", "SKU", "Our Price"]
        + supplier_price_columns
        + [
            "Cheapest Price",
            "Cheapest Supplier",
            "Saving €",
            "Saving %",
            "Status",
            "Matched Suppliers",
            "Realisation Summ",
            "Net Available Qty",
            "Days Since Last Sale",
            "MinStock",
            "MaxStock",
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

    return result, supplier_price_columns, duplicate_info

def parse_requested_codes(text):
    """
    Parse the Quick code filter.

    Supported input:
      CODE
      CODE<TAB>TARGET_PRICE

    This is designed for direct copy/paste from one or two Excel columns.
    Target price is optional, so the old one-code-per-line workflow still works.
    """
    if not text:
        return []

    requests = []

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # Excel separates copied cells with tabs. If there is no tab, keep the
        # whole line as the code so SKUs containing spaces are not broken.
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
    """
    Build code -> row-index dictionaries once per comparison instead of scanning
    every OurPrices row each time the user clicks Filter.
    """
    ean_lookup = {}
    sku_lookup = {}

    ean_keys = full_result["EAN"].map(normalize_ean)
    sku_keys = full_result["SKU"].map(normalize_sku)

    for idx, key in zip(full_result.index, ean_keys):
        if key and key not in ean_lookup:
            ean_lookup[key] = idx

    for idx, key in zip(full_result.index, sku_keys):
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
    """
    Return one row per pasted request, preserving the exact pasted order.

    Code matching:
      1. EAN first
      2. SKU fallback

    If a Target Price was pasted, list Our Price and every alternative supplier
    whose price is strictly below that target.
    """
    if not requested_codes:
        return pd.DataFrame()

    rows = []

    text_columns = {
        "EAN",
        "SKU",
        "Cheapest Supplier",
        "Status",
    }

    for request in requested_codes:
        # Backwards compatibility for a session created by an older app version.
        if isinstance(request, str):
            requested_code = request
            target_price = np.nan
        else:
            requested_code = request.get("code", "")
            target_price = request.get("target_price", np.nan)

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

    first_columns = [
        "Requested Code",
        "Target Price",
        "Below Target",
        "Lookup Status",
    ]
    remaining_columns = [
        col for col in full_result.columns
        if col not in first_columns
    ]

    return filtered[first_columns + remaining_columns]


def lightweight_supplier_configs(supplier_configs):
    """Do not keep uploaded supplier DataFrames in session state after comparison."""
    return [
        {
            key: value
            for key, value in config.items()
            if key != "dataframe"
        }
        for config in supplier_configs
    ]


def comparison_signature(our_bytes, our_sheet, supplier_configs):
    """
    Detect stale results when the user changes files or mapping settings after
    a comparison. This prevents old results being shown under new settings.
    """
    parts = [
        file_digest(our_bytes),
        str(our_sheet),
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


def create_excel(
    full_result,
    supplier_price_columns,
    supplier_configs,
    duplicate_info,
    our_file_name,
):
    """
    Export all OurPrices products. This function is called once per comparison
    and its bytes are then stored in session state.
    """
    export_result = full_result.drop(columns=["_relevant_for_display"]).copy()

    relevant_mask = (
        full_result["Realisation Summ"].fillna(0).gt(0)
        | full_result["Net Available Qty"].fillna(0).gt(0)
    )

    summary_rows = [
        ("OurPrices file", our_file_name),
        ("Suppliers uploaded", len(supplier_configs)),
        ("All OurPrices products exported", len(full_result)),
        ("Relevant products shown in browser", int(relevant_mask.sum())),
        (
            "Products matched with at least one supplier",
            int(full_result["Matched Suppliers"].gt(0).sum()),
        ),
        (
            "Products where at least one supplier is cheaper",
            int((full_result["Status"] == "CHEAPER").sum()),
        ),
        (
            "Products where cheapest supplier is more expensive",
            int((full_result["Status"] == "MORE EXPENSIVE").sum()),
        ),
        (
            "Products with same cheapest price",
            int((full_result["Status"] == "SAME PRICE").sum()),
        ),
        (
            "Products not found at any supplier",
            int((full_result["Status"] == "NOT FOUND").sum()),
        ),
    ]

    for config in supplier_configs:
        supplier_name = config["supplier_name"]
        price_col = f"{supplier_name} Price"

        supplier_summary_rows = [
            ("", ""),
            (f"{supplier_name} - source file", config["file_name"]),
            (f"{supplier_name} - sheet", config["sheet_name"]),
            (f"{supplier_name} - match method", config["match_method"]),
        ]

        if config["match_method"] == "EAN + SKU":
            supplier_summary_rows.extend(
                [
                    (f"{supplier_name} - EAN column", config["ean_column"]),
                    (f"{supplier_name} - SKU column", config["sku_column"]),
                    (f"{supplier_name} - priority", "EAN first, then SKU fallback"),
                ]
            )
        elif config["match_method"] == "EAN":
            supplier_summary_rows.append(
                (f"{supplier_name} - EAN column", config["ean_column"])
            )
        else:
            supplier_summary_rows.append(
                (f"{supplier_name} - SKU column", config["sku_column"])
            )

        supplier_summary_rows.extend(
            [
                (f"{supplier_name} - price column", config["price_column"]),
                (
                    f"{supplier_name} - matched products",
                    int(full_result[price_col].notna().sum()),
                ),
                (
                    f"{supplier_name} - cheaper than our current price",
                    int(
                        (
                            full_result[price_col].notna()
                            & full_result["Our Price"].notna()
                            & (full_result[price_col] < full_result["Our Price"])
                        ).sum()
                    ),
                ),
                (
                    f"{supplier_name} - duplicated identifiers",
                    duplicate_info.get(supplier_name, 0),
                ),
            ]
        )

        summary_rows.extend(supplier_summary_rows)

    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export_result.to_excel(
            writer,
            sheet_name="Price Comparison",
            index=False,
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        workbook = writer.book
        worksheet = writer.sheets["Price Comparison"]
        summary_ws = writer.sheets["Summary"]

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "bg_color": "#0F766E",
                "font_color": "#FFFFFF",
            }
        )

        money_fmt = workbook.add_format(
            {"num_format": '€#,##0.00;[Red]-€#,##0.00'}
        )

        percentage_fmt = workbook.add_format(
            {"num_format": "0.00%"}
        )

        decimal_fmt = workbook.add_format(
            {"num_format": "#,##0.00"}
        )

        integer_fmt = workbook.add_format(
            {"num_format": "0"}
        )

        supplier_cheaper_fmt = workbook.add_format(
            {
                "bg_color": "#DCFCE7",
                "font_color": "#166534",
            }
        )

        cheapest_fmt = workbook.add_format(
            {
                "bg_color": "#FEF3C7",
                "font_color": "#92400E",
                "bold": True,
            }
        )

        not_found_fmt = workbook.add_format(
            {
                "bg_color": "#F1F5F9",
                "font_color": "#475569",
            }
        )

        for col_idx, name in enumerate(export_result.columns):
            worksheet.write(0, col_idx, name, header_fmt)

        widths = {
            "EAN": 17,
            "SKU": 22,
            "Our Price": 14,
            "Cheapest Price": 23,
            "Cheapest Supplier": 22,
            "Saving €": 13,
            "Saving %": 12,
            "Status": 19,
            "Matched Suppliers": 17,
            "Realisation Summ": 18,
            "Net Available Qty": 18,
            "Days Since Last Sale": 21,
            "MinStock": 12,
            "MaxStock": 12,
        }

        positions = {
            name: export_result.columns.get_loc(name)
            for name in export_result.columns
        }

        for col_idx, name in enumerate(export_result.columns):
            default_width = 18 if name in supplier_price_columns else 15
            worksheet.set_column(
                col_idx,
                col_idx,
                widths.get(name, default_width),
            )

        money_columns = [
            "Our Price",
            *supplier_price_columns,
            "Cheapest Price",
            "Saving €",
        ]

        for name in money_columns:
            if name in positions:
                i = positions[name]
                worksheet.set_column(
                    i,
                    i,
                    widths.get(name, 18),
                    money_fmt,
                )

        if "Saving %" in positions:
            i = positions["Saving %"]
            worksheet.set_column(i, i, widths["Saving %"], percentage_fmt)

        if "Realisation Summ" in positions:
            i = positions["Realisation Summ"]
            worksheet.set_column(i, i, widths["Realisation Summ"], decimal_fmt)

        for name in [
            "Matched Suppliers",
            "Net Available Qty",
            "Days Since Last Sale",
            "MinStock",
            "MaxStock",
        ]:
            if name in positions:
                i = positions[name]
                worksheet.set_column(i, i, widths.get(name, 15), integer_fmt)

        # Freeze header + EAN/SKU/Our Price when scrolling.
        worksheet.freeze_panes(1, 3)

        if len(export_result):
            worksheet.autofilter(
                0,
                0,
                len(export_result),
                len(export_result.columns) - 1,
            )

            our_price_letter = xl_col_to_name(positions["Our Price"])

            for supplier_col in supplier_price_columns:
                supplier_idx = positions[supplier_col]
                supplier_letter = xl_col_to_name(supplier_idx)

                worksheet.conditional_format(
                    1,
                    supplier_idx,
                    len(export_result),
                    supplier_idx,
                    {
                        "type": "formula",
                        "criteria": (
                            f'=AND({supplier_letter}2<>"",'
                            f'${our_price_letter}2<>"",'
                            f'{supplier_letter}2<${our_price_letter}2)'
                        ),
                        "format": supplier_cheaper_fmt,
                    },
                )

            cheapest_price_idx = positions["Cheapest Price"]
            cheapest_supplier_idx = positions["Cheapest Supplier"]

            worksheet.conditional_format(
                1,
                cheapest_price_idx,
                len(export_result),
                cheapest_price_idx,
                {
                    "type": "no_blanks",
                    "format": cheapest_fmt,
                },
            )

            worksheet.conditional_format(
                1,
                cheapest_supplier_idx,
                len(export_result),
                cheapest_supplier_idx,
                {
                    "type": "no_blanks",
                    "format": cheapest_fmt,
                },
            )

            status_idx = positions["Status"]
            status_letter = xl_col_to_name(status_idx)
            worksheet.conditional_format(
                1,
                0,
                len(export_result),
                len(export_result.columns) - 1,
                {
                    "type": "formula",
                    "criteria": f'=${status_letter}2="NOT FOUND"',
                    "format": not_found_fmt,
                },
            )

        summary_ws.set_column("A:A", 48)
        summary_ws.set_column("B:B", 45)

        for col_idx, name in enumerate(summary_df.columns):
            summary_ws.write(0, col_idx, name, header_fmt)

        summary_ws.freeze_panes(1, 0)

    buffer.seek(0)
    return buffer.getvalue()


def style_browser_table(df, supplier_price_columns):
    """
    Light table styling.

    Normal comparison mode:
      - Supplier prices below Our Price are highlighted green.
      - Cheapest Price and Cheapest Supplier are highlighted yellow.

    Quick code filter rows with a valid Target Price:
      - Target Price is highlighted in soft blue.
      - Every price below Target Price uses the same blue palette.
      - Green "cheaper than Our Price" and yellow "cheapest" highlights are
        intentionally suppressed so target performance is the only price signal.
    """
    target_blue = "background-color: #e0f2fe; color: #075985; font-weight: 800;"

    def style_row(row):
        styles = pd.Series("", index=row.index, dtype=object)

        target_price = row.get("Target Price", np.nan)
        has_target = pd.notna(target_price)

        if row.get("Lookup Status", "") == "CODE NOT FOUND":
            styles[:] = "background-color: #f1f5f9; color: #475569;"
            if "Target Price" in styles.index and has_target:
                styles["Target Price"] = target_blue
            return styles

        our_price = row.get("Our Price", np.nan)
        cheapest_price = row.get("Cheapest Price", np.nan)
        cheapest_supplier = row.get("Cheapest Supplier", "")

        # -------------------------------------------------
        # Target mode: blue only
        # -------------------------------------------------
        if has_target:
            if "Target Price" in styles.index:
                styles["Target Price"] = target_blue

            price_columns = ["Our Price", *supplier_price_columns, "Cheapest Price"]
            for col in price_columns:
                if col not in styles.index:
                    continue

                value = row.get(col, np.nan)
                if pd.notna(value) and float(value) < float(target_price):
                    styles[col] = target_blue

            below_target = row.get("Below Target", "")
            if (
                "Below Target" in styles.index
                and pd.notna(below_target)
                and str(below_target).strip()
            ):
                styles["Below Target"] = target_blue

            # Do not add green/yellow comparison highlights to target rows.
            return styles

        # -------------------------------------------------
        # Normal comparison mode
        # -------------------------------------------------
        for col in supplier_price_columns:
            supplier_price = row.get(col, np.nan)
            if (
                pd.notna(supplier_price)
                and pd.notna(our_price)
                and float(supplier_price) < float(our_price)
            ):
                styles[col] = (
                    "background-color: #dcfce7; "
                    "color: #166534; font-weight: 600;"
                )

        if pd.notna(cheapest_price):
            if "Cheapest Price" in styles.index:
                styles["Cheapest Price"] = (
                    "background-color: #fef3c7; "
                    "color: #92400e; font-weight: 700;"
                )

            if "Cheapest Supplier" in styles.index and cheapest_supplier:
                styles["Cheapest Supplier"] = (
                    "background-color: #fef3c7; "
                    "color: #92400e; font-weight: 700;"
                )

        return styles

    formatters = {
        "Our Price": "€{:.2f}",
        "Target Price": "€{:.2f}",
        "Cheapest Price": "€{:.2f}",
        "Saving €": "€{:.2f}",
        "Saving %": "{:.2%}",
        "Realisation Summ": "{:.2f}",
    }

    for col in supplier_price_columns:
        formatters[col] = "€{:.2f}"

    formatters = {
        key: value
        for key, value in formatters.items()
        if key in df.columns
    }

    return df.style.apply(style_row, axis=1).format(formatters, na_rep="")

def reset_comparison_state():
    st.session_state.comparison_result = None
    st.session_state.comparison_supplier_columns = []
    st.session_state.comparison_supplier_configs = []
    st.session_state.comparison_duplicate_info = {}
    st.session_state.comparison_our_file_name = ""
    st.session_state.comparison_signature = ""
    st.session_state.comparison_excel_bytes = None
    st.session_state.comparison_ean_lookup = {}
    st.session_state.comparison_sku_lookup = {}
    st.session_state.comparison_code_filter_active = False
    st.session_state.comparison_code_filter_codes = []


# =========================================================
# SESSION STATE
# =========================================================

SESSION_DEFAULTS = {
    "comparison_result": None,
    "comparison_supplier_columns": [],
    "comparison_supplier_configs": [],
    "comparison_duplicate_info": {},
    "comparison_our_file_name": "",
    "comparison_signature": "",
    "comparison_excel_bytes": None,
    "comparison_ean_lookup": {},
    "comparison_sku_lookup": {},
    "comparison_code_filter_active": False,
    "comparison_code_filter_codes": [],
}

for key, default_value in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# =========================================================
# WEB INTERFACE
# =========================================================

st.markdown(
    """
    <div class="pc-hero">
        <div class="pc-hero-title">📊 Price List Comparison</div>
        <div class="pc-hero-subtitle">
            Compare your current prices with one or many supplier price lists,
            identify the cheapest offers, and export a complete Excel report.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="pc-note">
        The browser shows products where <b>Realisation Summ &gt; 0</b> or
        <b>Net Available Qty &gt; 0</b>. The downloaded Excel always contains
        <b>all OurPrices products</b>.
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# 1. OURPRICES
# ---------------------------------------------------------

section_title(1, "Our current price list")

with st.container(border=True):
    left, right = st.columns([2, 1])

    with left:
        our_file = st.file_uploader(
            "Upload OurPrices",
            type=["xlsx", "xlsm", "xls"],
            key="our_prices_upload",
            help="The fixed OurPrices structure is validated automatically.",
        )

    if our_file is None:
        with right:
            st.markdown(
                '<div class="pc-muted">Upload OurPrices to begin.</div>',
                unsafe_allow_html=True,
            )
        st.stop()

    our_bytes = our_file.getvalue()

    try:
        our_sheets = get_sheet_names(our_bytes)
    except Exception as exc:
        st.error(f"Could not open OurPrices: {exc}")
        st.stop()

    our_default_sheet_index = (
        our_sheets.index("Export") if "Export" in our_sheets else 0
    )

    with right:
        our_sheet = st.selectbox(
            "Sheet",
            our_sheets,
            index=our_default_sheet_index,
            key="our_prices_sheet",
        )

    try:
        our_df = read_excel_sheet(our_bytes, our_sheet)
    except Exception as exc:
        st.error(f"Could not read OurPrices sheet: {exc}")
        st.stop()

    missing_columns = [
        col for col in REQUIRED_OUR_COLUMNS if col not in our_df.columns
    ]

    if missing_columns:
        st.error(
            "OurPrices is missing required columns: "
            + ", ".join(missing_columns)
        )
        st.write("Columns found:", list(our_df.columns))
        st.stop()

    st.success(f"OurPrices validated · {len(our_df):,} rows")


# ---------------------------------------------------------
# 2. SUPPLIERS
# ---------------------------------------------------------

section_title(2, "Alternative suppliers")

with st.container(border=True):
    mode_col, info_col = st.columns([1.2, 2])

    with mode_col:
        upload_mode = st.radio(
            "Upload mode",
            ["Single supplier", "Bulk suppliers"],
            horizontal=True,
            key="supplier_upload_mode",
        )

    with info_col:
        st.caption(
            "Each supplier can match by SKU, EAN, or EAN + SKU. "
            "Combined matching always tries EAN first and then SKU."
        )

    if upload_mode == "Single supplier":
        single_file = st.file_uploader(
            "Alternative supplier price list",
            type=["xlsx", "xlsm", "xls"],
            key="single_supplier_upload",
        )
        supplier_files = [single_file] if single_file is not None else []
    else:
        supplier_files = st.file_uploader(
            "Alternative supplier price lists",
            type=["xlsx", "xlsm", "xls"],
            accept_multiple_files=True,
            key="bulk_supplier_upload",
        )

if not supplier_files:
    st.caption("Upload at least one alternative supplier price list to continue.")
    st.stop()


# ---------------------------------------------------------
# 3. CONFIGURE SUPPLIERS
# ---------------------------------------------------------

section_title(3, "Supplier mapping")

supplier_configs = []
configuration_errors = []
entered_supplier_names = []

for index, supplier_file in enumerate(supplier_files, start=1):
    file_bytes = supplier_file.getvalue()
    base_key = safe_widget_key(f"supplier_{index}_{supplier_file.name}")
    default_name = Path(supplier_file.name).stem

    try:
        supplier_sheets = get_sheet_names(file_bytes)
    except Exception as exc:
        configuration_errors.append(
            f"Could not open {supplier_file.name}: {exc}"
        )
        continue

    with st.expander(
        f"{index}. {supplier_file.name}",
        expanded=(len(supplier_files) <= 3),
    ):
        top1, top2 = st.columns([1.3, 1])

        with top1:
            supplier_name = st.text_input(
                "Supplier name",
                value=default_name,
                key=f"{base_key}_name",
                help="Used in result columns and in Cheapest Supplier.",
            ).strip()

        entered_supplier_names.append(supplier_name)

        with top2:
            sheet_name = st.selectbox(
                "Excel sheet",
                supplier_sheets,
                key=f"{base_key}_sheet",
            )

        try:
            supplier_df = read_excel_sheet(file_bytes, sheet_name)
        except Exception as exc:
            configuration_errors.append(
                f"Could not read {supplier_file.name} / {sheet_name}: {exc}"
            )
            continue

        if supplier_df.empty and len(supplier_df.columns) == 0:
            configuration_errors.append(
                f"{supplier_file.name} / {sheet_name} has no columns."
            )
            continue

        supplier_columns = list(supplier_df.columns)

        st.markdown("**Matching**")
        match_method = st.radio(
            "Match using",
            ["SKU", "EAN", "EAN + SKU"],
            horizontal=True,
            key=f"{base_key}_match",
            help=(
                "EAN + SKU: EAN is tried first. "
                "SKU is only used when EAN does not match."
            ),
            label_visibility="collapsed",
        )

        price_guess = guess_column(supplier_columns, "PRICE")
        ean_column = None
        sku_column = None

        if match_method == "EAN + SKU":
            ean_guess = guess_column(supplier_columns, "EAN")
            sku_guess = guess_column(supplier_columns, "SKU")

            col1, col2, col3 = st.columns(3)

            with col1:
                ean_column = st.selectbox(
                    "EAN column",
                    supplier_columns,
                    index=supplier_columns.index(ean_guess),
                    key=f"{base_key}_ean_col",
                )

            with col2:
                sku_column = st.selectbox(
                    "SKU column",
                    supplier_columns,
                    index=supplier_columns.index(sku_guess),
                    key=f"{base_key}_sku_col",
                )

            with col3:
                price_column = st.selectbox(
                    "Price column",
                    supplier_columns,
                    index=supplier_columns.index(price_guess),
                    key=f"{base_key}_price_col_both",
                )

            if ean_column == sku_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: "
                    "EAN and SKU columns cannot be the same."
                )

            if price_column in (ean_column, sku_column):
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: "
                    "Price column cannot also be an identifier column."
                )

        elif match_method == "EAN":
            ean_guess = guess_column(supplier_columns, "EAN")

            col1, col2 = st.columns(2)

            with col1:
                ean_column = st.selectbox(
                    "EAN column",
                    supplier_columns,
                    index=supplier_columns.index(ean_guess),
                    key=f"{base_key}_ean_col",
                )

            with col2:
                price_column = st.selectbox(
                    "Price column",
                    supplier_columns,
                    index=supplier_columns.index(price_guess),
                    key=f"{base_key}_price_col_ean",
                )

            if ean_column == price_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: "
                    "EAN and Price columns cannot be the same."
                )

        else:
            sku_guess = guess_column(supplier_columns, "SKU")

            col1, col2 = st.columns(2)

            with col1:
                sku_column = st.selectbox(
                    "SKU column",
                    supplier_columns,
                    index=supplier_columns.index(sku_guess),
                    key=f"{base_key}_sku_col",
                )

            with col2:
                price_column = st.selectbox(
                    "Price column",
                    supplier_columns,
                    index=supplier_columns.index(price_guess),
                    key=f"{base_key}_price_col_sku",
                )

            if sku_column == price_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: "
                    "SKU and Price columns cannot be the same."
                )

        preview_col, stat_col = st.columns([3, 1])

        with preview_col:
            with st.expander("Preview first 10 rows"):
                st.dataframe(
                    supplier_df.head(10),
                    use_container_width=True,
                    hide_index=True,
                    height=300,
                )

        with stat_col:
            valid_price_count = int(
                parse_numeric_series(supplier_df[price_column]).notna().sum()
            )
            st.metric("Rows", f"{len(supplier_df):,}")
            st.metric("Valid prices", f"{valid_price_count:,}")

        if valid_price_count == 0:
            configuration_errors.append(
                f"{supplier_name or supplier_file.name}: "
                "no valid numeric prices were found in the selected Price column."
            )

        if not supplier_name:
            configuration_errors.append(
                f"Supplier {index} ({supplier_file.name}) has no supplier name."
            )

        supplier_configs.append(
            {
                "supplier_name": supplier_name,
                "file_name": supplier_file.name,
                "file_digest": file_digest(file_bytes),
                "sheet_name": sheet_name,
                "match_method": match_method,
                "ean_column": ean_column,
                "sku_column": sku_column,
                "price_column": price_column,
                "dataframe": supplier_df,
            }
        )


# Case-insensitive duplicate-name check.
non_empty_names = [name.casefold() for name in entered_supplier_names if name]
if len(non_empty_names) != len(set(non_empty_names)):
    configuration_errors.append(
        "Supplier names must be unique (case-insensitive)."
    )

if configuration_errors:
    # Deduplicate identical messages created by rerun combinations.
    for error in dict.fromkeys(configuration_errors):
        st.error(error)


# ---------------------------------------------------------
# 4. COMPARE
# ---------------------------------------------------------

section_title(4, "Run comparison")

compare_disabled = bool(configuration_errors) or not supplier_configs
current_signature = comparison_signature(
    our_bytes,
    our_sheet,
    supplier_configs,
)

with st.container(border=True):
    action_col, note_col = st.columns([1, 2])

    with action_col:
        compare_clicked = st.button(
            "Compare prices",
            type="primary",
            use_container_width=True,
            disabled=compare_disabled,
        )

    with note_col:
        st.caption(
            "The result is stored in the browser session. "
            "Changing display filters will not rerun the comparison."
        )

if compare_clicked:
    try:
        with st.spinner("Comparing suppliers and preparing Excel..."):
            comparison_result, supplier_price_columns, duplicate_info = (
                compare_all_suppliers(
                    our_df,
                    supplier_configs,
                )
            )

            saved_configs = lightweight_supplier_configs(supplier_configs)
            ean_lookup, sku_lookup = build_code_lookups(comparison_result)

            # Generate the workbook once. Radio/filter changes reuse these bytes.
            excel_bytes = create_excel(
                full_result=comparison_result,
                supplier_price_columns=supplier_price_columns,
                supplier_configs=saved_configs,
                duplicate_info=duplicate_info,
                our_file_name=our_file.name,
            )

        st.session_state.comparison_result = comparison_result
        st.session_state.comparison_supplier_columns = supplier_price_columns
        st.session_state.comparison_supplier_configs = saved_configs
        st.session_state.comparison_duplicate_info = duplicate_info
        st.session_state.comparison_our_file_name = our_file.name
        st.session_state.comparison_signature = current_signature
        st.session_state.comparison_excel_bytes = excel_bytes
        st.session_state.comparison_ean_lookup = ean_lookup
        st.session_state.comparison_sku_lookup = sku_lookup
        st.session_state.comparison_code_filter_active = False
        st.session_state.comparison_code_filter_codes = []

        st.success("Comparison completed.")

    except Exception as exc:
        st.exception(exc)


# =========================================================
# 5. RESULTS
# =========================================================

full_result = st.session_state.comparison_result

if full_result is not None:
    # Prevent stale results from being presented after a file/mapping change.
    if st.session_state.comparison_signature != current_signature:
        st.warning(
            "Files or supplier mapping settings changed after the last comparison. "
            "Click **Compare prices** again to refresh the results."
        )
        st.stop()

    supplier_price_columns = st.session_state.comparison_supplier_columns
    saved_supplier_configs = st.session_state.comparison_supplier_configs
    duplicate_info = st.session_state.comparison_duplicate_info

    st.divider()
    section_title(5, "Comparison results")

    relevant_result = full_result[
        full_result["_relevant_for_display"]
    ].copy()

    relevant_matched = int(
        relevant_result["Matched Suppliers"].gt(0).sum()
    )
    relevant_cheaper = int((relevant_result["Status"] == "CHEAPER").sum())
    relevant_more_expensive = int(
        (relevant_result["Status"] == "MORE EXPENSIVE").sum()
    )
    relevant_not_found = int((relevant_result["Status"] == "NOT FOUND").sum())

    metric1, metric2, metric3, metric4, metric5 = st.columns(5)

    metric1.metric("Relevant products", f"{len(relevant_result):,}")
    metric2.metric("Matched", f"{relevant_matched:,}")
    metric3.metric("Supplier cheaper", f"{relevant_cheaper:,}")
    metric4.metric("More expensive", f"{relevant_more_expensive:,}")
    metric5.metric("Not found", f"{relevant_not_found:,}")

    duplicate_suppliers = {
        supplier: count
        for supplier, count in duplicate_info.items()
        if count > 0
    }

    if duplicate_suppliers:
        duplicate_text = " · ".join(
            f"{supplier}: {count}"
            for supplier, count in duplicate_suppliers.items()
        )
        st.warning(
            "Duplicate supplier identifiers were found. "
            "The lowest valid price was used. "
            + duplicate_text
        )

    # -----------------------------------------------------
    # Manual code filter
    # -----------------------------------------------------

    with st.container(border=True):
        st.markdown("#### Quick code filter")
        st.caption(
            "Paste one or two columns directly from Excel: **EAN/SKU** and an optional "
            "**Target Price**. Results keep the exact pasted order and search all "
            "compared OurPrices rows."
        )

        guide_left, guide_right = st.columns([1.3, 2.7])
        with guide_left:
            st.markdown("**Column 1:** EAN or SKU  \n**Column 2:** Target Price *(optional)*")
        with guide_right:
            st.caption(
                "When a target is supplied, the result lists **Our Price** and every "
                "alternative supplier whose price is below that target."
            )

        st.text_area(
            "Codes and optional target prices",
            height=155,
            placeholder=(
                "3606481482679\t10\n"
                "3606480790065\t5\n"
                "3606480789861\t5\n"
                "A9R35240\t12.50\n"
                "3606480303357"
            ),
            key="comparison_code_filter_input",
            label_visibility="collapsed",
        )

        filter_col, clear_col, spacer_col = st.columns([1, 1, 3])

        with filter_col:
            if st.button(
                "Apply code filter",
                type="primary",
                use_container_width=True,
                key="comparison_code_filter_button",
            ):
                requested_codes = parse_requested_codes(
                    st.session_state.get("comparison_code_filter_input", "")
                )

                if requested_codes:
                    st.session_state.comparison_code_filter_codes = requested_codes
                    st.session_state.comparison_code_filter_active = True
                else:
                    st.warning("Paste at least one EAN or SKU first.")

        with clear_col:
            if st.button(
                "Clear filter",
                use_container_width=True,
                key="comparison_code_filter_clear",
            ):
                st.session_state.comparison_code_filter_active = False
                st.session_state.comparison_code_filter_codes = []

    manual_filter_active = st.session_state.comparison_code_filter_active
    requested_codes = st.session_state.comparison_code_filter_codes

    # -----------------------------------------------------
    # Display filter
    # -----------------------------------------------------

    if manual_filter_active and requested_codes:
        shown = filter_result_by_codes(
            full_result,
            requested_codes,
            st.session_state.comparison_ean_lookup,
            st.session_state.comparison_sku_lookup,
            supplier_price_columns,
        )

        if "_relevant_for_display" in shown.columns:
            shown = shown.drop(columns=["_relevant_for_display"])

        code_not_found_count = int(
            (shown["Lookup Status"] == "CODE NOT FOUND").sum()
        )
        target_count = int(shown["Target Price"].notna().sum())
        target_hit_count = int(
            shown["Below Target"].fillna("").astype(str).str.strip().ne("").sum()
        )

        invalid_target_count = sum(
            1
            for request in requested_codes
            if isinstance(request, dict)
            and request.get("target_raw")
            and not request.get("target_valid")
        )

        info1, info2, info3, info4 = st.columns(4)
        info1.metric("Requested", f"{len(shown):,}")
        info2.metric("Not found", f"{code_not_found_count:,}")
        info3.metric("With target price", f"{target_count:,}")
        info4.metric(
            "Target met",
            f"{target_hit_count:,}",
            help="Our Price or at least one supplier price is below the pasted target.",
        )

        if invalid_target_count:
            st.warning(
                f"{invalid_target_count:,} pasted target price(s) could not be parsed "
                "or were not greater than zero. Those rows are still shown, but target "
                "comparison was skipped."
            )

    else:
        display_option = st.radio(
            "Display",
            [
                "All",
                "Only supplier cheaper",
                "Only matched",
                "Only not found",
            ],
            horizontal=True,
            key="comparison_display_filter",
        )

        if display_option == "Only supplier cheaper":
            shown = relevant_result[
                relevant_result["Status"] == "CHEAPER"
            ].copy()

        elif display_option == "Only matched":
            shown = relevant_result[
                relevant_result["Matched Suppliers"].gt(0)
            ].copy()

        elif display_option == "Only not found":
            shown = relevant_result[
                relevant_result["Matched Suppliers"].eq(0)
            ].copy()

        else:
            shown = relevant_result.copy()

        shown = shown.drop(columns=["_relevant_for_display"])

        st.caption(
            f"Showing {len(shown):,} of {len(relevant_result):,} relevant products."
        )

    st.dataframe(
        style_browser_table(shown, supplier_price_columns),
        use_container_width=True,
        hide_index=True,
        height=620,
    )

    # -----------------------------------------------------
    # Download
    # -----------------------------------------------------

    with st.container(border=True):
        download_col, info_col = st.columns([1.2, 2])

        with download_col:
            st.download_button(
                "Download full comparison Excel",
                data=st.session_state.comparison_excel_bytes,
                file_name=(
                    f"Price_Comparison_{datetime.now():%Y-%m-%d_%H-%M}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True,
            )

        with info_col:
            st.caption(
                f"Excel contains all {len(full_result):,} OurPrices products, "
                "including rows hidden by the browser relevance filter."
            )
