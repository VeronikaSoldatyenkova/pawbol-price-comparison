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
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Price List Comparison",
    page_icon="📊",
    layout="wide",
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


def safe_widget_key(text: str) -> str:
    """Create a Streamlit-safe widget key."""
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(text))



def normalize_ean(value):
    """Normalize EAN/GTIN values coming from Excel."""
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

    # Excel can turn an integer-looking EAN into "1234567890123.0"
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    return text.upper()



def normalize_sku(value):
    """Normalize SKU/article values."""
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
    """Parse common European/US Excel price formats into float."""
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
        # 1.234,56
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        # 1,234.56
        else:
            text = text.replace(",", "")
    elif "," in text:
        # 123,45
        text = text.replace(",", ".")

    text = re.sub(r"[^0-9.\-]", "", text)

    try:
        return float(text)
    except (TypeError, ValueError):
        return np.nan



def guess_column(columns, kind):
    """Try to preselect the most likely supplier column."""
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
        # More specific names first so a column such as SalesPrice is preferred
        # over another column merely containing the word "price".
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
    """
    Prepare ALL OurPrices rows.

    Important: no Realisation/Available filtering is performed here because
    the downloaded workbook must contain all products.
    """
    our = our_df.copy()

    for col in OUR_NUMERIC_COLUMNS:
        our[col] = our[col].map(parse_number)

    our["EAN"] = our["EAN"].map(normalize_ean)
    our["SKU"] = our["SKU"].map(
        lambda x: None if pd.isna(x) else str(x).strip().replace("\u00A0", "")
    )

    # This flag is used ONLY for the browser display.
    our["_relevant_for_display"] = (
        our["Realisation Summ"].fillna(0).gt(0)
        | our["Net Available Qty"].fillna(0).gt(0)
    )

    return our



def _build_identifier_price_map(supplier, id_column, normalizer):
    """
    Build normalized identifier -> lowest valid supplier price.

    `supplier` must already contain the parsed `_price` column.
    """
    temp = supplier[[id_column, "_price"]].copy()
    temp["_key"] = temp[id_column].map(normalizer)
    temp = temp.dropna(subset=["_key", "_price"]).copy()

    duplicate_count = int(
        temp.loc[
            temp["_key"].duplicated(keep=False),
            "_key",
        ].nunique()
    )

    # If the same identifier appears more than once, use the cheapest valid price.
    best = temp.groupby("_key", as_index=True)["_price"].min()

    return best.to_dict(), duplicate_count


def prepare_supplier_price_maps(supplier_df, config):
    """
    Prepare one or two supplier lookup maps.

    Supported match methods:
      - SKU
      - EAN
      - EAN + SKU

    For EAN + SKU the comparison always tries EAN first and only falls back
    to SKU when the EAN lookup does not find a supplier price.
    """
    supplier = supplier_df.copy()
    supplier["_price"] = supplier[config["price_column"]].map(parse_number)

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
    Compare every OurPrices row against all uploaded suppliers.

    Every supplier can independently use SKU, EAN, or EAN + SKU matching and
    can use differently named identifier/price columns.

    When EAN + SKU is selected, EAN has priority. SKU is used only if the EAN
    lookup fails for that product.
    """
    result = prepare_our_prices(our_df)

    # Retain only the columns wanted in the final result plus the display flag.
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
            "_relevant_for_display",
        ]
    ].copy()

    result = result.rename(columns={"Latest Pricelist Value": "Our Price"})

    supplier_price_columns = []
    duplicate_info = {}

    normalized_our_ean = result["EAN"].map(normalize_ean)
    normalized_our_sku = result["SKU"].map(normalize_sku)

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
            # EAN + SKU:
            # 1) try EAN
            # 2) only if no EAN match, try SKU
            ean_prices = normalized_our_ean.map(price_maps["EAN"])
            sku_prices = normalized_our_sku.map(price_maps["SKU"])
            result[price_col] = ean_prices.combine_first(sku_prices)

        supplier_price_columns.append(price_col)
        duplicate_info[supplier_name] = duplicate_count

    # -----------------------------------------------------
    # Best price across all suppliers
    # -----------------------------------------------------

    if supplier_price_columns:
        supplier_prices = result[supplier_price_columns].apply(
            pd.to_numeric,
            errors="coerce",
        )

        result["Cheapest Alternative Price"] = supplier_prices.min(
            axis=1,
            skipna=True,
        )

        def cheapest_supplier_for_row(row):
            valid = row.dropna()
            if valid.empty:
                return ""
            cheapest_col = valid.idxmin()
            return cheapest_col[: -len(" Price")]

        result["Cheapest Supplier"] = supplier_prices.apply(
            cheapest_supplier_for_row,
            axis=1,
        )

        result["Matched Suppliers"] = supplier_prices.notna().sum(axis=1)

    else:
        result["Cheapest Alternative Price"] = np.nan
        result["Cheapest Supplier"] = ""
        result["Matched Suppliers"] = 0

    # -----------------------------------------------------
    # Comparison versus our current price
    # -----------------------------------------------------

    result["Saving €"] = result["Our Price"] - result["Cheapest Alternative Price"]

    result["Saving %"] = np.where(
        result["Our Price"].notna() & result["Our Price"].ne(0),
        result["Saving €"] / result["Our Price"],
        np.nan,
    )

    def overall_status(row):
        alternative = row["Cheapest Alternative Price"]
        our_price = row["Our Price"]

        if pd.isna(alternative):
            return "NOT FOUND"

        if pd.isna(our_price):
            return "OUR PRICE MISSING"

        if alternative < our_price - 0.0000001:
            return "CHEAPER"

        if alternative > our_price + 0.0000001:
            return "MORE EXPENSIVE"

        return "SAME PRICE"

    result["Status"] = result.apply(overall_status, axis=1)

    # Arrange columns. All individual supplier prices remain visible.
    final_columns = (
        ["EAN", "SKU", "Our Price"]
        + supplier_price_columns
        + [
            "Cheapest Alternative Price",
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

    # Put useful/cheaper products first while preserving all products.
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
        )
        .drop(columns=["_sort_order"])
        .reset_index(drop=True)
    )

    return result, supplier_price_columns, duplicate_info


def parse_requested_codes(text):
    """Parse pasted EAN/SKU values while preserving the user's order."""
    if not text:
        return []

    # Primarily line-based, but commas/semicolons are also accepted.
    parts = re.split(r"[\r\n,;]+", text)
    return [part.strip() for part in parts if part.strip()]


def filter_result_by_codes(full_result, requested_codes):
    """
    Return exactly one row per requested code, in the same order as pasted.

    The lookup accepts an EAN, an SKU, or a mix. EAN is checked first and SKU
    second. The lookup uses the full comparison result, so a manually requested
    product can be shown even if it is normally hidden by the browser's
    Realisation/Available filter.

    If a code is not found, a placeholder row is retained with CODE NOT FOUND.
    """
    if not requested_codes:
        return pd.DataFrame()

    ean_lookup = {}
    sku_lookup = {}

    for idx, row in full_result.iterrows():
        ean_key = normalize_ean(row.get("EAN"))
        sku_key = normalize_sku(row.get("SKU"))

        # Keep the first OurPrices row for a duplicated identifier so that one
        # pasted code always produces one result row.
        if ean_key and ean_key not in ean_lookup:
            ean_lookup[ean_key] = idx

        if sku_key and sku_key not in sku_lookup:
            sku_lookup[sku_key] = idx

    rows = []

    text_columns = {
        "EAN",
        "SKU",
        "Cheapest Supplier",
        "Status",
    }

    for requested_code in requested_codes:
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
            row_dict = {}
            for col in full_result.columns:
                row_dict[col] = "" if col in text_columns else np.nan

            row_dict["_relevant_for_display"] = False
            lookup_status = "CODE NOT FOUND"

        row_dict["Requested Code"] = requested_code
        row_dict["Lookup Status"] = lookup_status
        rows.append(row_dict)

    filtered = pd.DataFrame(rows)

    first_columns = ["Requested Code", "Lookup Status"]
    remaining_columns = [
        col for col in full_result.columns
        if col not in first_columns
    ]

    return filtered[first_columns + remaining_columns]


def create_excel(
    full_result,
    supplier_price_columns,
    supplier_configs,
    duplicate_info,
    our_file_name,
):
    """
    Export ALL OurPrices products, not only products relevant for browser display.
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
        ("Products matched with at least one supplier", int(full_result["Cheapest Alternative Price"].notna().sum())),
        ("Products where at least one supplier is cheaper", int((full_result["Status"] == "CHEAPER").sum())),
        ("Products where cheapest supplier is more expensive", int((full_result["Status"] == "MORE EXPENSIVE").sum())),
        ("Products with same cheapest price", int((full_result["Status"] == "SAME PRICE").sum())),
        ("Products not found at any supplier", int((full_result["Status"] == "NOT FOUND").sum())),
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
                (f"{supplier_name} - matched products", int(full_result[price_col].notna().sum())),
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

        # -------------------------------------------------
        # Formats
        # -------------------------------------------------

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
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
                "bg_color": "#C6EFCE",
                "font_color": "#006100",
            }
        )

        cheapest_fmt = workbook.add_format(
            {
                "bg_color": "#FFF2CC",
                "font_color": "#7F6000",
                "bold": True,
            }
        )

        # Rewrite header with consistent formatting.
        for col_idx, name in enumerate(export_result.columns):
            worksheet.write(0, col_idx, name, header_fmt)

        # -------------------------------------------------
        # Column widths and number formats
        # -------------------------------------------------

        widths = {
            "EAN": 17,
            "SKU": 22,
            "Our Price": 14,
            "Cheapest Alternative Price": 23,
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
            "Cheapest Alternative Price",
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

        worksheet.freeze_panes(1, 0)

        if len(export_result):
            worksheet.autofilter(
                0,
                0,
                len(export_result),
                len(export_result.columns) - 1,
            )

            # -------------------------------------------------
            # Highlight every supplier price that is cheaper
            # than Our Price.
            # -------------------------------------------------

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

            # -------------------------------------------------
            # Highlight the chosen cheapest price + supplier.
            # This works whether or not that price beats ours.
            # -------------------------------------------------

            cheapest_price_idx = positions["Cheapest Alternative Price"]
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

        # Summary sheet formatting
        summary_ws.set_column("A:A", 48)
        summary_ws.set_column("B:B", 45)

        for col_idx, name in enumerate(summary_df.columns):
            summary_ws.write(0, col_idx, name, header_fmt)

        summary_ws.freeze_panes(1, 0)

    buffer.seek(0)
    return buffer.getvalue()



def style_browser_table(df, supplier_price_columns):
    """Highlight cheaper supplier cells and the cheapest supplier/price."""

    def style_row(row):
        styles = pd.Series("", index=row.index, dtype=object)

        if row.get("Lookup Status", "") == "CODE NOT FOUND":
            styles[:] = "background-color: #f8d7da; color: #842029;"
            return styles

        our_price = row.get("Our Price", np.nan)
        cheapest_price = row.get("Cheapest Alternative Price", np.nan)
        cheapest_supplier = row.get("Cheapest Supplier", "")

        # Mark each supplier price that is cheaper than our current price.
        for col in supplier_price_columns:
            supplier_price = row.get(col, np.nan)
            if (
                pd.notna(supplier_price)
                and pd.notna(our_price)
                and supplier_price < our_price
            ):
                styles[col] = "background-color: #d8f3dc; color: #006100;"

        # Mark the selected cheapest offer.
        if pd.notna(cheapest_price):
            if "Cheapest Alternative Price" in styles.index:
                styles["Cheapest Alternative Price"] = (
                    "background-color: #fff2cc; color: #7f6000; font-weight: bold;"
                )

            if "Cheapest Supplier" in styles.index and cheapest_supplier:
                styles["Cheapest Supplier"] = (
                    "background-color: #fff2cc; color: #7f6000; font-weight: bold;"
                )

        return styles

    formatters = {
        "Our Price": "€{:.2f}",
        "Cheapest Alternative Price": "€{:.2f}",
        "Saving €": "€{:.2f}",
        "Saving %": "{:.2%}",
        "Realisation Summ": "{:.2f}",
    }

    for col in supplier_price_columns:
        formatters[col] = "€{:.2f}"

    # Only use formatters for columns that currently exist.
    formatters = {
        key: value
        for key, value in formatters.items()
        if key in df.columns
    }

    return df.style.apply(style_row, axis=1).format(formatters, na_rep="")


# =========================================================
# SESSION STATE
# =========================================================

# The comparison is stored here. This is the important part that prevents the
# result from disappearing every time the user changes the Display radio.
if "comparison_result" not in st.session_state:
    st.session_state.comparison_result = None

if "comparison_supplier_columns" not in st.session_state:
    st.session_state.comparison_supplier_columns = []

if "comparison_supplier_configs" not in st.session_state:
    st.session_state.comparison_supplier_configs = []

if "comparison_duplicate_info" not in st.session_state:
    st.session_state.comparison_duplicate_info = {}

if "comparison_our_file_name" not in st.session_state:
    st.session_state.comparison_our_file_name = ""

if "comparison_code_filter_active" not in st.session_state:
    st.session_state.comparison_code_filter_active = False

if "comparison_code_filter_codes" not in st.session_state:
    st.session_state.comparison_code_filter_codes = []


# =========================================================
# WEB INTERFACE
# =========================================================

st.title("Price List Comparison")

st.write(
    "Compare your fixed **OurPrices** workbook with one or several alternative "
    "supplier price lists. Each supplier can use SKU, EAN, or combined EAN + SKU "
    "matching with its own identifier/price columns."
)

st.info(
    "The browser table shows only products where **Realisation Summ > 0** or "
    "**Net Available Qty > 0**. The downloaded Excel contains **all products** "
    "from OurPrices."
)

st.warning(
    "Make sure the supplier price and Our Price use the same basis "
    "(for example both net prices or both gross prices)."
)


# ---------------------------------------------------------
# 1. OURPRICES UPLOAD
# ---------------------------------------------------------

st.subheader("1. Upload OurPrices")

our_file = st.file_uploader(
    "Our current price list",
    type=["xlsx", "xlsm", "xls"],
    key="our_prices_upload",
)

if our_file is None:
    st.caption("Upload OurPrices to continue.")
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

our_sheet = st.selectbox(
    "OurPrices sheet",
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

st.success(f"OurPrices validated: {len(our_df):,} rows loaded.")


# ---------------------------------------------------------
# 2. SUPPLIER MODE
# ---------------------------------------------------------

st.subheader("2. Alternative suppliers")

upload_mode = st.radio(
    "Upload mode",
    ["Single supplier", "Bulk suppliers"],
    horizontal=True,
    key="supplier_upload_mode",
)

if upload_mode == "Single supplier":
    single_file = st.file_uploader(
        "Upload alternative supplier price list",
        type=["xlsx", "xlsm", "xls"],
        key="single_supplier_upload",
    )
    supplier_files = [single_file] if single_file is not None else []
else:
    supplier_files = st.file_uploader(
        "Upload alternative supplier price lists",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        key="bulk_supplier_upload",
    )

if not supplier_files:
    st.caption("Upload at least one alternative supplier price list to continue.")
    st.stop()


# ---------------------------------------------------------
# 3. CONFIGURE EACH SUPPLIER
# ---------------------------------------------------------

st.subheader("3. Configure supplier mapping")

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
        f"Supplier {index}: {supplier_file.name}",
        expanded=True,
    ):
        supplier_name = st.text_input(
            "Supplier name",
            value=default_name,
            key=f"{base_key}_name",
            help=(
                "This name will be used in the result columns, for example "
                "'Schneider Price', and in the Cheapest Supplier field."
            ),
        ).strip()

        entered_supplier_names.append(supplier_name)

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

        match_method = st.radio(
            "Match this supplier using",
            ["SKU", "EAN", "EAN + SKU"],
            horizontal=True,
            key=f"{base_key}_match",
            help=(
                "EAN + SKU tries EAN first. If the EAN is missing or not found "
                "at this supplier, the app then tries SKU."
            ),
        )

        price_guess = guess_column(supplier_columns, "PRICE")

        if match_method == "EAN + SKU":
            ean_guess = guess_column(supplier_columns, "EAN")
            sku_guess = guess_column(supplier_columns, "SKU")

            col1, col2, col3 = st.columns(3)

            with col1:
                ean_column = st.selectbox(
                    "Column corresponding to EAN",
                    supplier_columns,
                    index=supplier_columns.index(ean_guess),
                    key=f"{base_key}_ean_col",
                )

            with col2:
                sku_column = st.selectbox(
                    "Column corresponding to SKU",
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
                st.warning(
                    "EAN and SKU columns are the same. Check this supplier mapping."
                )

            if price_column in (ean_column, sku_column):
                st.warning(
                    "Price column is also selected as an identifier column. "
                    "Check this supplier mapping."
                )

        elif match_method == "EAN":
            ean_guess = guess_column(supplier_columns, "EAN")

            col1, col2 = st.columns(2)

            with col1:
                ean_column = st.selectbox(
                    "Column corresponding to EAN",
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

            sku_column = None

            if ean_column == price_column:
                st.warning(
                    "Identifier and price columns are the same. Check this supplier mapping."
                )

        else:
            sku_guess = guess_column(supplier_columns, "SKU")

            col1, col2 = st.columns(2)

            with col1:
                sku_column = st.selectbox(
                    "Column corresponding to SKU",
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

            ean_column = None

            if sku_column == price_column:
                st.warning(
                    "Identifier and price columns are the same. Check this supplier mapping."
                )

        with st.expander("Preview this supplier file"):
            st.dataframe(
                supplier_df.head(15),
                use_container_width=True,
                hide_index=True,
            )

        if not supplier_name:
            configuration_errors.append(
                f"Supplier {index} ({supplier_file.name}) has no supplier name."
            )

        supplier_configs.append(
            {
                "supplier_name": supplier_name,
                "file_name": supplier_file.name,
                "sheet_name": sheet_name,
                "match_method": match_method,
                "ean_column": ean_column,
                "sku_column": sku_column,
                "price_column": price_column,
                "dataframe": supplier_df,
            }
        )


# Supplier names must be unique because they become Excel/result column names.
non_empty_names = [name for name in entered_supplier_names if name]
if len(non_empty_names) != len(set(non_empty_names)):
    configuration_errors.append(
        "Supplier names must be unique. Two uploaded files currently use the same supplier name."
    )

if configuration_errors:
    for error in configuration_errors:
        st.error(error)


# ---------------------------------------------------------
# 4. RUN COMPARISON
# ---------------------------------------------------------

st.subheader("4. Compare")

compare_disabled = bool(configuration_errors) or not supplier_configs

if st.button(
    "Compare prices",
    type="primary",
    use_container_width=True,
    disabled=compare_disabled,
):
    try:
        comparison_result, supplier_price_columns, duplicate_info = (
            compare_all_suppliers(
                our_df,
                supplier_configs,
            )
        )

        # Persist the result so changing Display controls does NOT erase it.
        st.session_state.comparison_result = comparison_result
        st.session_state.comparison_supplier_columns = supplier_price_columns
        st.session_state.comparison_supplier_configs = supplier_configs
        st.session_state.comparison_duplicate_info = duplicate_info
        st.session_state.comparison_our_file_name = our_file.name
        st.session_state.comparison_code_filter_active = False
        st.session_state.comparison_code_filter_codes = []

        st.success("Comparison completed.")

    except Exception as exc:
        st.exception(exc)


# =========================================================
# RESULTS - RENDER FROM SESSION STATE
# =========================================================

full_result = st.session_state.comparison_result

if full_result is not None:
    supplier_price_columns = st.session_state.comparison_supplier_columns
    saved_supplier_configs = st.session_state.comparison_supplier_configs
    duplicate_info = st.session_state.comparison_duplicate_info

    st.divider()
    st.subheader("5. Comparison results")

    # Browser display is intentionally limited to active/relevant products.
    relevant_result = full_result[
        full_result["_relevant_for_display"]
    ].copy()

    relevant_matched = int(
        relevant_result["Cheapest Alternative Price"].notna().sum()
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
        duplicate_text = "; ".join(
            f"{supplier}: {count}"
            for supplier, count in duplicate_suppliers.items()
        )
        st.warning(
            "Duplicated supplier identifiers were found. The lowest valid price "
            f"was used for each duplicated SKU/EAN. {duplicate_text}"
        )

    # -----------------------------------------------------
    # Manual EAN / SKU code filter
    # -----------------------------------------------------

    st.markdown("#### Filter by EAN / SKU codes")

    st.caption(
        "Paste one code per line. You can mix EANs and SKUs. "
        "The filtered result follows exactly the order you pasted. "
        "Manual code lookup searches all compared OurPrices products, including "
        "products normally hidden by the Realisation/Available browser filter."
    )

    st.text_area(
        "Codes",
        height=160,
        placeholder=(
            "3606486365540\n"
            "3606486365588\n"
            "A9R35240\n"
            "3606480089459"
        ),
        key="comparison_code_filter_input",
    )

    filter_col, clear_col = st.columns([1, 1])

    with filter_col:
        if st.button(
            "Filter",
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
                st.warning("Paste at least one EAN or SKU before clicking Filter.")

    with clear_col:
        if st.button(
            "Clear code filter",
            use_container_width=True,
            key="comparison_code_filter_clear",
        ):
            st.session_state.comparison_code_filter_active = False
            st.session_state.comparison_code_filter_codes = []

    manual_filter_active = st.session_state.comparison_code_filter_active
    requested_codes = st.session_state.comparison_code_filter_codes

    # -----------------------------------------------------
    # Normal Display filter
    # -----------------------------------------------------

    if manual_filter_active and requested_codes:
        shown = filter_result_by_codes(
            full_result,
            requested_codes,
        )

        # Internal helper column is never shown to the user.
        if "_relevant_for_display" in shown.columns:
            shown = shown.drop(columns=["_relevant_for_display"])

        code_not_found_count = int(
            (shown["Lookup Status"] == "CODE NOT FOUND").sum()
        )

        st.info(
            f"Code filter active: showing {len(shown):,} pasted code(s) in pasted order. "
            f"Code not found: {code_not_found_count:,}."
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
                relevant_result["Cheapest Alternative Price"].notna()
            ].copy()

        elif display_option == "Only not found":
            shown = relevant_result[
                relevant_result["Cheapest Alternative Price"].isna()
            ].copy()

        else:
            shown = relevant_result.copy()

        # Internal helper column is never shown to the user.
        shown = shown.drop(columns=["_relevant_for_display"])

        st.caption(
            f"Showing {len(shown):,} of {len(relevant_result):,} relevant products. "
            "All uploaded supplier price columns remain visible."
        )

    st.dataframe(
        style_browser_table(shown, supplier_price_columns),
        use_container_width=True,
        hide_index=True,
        height=620,
    )

    # -----------------------------------------------------
    # Download ALL products
    # -----------------------------------------------------

    excel_bytes = create_excel(
        full_result=full_result,
        supplier_price_columns=supplier_price_columns,
        supplier_configs=saved_supplier_configs,
        duplicate_info=duplicate_info,
        our_file_name=st.session_state.comparison_our_file_name,
    )

    st.download_button(
        "Download full comparison Excel",
        data=excel_bytes,
        file_name=(
            f"Price_Comparison_{datetime.now():%Y-%m-%d_%H-%M}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
        help=(
            "The downloaded workbook contains ALL products from OurPrices, "
            "including products hidden from the browser by the Realisation/Available filter."
        ),
    )

    st.caption(
        f"Excel export contains all {len(full_result):,} OurPrices products."
    )
