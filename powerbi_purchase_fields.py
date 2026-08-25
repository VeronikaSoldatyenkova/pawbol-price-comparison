"""PowerBI-specific purchase fields and result layout.

The PowerBI current pricelist has two additional fixed fields:
- Последняя Цена Закупки
- Последняя Дата Закупки

The module also keeps all comparison-price columns together in the final
result. Operational PowerBI fields are moved to the far right so they do not
separate Our Price from supplier prices.
"""

import pandas as pd

import core as _core


LAST_PURCHASE_PRICE = "Последняя Цена Закупки"
LAST_PURCHASE_DATE = "Последняя Дата Закупки"

_applied = False
_original_prepare_current_pricelist = _core.prepare_current_pricelist
_original_dataframe_to_table_model = _core.dataframe_to_table_model
_original_compare_all = _core.compare_all


def _append_once(values, item):
    if item not in values:
        values.append(item)


def _prepare_current_pricelist(current_df, config):
    result, extra_columns = _original_prepare_current_pricelist(current_df, config)

    if config.get("type") == "PowerBI Pricelist" and LAST_PURCHASE_DATE in result.columns:
        # Keep this as a real datetime value so browser sorting and XLSX export
        # behave as dates instead of lexicographic text.
        result[LAST_PURCHASE_DATE] = pd.to_datetime(
            result[LAST_PURCHASE_DATE],
            errors="coerce",
        )

    return result, extra_columns


def _compare_all(current_df, current_config, supplier_items):
    """Run the normal comparison, then apply the PowerBI-specific column order."""
    result, supplier_price_columns, duplicate_info, current_extra_columns = (
        _original_compare_all(current_df, current_config, supplier_items)
    )

    if current_config.get("type") != "PowerBI Pricelist":
        return result, supplier_price_columns, duplicate_info, current_extra_columns

    # Keep every price/comparison field together.  PowerBI operational fields
    # are useful context, but belong at the right-hand side of the table.
    powerbi_tail = [
        col
        for col in _core.POWERBI_OUTPUT_COLUMNS
        if col in result.columns
    ]

    comparison_columns = [
        "EAN",
        "SKU",
        "Our Price",
        *[col for col in supplier_price_columns if col in result.columns],
        "Cheapest Price",
        "Cheapest Supplier",
        "Saving €",
        "Saving %",
        "Status",
        "Matched Suppliers",
    ]

    ordered = []
    for col in comparison_columns:
        if col in result.columns and col not in ordered:
            ordered.append(col)

    # Preserve any future comparison fields that are neither PowerBI context
    # columns nor the private display flag.
    for col in result.columns:
        if (
            col not in ordered
            and col not in powerbi_tail
            and col != "_relevant_for_display"
        ):
            ordered.append(col)

    ordered.extend(col for col in powerbi_tail if col not in ordered)
    if "_relevant_for_display" in result.columns:
        ordered.append("_relevant_for_display")

    result = result[ordered].copy()
    return result, supplier_price_columns, duplicate_info, current_extra_columns


def _dataframe_to_table_model(df, supplier_price_columns, target_mode=False):
    model = _original_dataframe_to_table_model(
        df,
        supplier_price_columns,
        target_mode=target_mode,
    )

    columns = model.get("columns", [])
    rows = model.get("rows", [])

    if LAST_PURCHASE_PRICE in columns:
        idx = columns.index(LAST_PURCHASE_PRICE)
        if LAST_PURCHASE_PRICE in df.columns:
            values = df[LAST_PURCHASE_PRICE].tolist()
            for row_cells, value in zip(rows, values):
                if pd.isna(value):
                    row_cells[idx]["value"] = ""
                else:
                    try:
                        row_cells[idx]["value"] = f"€{float(value):,.2f}"
                    except (TypeError, ValueError):
                        row_cells[idx]["value"] = str(value)

    if LAST_PURCHASE_DATE in columns:
        idx = columns.index(LAST_PURCHASE_DATE)
        if LAST_PURCHASE_DATE in df.columns:
            values = df[LAST_PURCHASE_DATE].tolist()
            for row_cells, value in zip(rows, values):
                if pd.isna(value):
                    row_cells[idx]["value"] = ""
                else:
                    parsed = pd.to_datetime(value, errors="coerce")
                    row_cells[idx]["value"] = (
                        parsed.strftime("%Y-%m-%d")
                        if pd.notna(parsed)
                        else str(value)
                    )

    return model


def apply():
    global _applied
    if _applied:
        return

    # Mutate the lists in place. main.py imports these list objects from core,
    # so in-place updates are visible there as well.
    _append_once(_core.POWERBI_REQUIRED_COLUMNS, LAST_PURCHASE_PRICE)
    _append_once(_core.POWERBI_REQUIRED_COLUMNS, LAST_PURCHASE_DATE)
    _append_once(_core.POWERBI_NUMERIC_COLUMNS, LAST_PURCHASE_PRICE)
    _append_once(_core.POWERBI_OUTPUT_COLUMNS, LAST_PURCHASE_PRICE)
    _append_once(_core.POWERBI_OUTPUT_COLUMNS, LAST_PURCHASE_DATE)

    _core.prepare_current_pricelist = _prepare_current_pricelist
    _core.compare_all = _compare_all
    _core.dataframe_to_table_model = _dataframe_to_table_model
    _applied = True
