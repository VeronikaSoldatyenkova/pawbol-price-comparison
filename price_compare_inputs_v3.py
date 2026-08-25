from pathlib import Path

import pandas as pd
import streamlit as st

from price_compare_common import (
    POWERBI_REQUIRED_COLUMNS,
    file_digest,
    get_sheet_names,
    guess_column,
    parse_numeric_series,
    read_excel_sheet,
    safe_widget_key,
)
from price_compare_theme import section_title


NONE_LABEL = "— Not selected —"


def _safe_index(options, value, fallback=0):
    try:
        return options.index(value)
    except (ValueError, TypeError):
        return fallback


def _valid_price_count(df, price_column):
    """Keep price validation isolated so a bad column never breaks the whole UI."""
    try:
        return int(parse_numeric_series(df[price_column]).notna().sum())
    except Exception:
        return 0


def render_current_pricelist():
    """
    Render current-pricelist inputs with a stable widget tree.

    Important stability rule: after a file is loaded, the same EAN/SKU/Price mapping
    widgets are always rendered. Switching PowerBI <-> Free format therefore does not
    create/remove a large group of widgets during a Streamlit rerun.
    """
    section_title(1, "Our current price list")

    with st.container(border=True):
        our_type = st.radio(
            "Pricelist type",
            ["PowerBI Pricelist", "Free format pricelist"],
            horizontal=True,
            key="our_pricelist_type",
            help=(
                "PowerBI Pricelist uses the fixed PowerBI export structure. "
                "Free format lets you map EAN/SKU/Price and carry extra columns."
            ),
        )

        left, right = st.columns([2, 1])
        with left:
            our_file = st.file_uploader(
                "Upload current pricelist",
                type=["xlsx", "xlsm", "xls"],
                key="our_prices_upload",
            )

        if our_file is None:
            with right:
                st.markdown(
                    '<div class="pc-muted">Upload a pricelist to begin.</div>',
                    unsafe_allow_html=True,
                )
            return None

        our_bytes = our_file.getvalue()

        try:
            our_sheets = get_sheet_names(our_bytes)
        except Exception as exc:
            st.error(f"Could not open current pricelist: {exc}")
            return None

        if not our_sheets:
            st.error("No worksheets were found in the current pricelist.")
            return None

        default_index = (
            our_sheets.index("Export")
            if our_type == "PowerBI Pricelist" and "Export" in our_sheets
            else 0
        )

        with right:
            our_sheet = st.selectbox(
                "Sheet",
                our_sheets,
                index=default_index,
                key="our_prices_sheet",
            )

        try:
            our_df = read_excel_sheet(our_bytes, our_sheet)
        except Exception as exc:
            st.error(f"Could not read current pricelist sheet: {exc}")
            return None

        if len(our_df.columns) == 0:
            st.error("The selected current-pricelist sheet has no columns.")
            return None

        our_columns = list(our_df.columns)
        errors = []

        # Stable mapping widgets: always present after the workbook is loaded.
        st.markdown("**Column mapping**")
        ean_guess = guess_column(our_columns, "EAN")
        sku_guess = guess_column(our_columns, "SKU")
        price_guess = guess_column(our_columns, "PRICE")

        ean_options = [NONE_LABEL] + our_columns
        sku_options = [NONE_LABEL] + our_columns

        if our_type == "PowerBI Pricelist":
            default_ean = "EAN" if "EAN" in our_columns else ean_guess
            default_sku = "SKU" if "SKU" in our_columns else sku_guess
            default_price = (
                "Latest Pricelist Value"
                if "Latest Pricelist Value" in our_columns
                else price_guess
            )
        else:
            default_ean = ean_guess
            default_sku = sku_guess
            default_price = price_guess

        col1, col2, col3 = st.columns(3)
        with col1:
            selected_ean = st.selectbox(
                "EAN column",
                ean_options,
                index=_safe_index(ean_options, default_ean),
                key="our_mapping_ean_column",
                help="For PowerBI mode this is fixed to EAN by validation.",
            )
        with col2:
            selected_sku = st.selectbox(
                "SKU column",
                sku_options,
                index=_safe_index(sku_options, default_sku),
                key="our_mapping_sku_column",
                help="For PowerBI mode this is fixed to SKU by validation.",
            )
        with col3:
            price_column = st.selectbox(
                "Price column",
                our_columns,
                index=_safe_index(our_columns, default_price),
                key="our_mapping_price_column",
            )

        ean_column = None if selected_ean == NONE_LABEL else selected_ean
        sku_column = None if selected_sku == NONE_LABEL else selected_sku

        excluded = {
            col for col in [ean_column, sku_column, price_column] if col is not None
        }
        extra_options = [col for col in our_columns if col not in excluded]
        extra_columns = st.multiselect(
            "Additional columns to display",
            options=extra_options,
            default=[],
            key="our_mapping_extra_columns",
            help=(
                "Used only for Free format pricelists. Selected columns are included "
                "in the browser result and downloaded Excel."
            ),
        )

        valid_prices = _valid_price_count(our_df, price_column)
        info1, info2 = st.columns(2)
        info1.metric("Rows", f"{len(our_df):,}")
        info2.metric("Valid prices", f"{valid_prices:,}")

        if our_type == "PowerBI Pricelist":
            missing = [col for col in POWERBI_REQUIRED_COLUMNS if col not in our_df.columns]
            if missing:
                errors.append(
                    "PowerBI Pricelist is missing required columns: " + ", ".join(missing)
                )

            # PowerBI behavior remains fixed regardless of what a stale widget value
            # may contain after switching from Free format mode.
            our_config = {
                "type": "PowerBI Pricelist",
                "ean_column": "EAN",
                "sku_column": "SKU",
                "price_column": "Latest Pricelist Value",
                "extra_columns": [],
            }

            if not errors:
                st.success(f"PowerBI Pricelist validated · {len(our_df):,} rows")
                st.caption(
                    "PowerBI mode keeps the existing Realisation / Net Available browser filter."
                )
        else:
            if not ean_column and not sku_column:
                errors.append(
                    "Free format pricelist: select at least one identifier column (EAN or SKU)."
                )
            if ean_column and sku_column and ean_column == sku_column:
                errors.append(
                    "Free format pricelist: EAN and SKU columns cannot be the same."
                )
            if price_column in {ean_column, sku_column}:
                errors.append(
                    "Free format pricelist: Price column cannot also be an identifier column."
                )
            if valid_prices == 0:
                errors.append(
                    "Free format pricelist: no valid numeric prices were found in the selected Price column."
                )

            our_config = {
                "type": "Free format pricelist",
                "ean_column": ean_column,
                "sku_column": sku_column,
                "price_column": price_column,
                "extra_columns": list(extra_columns),
            }

            if not errors:
                identifiers = " + ".join(
                    name
                    for name, value in [("EAN", ean_column), ("SKU", sku_column)]
                    if value
                )
                st.success(
                    f"Free format pricelist mapped · {len(our_df):,} rows · "
                    f"Identifiers: {identifiers}"
                )

        for error in dict.fromkeys(errors):
            st.error(error)

    return {
        "type": our_type,
        "file": our_file,
        "bytes": our_bytes,
        "sheet": our_sheet,
        "dataframe": our_df,
        "config": our_config,
        "errors": errors,
    }


def render_supplier_inputs(initial_errors=None):
    """
    Render suppliers without expanders/popovers or branch-dependent selectboxes.

    Every supplier always owns exactly these widgets:
      supplier name, sheet, match method, EAN column, SKU column, price column.
    Changing SKU/EAN/EAN+SKU therefore keeps the React/Streamlit element tree stable.
    """
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
                "EAN + SKU always tries EAN first and then SKU."
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

    section_title(3, "Supplier mapping")
    supplier_configs = []
    configuration_errors = list(initial_errors or [])
    entered_names = []

    if not supplier_files:
        st.caption("Upload at least one alternative supplier price list to continue.")
        return supplier_configs, configuration_errors

    for index, supplier_file in enumerate(supplier_files, start=1):
        file_bytes = supplier_file.getvalue()
        # Include a digest fragment so replacing a file with the same filename does
        # not leave incompatible widget state behind.
        digest = file_digest(file_bytes)
        base_key = safe_widget_key(
            f"supplier_{index}_{supplier_file.name}_{digest[:10]}"
        )
        default_name = Path(supplier_file.name).stem

        try:
            sheets = get_sheet_names(file_bytes)
        except Exception as exc:
            configuration_errors.append(f"Could not open {supplier_file.name}: {exc}")
            continue

        if not sheets:
            configuration_errors.append(f"{supplier_file.name}: no worksheets found.")
            continue

        with st.container(border=True):
            st.markdown(f"#### {index}. {supplier_file.name}")

            top1, top2 = st.columns([1.3, 1])
            with top1:
                supplier_name = st.text_input(
                    "Supplier name",
                    value=default_name,
                    key=f"{base_key}_name",
                    help="Used in result columns and in Cheapest Supplier.",
                ).strip()
            entered_names.append(supplier_name)

            with top2:
                sheet_name = st.selectbox(
                    "Excel sheet",
                    sheets,
                    key=f"{base_key}_sheet",
                )

            try:
                supplier_df = read_excel_sheet(file_bytes, sheet_name)
            except Exception as exc:
                configuration_errors.append(
                    f"Could not read {supplier_file.name} / {sheet_name}: {exc}"
                )
                continue

            if len(supplier_df.columns) == 0:
                configuration_errors.append(
                    f"{supplier_file.name} / {sheet_name} has no columns."
                )
                continue

            columns = list(supplier_df.columns)
            match_method = st.radio(
                "Match using",
                ["SKU", "EAN", "EAN + SKU"],
                horizontal=True,
                key=f"{base_key}_match",
                help="EAN + SKU tries EAN first and uses SKU only if EAN does not match.",
            )

            # IMPORTANT: all three mapping selectboxes are ALWAYS rendered.
            # We only decide later which identifiers are actually used.
            ean_guess = guess_column(columns, "EAN")
            sku_guess = guess_column(columns, "SKU")
            price_guess = guess_column(columns, "PRICE")

            col1, col2, col3 = st.columns(3)
            with col1:
                selected_ean_column = st.selectbox(
                    "EAN column",
                    columns,
                    index=_safe_index(columns, ean_guess),
                    key=f"{base_key}_ean_col",
                    help="Used when match mode is EAN or EAN + SKU.",
                )
            with col2:
                selected_sku_column = st.selectbox(
                    "SKU column",
                    columns,
                    index=_safe_index(columns, sku_guess),
                    key=f"{base_key}_sku_col",
                    help="Used when match mode is SKU or EAN + SKU.",
                )
            with col3:
                price_column = st.selectbox(
                    "Price column",
                    columns,
                    index=_safe_index(columns, price_guess),
                    key=f"{base_key}_price_col",
                )

            ean_column = (
                selected_ean_column
                if match_method in ("EAN", "EAN + SKU")
                else None
            )
            sku_column = (
                selected_sku_column
                if match_method in ("SKU", "EAN + SKU")
                else None
            )

            if match_method == "EAN + SKU" and selected_ean_column == selected_sku_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: EAN and SKU columns cannot be the same."
                )
            if match_method in ("EAN", "EAN + SKU") and price_column == selected_ean_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: Price column cannot also be the EAN column."
                )
            if match_method in ("SKU", "EAN + SKU") and price_column == selected_sku_column:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: Price column cannot also be the SKU column."
                )

            valid_price_count = _valid_price_count(supplier_df, price_column)
            stat1, stat2, stat3 = st.columns(3)
            stat1.metric("Rows", f"{len(supplier_df):,}")
            stat2.metric("Columns", f"{len(columns):,}")
            stat3.metric("Valid prices", f"{valid_price_count:,}")

            # Preview was intentionally removed from this highly interactive area.
            # It was not needed for comparison logic and was a second source of
            # frontend reconciliation failures during mapping reruns.
            st.caption(
                "Mapping controls are kept static for stability. "
                "Changing match mode no longer creates or removes UI components."
            )

            if valid_price_count == 0:
                configuration_errors.append(
                    f"{supplier_name or supplier_file.name}: no valid numeric prices were found in the selected Price column."
                )
            if not supplier_name:
                configuration_errors.append(
                    f"Supplier {index} ({supplier_file.name}) has no supplier name."
                )

            supplier_configs.append(
                {
                    "supplier_name": supplier_name,
                    "file_name": supplier_file.name,
                    "file_digest": digest,
                    "sheet_name": sheet_name,
                    "match_method": match_method,
                    "ean_column": ean_column,
                    "sku_column": sku_column,
                    "price_column": price_column,
                    "dataframe": supplier_df,
                }
            )

    non_empty_names = [name.casefold() for name in entered_names if name]
    if len(non_empty_names) != len(set(non_empty_names)):
        configuration_errors.append("Supplier names must be unique (case-insensitive).")

    if configuration_errors:
        for error in dict.fromkeys(configuration_errors):
            st.error(error)

    return supplier_configs, configuration_errors
