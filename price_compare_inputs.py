from pathlib import Path

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


def render_current_pricelist():
    section_title(1, "Our current price list")

    with st.container(border=True):
        our_type = st.radio(
            "Pricelist type",
            ["PowerBI Pricelist", "Free format pricelist"],
            horizontal=True,
            key="our_pricelist_type",
            help=(
                "PowerBI Pricelist uses the existing fixed export structure. "
                "Free format lets you map SKU/EAN/Price and select extra columns."
            ),
        )

        if our_type == "PowerBI Pricelist":
            st.markdown(
                """
                <div class="pc-note">
                    Keeps the existing PowerBI behavior. The browser shows only products where
                    <b>Realisation Summ &gt; 0</b> or <b>Net Available Qty &gt; 0</b>.
                    Excel export still contains all rows.
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                """
                <div class="pc-note">
                    Map at least one identifier (<b>EAN</b> or <b>SKU</b>) plus the
                    <b>Price</b> column. Free-format files do not use stock, realisation,
                    or last-sale filters, so every row is shown.
                </div>
                """,
                unsafe_allow_html=True,
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
            st.stop()

        our_bytes = our_file.getvalue()

        try:
            our_sheets = get_sheet_names(our_bytes)
        except Exception as exc:
            st.error(f"Could not open current pricelist: {exc}")
            st.stop()

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
            st.stop()

        if our_df.empty and len(our_df.columns) == 0:
            st.error("The selected current-pricelist sheet has no columns.")
            st.stop()

        our_columns = list(our_df.columns)
        errors = []

        if our_type == "PowerBI Pricelist":
            missing = [col for col in POWERBI_REQUIRED_COLUMNS if col not in our_df.columns]
            if missing:
                st.error(
                    "PowerBI Pricelist is missing required columns: " + ", ".join(missing)
                )
                st.write("Columns found:", our_columns)
                st.stop()

            our_config = {
                "type": "PowerBI Pricelist",
                "ean_column": "EAN",
                "sku_column": "SKU",
                "price_column": "Latest Pricelist Value",
                "extra_columns": [],
            }
            st.success(f"PowerBI Pricelist validated · {len(our_df):,} rows")

        else:
            st.markdown("**Map free-format columns**")
            none_label = "— Not selected —"
            ean_guess = guess_column(our_columns, "EAN")
            sku_guess = guess_column(our_columns, "SKU")
            price_guess = guess_column(our_columns, "PRICE")

            col1, col2, col3 = st.columns(3)
            with col1:
                ean_options = [none_label] + our_columns
                ean_index = ean_options.index(ean_guess) if ean_guess in ean_options else 0
                selected_ean = st.selectbox(
                    "EAN column",
                    ean_options,
                    index=ean_index,
                    key="free_our_ean_column",
                )

            with col2:
                sku_options = [none_label] + our_columns
                sku_index = sku_options.index(sku_guess) if sku_guess in sku_options else 0
                selected_sku = st.selectbox(
                    "SKU column",
                    sku_options,
                    index=sku_index,
                    key="free_our_sku_column",
                )

            with col3:
                price_column = st.selectbox(
                    "Price column",
                    our_columns,
                    index=our_columns.index(price_guess),
                    key="free_our_price_column",
                )

            ean_column = None if selected_ean == none_label else selected_ean
            sku_column = None if selected_sku == none_label else selected_sku

            excluded = {col for col in [ean_column, sku_column, price_column] if col is not None}
            extra_options = [col for col in our_columns if col not in excluded]
            extra_columns = st.multiselect(
                "Additional columns to display",
                options=extra_options,
                default=[],
                key="free_our_extra_columns",
                help=(
                    "Optional. Selected columns are carried into both the browser "
                    "comparison and the downloaded Excel file."
                ),
            )

            if not ean_column and not sku_column:
                errors.append(
                    "Free format pricelist: select at least one identifier column (EAN or SKU)."
                )
            if ean_column and sku_column and ean_column == sku_column:
                errors.append("Free format pricelist: EAN and SKU columns cannot be the same.")
            if price_column in {ean_column, sku_column}:
                errors.append(
                    "Free format pricelist: Price column cannot also be an identifier column."
                )

            valid_prices = int(parse_numeric_series(our_df[price_column]).notna().sum())
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

            preview_col, stat_col = st.columns([3, 1])
            with preview_col:
                show_preview = st.checkbox(
                    "Show first 10 rows",
                    key="free_our_preview_toggle",
                )
                if show_preview:
                    st.dataframe(
                        our_df.head(10),
                        use_container_width=True,
                        hide_index=True,
                        height=390,
                    )

            with stat_col:
                st.metric("Rows", f"{len(our_df):,}")
                st.metric("Valid prices", f"{valid_prices:,}")

            if errors:
                for error in errors:
                    st.error(error)
            else:
                identifiers = " + ".join(
                    name
                    for name, value in [("EAN", ean_column), ("SKU", sku_column)]
                    if value
                )
                st.success(
                    f"Free format pricelist mapped · {len(our_df):,} rows · "
                    f"Identifiers: {identifiers}"
                )

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

    section_title(3, "Supplier mapping")
    supplier_configs = []
    configuration_errors = list(initial_errors or [])
    entered_names = []

    for index, supplier_file in enumerate(supplier_files, start=1):
        file_bytes = supplier_file.getvalue()
        base_key = safe_widget_key(f"supplier_{index}_{supplier_file.name}")
        default_name = Path(supplier_file.name).stem

        try:
            sheets = get_sheet_names(file_bytes)
        except Exception as exc:
            configuration_errors.append(f"Could not open {supplier_file.name}: {exc}")
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
            entered_names.append(supplier_name)

            with top2:
                sheet_name = st.selectbox(
                    "Excel sheet", sheets, key=f"{base_key}_sheet"
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

            columns = list(supplier_df.columns)
            st.markdown("**Matching**")
            match_method = st.radio(
                "Match using",
                ["SKU", "EAN", "EAN + SKU"],
                horizontal=True,
                key=f"{base_key}_match",
                help="EAN + SKU tries EAN first and uses SKU only if EAN does not match.",
                label_visibility="collapsed",
            )

            price_guess = guess_column(columns, "PRICE")
            ean_column = None
            sku_column = None

            if match_method == "EAN + SKU":
                ean_guess = guess_column(columns, "EAN")
                sku_guess = guess_column(columns, "SKU")
                col1, col2, col3 = st.columns(3)
                with col1:
                    ean_column = st.selectbox(
                        "EAN column",
                        columns,
                        index=columns.index(ean_guess),
                        key=f"{base_key}_ean_col",
                    )
                with col2:
                    sku_column = st.selectbox(
                        "SKU column",
                        columns,
                        index=columns.index(sku_guess),
                        key=f"{base_key}_sku_col",
                    )
                with col3:
                    price_column = st.selectbox(
                        "Price column",
                        columns,
                        index=columns.index(price_guess),
                        key=f"{base_key}_price_col_both",
                    )

                if ean_column == sku_column:
                    configuration_errors.append(
                        f"{supplier_name or supplier_file.name}: EAN and SKU columns cannot be the same."
                    )
                if price_column in (ean_column, sku_column):
                    configuration_errors.append(
                        f"{supplier_name or supplier_file.name}: Price column cannot also be an identifier column."
                    )

            elif match_method == "EAN":
                ean_guess = guess_column(columns, "EAN")
                col1, col2 = st.columns(2)
                with col1:
                    ean_column = st.selectbox(
                        "EAN column",
                        columns,
                        index=columns.index(ean_guess),
                        key=f"{base_key}_ean_col",
                    )
                with col2:
                    price_column = st.selectbox(
                        "Price column",
                        columns,
                        index=columns.index(price_guess),
                        key=f"{base_key}_price_col_ean",
                    )
                if ean_column == price_column:
                    configuration_errors.append(
                        f"{supplier_name or supplier_file.name}: EAN and Price columns cannot be the same."
                    )

            else:
                sku_guess = guess_column(columns, "SKU")
                col1, col2 = st.columns(2)
                with col1:
                    sku_column = st.selectbox(
                        "SKU column",
                        columns,
                        index=columns.index(sku_guess),
                        key=f"{base_key}_sku_col",
                    )
                with col2:
                    price_column = st.selectbox(
                        "Price column",
                        columns,
                        index=columns.index(price_guess),
                        key=f"{base_key}_price_col_sku",
                    )
                if sku_column == price_column:
                    configuration_errors.append(
                        f"{supplier_name or supplier_file.name}: SKU and Price columns cannot be the same."
                    )

            preview_col, stat_col = st.columns([3, 1])
            with preview_col:
                show_preview = st.checkbox(
                    "Show first 10 rows", key=f"{base_key}_preview_toggle"
                )
                if show_preview:
                    st.dataframe(
                        supplier_df.head(10),
                        use_container_width=True,
                        hide_index=True,
                        height=390,
                    )

            with stat_col:
                valid_price_count = int(
                    parse_numeric_series(supplier_df[price_column]).notna().sum()
                )
                st.metric("Rows", f"{len(supplier_df):,}")
                st.metric("Valid prices", f"{valid_price_count:,}")

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
                    "file_digest": file_digest(file_bytes),
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
