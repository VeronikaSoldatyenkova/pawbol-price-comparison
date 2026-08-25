from datetime import datetime

import streamlit as st

from price_compare_compare import filter_result_by_codes, parse_requested_codes
from price_compare_export import style_browser_table
from price_compare_theme import section_title


def render_results(current_signature):
    full_result = st.session_state.comparison_result
    if full_result is None:
        return

    results_are_stale = (
        st.session_state.comparison_signature != current_signature
    )

    if results_are_stale:
        st.divider()
        section_title(5, "Comparison results")
        st.warning(
            "Files or mapping settings changed after the last comparison. "
            "The previous result is hidden. Click **Compare prices** again to refresh it."
        )
        st.caption(
            "You can continue editing settings normally; the page is not stopped or truncated."
        )
        return

    supplier_price_columns = st.session_state.comparison_supplier_columns
    duplicate_info = st.session_state.comparison_duplicate_info
    our_config = st.session_state.comparison_our_config

    st.divider()
    section_title(5, "Comparison results")

    relevant_result = full_result[full_result["_relevant_for_display"]].copy()
    relevant_matched = int(relevant_result["Matched Suppliers"].gt(0).sum())
    relevant_cheaper = int((relevant_result["Status"] == "CHEAPER").sum())
    relevant_more_expensive = int(
        (relevant_result["Status"] == "MORE EXPENSIVE").sum()
    )
    relevant_not_found = int((relevant_result["Status"] == "NOT FOUND").sum())

    metric1, metric2, metric3, metric4, metric5 = st.columns(5)
    if our_config.get("type") == "PowerBI Pricelist":
        metric1.metric("Relevant products", f"{len(relevant_result):,}")
    else:
        metric1.metric("Products", f"{len(relevant_result):,}")
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
            "The lowest valid price was used. " + duplicate_text
        )

    with st.container(border=True):
        st.markdown("#### Quick code filter")
        st.caption(
            "Paste one or two columns directly from Excel: **EAN/SKU** and an optional "
            "**Target Price**. Results keep the exact pasted order and search all "
            "current-pricelist rows."
        )

        guide_left, guide_right = st.columns([1.3, 2.7])
        with guide_left:
            st.markdown(
                "**Column 1:** EAN or SKU  \n"
                "**Column 2:** Target Price *(optional)*"
            )
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
                "A9R35240\t12.50\n"
                "3606480303357"
            ),
            key="comparison_code_filter_input",
            label_visibility="collapsed",
        )

        filter_col, clear_col, _ = st.columns([1, 1, 3])
        with filter_col:
            if st.button(
                "Apply code filter",
                type="primary",
                use_container_width=True,
                key="comparison_code_filter_button",
            ):
                requested = parse_requested_codes(
                    st.session_state.get("comparison_code_filter_input", "")
                )
                if requested:
                    st.session_state.comparison_code_filter_codes = requested
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

    manual_active = st.session_state.comparison_code_filter_active
    requested_codes = st.session_state.comparison_code_filter_codes

    if manual_active and requested_codes:
        shown = filter_result_by_codes(
            full_result,
            requested_codes,
            st.session_state.comparison_ean_lookup,
            st.session_state.comparison_sku_lookup,
            supplier_price_columns,
        )
        if "_relevant_for_display" in shown.columns:
            shown = shown.drop(columns=["_relevant_for_display"])

        not_found = int((shown["Lookup Status"] == "CODE NOT FOUND").sum())
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
        info2.metric("Not found", f"{not_found:,}")
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
            ["All", "Only supplier cheaper", "Only matched", "Only not found"],
            horizontal=True,
            key="comparison_display_filter",
        )

        if display_option == "Only supplier cheaper":
            shown = relevant_result[relevant_result["Status"] == "CHEAPER"].copy()
        elif display_option == "Only matched":
            shown = relevant_result[relevant_result["Matched Suppliers"].gt(0)].copy()
        elif display_option == "Only not found":
            shown = relevant_result[relevant_result["Matched Suppliers"].eq(0)].copy()
        else:
            shown = relevant_result.copy()

        shown = shown.drop(columns=["_relevant_for_display"])
        if our_config.get("type") == "PowerBI Pricelist":
            st.caption(
                f"Showing {len(shown):,} of {len(relevant_result):,} relevant products."
            )
        else:
            st.caption(
                f"Showing {len(shown):,} of {len(relevant_result):,} products."
            )

    st.dataframe(
        style_browser_table(shown, supplier_price_columns),
        use_container_width=True,
        hide_index=True,
        height=620,
    )

    with st.container(border=True):
        download_col, info_col = st.columns([1.2, 2])
        with download_col:
            st.download_button(
                "Download full comparison Excel",
                data=st.session_state.comparison_excel_bytes,
                file_name=f"Price_Comparison_{datetime.now():%Y-%m-%d_%H-%M}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True,
            )

        with info_col:
            if our_config.get("type") == "PowerBI Pricelist":
                st.caption(
                    f"Excel contains all {len(full_result):,} PowerBI Pricelist products, "
                    "including rows hidden by the browser relevance filter."
                )
            else:
                st.caption(
                    f"Excel contains all {len(full_result):,} Free format pricelist products "
                    "and the additional columns selected during mapping."
                )
