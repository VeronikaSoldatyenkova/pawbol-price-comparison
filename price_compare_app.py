import streamlit as st

from price_compare_compare import (
    build_code_lookups,
    compare_all_suppliers,
    comparison_signature,
    lightweight_supplier_configs,
)
from price_compare_export import create_excel
from price_compare_inputs_v3 import render_current_pricelist, render_supplier_inputs
from price_compare_results import render_results
from price_compare_theme import configure_theme, render_hero, section_title


def initialize_state():
    defaults = {
        "comparison_result": None,
        "comparison_supplier_columns": [],
        "comparison_supplier_configs": [],
        "comparison_duplicate_info": {},
        "comparison_our_file_name": "",
        "comparison_our_config": {},
        "comparison_signature": "",
        "comparison_excel_bytes": None,
        "comparison_ean_lookup": {},
        "comparison_sku_lookup": {},
        "comparison_code_filter_active": False,
        "comparison_code_filter_codes": [],
    }
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value


def _run_main_ui():
    initialize_state()
    render_hero()

    our = render_current_pricelist()
    if our is None:
        return

    supplier_configs, configuration_errors = render_supplier_inputs(our["errors"])

    section_title(4, "Run comparison")
    current_signature = comparison_signature(
        our["bytes"],
        our["sheet"],
        our["config"],
        supplier_configs,
    )
    compare_disabled = bool(configuration_errors) or not supplier_configs

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
            if our["type"] == "PowerBI Pricelist":
                st.caption(
                    "PowerBI mode keeps the existing stock/realisation browser filter. "
                    "The complete workbook is still exported."
                )
            else:
                st.caption(
                    "Free-format mode shows every current-pricelist row and only the "
                    "additional columns you selected."
                )

    if compare_clicked:
        try:
            with st.spinner("Comparing suppliers and preparing Excel..."):
                (
                    comparison_result,
                    supplier_price_columns,
                    duplicate_info,
                    _,
                ) = compare_all_suppliers(
                    our["dataframe"],
                    our["config"],
                    supplier_configs,
                )

                saved_suppliers = lightweight_supplier_configs(supplier_configs)
                saved_our_config = dict(our["config"])
                ean_lookup, sku_lookup = build_code_lookups(comparison_result)

                excel_bytes = create_excel(
                    full_result=comparison_result,
                    supplier_price_columns=supplier_price_columns,
                    supplier_configs=saved_suppliers,
                    duplicate_info=duplicate_info,
                    our_file_name=our["file"].name,
                    our_config=saved_our_config,
                )

            st.session_state.comparison_result = comparison_result
            st.session_state.comparison_supplier_columns = supplier_price_columns
            st.session_state.comparison_supplier_configs = saved_suppliers
            st.session_state.comparison_duplicate_info = duplicate_info
            st.session_state.comparison_our_file_name = our["file"].name
            st.session_state.comparison_our_config = saved_our_config
            st.session_state.comparison_signature = current_signature
            st.session_state.comparison_excel_bytes = excel_bytes
            st.session_state.comparison_ean_lookup = ean_lookup
            st.session_state.comparison_sku_lookup = sku_lookup
            st.session_state.comparison_code_filter_active = False
            st.session_state.comparison_code_filter_codes = []

            st.success("Comparison completed.")
        except Exception as exc:
            st.error("Comparison failed. The page is still usable; see details below.")
            st.exception(exc)

    render_results(current_signature)


def run_app():
    # Configure the page before any other Streamlit output.
    configure_theme()

    # A Python-side input/render exception must never leave the user with a blank page.
    # Streamlit will render the exception block instead of terminating the whole UI.
    try:
        _run_main_ui()
    except Exception as exc:
        st.error(
            "The interface encountered an unexpected error while rendering. "
            "The app did not intentionally stop; reload is not required to see the error."
        )
        st.exception(exc)


run_app()
