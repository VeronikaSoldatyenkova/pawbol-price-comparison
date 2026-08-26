"""Special supplier-price workflow for Quick code filter.

Adds a second quick-filter mode where pasted EAN/SKU + price pairs become a
Special Price reference.  In Legrand mode the pasted price is adjusted toward
Our Price using:

    Special + ((Our Price - Special) * 0.3 * 1.25)

The adjusted value is the value displayed and used for all comparisons.
"""

from __future__ import annotations

import io
import time

import numpy as np
import pandas as pd
from fastapi import HTTPException, Request
from fastapi.responses import Response

import core
import main as _main
import runtime_main as _runtime


XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SPECIAL_PRICE = "Special Price"
BELOW_SPECIAL = "Below Special Price"


def _is_checked(value) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def _recalculate_legrand_price(raw_special, our_price):
    if pd.isna(raw_special) or pd.isna(our_price):
        return raw_special
    special = float(raw_special)
    our = float(our_price)
    return special + ((our - special) * 0.3 * 1.25)


def build_special_price_result(
    full_result: pd.DataFrame,
    requests_list: list[dict],
    ean_lookup: dict,
    sku_lookup: dict,
    supplier_price_columns: list[str],
    legrand: bool = False,
):
    """Build Quick Filter rows using Special Price instead of Target Price."""
    # Reuse the established code matching/order/not-found behaviour first.
    shown = core.filter_result_by_codes(
        full_result,
        requests_list,
        ean_lookup,
        sku_lookup,
        supplier_price_columns,
    ).copy()

    shown = shown.rename(
        columns={
            "Target Price": SPECIAL_PRICE,
            "Below Target": BELOW_SPECIAL,
        }
    )

    legrand_missing_our_price = 0
    final_special_prices = []
    below_values = []

    for row_idx, (_, row) in enumerate(shown.iterrows()):
        request_data = requests_list[row_idx] if row_idx < len(requests_list) else {}
        raw_special = request_data.get("target_price", np.nan)
        matched = str(row.get("Lookup Status", "")) != "CODE NOT FOUND"
        our_price = row.get("Our Price", np.nan)

        special_price = raw_special
        if legrand and matched and pd.notna(raw_special):
            if pd.notna(our_price):
                special_price = _recalculate_legrand_price(raw_special, our_price)
            else:
                # The formula cannot be evaluated without Our Price. Keep the
                # entered value visible and tell the user how many rows skipped
                # the Legrand adjustment.
                legrand_missing_our_price += 1

        final_special_prices.append(special_price)

        below = []
        if matched and pd.notna(special_price):
            if pd.notna(our_price) and float(our_price) < float(special_price):
                below.append("Our Price")
            for price_col in supplier_price_columns:
                value = row.get(price_col, np.nan)
                if pd.notna(value) and float(value) < float(special_price):
                    below.append(
                        price_col[:-len(" Price")]
                        if price_col.endswith(" Price")
                        else price_col
                    )
        below_values.append(", ".join(below))

    shown[SPECIAL_PRICE] = final_special_prices
    shown[BELOW_SPECIAL] = below_values

    first = ["Requested Code", SPECIAL_PRICE, BELOW_SPECIAL, "Lookup Status"]
    remaining = [col for col in shown.columns if col not in first]
    shown = shown[first + remaining]
    return shown, legrand_missing_our_price


def _special_table_model(df, supplier_price_columns):
    """Reuse Target Price browser styling but retain Special Price headings."""
    temporary = df.rename(
        columns={
            SPECIAL_PRICE: "Target Price",
            BELOW_SPECIAL: "Below Target",
        }
    )
    model = core.dataframe_to_table_model(
        temporary,
        supplier_price_columns,
        target_mode=True,
    )
    model["columns"] = [
        SPECIAL_PRICE if col == "Target Price" else BELOW_SPECIAL if col == "Below Target" else col
        for col in model["columns"]
    ]
    return model


def _render_special_results(
    request,
    ws,
    shown,
    requests_list,
    legrand=False,
    legrand_missing_our_price=0,
):
    supplier_cols = ws["supplier_price_columns"]
    table_truncated = len(shown) > _main.MAX_TABLE_ROWS
    table_source = shown.head(_main.MAX_TABLE_ROWS) if table_truncated else shown
    table = _special_table_model(table_source, supplier_cols)

    invalid_prices = sum(
        1
        for req in requests_list
        if req.get("target_raw") and not req.get("target_valid")
    )
    quick_metrics = {
        "requested": len(shown),
        "not_found": int((shown["Lookup Status"] == "CODE NOT FOUND").sum()),
        "with_target": int(shown[SPECIAL_PRICE].notna().sum()),
        "target_met": int(
            shown[BELOW_SPECIAL].fillna("").astype(str).str.strip().ne("").sum()
        ),
    }

    duplicate_text = ", ".join(
        f"{name}: {count}"
        for name, count in ws.get("duplicate_info", {}).items()
        if count > 0
    )

    return _main.templates.TemplateResponse(
        request=request,
        name="results.html",
        context={
            "workspace_id": ws["id"],
            "metrics": _main._result_metrics(
                ws["result"],
                ws["current_config"]["type"],
            ),
            "table": table,
            "view": "all",
            "quick": True,
            "quick_mode": "special",
            "legrand": bool(legrand),
            "legrand_missing_our_price": int(legrand_missing_our_price),
            "quick_codes_text": _runtime._requests_to_text(requests_list),
            "quick_metrics": quick_metrics,
            "invalid_targets": invalid_prices,
            "duplicate_text": duplicate_text,
            "shown_count": len(shown),
            "table_truncated": table_truncated,
            "max_table_rows": _main.MAX_TABLE_ROWS,
        },
    )


def _special_displayed_excel_bytes(df: pd.DataFrame, supplier_price_columns) -> bytes:
    export = df.drop(columns=["_relevant_for_display"], errors="ignore").copy()
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export.to_excel(writer, sheet_name="Displayed Result", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Displayed Result"]

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
        money_fmt = workbook.add_format({"num_format": '€#,##0.00;-€#,##0.00'})
        pct_fmt = workbook.add_format({"num_format": "0.00%"})
        blue_fmt = workbook.add_format(
            {"bg_color": "#E0F2FE", "font_color": "#075985", "bold": True}
        )

        for col_idx, name in enumerate(export.columns):
            worksheet.write(0, col_idx, name, header_fmt)

        positions = {name: export.columns.get_loc(name) for name in export.columns}
        for col_idx, name in enumerate(export.columns):
            if name == "EAN":
                width = 17
            elif name in {"SKU", "Cheapest Supplier", BELOW_SPECIAL}:
                width = 26
            elif name in {"Requested Code", "Lookup Status"}:
                width = 20
            else:
                width = 18
            worksheet.set_column(col_idx, col_idx, width)

        for name in [
            "Our Price",
            SPECIAL_PRICE,
            *supplier_price_columns,
            "Cheapest Price",
            "Saving €",
        ]:
            if name in positions:
                idx = positions[name]
                worksheet.set_column(idx, idx, 18, money_fmt)

        if "Saving %" in positions:
            worksheet.set_column(positions["Saving %"], positions["Saving %"], 13, pct_fmt)

        worksheet.freeze_panes(1, min(3, len(export.columns)))
        if len(export) and len(export.columns):
            worksheet.autofilter(0, 0, len(export), len(export.columns) - 1)

            if SPECIAL_PRICE in positions:
                special_idx = positions[SPECIAL_PRICE]
                special_letter = _runtime._xlsx_col_letter(special_idx)
                worksheet.conditional_format(
                    1,
                    special_idx,
                    len(export),
                    special_idx,
                    {"type": "no_blanks", "format": blue_fmt},
                )

                for name in ["Our Price", *supplier_price_columns, "Cheapest Price"]:
                    if name not in positions:
                        continue
                    idx = positions[name]
                    letter = _runtime._xlsx_col_letter(idx)
                    worksheet.conditional_format(
                        1,
                        idx,
                        len(export),
                        idx,
                        {
                            "type": "formula",
                            "criteria": (
                                f'=AND(${special_letter}2<>"",{letter}2<>"",'
                                f'{letter}2<${special_letter}2)'
                            ),
                            "format": blue_fmt,
                        },
                    )

                if BELOW_SPECIAL in positions:
                    idx = positions[BELOW_SPECIAL]
                    worksheet.conditional_format(
                        1,
                        idx,
                        len(export),
                        idx,
                        {"type": "no_blanks", "format": blue_fmt},
                    )

    buffer.seek(0)
    return buffer.getvalue()


def install(app):
    @app.post("/results/{workspace_id}/quick-filter")
    async def quick_filter(request: Request, workspace_id: str):
        try:
            ws = _main._touch(workspace_id)
        except HTTPException as exc:
            return _main._render_error(request, "Workspace expired", exc.detail, 404)
        if "result" not in ws:
            return _main.RedirectResponse(
                url=f"/configure/{workspace_id}",
                status_code=303,
            )

        form = await request.form()
        mode = str(form.get("quick_mode", "default") or "default")
        codes_text = str(form.get("codes", "") or "")
        requests_list = core.parse_requested_codes(codes_text)

        if not requests_list:
            return _main._render_results_page(
                request,
                ws,
                _main._select_view(ws["result"], "all"),
                view="all",
                quick=False,
            )

        if mode != "special":
            shown = core.filter_result_by_codes(
                ws["result"],
                requests_list,
                ws["ean_lookup"],
                ws["sku_lookup"],
                ws["supplier_price_columns"],
            )
            return _main._render_results_page(
                request,
                ws,
                shown,
                view="all",
                quick=True,
                requests_list=requests_list,
            )

        legrand = _is_checked(form.get("legrand"))
        shown, skipped = build_special_price_result(
            ws["result"],
            requests_list,
            ws["ean_lookup"],
            ws["sku_lookup"],
            ws["supplier_price_columns"],
            legrand=legrand,
        )
        return _render_special_results(
            request,
            ws,
            shown,
            requests_list,
            legrand=legrand,
            legrand_missing_our_price=skipped,
        )

    @app.post("/download/{workspace_id}/shown-smart")
    async def download_displayed_smart(request: Request, workspace_id: str):
        ws = _main._touch(workspace_id)
        if "result" not in ws:
            raise HTTPException(status_code=404, detail="Run a comparison before downloading.")

        form = await request.form()
        mode = str(form.get("quick_mode", "default") or "default")
        view = str(form.get("view", "all") or "all")
        codes_text = str(form.get("codes", "") or "")
        sort_col = str(form.get("sort_col", "") or "")
        sort_dir = str(form.get("sort_dir", "asc") or "asc")

        if codes_text.strip():
            requests_list = core.parse_requested_codes(codes_text)
            if mode == "special":
                shown, _ = build_special_price_result(
                    ws["result"],
                    requests_list,
                    ws["ean_lookup"],
                    ws["sku_lookup"],
                    ws["supplier_price_columns"],
                    legrand=_is_checked(form.get("legrand")),
                )
                content = None
                label = "Special_Prices"
            else:
                shown = core.filter_result_by_codes(
                    ws["result"],
                    requests_list,
                    ws["ean_lookup"],
                    ws["sku_lookup"],
                    ws["supplier_price_columns"],
                )
                content = None
                label = "Filtered"
        else:
            shown = _main._select_view(ws["result"], view)
            content = None
            label = {
                "all": "Displayed",
                "cheaper": "Supplier_Cheaper",
                "matched": "Matched",
                "not_found": "Not_Found",
            }.get(view, "Displayed")

        shown = _runtime._sort_displayed_result(shown, sort_col, sort_dir)
        if mode == "special" and codes_text.strip():
            content = _special_displayed_excel_bytes(
                shown,
                ws["supplier_price_columns"],
            )
        else:
            content = _runtime._displayed_excel_bytes(
                shown,
                ws["supplier_price_columns"],
            )

        filename = f"Price_Comparison_{label}_{time.strftime('%Y-%m-%d_%H-%M')}.xlsx"
        return Response(
            content=content,
            media_type=XLSX_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
