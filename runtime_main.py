"""Production entrypoint and runtime hardening for Price Comparison."""

import io
import time
from pathlib import Path

import pandas as pd
from fastapi import HTTPException, Request
from fastapi.responses import Response

import main as _main
from core import dataframe_to_table_model, filter_result_by_codes, parse_requested_codes


METADATA_SCAN_ROWS = 250
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _robust_workbook_metadata(path: str) -> dict:
    try:
        xls = pd.ExcelFile(path)
    except Exception as exc:
        raise ValueError(f"Could not open workbook: {exc}") from exc

    sheets = []
    for sheet in xls.sheet_names:
        try:
            # nrows=0 can lose columns whose first-row header is blank.
            # A small data scan reveals the actual used worksheet width while
            # remaining much cheaper than loading an entire large pricelist.
            sample = pd.read_excel(
                path,
                sheet_name=sheet,
                nrows=METADATA_SCAN_ROWS,
                dtype=object,
            )
            sample = _main._normalise_excel_columns(sample)
            columns = [str(c) for c in sample.columns]
        except Exception:
            columns = []
        sheets.append({"name": str(sheet), "columns": columns})
    return {"sheets": sheets}


def _authoritative_sheet_columns(file_meta, sheet_name):
    """Return the same normalized columns that comparison will actually use."""
    try:
        df = _main._read_sheet(file_meta, sheet_name)
        return [str(c) for c in df.columns]
    except Exception:
        for sheet in file_meta.get("metadata", {}).get("sheets", []):
            if sheet.get("name") == sheet_name:
                return list(sheet.get("columns") or [])
        return []


def _requests_to_text(requests_list) -> str:
    lines = []
    for request in requests_list or []:
        code = str(request.get("code", "")).strip()
        target_raw = str(request.get("target_raw", "") or "").strip()
        if not code:
            continue
        lines.append(f"{code}\t{target_raw}" if target_raw else code)
    return "\n".join(lines)


def _render_results_page(
    request,
    ws,
    shown,
    view,
    quick=False,
    requests_list=None,
):
    """Render result page with enough state to export exactly the displayed set."""
    result = ws["result"]
    supplier_cols = ws["supplier_price_columns"]
    table_truncated = len(shown) > _main.MAX_TABLE_ROWS
    table_source = shown.head(_main.MAX_TABLE_ROWS) if table_truncated else shown
    table = dataframe_to_table_model(table_source, supplier_cols, target_mode=quick)

    quick_metrics = None
    invalid_targets = 0
    if quick:
        quick_metrics = {
            "requested": len(shown),
            "not_found": int((shown["Lookup Status"] == "CODE NOT FOUND").sum()),
            "with_target": int(shown["Target Price"].notna().sum()),
            "target_met": int(
                shown["Below Target"].fillna("").astype(str).str.strip().ne("").sum()
            ),
        }
        invalid_targets = sum(
            1
            for req in (requests_list or [])
            if req.get("target_raw") and not req.get("target_valid")
        )

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
                result,
                ws["current_config"]["type"],
            ),
            "table": table,
            "view": view,
            "quick": quick,
            "quick_codes_text": _requests_to_text(requests_list) if quick else "",
            "quick_metrics": quick_metrics,
            "invalid_targets": invalid_targets,
            "duplicate_text": duplicate_text,
            "shown_count": len(shown),
            "table_truncated": table_truncated,
            "max_table_rows": _main.MAX_TABLE_ROWS,
        },
    )


def _sort_displayed_result(df: pd.DataFrame, sort_col: str, sort_dir: str) -> pd.DataFrame:
    if not sort_col or sort_col not in df.columns:
        return df

    ascending = str(sort_dir).lower() != "desc"
    result = df.copy()
    series = result[sort_col]

    # Keep numeric price/quantity sorting numeric even when a dataframe column
    # arrived as object because of blanks or mixed Excel cells.
    numeric = pd.to_numeric(series, errors="coerce")
    numeric_share = float(numeric.notna().mean()) if len(numeric) else 0.0
    if numeric_share >= 0.8:
        result["__display_sort"] = numeric
    else:
        result["__display_sort"] = series.fillna("").astype(str).str.casefold()

    result = (
        result.sort_values(
            "__display_sort",
            ascending=ascending,
            na_position="last",
            kind="stable",
        )
        .drop(columns="__display_sort")
        .reset_index(drop=True)
    )
    return result


def _displayed_excel_bytes(df: pd.DataFrame, supplier_price_columns) -> bytes:
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
        target_fmt = workbook.add_format(
            {"bg_color": "#E0F2FE", "font_color": "#075985", "bold": True}
        )
        green_fmt = workbook.add_format(
            {"bg_color": "#DCFCE7", "font_color": "#166534"}
        )
        yellow_fmt = workbook.add_format(
            {"bg_color": "#FEF3C7", "font_color": "#92400E", "bold": True}
        )

        for col_idx, name in enumerate(export.columns):
            worksheet.write(0, col_idx, name, header_fmt)

        positions = {name: export.columns.get_loc(name) for name in export.columns}
        for col_idx, name in enumerate(export.columns):
            if name == "EAN":
                width = 17
            elif name in {"SKU", "Cheapest Supplier", "Below Target"}:
                width = 24
            elif name in {"Requested Code", "Lookup Status"}:
                width = 20
            else:
                width = 18
            worksheet.set_column(col_idx, col_idx, width)

        money_columns = [
            "Our Price",
            "Target Price",
            *supplier_price_columns,
            "Cheapest Price",
            "Saving €",
        ]
        for name in money_columns:
            if name in positions:
                idx = positions[name]
                worksheet.set_column(idx, idx, 18, money_fmt)

        if "Saving %" in positions:
            idx = positions["Saving %"]
            worksheet.set_column(idx, idx, 13, pct_fmt)

        worksheet.freeze_panes(1, min(3, len(export.columns)))
        if len(export) and len(export.columns):
            worksheet.autofilter(0, 0, len(export), len(export.columns) - 1)

            if "Target Price" in positions:
                target_idx = positions["Target Price"]
                worksheet.set_column(target_idx, target_idx, 18, money_fmt)
                worksheet.conditional_format(
                    1,
                    target_idx,
                    len(export),
                    target_idx,
                    {"type": "no_blanks", "format": target_fmt},
                )

            # If the displayed result is in target mode, mirror the browser's
            # blue target highlighting. Otherwise retain the normal comparison
            # green/yellow cues.
            has_target_mode = "Target Price" in positions and export["Target Price"].notna().any()
            if has_target_mode:
                target_letter = _xlsx_col_letter(positions["Target Price"])
                for name in ["Our Price", *supplier_price_columns, "Cheapest Price"]:
                    if name not in positions:
                        continue
                    idx = positions[name]
                    letter = _xlsx_col_letter(idx)
                    worksheet.conditional_format(
                        1,
                        idx,
                        len(export),
                        idx,
                        {
                            "type": "formula",
                            "criteria": (
                                f'=AND(${target_letter}2<>"",{letter}2<>"",'
                                f'{letter}2<${target_letter}2)'
                            ),
                            "format": target_fmt,
                        },
                    )
                if "Below Target" in positions:
                    idx = positions["Below Target"]
                    worksheet.conditional_format(
                        1,
                        idx,
                        len(export),
                        idx,
                        {"type": "no_blanks", "format": target_fmt},
                    )
            else:
                if "Our Price" in positions:
                    our_letter = _xlsx_col_letter(positions["Our Price"])
                    for name in supplier_price_columns:
                        if name not in positions:
                            continue
                        idx = positions[name]
                        letter = _xlsx_col_letter(idx)
                        worksheet.conditional_format(
                            1,
                            idx,
                            len(export),
                            idx,
                            {
                                "type": "formula",
                                "criteria": (
                                    f'=AND({letter}2<>"",${our_letter}2<>"",'
                                    f'{letter}2<${our_letter}2)'
                                ),
                                "format": green_fmt,
                            },
                        )
                for name in ["Cheapest Price", "Cheapest Supplier"]:
                    if name in positions:
                        idx = positions[name]
                        worksheet.conditional_format(
                            1,
                            idx,
                            len(export),
                            idx,
                            {"type": "no_blanks", "format": yellow_fmt},
                        )

    buffer.seek(0)
    return buffer.getvalue()


def _xlsx_col_letter(index: int) -> str:
    """Zero-based Excel column index to A/AA/... without another dependency."""
    index += 1
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _xlsx_response(content: bytes, filename: str) -> Response:
    return Response(
        content=content,
        media_type=XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


_main._workbook_metadata = _robust_workbook_metadata
_main._sheet_columns = _authoritative_sheet_columns
_main._render_results_page = _render_results_page

app = _main.app


@app.post("/download/{workspace_id}/shown")
async def download_displayed_table(request: Request, workspace_id: str):
    ws = _main._touch(workspace_id)
    if "result" not in ws:
        raise HTTPException(status_code=404, detail="Run a comparison before downloading.")

    form = await request.form()
    view = str(form.get("view", "all") or "all")
    codes_text = str(form.get("codes", "") or "")
    sort_col = str(form.get("sort_col", "") or "")
    sort_dir = str(form.get("sort_dir", "asc") or "asc")

    if codes_text.strip():
        requests_list = parse_requested_codes(codes_text)
        shown = filter_result_by_codes(
            ws["result"],
            requests_list,
            ws["ean_lookup"],
            ws["sku_lookup"],
            ws["supplier_price_columns"],
        )
        label = "Filtered"
    else:
        shown = _main._select_view(ws["result"], view)
        labels = {
            "all": "Displayed",
            "cheaper": "Supplier_Cheaper",
            "matched": "Matched",
            "not_found": "Not_Found",
        }
        label = labels.get(view, "Displayed")

    shown = _sort_displayed_result(shown, sort_col, sort_dir)
    content = _displayed_excel_bytes(shown, ws["supplier_price_columns"])
    filename = f"Price_Comparison_{label}_{time.strftime('%Y-%m-%d_%H-%M')}.xlsx"
    return _xlsx_response(content, filename)
