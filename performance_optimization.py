"""Performance layer for large multi-supplier comparisons.

Goals:
- never keep a long-running Excel comparison behind one HTTP request;
- read each selected worksheet only once for the actual comparison;
- read only mapped columns from supplier files;
- prefer the fast Rust-backed calamine reader, with pandas defaults as fallback;
- generate the full comparison XLSX lazily only when the user downloads it.
"""

from __future__ import annotations

import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from fastapi import HTTPException, Request
from fastapi.responses import JSONResponse, RedirectResponse, Response

import core
import main as _main


_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="price-compare")
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _metadata_columns(file_meta: dict, sheet_name: str) -> list[str]:
    for sheet in file_meta.get("metadata", {}).get("sheets", []):
        if sheet.get("name") == sheet_name:
            return [str(c) for c in sheet.get("columns", [])]
    return []


def _read_excel(path: str, sheet_name: str, usecols=None, nrows=None) -> pd.DataFrame:
    kwargs = {
        "sheet_name": sheet_name,
        "dtype": object,
    }
    if usecols is not None:
        kwargs["usecols"] = usecols
    if nrows is not None:
        kwargs["nrows"] = nrows

    # Calamine is considerably faster on large xlsx/xls workbooks and supports
    # both formats. Keep a fallback so an unusual workbook does not fail only
    # because the faster reader cannot parse it.
    try:
        return pd.read_excel(path, engine="calamine", **kwargs)
    except Exception:
        return pd.read_excel(path, **kwargs)


def _fast_read_sheet(file_meta: dict, sheet_name: str) -> pd.DataFrame:
    try:
        df = _read_excel(file_meta["path"], sheet_name)
        return _main._normalise_excel_columns(df)
    except Exception as exc:
        raise ValueError(
            f"Could not read {file_meta['name']} / {sheet_name}: {exc}"
        ) from exc


def _read_selected_columns(
    file_meta: dict,
    sheet_name: str,
    requested_columns: list[str],
) -> pd.DataFrame:
    """Read only the columns required by the selected mapping when possible."""
    requested_columns = list(dict.fromkeys(c for c in requested_columns if c))
    metadata = _metadata_columns(file_meta, sheet_name)

    if requested_columns and metadata and all(c in metadata for c in requested_columns):
        ordered = sorted((metadata.index(c), c) for c in requested_columns)
        indices = [idx for idx, _ in ordered]
        names = [name for _, name in ordered]
        try:
            df = _read_excel(file_meta["path"], sheet_name, usecols=indices)
            if len(df.columns) == len(names):
                # Blank Excel headers are renamed consistently according to the
                # full-sheet metadata rather than according to the subset read.
                df.columns = names
                return df
        except Exception:
            pass

    # Safe fallback for unusual workbooks / mappings.
    full = _fast_read_sheet(file_meta, sheet_name)
    missing = [c for c in requested_columns if c not in full.columns]
    if missing:
        raise ValueError(
            f"{file_meta['name']} / {sheet_name}: selected column(s) not found: "
            + ", ".join(missing)
        )
    return full[requested_columns].copy() if requested_columns else full


def _job_update(ws: dict, *, status=None, stage=None, progress=None, error=None):
    job = ws.setdefault("compare_job", {})
    if status is not None:
        job["status"] = status
    if stage is not None:
        job["stage"] = stage
    if progress is not None:
        job["progress"] = int(max(0, min(100, progress)))
    if error is not None:
        job["error"] = str(error)
    job["updated"] = time.time()


def _validate_current_df(df: pd.DataFrame, config: dict):
    if config["type"] == "PowerBI Pricelist":
        missing = [c for c in core.POWERBI_REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                "PowerBI Pricelist is missing required columns: " + ", ".join(missing)
            )
        return

    ean = config.get("ean_column")
    sku = config.get("sku_column")
    price = config.get("price_column")
    if not ean and not sku:
        raise ValueError("Free format current pricelist requires EAN and/or SKU.")
    for name, label in [(ean, "EAN"), (sku, "SKU"), (price, "Price")]:
        if name and name not in df.columns:
            raise ValueError(f"Current pricelist: selected {label} column '{name}' was not found.")


def _validate_supplier_df(df: pd.DataFrame, config: dict):
    method = config["match_method"]
    required = [config["price_column"]]
    if method in {"EAN", "EAN + SKU"}:
        required.append(config["ean_column"])
    if method in {"SKU", "EAN + SKU"}:
        required.append(config["sku_column"])
    missing = [c for c in required if not c or c not in df.columns]
    if missing:
        raise ValueError(
            f"{config['supplier_name']}: selected mapping column(s) were not found: "
            + ", ".join(str(c) for c in missing)
        )


def _run_comparison_job(workspace_id: str, spec: dict):
    ws = _main.WORKSPACES.get(workspace_id)
    if not ws:
        return

    try:
        started = time.perf_counter()
        _job_update(ws, status="running", stage="Reading current pricelist…", progress=4)

        current_config = spec["current_config"]
        current_sheet = spec["current_sheet"]
        if current_config["type"] == "PowerBI Pricelist":
            current_needed = list(core.POWERBI_REQUIRED_COLUMNS)
        else:
            current_needed = [
                current_config.get("ean_column"),
                current_config.get("sku_column"),
                current_config.get("price_column"),
                *list(current_config.get("extra_columns") or []),
            ]
        current_df = _read_selected_columns(ws["current"], current_sheet, current_needed)
        _validate_current_df(current_df, current_config)

        supplier_configs = spec["supplier_configs"]
        supplier_items = []
        total = max(1, len(supplier_configs))

        for idx, (supplier_meta, config) in enumerate(
            zip(ws["suppliers"], supplier_configs),
            start=1,
        ):
            _job_update(
                ws,
                stage=f"Reading supplier {idx} of {total}: {config['supplier_name']}…",
                progress=8 + int(58 * (idx - 1) / total),
            )
            needed = [config.get("price_column")]
            if config["match_method"] in {"EAN", "EAN + SKU"}:
                needed.append(config.get("ean_column"))
            if config["match_method"] in {"SKU", "EAN + SKU"}:
                needed.append(config.get("sku_column"))
            supplier_df = _read_selected_columns(
                supplier_meta,
                config["sheet_name"],
                needed,
            )
            _validate_supplier_df(supplier_df, config)
            supplier_items.append((supplier_df, config))

        _job_update(ws, stage="Comparing prices…", progress=72)
        result, supplier_price_columns, duplicate_info, _ = core.compare_all(
            current_df,
            current_config,
            supplier_items,
        )

        _job_update(ws, stage="Preparing result table…", progress=91)
        ean_lookup, sku_lookup = core.build_code_lookups(result)

        # Do not build the full XLSX here. It is generated on demand when the
        # user clicks Download full comparison, which shortens Compare greatly.
        ws.update(
            {
                "current_sheet": current_sheet,
                "current_config": current_config,
                "supplier_configs": supplier_configs,
                "result": result,
                "supplier_price_columns": supplier_price_columns,
                "duplicate_info": duplicate_info,
                "ean_lookup": ean_lookup,
                "sku_lookup": sku_lookup,
            }
        )
        ws.pop("excel_bytes", None)
        ws.pop("full_excel_bytes", None)

        elapsed = time.perf_counter() - started
        job = ws.setdefault("compare_job", {})
        job["elapsed_seconds"] = round(elapsed, 1)
        _job_update(ws, status="done", stage="Comparison complete", progress=100)
    except Exception as exc:
        _job_update(
            ws,
            status="error",
            stage="Comparison failed",
            progress=100,
            error=str(exc),
        )


def _parse_form(ws: dict, form) -> tuple[dict | None, list[str]]:
    errors = []
    current_sheet = str(form.get("current_sheet", "")).strip()
    if not current_sheet:
        errors.append("Select a current-pricelist worksheet.")

    if ws["current_type"] == "PowerBI Pricelist":
        current_config = {
            "type": "PowerBI Pricelist",
            "ean_column": "EAN",
            "sku_column": "SKU",
            "price_column": "Latest Pricelist Value",
            "extra_columns": [],
        }
    else:
        ean = str(form.get("current_ean", "")).strip() or None
        sku = str(form.get("current_sku", "")).strip() or None
        price = str(form.get("current_price", "")).strip()
        extras = [str(x) for x in form.getlist("current_extras")]
        if not ean and not sku:
            errors.append("Free format current pricelist: select EAN and/or SKU.")
        if not price:
            errors.append("Free format current pricelist: select a Price column.")
        current_config = {
            "type": "Free format pricelist",
            "ean_column": ean,
            "sku_column": sku,
            "price_column": price,
            "extra_columns": extras,
        }

    supplier_configs = []
    names_seen = set()
    for idx, supplier in enumerate(ws["suppliers"]):
        name = str(form.get(f"supplier_{idx}_name", "")).strip()
        sheet = str(form.get(f"supplier_{idx}_sheet", "")).strip()
        method = str(form.get(f"supplier_{idx}_method", "SKU")).strip()
        ean = str(form.get(f"supplier_{idx}_ean", "")).strip() or None
        sku = str(form.get(f"supplier_{idx}_sku", "")).strip() or None
        price = str(form.get(f"supplier_{idx}_price", "")).strip()

        label = name or supplier["name"]
        if not name:
            errors.append(f"Supplier {idx + 1}: enter a supplier name.")
        elif name.casefold() in names_seen:
            errors.append(f"Supplier names must be unique: {name} is duplicated.")
        else:
            names_seen.add(name.casefold())
        if not sheet:
            errors.append(f"{label}: select a worksheet.")
        if method not in {"SKU", "EAN", "EAN + SKU"}:
            errors.append(f"{label}: invalid match method.")
        if not price:
            errors.append(f"{label}: select a Price column.")
        if method in {"EAN", "EAN + SKU"} and not ean:
            errors.append(f"{label}: select an EAN column.")
        if method in {"SKU", "EAN + SKU"} and not sku:
            errors.append(f"{label}: select a SKU column.")
        if method == "EAN + SKU" and ean == sku:
            errors.append(f"{label}: EAN and SKU columns cannot be the same.")
        if price and price in {ean, sku}:
            errors.append(f"{label}: Price column cannot also be an identifier.")

        supplier_configs.append(
            {
                "supplier_name": name,
                "file_name": supplier["name"],
                "sheet_name": sheet,
                "match_method": method,
                "ean_column": ean if method in {"EAN", "EAN + SKU"} else None,
                "sku_column": sku if method in {"SKU", "EAN + SKU"} else None,
                "price_column": price,
            }
        )

    if errors:
        return None, errors
    return {
        "current_sheet": current_sheet,
        "current_config": current_config,
        "supplier_configs": supplier_configs,
    }, []


def install(app):
    # Faster reader for preview and any legacy code path.
    _main._read_sheet = _fast_read_sheet
    # Avoid the previous runtime patch that read the complete sheet only to
    # validate column names. Robust upload metadata already contains the width.
    _main._sheet_columns = _metadata_columns

    @app.post("/compare-fast/{workspace_id}")
    async def compare_fast(request: Request, workspace_id: str):
        try:
            ws = _main._touch(workspace_id)
        except HTTPException as exc:
            return _main._render_error(request, "Workspace expired", exc.detail, 404)

        form = await request.form()
        spec, errors = _parse_form(ws, form)
        if errors:
            return _main.templates.TemplateResponse(
                request=request,
                name="configure.html",
                context=_main._config_context(ws, errors=errors),
                status_code=400,
            )

        # Remove stale results before a new run starts.
        for key in [
            "result",
            "supplier_price_columns",
            "duplicate_info",
            "ean_lookup",
            "sku_lookup",
            "excel_bytes",
            "full_excel_bytes",
        ]:
            ws.pop(key, None)

        ws["compare_job"] = {
            "status": "queued",
            "stage": "Starting comparison…",
            "progress": 1,
            "started": time.time(),
            "error": "",
        }
        _executor.submit(_run_comparison_job, workspace_id, spec)
        return RedirectResponse(url=f"/compare-progress/{workspace_id}", status_code=303)

    @app.get("/compare-progress/{workspace_id}")
    def compare_progress(request: Request, workspace_id: str):
        try:
            ws = _main._touch(workspace_id)
        except HTTPException as exc:
            return _main._render_error(request, "Workspace expired", exc.detail, 404)
        return _main.templates.TemplateResponse(
            request=request,
            name="compare_progress.html",
            context={"workspace_id": workspace_id},
        )

    @app.get("/api/compare-progress/{workspace_id}")
    def compare_progress_api(workspace_id: str):
        ws = _main._touch(workspace_id)
        job = ws.get("compare_job") or {
            "status": "error",
            "stage": "No comparison job is running.",
            "progress": 100,
            "error": "No comparison job is running.",
        }
        return JSONResponse(job)

    @app.get("/download-full/{workspace_id}")
    def download_full(workspace_id: str):
        ws = _main._touch(workspace_id)
        if "result" not in ws:
            raise HTTPException(status_code=404, detail="Run a comparison before downloading.")

        content = ws.get("full_excel_bytes")
        if content is None:
            content = _main.create_excel(
                ws["result"],
                ws["supplier_price_columns"],
                ws["supplier_configs"],
                ws["duplicate_info"],
                ws["current"]["name"],
                ws["current_config"],
            )
            ws["full_excel_bytes"] = content

        filename = f"Price_Comparison_{time.strftime('%Y-%m-%d_%H-%M')}.xlsx"
        return Response(
            content=content,
            media_type=XLSX_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
