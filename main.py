import json
import shutil
import tempfile
import time
import uuid
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from core import (
    POWERBI_REQUIRED_COLUMNS,
    build_code_lookups,
    compare_all,
    create_excel,
    dataframe_to_table_model,
    filter_result_by_codes,
    parse_requested_codes,
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
WORKSPACE_ROOT = Path(tempfile.gettempdir()) / "pawbol-price-comparison"
WORKSPACE_ROOT.mkdir(parents=True, exist_ok=True)
WORKSPACE_TTL_SECONDS = 6 * 60 * 60
MAX_TABLE_ROWS = 500
PREVIEW_ROWS = 10
PREVIEW_COLUMNS = 30

app = FastAPI(title="Price List Comparison")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
WORKSPACES: dict[str, dict] = {}


def _cleanup_old_workspaces():
    now = time.time()
    stale = [
        wid
        for wid, ws in WORKSPACES.items()
        if now - ws.get("last_access", now) > WORKSPACE_TTL_SECONDS
    ]
    for wid in stale:
        ws = WORKSPACES.pop(wid, None)
        if ws:
            shutil.rmtree(ws.get("dir", ""), ignore_errors=True)


def _touch(workspace_id: str) -> dict:
    _cleanup_old_workspaces()
    ws = WORKSPACES.get(workspace_id)
    if not ws:
        raise HTTPException(
            status_code=404,
            detail="Workspace expired or does not exist. Start a new comparison.",
        )
    ws["last_access"] = time.time()
    return ws


def _safe_filename(name: str) -> str:
    name = Path(name or "upload.xlsx").name
    cleaned = "".join(ch for ch in name if ch.isalnum() or ch in " ._-()[]")
    return cleaned[:180] or "upload.xlsx"


async def _save_upload(upload: UploadFile, directory: Path, prefix: str) -> dict:
    filename = _safe_filename(upload.filename or "upload.xlsx")
    path = directory / f"{prefix}_{filename}"
    content = await upload.read()
    path.write_bytes(content)
    return {"name": filename, "path": str(path), "size": len(content)}


def _is_blank_excel_header(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return not text or text.lower().startswith("unnamed:")


def _normalise_excel_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Give empty Excel headers deterministic names used everywhere in the app."""
    if df is None:
        return df

    renamed = []
    used = set()
    blank_number = 1

    for raw_name in df.columns:
        if _is_blank_excel_header(raw_name):
            while f"Column{blank_number}" in used:
                blank_number += 1
            name = f"Column{blank_number}"
            blank_number += 1
        else:
            name = str(raw_name).strip()

        if name in used:
            base = name
            suffix = 2
            while f"{base}_{suffix}" in used:
                suffix += 1
            name = f"{base}_{suffix}"

        used.add(name)
        renamed.append(name)

    result = df.copy()
    result.columns = renamed
    return result


def _workbook_metadata(path: str) -> dict:
    try:
        xls = pd.ExcelFile(path)
    except Exception as exc:
        raise ValueError(f"Could not open workbook: {exc}") from exc

    sheets = []
    for sheet in xls.sheet_names:
        try:
            header = pd.read_excel(path, sheet_name=sheet, nrows=0, dtype=object)
            header = _normalise_excel_columns(header)
            columns = [str(c) for c in header.columns]
        except Exception:
            columns = []
        sheets.append({"name": str(sheet), "columns": columns})
    return {"sheets": sheets}


def _sheet_columns(file_meta, sheet_name):
    for sheet in file_meta["metadata"]["sheets"]:
        if sheet["name"] == sheet_name:
            return sheet["columns"]
    return []


def _read_sheet(file_meta, sheet_name):
    try:
        df = pd.read_excel(file_meta["path"], sheet_name=sheet_name, dtype=object)
        return _normalise_excel_columns(df)
    except Exception as exc:
        raise ValueError(
            f"Could not read {file_meta['name']} / {sheet_name}: {exc}"
        ) from exc


def _preview_payload(df: pd.DataFrame) -> dict:
    preview_df = df.iloc[:PREVIEW_ROWS, :PREVIEW_COLUMNS].copy()
    rows = json.loads(preview_df.to_json(orient="records", date_format="iso"))
    return {
        "columns": [str(c) for c in preview_df.columns],
        "rows": rows,
        "rows_total": int(len(df)),
        "columns_total": int(len(df.columns)),
        "rows_shown": int(len(preview_df)),
        "columns_shown": int(len(preview_df.columns)),
    }


def _render_error(request, title, message, status_code=400):
    return templates.TemplateResponse(
        request=request,
        name="error.html",
        context={"title": title, "message": message},
        status_code=status_code,
    )


def _config_context(ws, errors=None):
    current, suppliers = ws["current"], ws["suppliers"]
    current_default_sheet = None
    current_sheets = current["metadata"]["sheets"]
    if current_sheets:
        names = [s["name"] for s in current_sheets]
        current_default_sheet = (
            "Export"
            if ws["current_type"] == "PowerBI Pricelist" and "Export" in names
            else names[0]
        )

    supplier_defaults = []
    for idx, supplier in enumerate(suppliers):
        first_sheet = (
            supplier["metadata"]["sheets"][0]["name"]
            if supplier["metadata"]["sheets"]
            else ""
        )
        supplier_defaults.append(
            {
                "index": idx,
                "default_name": Path(supplier["name"]).stem,
                "default_sheet": first_sheet,
            }
        )

    metadata_json = {
        "current": current["metadata"],
        "suppliers": [s["metadata"] for s in suppliers],
    }
    return {
        "workspace_id": ws["id"],
        "current_type": ws["current_type"],
        "current": current,
        "suppliers": suppliers,
        "current_default_sheet": current_default_sheet,
        "supplier_defaults": supplier_defaults,
        "metadata_json": json.dumps(metadata_json, ensure_ascii=False),
        "errors": errors or [],
    }


def _result_metrics(result, current_type):
    relevant = result[result["_relevant_for_display"].fillna(False).astype(bool)]
    return {
        "product_label": (
            "Relevant products" if current_type == "PowerBI Pricelist" else "Products"
        ),
        "products": len(relevant),
        "matched": int(relevant["Matched Suppliers"].gt(0).sum()),
        "cheaper": int((relevant["Status"] == "CHEAPER").sum()),
        "more_expensive": int((relevant["Status"] == "MORE EXPENSIVE").sum()),
        "not_found": int((relevant["Status"] == "NOT FOUND").sum()),
    }


def _select_view(result, view):
    relevant = result[result["_relevant_for_display"].fillna(False).astype(bool)].copy()
    if view == "cheaper":
        return relevant[relevant["Status"] == "CHEAPER"].copy()
    if view == "matched":
        return relevant[relevant["Matched Suppliers"].gt(0)].copy()
    if view == "not_found":
        return relevant[relevant["Matched Suppliers"].eq(0)].copy()
    return relevant


def _render_results_page(request, ws, shown, view, quick=False, requests_list=None):
    result, supplier_cols = ws["result"], ws["supplier_price_columns"]
    table_truncated = len(shown) > MAX_TABLE_ROWS
    table_source = shown.head(MAX_TABLE_ROWS) if table_truncated else shown
    table = dataframe_to_table_model(table_source, supplier_cols, target_mode=quick)
    quick_metrics, invalid_targets = None, 0

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
    return templates.TemplateResponse(
        request=request,
        name="results.html",
        context={
            "workspace_id": ws["id"],
            "metrics": _result_metrics(result, ws["current_config"]["type"]),
            "table": table,
            "view": view,
            "quick": quick,
            "quick_metrics": quick_metrics,
            "invalid_targets": invalid_targets,
            "duplicate_text": duplicate_text,
            "shown_count": len(shown),
            "table_truncated": table_truncated,
            "max_table_rows": MAX_TABLE_ROWS,
        },
    )


@app.exception_handler(Exception)
async def unhandled_exception(request: Request, exc: Exception):
    return _render_error(request, "Unexpected application error", str(exc), 500)


@app.get("/health")
def health():
    return {"status": "ok", "app": "price-comparison"}


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    _cleanup_old_workspaces()
    return templates.TemplateResponse(request=request, name="index.html", context={})


@app.post("/setup", response_class=HTMLResponse)
async def setup(
    request: Request,
    current_type: str = Form(...),
    current_file: UploadFile = File(...),
    supplier_files: list[UploadFile] = File(...),
):
    if current_type not in {"PowerBI Pricelist", "Free format pricelist"}:
        return _render_error(
            request,
            "Invalid pricelist type",
            "Choose PowerBI Pricelist or Free format pricelist.",
        )
    if not supplier_files:
        return _render_error(
            request,
            "No suppliers uploaded",
            "Upload at least one alternative supplier pricelist.",
        )

    workspace_id = uuid.uuid4().hex
    workspace_dir = WORKSPACE_ROOT / workspace_id
    workspace_dir.mkdir(parents=True, exist_ok=True)

    try:
        current = await _save_upload(current_file, workspace_dir, "current")
        current["metadata"] = _workbook_metadata(current["path"])
        suppliers = []
        for idx, upload in enumerate(supplier_files):
            item = await _save_upload(upload, workspace_dir, f"supplier_{idx}")
            item["metadata"] = _workbook_metadata(item["path"])
            suppliers.append(item)
    except Exception as exc:
        shutil.rmtree(workspace_dir, ignore_errors=True)
        return _render_error(request, "Could not read uploaded files", str(exc))

    ws = {
        "id": workspace_id,
        "dir": str(workspace_dir),
        "created": time.time(),
        "last_access": time.time(),
        "current_type": current_type,
        "current": current,
        "suppliers": suppliers,
    }
    WORKSPACES[workspace_id] = ws
    return templates.TemplateResponse(
        request=request,
        name="configure.html",
        context=_config_context(ws),
    )


@app.get("/configure/{workspace_id}", response_class=HTMLResponse)
def configure(request: Request, workspace_id: str):
    try:
        ws = _touch(workspace_id)
    except HTTPException as exc:
        return _render_error(request, "Workspace expired", exc.detail, 404)
    return templates.TemplateResponse(
        request=request,
        name="configure.html",
        context=_config_context(ws),
    )


@app.get("/preview/{workspace_id}")
def preview(
    workspace_id: str,
    file_type: str,
    index: int = 0,
    sheet: str = "",
):
    ws = _touch(workspace_id)

    if file_type == "current":
        file_meta = ws["current"]
    elif file_type == "supplier":
        if index < 0 or index >= len(ws["suppliers"]):
            raise HTTPException(status_code=400, detail="Supplier index is invalid.")
        file_meta = ws["suppliers"][index]
    else:
        raise HTTPException(status_code=400, detail="Unknown preview file type.")

    if sheet not in [s["name"] for s in file_meta["metadata"]["sheets"]]:
        raise HTTPException(status_code=400, detail="Selected worksheet was not found.")

    try:
        df = _read_sheet(file_meta, sheet)
        payload = _preview_payload(df)
        payload.update({"filename": file_meta["name"], "sheet": sheet})
        return payload
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/compare/{workspace_id}", response_class=HTMLResponse)
async def compare(request: Request, workspace_id: str):
    try:
        ws = _touch(workspace_id)
    except HTTPException as exc:
        return _render_error(request, "Workspace expired", exc.detail, 404)

    form = await request.form()
    errors = []
    current_sheet = str(form.get("current_sheet", ""))
    current_columns = _sheet_columns(ws["current"], current_sheet)

    if not current_columns:
        errors.append("Select a valid current-pricelist worksheet.")

    if ws["current_type"] == "PowerBI Pricelist":
        missing = [c for c in POWERBI_REQUIRED_COLUMNS if c not in current_columns]
        if missing:
            errors.append(
                "PowerBI Pricelist is missing required columns: " + ", ".join(missing)
            )
        current_config = {
            "type": "PowerBI Pricelist",
            "ean_column": "EAN",
            "sku_column": "SKU",
            "price_column": "Latest Pricelist Value",
            "extra_columns": [],
        }
    else:
        ean_col = str(form.get("current_ean", "")).strip() or None
        sku_col = str(form.get("current_sku", "")).strip() or None
        price_col = str(form.get("current_price", "")).strip()
        extra_cols = [str(x) for x in form.getlist("current_extras")]

        if not ean_col and not sku_col:
            errors.append("Free format current pricelist: select EAN and/or SKU.")
        if ean_col and ean_col not in current_columns:
            errors.append("Selected current EAN column is not in the selected sheet.")
        if sku_col and sku_col not in current_columns:
            errors.append("Selected current SKU column is not in the selected sheet.")
        if not price_col or price_col not in current_columns:
            errors.append("Select a valid current Price column.")
        if price_col and price_col in {ean_col, sku_col}:
            errors.append("Current Price column cannot also be EAN or SKU.")

        current_config = {
            "type": "Free format pricelist",
            "ean_column": ean_col,
            "sku_column": sku_col,
            "price_column": price_col,
            "extra_columns": extra_cols,
        }

    supplier_configs = []
    names_seen = set()

    for idx, supplier in enumerate(ws["suppliers"]):
        name = str(form.get(f"supplier_{idx}_name", "")).strip()
        sheet = str(form.get(f"supplier_{idx}_sheet", "")).strip()
        method = str(form.get(f"supplier_{idx}_method", "SKU")).strip()
        ean_col = str(form.get(f"supplier_{idx}_ean", "")).strip() or None
        sku_col = str(form.get(f"supplier_{idx}_sku", "")).strip() or None
        price_col = str(form.get(f"supplier_{idx}_price", "")).strip()
        columns = _sheet_columns(supplier, sheet)

        if not name:
            errors.append(f"Supplier {idx + 1}: enter a supplier name.")
        elif name.casefold() in names_seen:
            errors.append(f"Supplier names must be unique: {name} is duplicated.")
        else:
            names_seen.add(name.casefold())

        if method not in {"SKU", "EAN", "EAN + SKU"}:
            errors.append(f"{name or supplier['name']}: invalid match method.")
        if not columns:
            errors.append(f"{name or supplier['name']}: select a valid worksheet.")
        if not price_col or price_col not in columns:
            errors.append(f"{name or supplier['name']}: select a valid Price column.")
        if method in {"EAN", "EAN + SKU"} and (
            not ean_col or ean_col not in columns
        ):
            errors.append(f"{name or supplier['name']}: select a valid EAN column.")
        if method in {"SKU", "EAN + SKU"} and (
            not sku_col or sku_col not in columns
        ):
            errors.append(f"{name or supplier['name']}: select a valid SKU column.")
        if method == "EAN + SKU" and ean_col == sku_col:
            errors.append(
                f"{name or supplier['name']}: EAN and SKU columns cannot be the same."
            )

        used_ids = {
            ean_col if method in {"EAN", "EAN + SKU"} else None,
            sku_col if method in {"SKU", "EAN + SKU"} else None,
        }
        if price_col and price_col in used_ids:
            errors.append(
                f"{name or supplier['name']}: Price column cannot also be an identifier."
            )

        supplier_configs.append(
            {
                "supplier_name": name,
                "file_name": supplier["name"],
                "sheet_name": sheet,
                "match_method": method,
                "ean_column": (
                    ean_col if method in {"EAN", "EAN + SKU"} else None
                ),
                "sku_column": (
                    sku_col if method in {"SKU", "EAN + SKU"} else None
                ),
                "price_column": price_col,
            }
        )

    if errors:
        return templates.TemplateResponse(
            request=request,
            name="configure.html",
            context=_config_context(ws, errors=errors),
            status_code=400,
        )

    try:
        current_df = _read_sheet(ws["current"], current_sheet)
        supplier_items = []
        for supplier, config in zip(ws["suppliers"], supplier_configs):
            supplier_items.append(
                (_read_sheet(supplier, config["sheet_name"]), config)
            )

        result, supplier_price_columns, duplicate_info, _ = compare_all(
            current_df,
            current_config,
            supplier_items,
        )
        ean_lookup, sku_lookup = build_code_lookups(result)
        excel_bytes = create_excel(
            result,
            supplier_price_columns,
            supplier_configs,
            duplicate_info,
            ws["current"]["name"],
            current_config,
        )
    except Exception as exc:
        return _render_error(request, "Comparison failed", str(exc), 400)

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
            "excel_bytes": excel_bytes,
        }
    )
    return RedirectResponse(url=f"/results/{workspace_id}", status_code=303)


@app.get("/results/{workspace_id}", response_class=HTMLResponse)
def results(request: Request, workspace_id: str, view: str = "all"):
    try:
        ws = _touch(workspace_id)
    except HTTPException as exc:
        return _render_error(request, "Workspace expired", exc.detail, 404)
    if "result" not in ws:
        return RedirectResponse(url=f"/configure/{workspace_id}", status_code=303)
    return _render_results_page(
        request,
        ws,
        _select_view(ws["result"], view),
        view=view,
        quick=False,
    )


@app.post("/results/{workspace_id}/quick", response_class=HTMLResponse)
async def quick_results(request: Request, workspace_id: str):
    try:
        ws = _touch(workspace_id)
    except HTTPException as exc:
        return _render_error(request, "Workspace expired", exc.detail, 404)
    if "result" not in ws:
        return RedirectResponse(url=f"/configure/{workspace_id}", status_code=303)

    form = await request.form()
    requests_list = parse_requested_codes(str(form.get("codes", "")))
    if not requests_list:
        return _render_results_page(
            request,
            ws,
            _select_view(ws["result"], "all"),
            view="all",
            quick=False,
        )

    shown = filter_result_by_codes(
        ws["result"],
        requests_list,
        ws["ean_lookup"],
        ws["sku_lookup"],
        ws["supplier_price_columns"],
    )
    return _render_results_page(
        request,
        ws,
        shown,
        view="all",
        quick=True,
        requests_list=requests_list,
    )


@app.get("/download/{workspace_id}")
def download(workspace_id: str):
    ws = _touch(workspace_id)
    if "excel_bytes" not in ws:
        raise HTTPException(
            status_code=404,
            detail="Run a comparison before downloading.",
        )
    filename = f"Price_Comparison_{time.strftime('%Y-%m-%d_%H-%M')}.xlsx"
    return Response(
        content=ws["excel_bytes"],
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        },
    )
