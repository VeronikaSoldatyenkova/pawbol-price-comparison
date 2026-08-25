import json
import os
import pickle
import re
import shutil
import time
import uuid
from pathlib import Path

import numpy as np
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file, session

from price_compare_common import (
    POWERBI_REQUIRED_COLUMNS,
    file_digest,
    get_sheet_names,
    guess_column,
    parse_numeric_series,
    read_excel_sheet,
)
from price_compare_compare import (
    build_code_lookups,
    compare_all_suppliers,
    filter_result_by_codes,
    lightweight_supplier_configs,
    parse_requested_codes,
)
from price_compare_export import create_excel


APP_VERSION = "2.0.0-flask"
BASE_TMP = Path(os.environ.get("PRICE_COMPARE_TMP", "/tmp/pawbol-price-comparison"))
MAX_UPLOAD_BYTES = 300 * 1024 * 1024
SESSION_TTL_SECONDS = 24 * 60 * 60
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get(
        "PRICE_COMPARE_SECRET_KEY",
        "pawbol-price-comparison-internal-v2-change-me",
    ),
    MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)
BASE_TMP.mkdir(parents=True, exist_ok=True)

_last_cleanup = 0.0


def _cleanup_old_sessions():
    global _last_cleanup
    now = time.time()
    if now - _last_cleanup < 3600:
        return
    _last_cleanup = now

    try:
        for folder in BASE_TMP.iterdir():
            if not folder.is_dir():
                continue
            try:
                if now - folder.stat().st_mtime > SESSION_TTL_SECONDS:
                    shutil.rmtree(folder, ignore_errors=True)
            except OSError:
                pass
    except OSError:
        pass


def _sid():
    _cleanup_old_sessions()
    sid = session.get("sid")
    if not sid or not re.fullmatch(r"[a-f0-9]{32}", str(sid)):
        sid = uuid.uuid4().hex
        session["sid"] = sid
    folder = BASE_TMP / sid
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "uploads").mkdir(parents=True, exist_ok=True)
    try:
        os.utime(folder, None)
    except OSError:
        pass
    return sid


def _session_dir():
    return BASE_TMP / _sid()


def _file_meta_path(file_id):
    if not re.fullmatch(r"[a-f0-9]{32}", str(file_id or "")):
        raise ValueError("Invalid file reference.")
    return _session_dir() / "uploads" / f"{file_id}.json"


def _load_file_meta(file_id):
    meta_path = _file_meta_path(file_id)
    if not meta_path.exists():
        raise FileNotFoundError(
            "Uploaded file is no longer available. Please upload it again."
        )
    with meta_path.open("r", encoding="utf-8") as fh:
        meta = json.load(fh)
    file_path = Path(meta["path"])
    if not file_path.exists():
        raise FileNotFoundError(
            "Uploaded file is no longer available. Please upload it again."
        )
    return meta, file_path


def _load_file_bytes(file_id):
    meta, file_path = _load_file_meta(file_id)
    return meta, file_path.read_bytes()


def _json_records(df):
    if df is None or df.empty:
        return []
    safe = df.copy()
    safe = safe.replace([np.inf, -np.inf], np.nan)
    return json.loads(safe.to_json(orient="records", date_format="iso"))


def _safe_columns(df):
    return [str(col) for col in df.columns]


def _require_column(df, column, label):
    if not column or column not in df.columns:
        raise ValueError(f"{label}: selected column '{column}' was not found.")


def _validate_our(df, config):
    our_type = config.get("type")
    if our_type == "PowerBI Pricelist":
        missing = [col for col in POWERBI_REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            raise ValueError(
                "PowerBI Pricelist is missing required columns: " + ", ".join(missing)
            )
        return

    if our_type != "Free format pricelist":
        raise ValueError("Unknown current-pricelist type.")

    ean_col = config.get("ean_column") or None
    sku_col = config.get("sku_column") or None
    price_col = config.get("price_column")

    if not ean_col and not sku_col:
        raise ValueError("Free format pricelist requires at least EAN or SKU.")
    if ean_col:
        _require_column(df, ean_col, "Free format EAN")
    if sku_col:
        _require_column(df, sku_col, "Free format SKU")
    _require_column(df, price_col, "Free format Price")

    if ean_col and sku_col and ean_col == sku_col:
        raise ValueError("EAN and SKU columns cannot be the same.")
    if price_col in {ean_col, sku_col}:
        raise ValueError("Price column cannot also be an identifier column.")

    extras = list(config.get("extra_columns") or [])
    bad_extras = [col for col in extras if col not in df.columns]
    if bad_extras:
        raise ValueError(
            "Additional column(s) not found: " + ", ".join(map(str, bad_extras))
        )
    config["extra_columns"] = [
        col for col in extras if col not in {ean_col, sku_col, price_col}
    ]

    if int(parse_numeric_series(df[price_col]).notna().sum()) == 0:
        raise ValueError("No valid numeric values were found in Our Price column.")


def _validate_supplier(df, config, index):
    name = str(config.get("supplier_name") or "").strip()
    if not name:
        raise ValueError(f"Supplier {index}: supplier name is required.")

    method = config.get("match_method")
    if method not in {"SKU", "EAN", "EAN + SKU"}:
        raise ValueError(f"{name}: invalid matching method.")

    price_col = config.get("price_column")
    _require_column(df, price_col, f"{name} Price")

    if method in {"EAN", "EAN + SKU"}:
        _require_column(df, config.get("ean_column"), f"{name} EAN")
    if method in {"SKU", "EAN + SKU"}:
        _require_column(df, config.get("sku_column"), f"{name} SKU")

    ean_col = config.get("ean_column")
    sku_col = config.get("sku_column")
    if method == "EAN + SKU" and ean_col == sku_col:
        raise ValueError(f"{name}: EAN and SKU columns cannot be the same.")
    if method in {"EAN", "EAN + SKU"} and price_col == ean_col:
        raise ValueError(f"{name}: Price column cannot also be the EAN column.")
    if method in {"SKU", "EAN + SKU"} and price_col == sku_col:
        raise ValueError(f"{name}: Price column cannot also be the SKU column.")

    if int(parse_numeric_series(df[price_col]).notna().sum()) == 0:
        raise ValueError(f"{name}: no valid numeric prices were found.")


def _comparison_paths():
    folder = _session_dir()
    return {
        "result": folder / "result.pkl",
        "meta": folder / "comparison.json",
        "lookups": folder / "lookups.pkl",
        "excel": folder / "Price_Comparison.xlsx",
    }


def _load_comparison():
    paths = _comparison_paths()
    if not paths["result"].exists() or not paths["meta"].exists():
        raise FileNotFoundError("No comparison is available. Run Compare prices first.")
    result = pd.read_pickle(paths["result"])
    with paths["meta"].open("r", encoding="utf-8") as fh:
        meta = json.load(fh)
    return result, meta, paths


def _metrics(full_result, our_config):
    relevant = full_result[full_result["_relevant_for_display"]].copy()
    return {
        "primary_label": (
            "Relevant products"
            if our_config.get("type") == "PowerBI Pricelist"
            else "Products"
        ),
        "primary": int(len(relevant)),
        "matched": int(relevant["Matched Suppliers"].gt(0).sum()),
        "cheaper": int((relevant["Status"] == "CHEAPER").sum()),
        "more_expensive": int((relevant["Status"] == "MORE EXPENSIVE").sum()),
        "not_found": int((relevant["Status"] == "NOT FOUND").sum()),
    }


@app.get("/")
def index():
    _sid()
    return render_template("index.html", app_version=APP_VERSION)


@app.get("/health")
def health():
    return jsonify({"status": "ok", "version": APP_VERSION})


@app.post("/api/reset")
def reset_session_data():
    sid = _sid()
    folder = BASE_TMP / sid
    for name in ["result.pkl", "comparison.json", "lookups.pkl", "Price_Comparison.xlsx"]:
        try:
            (folder / name).unlink(missing_ok=True)
        except OSError:
            pass
    return jsonify({"ok": True})


@app.post("/api/upload")
def upload_file():
    uploaded = request.files.get("file")
    if uploaded is None or not uploaded.filename:
        return jsonify({"error": "Choose an Excel file first."}), 400

    ext = Path(uploaded.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Supported formats: .xlsx, .xlsm, .xls"}), 400

    file_id = uuid.uuid4().hex
    upload_dir = _session_dir() / "uploads"
    file_path = upload_dir / f"{file_id}{ext}"
    uploaded.save(file_path)

    try:
        file_bytes = file_path.read_bytes()
        sheets = get_sheet_names(file_bytes)
        if not sheets:
            raise ValueError("No worksheets were found in this workbook.")
    except Exception as exc:
        file_path.unlink(missing_ok=True)
        return jsonify({"error": f"Could not read workbook: {exc}"}), 400

    meta = {
        "id": file_id,
        "filename": uploaded.filename,
        "path": str(file_path),
        "digest": file_digest(file_bytes),
        "sheets": sheets,
    }
    with (upload_dir / f"{file_id}.json").open("w", encoding="utf-8") as fh:
        json.dump(meta, fh, ensure_ascii=False)

    return jsonify(
        {
            "id": file_id,
            "filename": uploaded.filename,
            "sheets": sheets,
        }
    )


@app.post("/api/file-info")
def file_info():
    payload = request.get_json(silent=True) or {}
    try:
        meta, file_bytes = _load_file_bytes(payload.get("file_id"))
        sheet = payload.get("sheet")
        if sheet not in meta["sheets"]:
            raise ValueError("Selected worksheet was not found.")
        df = read_excel_sheet(file_bytes, sheet)
        columns = _safe_columns(df)
        return jsonify(
            {
                "filename": meta["filename"],
                "sheet": sheet,
                "rows": int(len(df)),
                "columns": columns,
                "guesses": {
                    "ean": guess_column(columns, "EAN"),
                    "sku": guess_column(columns, "SKU"),
                    "price": guess_column(columns, "PRICE"),
                },
                "powerbi_missing": [
                    col for col in POWERBI_REQUIRED_COLUMNS if col not in columns
                ],
            }
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/preview")
def preview():
    payload = request.get_json(silent=True) or {}
    try:
        meta, file_bytes = _load_file_bytes(payload.get("file_id"))
        sheet = payload.get("sheet")
        if sheet not in meta["sheets"]:
            raise ValueError("Selected worksheet was not found.")
        df = read_excel_sheet(file_bytes, sheet)
        preview_df = df.iloc[:10, :30].copy()
        return jsonify(
            {
                "filename": meta["filename"],
                "sheet": sheet,
                "rows_total": int(len(df)),
                "columns_total": int(len(df.columns)),
                "columns": _safe_columns(preview_df),
                "rows": _json_records(preview_df),
            }
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/compare")
def compare():
    payload = request.get_json(silent=True) or {}
    our_payload = dict(payload.get("our") or {})
    supplier_payloads = list(payload.get("suppliers") or [])

    try:
        if not supplier_payloads:
            raise ValueError("Upload at least one alternative supplier pricelist.")

        our_meta, our_bytes = _load_file_bytes(our_payload.get("file_id"))
        our_sheet = our_payload.get("sheet")
        if our_sheet not in our_meta["sheets"]:
            raise ValueError("Our current pricelist worksheet was not found.")
        our_df = read_excel_sheet(our_bytes, our_sheet)

        our_type = our_payload.get("type")
        if our_type == "PowerBI Pricelist":
            our_config = {
                "type": "PowerBI Pricelist",
                "ean_column": "EAN",
                "sku_column": "SKU",
                "price_column": "Latest Pricelist Value",
                "extra_columns": [],
            }
        else:
            our_config = {
                "type": "Free format pricelist",
                "ean_column": our_payload.get("ean_column") or None,
                "sku_column": our_payload.get("sku_column") or None,
                "price_column": our_payload.get("price_column"),
                "extra_columns": list(our_payload.get("extra_columns") or []),
            }
        _validate_our(our_df, our_config)

        names_seen = set()
        supplier_configs = []
        for index, raw in enumerate(supplier_payloads, start=1):
            raw = dict(raw or {})
            supplier_meta, supplier_bytes = _load_file_bytes(raw.get("file_id"))
            sheet = raw.get("sheet")
            if sheet not in supplier_meta["sheets"]:
                raise ValueError(
                    f"Supplier {index}: selected worksheet was not found."
                )

            supplier_df = read_excel_sheet(supplier_bytes, sheet)
            supplier_name = str(raw.get("supplier_name") or "").strip()
            folded = supplier_name.casefold()
            if folded in names_seen:
                raise ValueError("Supplier names must be unique (case-insensitive).")
            names_seen.add(folded)

            config = {
                "supplier_name": supplier_name,
                "file_name": supplier_meta["filename"],
                "file_digest": supplier_meta["digest"],
                "sheet_name": sheet,
                "match_method": raw.get("match_method"),
                "ean_column": raw.get("ean_column") or None,
                "sku_column": raw.get("sku_column") or None,
                "price_column": raw.get("price_column"),
                "dataframe": supplier_df,
            }
            _validate_supplier(supplier_df, config, index)
            supplier_configs.append(config)

        (
            comparison_result,
            supplier_price_columns,
            duplicate_info,
            _,
        ) = compare_all_suppliers(our_df, our_config, supplier_configs)

        saved_suppliers = lightweight_supplier_configs(supplier_configs)
        ean_lookup, sku_lookup = build_code_lookups(comparison_result)
        excel_bytes = create_excel(
            full_result=comparison_result,
            supplier_price_columns=supplier_price_columns,
            supplier_configs=saved_suppliers,
            duplicate_info=duplicate_info,
            our_file_name=our_meta["filename"],
            our_config=our_config,
        )

        paths = _comparison_paths()
        comparison_result.to_pickle(paths["result"])
        with paths["lookups"].open("wb") as fh:
            pickle.dump((ean_lookup, sku_lookup), fh, protocol=pickle.HIGHEST_PROTOCOL)
        paths["excel"].write_bytes(excel_bytes)

        meta = {
            "app_version": APP_VERSION,
            "created_at": int(time.time()),
            "supplier_price_columns": supplier_price_columns,
            "duplicate_info": duplicate_info,
            "supplier_configs": saved_suppliers,
            "our_config": our_config,
            "our_file_name": our_meta["filename"],
            "metrics": _metrics(comparison_result, our_config),
        }
        with paths["meta"].open("w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, default=str)

        return jsonify(
            {
                "ok": True,
                "metrics": meta["metrics"],
                "duplicate_info": duplicate_info,
                "supplier_price_columns": supplier_price_columns,
                "our_type": our_config["type"],
            }
        )
    except Exception as exc:
        app.logger.exception("Comparison failed")
        return jsonify({"error": str(exc)}), 400


@app.get("/api/results")
def results():
    try:
        full_result, meta, _ = _load_comparison()
        relevant = full_result[full_result["_relevant_for_display"]].copy()
        display_filter = request.args.get("filter", "all")

        if display_filter == "cheaper":
            shown = relevant[relevant["Status"] == "CHEAPER"].copy()
        elif display_filter == "matched":
            shown = relevant[relevant["Matched Suppliers"].gt(0)].copy()
        elif display_filter == "not_found":
            shown = relevant[relevant["Matched Suppliers"].eq(0)].copy()
        else:
            shown = relevant.copy()

        shown = shown.drop(columns=["_relevant_for_display"], errors="ignore")
        page = max(1, int(request.args.get("page", 1)))
        page_size = int(request.args.get("page_size", 100))
        page_size = min(max(page_size, 25), 500)
        total = int(len(shown))
        start = (page - 1) * page_size
        end = start + page_size
        page_df = shown.iloc[start:end]

        return jsonify(
            {
                "columns": _safe_columns(page_df if len(page_df.columns) else shown),
                "rows": _json_records(page_df),
                "total": total,
                "page": page,
                "page_size": page_size,
                "supplier_price_columns": meta["supplier_price_columns"],
                "metrics": meta["metrics"],
            }
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/quick-filter")
def quick_filter():
    payload = request.get_json(silent=True) or {}
    try:
        full_result, meta, paths = _load_comparison()
        text = str(payload.get("text") or "")
        requested = parse_requested_codes(text)
        if not requested:
            raise ValueError("Paste at least one EAN or SKU first.")
        if len(requested) > 5000:
            raise ValueError("Quick Filter supports up to 5,000 pasted rows at once.")

        if paths["lookups"].exists():
            with paths["lookups"].open("rb") as fh:
                ean_lookup, sku_lookup = pickle.load(fh)
        else:
            ean_lookup, sku_lookup = build_code_lookups(full_result)

        shown = filter_result_by_codes(
            full_result,
            requested,
            ean_lookup,
            sku_lookup,
            meta["supplier_price_columns"],
        ).drop(columns=["_relevant_for_display"], errors="ignore")

        invalid_target_count = sum(
            1
            for item in requested
            if item.get("target_raw") and not item.get("target_valid")
        )
        return jsonify(
            {
                "columns": _safe_columns(shown),
                "rows": _json_records(shown),
                "supplier_price_columns": meta["supplier_price_columns"],
                "summary": {
                    "requested": int(len(shown)),
                    "not_found": int(
                        (shown["Lookup Status"] == "CODE NOT FOUND").sum()
                    ),
                    "with_target": int(shown["Target Price"].notna().sum()),
                    "target_met": int(
                        shown["Below Target"]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .ne("")
                        .sum()
                    ),
                    "invalid_targets": int(invalid_target_count),
                },
            }
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/download")
def download():
    try:
        _, meta, paths = _load_comparison()
        if not paths["excel"].exists():
            raise FileNotFoundError("Excel export is not available. Run comparison again.")
        return send_file(
            paths["excel"],
            as_attachment=True,
            download_name="Price_Comparison.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.errorhandler(413)
def too_large(_exc):
    return jsonify({"error": "File is larger than the 300 MB upload limit."}), 413


@app.errorhandler(500)
def internal_error(exc):
    app.logger.exception("Unhandled server error", exc_info=exc)
    return jsonify({"error": "Unexpected server error. Check application logs."}), 500
