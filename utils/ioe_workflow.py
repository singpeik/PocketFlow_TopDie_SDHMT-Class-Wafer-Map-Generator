import json
import re
import os
from tempfile import NamedTemporaryFile
from decimal import Decimal, InvalidOperation
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from utils.file_io import read_yaml_config
from utils.pyuber_client import execute_query
from utils.data_wrangling import wrangle_results

DEFAULT_IOE_WAFER_MAP_SOURCE = Path(
    "domain-knowledge/wafer-map/IOE wafer map file.xlsx"
)


RUN_DATE_STAMP = date.today().strftime("%Y%m%d")


def add_date_stamp_to_path(path_like, separator="_"):
    path = Path(path_like)
    suffix = path.suffix
    stem = path.stem
    date_pattern = rf"{re.escape(separator)}\d{{8}}$" if separator else r"\d{8}$"
    if re.search(date_pattern, stem):
        return path
    return path.with_name(f"{stem}{separator}{RUN_DATE_STAMP}{suffix}")


DEFAULT_IOE_WAFER_MAP_OUTPUT = Path("domain-knowledge/wafer-map/ioe_wafer_map.json")
DEFAULT_GPU_WAFER_MAP_SOURCE = Path("domain-knowledge/wafer-map/GPU wafer map file.xlsx")
DEFAULT_GPU_WAFER_MAP_OUTPUT = Path("domain-knowledge/wafer-map/gpu_wafer_map.json")
DEFAULT_ALL_PRODUCT_MAPS_OUTPUT = Path("domain-knowledge/wafer-map/all_product_maps.json")
DEFAULT_OUTPUT_DIR = Path("outputs")
DEFAULT_FINAL_WORKBOOK = add_date_stamp_to_path(DEFAULT_OUTPUT_DIR / "Final.xlsx")
DEFAULT_VID_OUTPUT = Path("VID.csv")
DEFAULT_WLA_SDHMT_OUTPUT = Path("WLA lots SDHMT.xlsx")
CATTS_CLEANED_SHEET = "CATTS-cleaned"
VID_SHEET = "VID"
WLA_LOTS_SHEET = "WLA lots"
VID_TO_IOE_WLA_SHEET = "VID to Wafer&WLA"
VID_TO_WLA_SHEET = "VID to WLA"
WLA_SDHMT_SHEET = "WLA lots SDHMT"
AQUA_CLASS_SHEET = "AQUA class_7820"
AQUA_CLASS_SELECTION_FLAGS = ('LATO_Valid_Flag', 'Operation_Final_Latest_Flag')
CATTS_REQUIRED_COLUMNS = [
    "FAB_LOT",
    "FAB_WAFER",
    "X_LOCATION",
    "Y_LOCATION",
    "VISUAL_ID",
    "INTEL_DIE_INDEX",
    "SS_ID",
    "LOT_AT_1204",
    "LOT_AT_216365",
    "LOT_AT_852060",
]
WLA_LOOKUP_COLUMNS = ["WLA lot", "WLA wafer", "WLA die index"]
CATTS_ENRICHED_COLUMNS = CATTS_REQUIRED_COLUMNS + WLA_LOOKUP_COLUMNS
VID_IOE_WLA_COLUMNS = [
    "FAB_LOT",
    "FAB_WAFER",
    "X_LOCATION",
    "Y_LOCATION",
    "VISUAL_ID",
    "INTEL_DIE_INDEX",
    "SS_ID",
    "LOT_AT_1204",
]
WLA_REQUIRED_COLUMNS = [
    "LOT",
    "OPERATION",
    "DEVREVSTEP",
    "TEST_END_DATE",
    "SESSION_LATEST_FLAG",
    "DEVICE_LATEST_FLAG",
    "PROGRAM_NAME",
    "WAFER_ID",
    "SORT_LOT",
    "WAFER_SCRIBE",
    "SORT_WAFER_ID",
    "SORT_X",
    "SORT_Y",
    "INTERFACE_BIN",
    "FUNCTIONAL_BIN",
    "WLA die index",
]
SHORT_WID_FIELDS = ["FAB_WAFER", "WAFER_ID", "SORT_WAFER_ID"]
LONG_WAFER_ID_FIELDS = ["WAFER_SCRIBE"]


def normalize_lot7_wid3(user_input):
    raw_value = str(user_input or "").strip().upper()
    if not raw_value:
        raise ValueError("Wafer Lot7_WID3 is required.")

    match = re.match(r"^([A-Z0-9]{7,8})[_\-\s]?([0-9]{1,3})$", raw_value)
    if not match:
        raise ValueError(
            "Expected wafer input in the format LOT7_WID3 or LOT8_WID3, "
            "for example C442Y19_243 or C525Y9K0_102."
        )

    lot_token = match.group(1)
    lot7 = lot_token[:7]
    wid3 = match.group(2).zfill(3)
    return {"lot7": lot7, "wid3": wid3, "lot7_wid3": f"{lot7}_{wid3}"}


def normalize_lot7_wid3_list(user_input):
    raw_value = str(user_input or "").strip().upper()
    if not raw_value:
        raise ValueError("At least one wafer Lot7_WID3 is required.")

    tokens = [token for token in re.split(r"[\s,;]+", raw_value) if token]
    normalized_items = [normalize_lot7_wid3(token) for token in tokens]
    deduped_items = list({item["lot7_wid3"]: item for item in normalized_items}.values())
    return {
        "lot7_wid3_values": [item["lot7_wid3"] for item in deduped_items],
        "selected_wafers": deduped_items,
        "wafer_count": len(deduped_items),
        "lot7_wid3": deduped_items[0]["lot7_wid3"],
    }


def _normalize_column_name(name):
    cleaned = str(name or "").strip().upper()
    return re.sub(r"[^A-Z0-9]+", "_", cleaned).strip("_")


def _find_column(df, candidates, fallback_index=None):
    normalized_lookup = {_normalize_column_name(col): col for col in df.columns}
    for candidate in candidates:
        column = normalized_lookup.get(_normalize_column_name(candidate))
        if column is not None:
            return column

    if fallback_index is not None and fallback_index < len(df.columns):
        return df.columns[fallback_index]

    raise KeyError(f"Unable to find required column from candidates: {candidates}")


def cleanup_ioe_intermediate_files(summary=None, output_dir=None):
    summary = summary or {}
    cleanup_targets = []

    for key in ["filtered_report", "summary_file"]:
        target = summary.get(key)
        if target:
            cleanup_targets.append(Path(target))

    if output_dir:
        cleanup_targets.append(Path(output_dir) / "step6_temp")

    removed_paths = []
    for target_path in cleanup_targets:
        if not target_path.exists():
            continue
        if target_path.is_dir():
            for child in sorted(target_path.rglob('*'), reverse=True):
                if child.is_file():
                    child.unlink()
                elif child.is_dir():
                    child.rmdir()
            target_path.rmdir()
        else:
            target_path.unlink()
        removed_paths.append(str(target_path))

    return removed_paths


def _coerce_wid3(value):
    if pd.isna(value):
        return None

    numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(numeric_value):
        if float(numeric_value).is_integer():
            return str(int(numeric_value)).zfill(3)

    raw_value = str(value).strip()
    if not raw_value:
        return None

    digits = re.findall(r"\d+", raw_value)
    if not digits:
        return raw_value.upper()

    joined = "".join(digits)
    if len(joined) >= 3:
        return joined[-3:].zfill(3)
    return joined.zfill(3)


def _normalize_short_wid(value, field_name):
    normalized = _coerce_wid3(value)
    if normalized is None:
        return None

    if not re.fullmatch(r"\d{3}", str(normalized)):
        raise ValueError(
            f"{field_name} must normalize to a 3-digit WID, but received value: {value!r}"
        )
    return str(normalized)


def _normalize_short_wid_series(series, field_name):
    normalized = series.apply(lambda value: _normalize_short_wid(value, field_name) if pd.notna(value) else None)
    invalid_mask = series.notna() & normalized.isna()
    if invalid_mask.any():
        sample_values = series[invalid_mask].astype(str).head(5).tolist()
        raise ValueError(
            f"{field_name} contains values that cannot be normalized to a 3-digit WID: {sample_values}"
        )
    return normalized


def _validate_long_wafer_id_series(series, field_name):
    non_null = series.dropna().astype(str).str.strip()
    invalid = non_null[non_null.str.len() != 12]
    if not invalid.empty:
        sample_values = invalid.head(5).tolist()
        raise ValueError(
            f"{field_name} must remain a 12-character long wafer ID. Invalid values: {sample_values}"
        )


def _ensure_parent_dir(path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)


def _replace_excel_sheet(workbook_path, sheet_name, df):
    workbook = Path(workbook_path)
    _ensure_parent_dir(workbook)
    if workbook.exists():
        with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(workbook, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def _is_excel_file(path):
    return Path(path).suffix.lower() in {".xlsx", ".xlsm", ".xltx", ".xltm"}


def _is_csv_file(path):
    return Path(path).suffix.lower() == ".csv"


def _format_catts_visual_id_excel_column(raw_report_path):
    """Best-effort Excel formatting for CATTS VISUAL_ID to avoid scientific-notation display."""
    raw_path = Path(raw_report_path)
    supported_suffixes = {".xlsx", ".xlsm", ".xltx", ".xltm"}
    if raw_path.suffix.lower() not in supported_suffixes:
        return {
            "applied": False,
            "reason": f"unsupported Excel format for openpyxl: {raw_path.suffix}",
        }

    workbook = load_workbook(raw_path)
    sheet = workbook[workbook.sheetnames[0]]

    header_map = {
        _normalize_column_name(cell.value): cell.column
        for cell in sheet[1]
        if cell.value is not None
    }
    visual_id_column = header_map.get("VISUAL_ID")
    if visual_id_column is None:
        # Fallback to CATTS convention where VISUAL_ID is usually column E.
        visual_id_column = 5 if sheet.max_column >= 5 else None

    if visual_id_column is None:
        workbook.save(raw_path)
        return {
            "applied": False,
            "reason": "unable to locate VISUAL_ID column",
        }

    formatted_cells = 0
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=visual_id_column)
        if cell.value in (None, ""):
            continue
        # Equivalent to Format Cells -> Number -> Decimal places: 0.
        cell.number_format = "0"
        formatted_cells += 1

    workbook.save(raw_path)
    return {
        "applied": True,
        "column_index": int(visual_id_column),
        "formatted_cells": int(formatted_cells),
        "reason": "ok",
    }


def _format_catts_visual_id_csv_column(raw_report_path):
    """Best-effort VISUAL_ID normalization for CSV CATTS reports."""
    raw_path = Path(raw_report_path)
    frame = pd.read_csv(raw_path, dtype=str, keep_default_na=False)
    if frame.empty:
        frame.to_csv(raw_path, index=False)
        return {
            "applied": False,
            "reason": "empty csv",
        }

    visual_id_column = "VISUAL_ID" if "VISUAL_ID" in frame.columns else (frame.columns[4] if len(frame.columns) >= 5 else None)
    if visual_id_column is None:
        frame.to_csv(raw_path, index=False)
        return {
            "applied": False,
            "reason": "unable to locate VISUAL_ID column",
        }

    formatted_cells = 0

    def _normalize_for_csv(value):
        nonlocal formatted_cells
        if value is None:
            return value
        raw_value = str(value).strip()
        if not raw_value:
            return raw_value
        if "_" in raw_value:
            return raw_value
        normalized = _normalize_visual_id(raw_value)
        if normalized is None:
            return raw_value
        formatted_cells += 1
        return normalized

    frame[visual_id_column] = frame[visual_id_column].map(_normalize_for_csv)
    frame.to_csv(raw_path, index=False)
    return {
        "applied": True,
        "column_index": int(list(frame.columns).index(visual_id_column) + 1),
        "formatted_cells": int(formatted_cells),
        "reason": "ok",
    }


def _format_catts_visual_id_source(raw_report_path):
    if _is_excel_file(raw_report_path):
        return _format_catts_visual_id_excel_column(raw_report_path)
    if _is_csv_file(raw_report_path):
        return _format_catts_visual_id_csv_column(raw_report_path)
    return {
        "applied": False,
        "reason": f"unsupported CATTS report format: {Path(raw_report_path).suffix}",
    }


def _looks_like_precision_lost_scientific_notation(value):
    raw_value = str(value or "").strip().upper()
    match = re.fullmatch(r"([+-]?\d*\.?\d+)[E]([+-]?\d+)", raw_value)
    if not match:
        return False

    mantissa_digits = re.sub(r"\D", "", match.group(1)).lstrip("0")
    # 13-digit VID should carry 13 significant digits in scientific notation.
    return len(mantissa_digits) < 13


def _analyze_true_vid_quality(series):
    cleaned_series = series.astype(str).str.strip()
    true_vid_mask = series.notna() & (cleaned_series != "") & ~cleaned_series.str.contains("_", regex=False)
    true_vid_series = series[true_vid_mask]

    normalized_series = true_vid_series.map(_normalize_visual_id)
    # Accept 13-character alphanumeric OR pure 13-digit numeric format
    is_valid_format_mask = pd.Series(
        [bool(re.fullmatch(r"[\dA-Z]{13}", str(value or "").upper())) for value in normalized_series],
        index=normalized_series.index,
        dtype=bool,
    )
    precision_loss_mask = pd.Series(
        [_looks_like_precision_lost_scientific_notation(value) for value in true_vid_series],
        index=true_vid_series.index,
        dtype=bool,
    )

    valid_mask = is_valid_format_mask & ~precision_loss_mask
    invalid_values = true_vid_series[~valid_mask]

    return {
        "true_vid_rows": int(len(true_vid_series)),
        "valid_13_character_rows": int(valid_mask.sum()),
        "invalid_13_character_rows": int((~valid_mask).sum()),
        "precision_loss_rows": int(precision_loss_mask.sum()),
        "invalid_13_character_samples": [str(value) for value in invalid_values.head(10).tolist()],
    }


def _warn_and_block_on_invalid_true_vids(raw_path, visual_id_quality):
    invalid_count = int(visual_id_quality.get("invalid_13_character_rows", 0))
    if invalid_count <= 0:
        return

    sample_values = visual_id_quality.get("invalid_13_character_samples", [])
    sample_text = "\n".join(sample_values[:5]) if sample_values else "(no sample available)"
    message = (
        "CATTS report VID check failed.\n\n"
        f"File: {raw_path}\n"
        f"True VID rows checked: {visual_id_quality.get('true_vid_rows', 0)}\n"
        f"Rows not recoverable as 13-digit VID: {invalid_count}\n\n"
        "Examples:\n"
        f"{sample_text}\n\n"
        "Please fix VISUAL_ID formatting in CATTS file (Column E) and reload."
    )

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showwarning("CATTS VID Format Warning", message, parent=root)
        root.destroy()
    except Exception:
        pass

    raise ValueError(message)


def _apply_short_wid_excel_format(workbook_path, sheet_name, short_wid_columns):
    workbook = load_workbook(workbook_path)
    if sheet_name not in workbook.sheetnames:
        workbook.save(workbook_path)
        return

    sheet = workbook[sheet_name]
    header_map = {cell.value: cell.column for cell in sheet[1] if cell.value}

    for column_name in short_wid_columns:
        column_index = header_map.get(column_name)
        if column_index is None:
            continue

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value in (None, ""):
                continue

            normalized = _normalize_short_wid(cell.value, column_name)
            cell.value = normalized
            cell.number_format = "@"

    workbook.save(workbook_path)


def _normalize_wla_lot7_pattern(value):
    if pd.isna(value):
        return None

    raw_value = str(value).strip().upper()
    if not raw_value.startswith("W"):
        return None

    return f"{raw_value[:7]}%"


def _build_filter_condition(lot_patterns, product="%8P%", days=600):
    processed_lots = []
    for lot in lot_patterns:
        normalized = str(lot or "").strip().upper()
        if not normalized:
            continue
        if "%" in normalized:
            processed_lots.append(f"v0.lot LIKE '{normalized}'")
        else:
            processed_lots.append(f"v0.lot LIKE '{normalized}%'")

    lot_condition = "(" + " OR ".join(processed_lots) + ")" if processed_lots else "1=1"

    if "%" in product:
        product_condition = f"v0.devrevstep LIKE '{product}'"
    else:
        product_condition = f"v0.devrevstep = '{product}'"

    return (
        f"{lot_condition} AND {product_condition} "
        f"AND v0.test_end_date_time >= TRUNC(SYSDATE) - {int(days)}"
    )


def _load_query_with_filters(query_path, lot_patterns, product="%8P%", days=600):
    query_text = Path(query_path).read_text(encoding="utf-8")
    filter_condition = _build_filter_condition(lot_patterns, product=product, days=days)
    if "{FILTER_CONDITION}" not in query_text:
        raise ValueError("query.sql must contain the {FILTER_CONDITION} placeholder.")
    return query_text.replace("{FILTER_CONDITION}", filter_condition), filter_condition


def _compose_wla_die_index(wafer_scribe, sort_x, sort_y):
    if pd.isna(wafer_scribe) or pd.isna(sort_x) or pd.isna(sort_y):
        return None

    x_value = int(pd.to_numeric(sort_x, errors="coerce"))
    y_value = int(pd.to_numeric(sort_y, errors="coerce"))
    return f"{str(wafer_scribe).strip()}_{x_value}_{y_value}"


def _normalize_visual_id(value):
    if pd.isna(value):
        return None

    normalized = str(value).strip()
    if not normalized:
        return None

    normalized = normalized.replace(",", "")
    try:
        decimal_value = Decimal(normalized)
    except (InvalidOperation, ValueError):
        return normalized

    if not decimal_value.is_finite():
        return None

    if decimal_value == decimal_value.to_integral_value():
        return str(int(decimal_value))

    return format(decimal_value.normalize(), "f").rstrip("0").rstrip(".")


def _build_vid_cross_reference_tables(canonical_df, selected_tokens, visual_id_col, fab_lot_col, fab_wafer_col):
    working_df = canonical_df.copy()
    working_df["VISUAL_ID"] = working_df[visual_id_col].map(_normalize_visual_id)
    working_df = working_df.dropna(subset=["VISUAL_ID"])

    selected_visual_ids = set(working_df.loc[working_df["Lot7_WID3"].isin(selected_tokens), "VISUAL_ID"].tolist())
    if not selected_visual_ids:
        empty_cross_ref = pd.DataFrame(columns=VID_IOE_WLA_COLUMNS)
        empty_wla = pd.DataFrame(columns=VID_IOE_WLA_COLUMNS)
        return empty_cross_ref, empty_wla

    working_df = working_df.loc[working_df["VISUAL_ID"].isin(selected_visual_ids)].copy()
    working_df["FAB_LOT"] = working_df[fab_lot_col].map(lambda value: str(value).strip().upper() if pd.notna(value) else None)
    working_df["FAB_WAFER"] = _normalize_short_wid_series(working_df[fab_wafer_col], "FAB_WAFER")
    working_df["ROW_TYPE"] = None
    working_df.loc[working_df["Lot7_WID3"].isin(selected_tokens), "ROW_TYPE"] = "IOE"
    working_df.loc[working_df["FAB_LOT"].fillna("").str.startswith("W"), "ROW_TYPE"] = "WLA"
    working_df = working_df.loc[working_df["ROW_TYPE"].isin(["IOE", "WLA"])].copy()

    token_rank = {token: index for index, token in enumerate(selected_tokens)}
    working_df["ROW_PRIORITY"] = working_df["ROW_TYPE"].map({"IOE": 0, "WLA": 1})
    working_df["LOT7_WID3_ORDER"] = working_df["Lot7_WID3"].map(token_rank).fillna(len(token_rank)).astype(int)
    working_df = working_df.sort_values(
        by=["VISUAL_ID", "ROW_PRIORITY", "LOT7_WID3_ORDER", "FAB_LOT", "FAB_WAFER", "INTEL_DIE_INDEX"],
        kind="stable",
    )

    ioe_rows = (
        working_df.loc[working_df["ROW_TYPE"] == "IOE"]
        .drop_duplicates(subset=["VISUAL_ID"], keep="first")
    )
    wla_rows = (
        working_df.loc[working_df["ROW_TYPE"] == "WLA"]
        .drop_duplicates(subset=["VISUAL_ID"], keep="first")
    )

    cross_ref_df = pd.concat([ioe_rows, wla_rows], ignore_index=True)
    cross_ref_df["ROW_PRIORITY"] = cross_ref_df["ROW_TYPE"].map({"IOE": 0, "WLA": 1})
    cross_ref_df = cross_ref_df.sort_values(by=["VISUAL_ID", "ROW_PRIORITY"], kind="stable")
    cross_ref_df = _align_columns(cross_ref_df, VID_IOE_WLA_COLUMNS)

    vid_to_wla_df = _align_columns(wla_rows.copy(), VID_IOE_WLA_COLUMNS)
    vid_to_wla_df = vid_to_wla_df.sort_values(by=["VISUAL_ID", "FAB_LOT", "FAB_WAFER"], kind="stable")
    return cross_ref_df.reset_index(drop=True), vid_to_wla_df.reset_index(drop=True)


def _align_columns(df, ordered_columns):
    aligned = pd.DataFrame()
    for column in ordered_columns:
        aligned[column] = df[column] if column in df.columns else pd.Series([None] * len(df))
    return aligned


def _load_existing_ioe_wafer_map(target_path):
    if not target_path.exists():
        return None

    payload = json.loads(target_path.read_text(encoding="utf-8"))
    active_dies = payload.get("active_dies") or []
    normalized_active_dies = []
    for die in active_dies:
        if not isinstance(die, (list, tuple)) or len(die) != 2:
            continue
        x_val, y_val = die
        try:
            normalized_active_dies.append([int(x_val), int(y_val)])
        except (TypeError, ValueError):
            continue

    payload["map_name"] = payload.get("map_name") or "IOE"
    payload["source_file"] = payload.get("source_file") or str(target_path)
    payload["active_dies"] = normalized_active_dies
    payload["active_die_count"] = len(normalized_active_dies)
    payload["wafer_count"] = int(payload.get("wafer_count") or 0)
    return payload


def build_ioe_wafer_map(source_path=None, output_path=None, all_maps_path=None, force_rebuild=False, map_name="IOE"):
    map_name = str(map_name).strip().upper()
    if map_name not in {"IOE", "GPU"}:
        raise ValueError("Wafer geometry must be IOE or GPU.")
    default_source = DEFAULT_GPU_WAFER_MAP_SOURCE if map_name == "GPU" else DEFAULT_IOE_WAFER_MAP_SOURCE
    default_output = DEFAULT_GPU_WAFER_MAP_OUTPUT if map_name == "GPU" else DEFAULT_IOE_WAFER_MAP_OUTPUT
    source = Path(source_path or default_source)
    target = Path(output_path or default_output)
    all_maps_target = Path(all_maps_path or DEFAULT_ALL_PRODUCT_MAPS_OUTPUT)

    # Prefer a previously generated IOE map JSON for stable, repeatable runs.
    # Rebuild from the Excel source only when explicitly forced.
    existing_payload = None
    if not force_rebuild:
        existing_payload = _load_existing_ioe_wafer_map(target)

    if existing_payload is not None:
        if existing_payload["map_name"] != map_name:
            raise ValueError(f"Expected {map_name} geometry, but {target} contains {existing_payload['map_name']}.")
        payload = existing_payload
        payload["build_mode"] = "reused_json"
    elif source.exists():
        df = pd.read_excel(source)

        x_col = _find_column(df, ["X_LOCATION", "X", "DIE_X"], fallback_index=2)
        y_col = _find_column(df, ["Y_LOCATION", "Y", "DIE_Y"], fallback_index=3)

        wafer_columns = [col for col in df.columns if _normalize_column_name(col) in {"FAB_WAFER", "WAFER", "WAFER_ID"}]
        wafer_count = int(df[wafer_columns[0]].nunique()) if wafer_columns else 1
        working = df[[x_col, y_col]].copy()
        working.columns = ["X_LOCATION", "Y_LOCATION"]
        working["X_LOCATION"] = pd.to_numeric(working["X_LOCATION"], errors="coerce")
        working["Y_LOCATION"] = pd.to_numeric(working["Y_LOCATION"], errors="coerce")
        if working.empty or working.isna().any().any():
            raise ValueError(f"{map_name} geometry must contain valid X/Y coordinates on every row.")
        if ((working % 1) != 0).any().any():
            raise ValueError(f"{map_name} geometry coordinates must be integers.")
        working["X_LOCATION"] = working["X_LOCATION"].astype(int)
        working["Y_LOCATION"] = working["Y_LOCATION"].astype(int)

        active_dies = sorted(
            {(int(row.X_LOCATION), int(row.Y_LOCATION)) for row in working.itertuples()},
            key=lambda item: (item[1], item[0]),
        )

        payload = {
            "map_name": map_name,
            "source_file": str(source),
            "wafer_count": wafer_count,
            "active_die_count": len(active_dies),
            "active_dies": [[x, y] for x, y in active_dies],
            "build_mode": "rebuilt_from_source",
        }
    else:
        payload = _load_existing_ioe_wafer_map(target)
        if payload is None:
            raise FileNotFoundError(
                f"{map_name} wafer map source file not found: {source}. "
                f"Also could not find existing {map_name} map JSON at: {target}"
            )
        payload["build_mode"] = "reused_json"

    if not payload["active_dies"]:
        raise ValueError(f"{map_name} geometry contains no active dies.")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    if all_maps_target.exists():
        maps_payload = json.loads(all_maps_target.read_text(encoding="utf-8"))
    else:
        all_maps_target.parent.mkdir(parents=True, exist_ok=True)
        maps_payload = {}

    maps_payload[map_name] = {"active_dies": payload["active_dies"]}
    all_maps_target.write_text(json.dumps(maps_payload, indent=2), encoding="utf-8")

    return payload


def collect_ioe_user_inputs(initial_dir=None):
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError as exc:
        raise RuntimeError("Tkinter is required for the wafer workflow UI.") from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    dialog = tk.Toplevel(root)
    dialog.title("Wafer Test Result Analyzer - Wafer Selection")
    dialog.attributes("-topmost", True)
    dialog.geometry("560x520+200+200")

    selection = {"value": None}
    product = tk.StringVar(value="IOE")
    product_frame = tk.Frame(dialog)
    product_frame.pack(fill="x", padx=12, pady=(12, 0))
    tk.Label(product_frame, text="Wafer geometry:").pack(side="left")
    for map_name in ("IOE", "GPU"):
        tk.Radiobutton(product_frame, text=map_name, variable=product, value=map_name).pack(side="left", padx=8)

    instruction = tk.Label(
        dialog,
        text=(
            "Please input the wafer list that you want for analysis.\n\n"
            "Acceptable input methods:\n"
            "1. Single Lot7_WID3: C442Y19_017\n"
            "2. Single Lot8_WID3: C525Y9K0_102 -> auto-convert to C525Y9K_102\n"
            "3. Multi-wafer paste from Excel column: one wafer per line\n"
            "4. Multi-wafer paste separated by tab, space, comma, or semicolon\n\n"
            "Accepted examples:\n"
            "C442Y19_017\n"
            "C525Y9K0_102\n"
            "C442Y19_017\nC442Y19_018"
        ),
        justify="left",
        anchor="w",
    )
    instruction.pack(fill="x", padx=12, pady=(12, 6))

    input_box = tk.Text(dialog, height=12, width=50, wrap="word")
    input_box.pack(fill="both", expand=True, padx=12, pady=6)
    input_box.focus_set()

    button_frame = tk.Frame(dialog)
    button_frame.pack(fill="x", padx=12, pady=(0, 12))

    def submit_selection():
        selection["value"] = input_box.get("1.0", tk.END).strip()
        selection["product"] = product.get()
        dialog.destroy()

    def cancel_selection():
        selection["value"] = None
        dialog.destroy()

    tk.Button(button_frame, text="OK", width=12, command=submit_selection).pack(side="right", padx=(6, 0))
    tk.Button(button_frame, text="Cancel", width=12, command=cancel_selection).pack(side="right")

    dialog.protocol("WM_DELETE_WINDOW", cancel_selection)
    root.wait_window(dialog)

    lot7_wid3_input = selection["value"]
    if lot7_wid3_input is None:
        root.destroy()
        raise RuntimeError("User cancelled wafer Lot7_WID3 input.")

    normalized = normalize_lot7_wid3_list(lot7_wid3_input)
    normalized["product"] = selection["product"]
    product_name = normalized["product"]

    messagebox.showinfo(
        "CATTS Raw Report",
        (
            "1. Go to CATTS --> Extended report --> MRB PRT Data extract reports --> "
            f"QueryUnitData_UsingWafer --> input the {product_name} wafers with Lot7_WID3 --> "
            f"download all {product_name} wafers die index report to get every die index for accurate ULT data pull.\n\n"
            f"2. Go to CATTS --> Query --> lot to unit integration --> input all {product_name} die index from step 1 --> "
            "make sure to select and manually type ACW with 2 operations: 852060 and 216365 --> "
            "download raw Excel report, then upload it into this app."
        ),
        parent=root,
    )

    raw_report_path = filedialog.askopenfilename(
        parent=root,
        title="Select the CATTS raw excel report",
        initialdir=str(initial_dir) if initial_dir else None,
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    if not raw_report_path:
        root.destroy()
        raise RuntimeError("User cancelled CATTS raw report selection.")

    root.destroy()
    normalized["raw_report_path"] = raw_report_path
    return normalized


def parse_catts_raw_report(raw_report_path, lot7=None, wid3=None, output_dir=None, lot7_wid3_values=None, final_workbook_path=None):
    raw_path = Path(raw_report_path)
    if not raw_path.exists():
        raise FileNotFoundError(f"CATTS raw report not found: {raw_path}")

    visual_id_excel_formatting = {"applied": False, "reason": "not-run"}
    try:
        visual_id_excel_formatting = _format_catts_visual_id_source(raw_path)
    except Exception as exc:
        visual_id_excel_formatting = {
            "applied": False,
            "reason": f"formatting failed: {exc}",
        }

    if lot7_wid3_values:
        selected_tokens = [normalize_lot7_wid3(item)["lot7_wid3"] for item in lot7_wid3_values]
    elif lot7 is not None and wid3 is not None:
        selected_tokens = [f"{str(lot7).upper()}_{str(wid3).zfill(3)}"]
    else:
        raise ValueError("At least one Lot7_WID3 selection is required for CATTS filtering.")

    if _is_excel_file(raw_path):
        df = pd.read_excel(raw_path)
    elif _is_csv_file(raw_path):
        df = pd.read_csv(raw_path)
    else:
        raise ValueError(
            "CATTS raw report must be an Excel (.xlsx/.xlsm/.xltx/.xltm) or CSV (.csv) file."
        )
    required_input_columns = ["FAB_LOT", "FAB_WAFER", "VISUAL_ID", "INTEL_DIE_INDEX"]
    missing_input_columns = [column for column in required_input_columns if column not in df.columns]
    if missing_input_columns:
        raise ValueError(
            "CATTS raw report is missing required columns: " + ", ".join(missing_input_columns)
        )

    fab_lot_col = df.columns[0]
    fab_wafer_col = df.columns[1]
    visual_id_col = df.columns[4]

    canonical_df = df.copy()
    visual_id_quality = _analyze_true_vid_quality(canonical_df["VISUAL_ID"])
    _warn_and_block_on_invalid_true_vids(raw_path, visual_id_quality)
    canonical_df["LOT7"] = canonical_df[fab_lot_col].astype(str).str.strip().str.upper().str[:7]
    canonical_df["WID3"] = _normalize_short_wid_series(canonical_df[fab_wafer_col], "FAB_WAFER")
    canonical_df["Lot7_WID3"] = canonical_df["LOT7"] + "_" + canonical_df["WID3"].fillna("")

    filtered_df = canonical_df[canonical_df["Lot7_WID3"].isin(selected_tokens)].copy()

    filtered_df["IS_FAKE_VID"] = filtered_df[visual_id_col].astype(str).str.contains("_", regex=False, na=False)
    filtered_df["IS_TRUE_VID"] = ~filtered_df["IS_FAKE_VID"] & filtered_df[visual_id_col].notna()

    filtered_df["FAB_WAFER"] = _normalize_short_wid_series(filtered_df[fab_wafer_col], "FAB_WAFER")
    cleaned_df = _align_columns(filtered_df, CATTS_REQUIRED_COLUMNS)

    vid_to_ioe_wla_df, vid_to_wla_df = _build_vid_cross_reference_tables(
        canonical_df=canonical_df,
        selected_tokens=selected_tokens,
        visual_id_col=visual_id_col,
        fab_lot_col=fab_lot_col,
        fab_wafer_col=fab_wafer_col,
    )
    wla_lookup_df = (
        vid_to_wla_df[["VISUAL_ID", "FAB_LOT", "FAB_WAFER", "INTEL_DIE_INDEX"]]
        .dropna(subset=["VISUAL_ID"])
        .drop_duplicates(subset=["VISUAL_ID"], keep="first")
        .rename(columns={
            "FAB_LOT": "WLA lot",
            "FAB_WAFER": "WLA wafer",
            "INTEL_DIE_INDEX": "WLA die index",
        })
    )
    cleaned_df = cleaned_df.drop(columns=[column for column in WLA_LOOKUP_COLUMNS if column in cleaned_df.columns])
    cleaned_df["VISUAL_ID"] = cleaned_df["VISUAL_ID"].map(_normalize_visual_id)
    cleaned_df = cleaned_df.merge(wla_lookup_df, how="left", on="VISUAL_ID")
    cleaned_df = _align_columns(cleaned_df, CATTS_ENRICHED_COLUMNS)

    lot_columns = [
        column
        for column in canonical_df.columns
        if column.startswith("LOT_AT_") or column in ["LOT_AT_216365", "LOT_AT_852060"]
    ]
    extracted_wla_lots = set()
    for column in lot_columns:
        if column in filtered_df.columns:
            for value in filtered_df[column].dropna().tolist():
                val_str = str(value).strip().upper()
                if val_str.startswith("W"):
                    extracted_wla_lots.add(val_str)

    if not vid_to_wla_df.empty and "FAB_LOT" in vid_to_wla_df.columns:
        for value in vid_to_wla_df["FAB_LOT"].dropna().tolist():
            val_str = str(value).strip().upper()
            if val_str.startswith("W"):
                extracted_wla_lots.add(val_str)

    if "WLA lot" in cleaned_df.columns:
        for value in cleaned_df["WLA lot"].dropna().tolist():
            val_str = str(value).strip().upper()
            if val_str.startswith("W"):
                extracted_wla_lots.add(val_str)

    wla_lots = sorted(extracted_wla_lots)
    wla_lot_patterns = sorted(filter(None, {_normalize_wla_lot7_pattern(value) for value in wla_lots}))

    final_workbook_path = Path(final_workbook_path or DEFAULT_FINAL_WORKBOOK)
    true_vid_df = (
        filtered_df.loc[filtered_df["IS_TRUE_VID"], [visual_id_col]]
        .rename(columns={visual_id_col: "VISUAL_ID"})
        .dropna(subset=["VISUAL_ID"])
        .assign(VISUAL_ID=lambda frame: frame["VISUAL_ID"].map(_normalize_visual_id))
        .dropna(subset=["VISUAL_ID"])
        .loc[lambda frame: frame["VISUAL_ID"] != ""]
        .drop_duplicates()
        .sort_values("VISUAL_ID")
        .reset_index(drop=True)
    )
    lot_patterns_df = pd.DataFrame({"WLA_LOT7_PATTERN": wla_lot_patterns})

    output_base = Path(output_dir or DEFAULT_OUTPUT_DIR)
    output_base.mkdir(parents=True, exist_ok=True)
    final_workbook_path.parent.mkdir(parents=True, exist_ok=True)

    safe_token = selected_tokens[0] if len(selected_tokens) == 1 else f"multi_{len(selected_tokens)}_wafers"
    filtered_report_path = output_base / f"filtered_catts_raw_{safe_token}.xlsx"
    summary_path = output_base / f"filtered_catts_summary_{safe_token}.json"

    cleaned_df.to_excel(filtered_report_path, index=False)
    _replace_excel_sheet(final_workbook_path, CATTS_CLEANED_SHEET, cleaned_df)
    _replace_excel_sheet(final_workbook_path, VID_SHEET, true_vid_df)
    _replace_excel_sheet(final_workbook_path, WLA_LOTS_SHEET, lot_patterns_df)
    _replace_excel_sheet(final_workbook_path, VID_TO_IOE_WLA_SHEET, vid_to_ioe_wla_df)
    _replace_excel_sheet(final_workbook_path, VID_TO_WLA_SHEET, vid_to_wla_df)
    _apply_short_wid_excel_format(final_workbook_path, CATTS_CLEANED_SHEET, ["FAB_WAFER", "WLA wafer"])
    _apply_short_wid_excel_format(final_workbook_path, VID_TO_IOE_WLA_SHEET, ["FAB_WAFER"])
    _apply_short_wid_excel_format(final_workbook_path, VID_TO_WLA_SHEET, ["FAB_WAFER"])

    summary = {
        "lot7": str(lot7).upper() if lot7 is not None else None,
        "wid3": str(wid3).zfill(3) if wid3 is not None else None,
        "lot7_wid3": safe_token,
        "lot7_wid3_values": selected_tokens,
        "selected_wafer_count": len(selected_tokens),
        "source_file": str(raw_path),
        "final_workbook": str(final_workbook_path),
        "vid_sheet": VID_SHEET,
        "filtered_report": str(filtered_report_path),
        "total_input_rows": int(len(canonical_df)),
        "filtered_rows": int(len(cleaned_df)),
        "filtered_rows_by_wafer": {
            lot7_wid3_value: int(count)
            for lot7_wid3_value, count in filtered_df["Lot7_WID3"].value_counts().sort_index().items()
        },
        "fake_vid_rows": int(filtered_df["IS_FAKE_VID"].sum()),
        "true_vid_rows": int(filtered_df["IS_TRUE_VID"].sum()),
        "wla_lot_columns": lot_columns,
        "wla_lots": wla_lots,
        "wla_lot7_patterns": wla_lot_patterns,
        "wla_lot7_count": len(wla_lot_patterns),
        "vid_to_ioe_wla_sheet": VID_TO_IOE_WLA_SHEET,
        "vid_to_wla_sheet": VID_TO_WLA_SHEET,
        "vid_to_ioe_wla_rows": int(len(vid_to_ioe_wla_df)),
        "vid_to_wla_rows": int(len(vid_to_wla_df)),
        "wla_lookup_matches": int(cleaned_df["WLA die index"].notna().sum()),
        "visual_id_column": str(visual_id_col),
        "catts_cleaned_sheet": CATTS_CLEANED_SHEET,
        "vid_sheet_name": VID_SHEET,
        "wla_lots_sheet": WLA_LOTS_SHEET,
        "visual_id_excel_formatting": visual_id_excel_formatting,
        "visual_id_quality": visual_id_quality,
    }
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    summary["summary_file"] = str(summary_path)
    return cleaned_df, summary


def fetch_sdhmt_for_wla_lots(summary, config_path="config.yaml", query_path="query.sql", output_path=None):
    lot_patterns = summary.get("wla_lot7_patterns", [])
    if not lot_patterns:
        raise ValueError("No WLA lot7% patterns were derived from the CATTS-cleaned data.")

    config = read_yaml_config(config_path)
    datasources = config.get("data_source", [])
    if isinstance(datasources, str):
        datasources = [datasources]

    query, filter_condition = _load_query_with_filters(query_path, lot_patterns, product="%8P%", days=600)

    all_results = []
    datasource_errors = []
    for datasource in datasources:
        try:
            results = execute_query(query, datasource)
            all_results.extend(results)
            print(f"Fetched {len(results)} records from {datasource}.")
        except Exception as exc:
            datasource_errors.append(f"{datasource}: {exc}")
            print(f"PyUber datasource failed: {datasource}: {exc}")

    if datasource_errors and not all_results:
        raise RuntimeError(
            "SDHMT query failed for all configured datasources. "
            "Cannot continue without data.\n"
            + "\n".join(datasource_errors)
        )
    elif datasource_errors:
        print(
            "Warning: SDHMT query failed for one or more datasources, "
            "but continuing with available data from successful datasources:\n"
            + "\n".join(datasource_errors)
        )

    sdhmt_df = wrangle_results(all_results)
    required_query_columns = [
        "LOT",
        "WAFER_ID",
        "WAFER_SCRIBE",
        "SORT_WAFER_ID",
        "SORT_X",
        "SORT_Y",
        "INTERFACE_BIN",
        "FUNCTIONAL_BIN",
    ]
    missing_query_columns = [column for column in required_query_columns if column not in sdhmt_df.columns]
    if missing_query_columns:
        raise ValueError(
            "SDHMT query results are missing required columns: " + ", ".join(missing_query_columns)
        )

    sdhmt_df["WAFER_ID"] = _normalize_short_wid_series(sdhmt_df["WAFER_ID"], "WAFER_ID")
    sdhmt_df["SORT_WAFER_ID"] = _normalize_short_wid_series(sdhmt_df["SORT_WAFER_ID"], "SORT_WAFER_ID")
    _validate_long_wafer_id_series(sdhmt_df["WAFER_SCRIBE"], "WAFER_SCRIBE")

    sdhmt_df["WLA die index"] = sdhmt_df.apply(
        lambda row: _compose_wla_die_index(row["WAFER_SCRIBE"], row["SORT_X"], row["SORT_Y"]),
        axis=1,
    )
    sdhmt_df = sdhmt_df.drop_duplicates().reset_index(drop=True)

    sdhmt_export_df = _align_columns(sdhmt_df, WLA_REQUIRED_COLUMNS)

    final_workbook_path = Path(summary["final_workbook"])
    cleaned_df = pd.read_excel(final_workbook_path, sheet_name=CATTS_CLEANED_SHEET)
    lookup_df = (
        sdhmt_export_df[["WLA die index", "INTERFACE_BIN", "FUNCTIONAL_BIN"]]
        .dropna(subset=["WLA die index"])
        .drop_duplicates(subset=["WLA die index"], keep="first")
        .rename(columns={
            "INTERFACE_BIN": "SDHMT IB",
            "FUNCTIONAL_BIN": "SDHMT FB",
        })
    )

    merged_df = cleaned_df.drop(columns=[column for column in ["SDHMT IB", "SDHMT FB"] if column in cleaned_df.columns])
    merged_df["WLA die index"] = merged_df["WLA die index"].map(_normalize_visual_id)
    merged_df = merged_df.merge(lookup_df, how="left", on="WLA die index")
    if "WLA die index" in merged_df.columns:
        merged_df["WLA die index"] = merged_df["WLA die index"].map(_normalize_visual_id)

    ordered_columns = CATTS_ENRICHED_COLUMNS + ["SDHMT IB", "SDHMT FB"]
    merged_df = _align_columns(merged_df, ordered_columns)

    _replace_excel_sheet(final_workbook_path, CATTS_CLEANED_SHEET, merged_df)
    _replace_excel_sheet(final_workbook_path, WLA_SDHMT_SHEET, sdhmt_export_df)
    _apply_short_wid_excel_format(final_workbook_path, CATTS_CLEANED_SHEET, ["FAB_WAFER"])
    _apply_short_wid_excel_format(final_workbook_path, WLA_SDHMT_SHEET, ["WAFER_ID", "SORT_WAFER_ID"])

    step4_summary = {
        "config_path": str(config_path),
        "query_path": str(query_path),
        "filter_condition": filter_condition,
        "final_workbook": str(final_workbook_path),
        "sdhmt_rows": int(len(sdhmt_export_df)),
        "matched_wla_die_indexes": int(merged_df["SDHMT IB"].notna().sum()),
        "wla_sdhmt_sheet": WLA_SDHMT_SHEET,
    }
    step4_summary["matched_visual_ids"] = step4_summary["matched_wla_die_indexes"]
    return merged_df, sdhmt_export_df, step4_summary


def consolidate_final_workbooks(input_paths, output_path):
    paths = [Path(path) for path in input_paths]
    output = Path(output_path)
    if not paths:
        raise ValueError('At least one finalized workbook is required.')
    if output.resolve() in {path.resolve() for path in paths}:
        raise ValueError('Consolidated output must not overwrite an input workbook.')
    if output.exists():
        raise FileExistsError(f'Consolidated output already exists: {output}')
    sheets = {}
    source_rows = []
    for path in paths:
        with pd.ExcelFile(path) as source:
            if CATTS_CLEANED_SHEET not in source.sheet_names:
                raise ValueError(f'{path} does not contain {CATTS_CLEANED_SHEET}.')
            for name in source.sheet_names:
                if name == 'Result Summary':
                    continue
                frame = pd.read_excel(source, sheet_name=name)
                sheets.setdefault(name, []).append(frame)
                if name == CATTS_CLEANED_SHEET:
                    source_rows.append({'Source Workbook': path.name, 'Die Rows': len(frame)})
    cleaned = pd.concat(sheets[CATTS_CLEANED_SHEET], ignore_index=True)
    wafer_keys = cleaned['FAB_LOT'].astype(str).str.strip().str.upper().str[:7] + '_' + _normalize_short_wid_series(cleaned['FAB_WAFER'], 'FAB_WAFER')
    def die_key(row):
        wafer = wafer_keys.loc[row.name]
        if pd.notna(row.get('X_LOCATION')) and pd.notna(row.get('Y_LOCATION')):
            return (wafer, 'xy', int(row['X_LOCATION']), int(row['Y_LOCATION']))
        for column in ('INTEL_DIE_INDEX', 'VISUAL_ID'):
            value = row.get(column)
            if pd.notna(value) and str(value).strip():
                return (wafer, column, str(value).strip())
        return (wafer, 'row', row.name)
    keys = cleaned.apply(die_key, axis=1)
    cleaned = cleaned.loc[~keys.duplicated(keep='last')].copy()
    sheets[CATTS_CLEANED_SHEET] = [cleaned]
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook(write_only=True)
    temporary_path = None
    try:
        for name, frames in sheets.items():
            merged = pd.concat(frames, ignore_index=True)
            if name == VID_SHEET and 'VISUAL_ID' in merged.columns:
                merged = merged.drop_duplicates(subset=['VISUAL_ID'], keep='last')
            sheet = workbook.create_sheet(name)
            sheet.append(list(merged.columns))
            for values in merged.itertuples(index=False, name=None):
                sheet.append([None if pd.isna(value) else value for value in values])
        source_name = 'Consolidation Sources'
        while source_name in workbook.sheetnames:
            source_name += '_'
        sources = workbook.create_sheet(source_name)
        sources.append(['Source Workbook', 'Die Rows'])
        for row in source_rows:
            sources.append([row['Source Workbook'], row['Die Rows']])
        with NamedTemporaryFile(dir=output.parent, suffix='.xlsx', delete=False) as temporary:
            temporary_path = Path(temporary.name)
        workbook.save(temporary_path)
        workbook.close()
        os.replace(temporary_path, output)
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return {'output_workbook': str(output), 'source_workbooks': [str(path) for path in paths],
            'die_rows': len(cleaned), 'duplicate_die_rows_removed': len(keys) - len(cleaned),
            'wafer_count': int(wafer_keys.loc[cleaned.index].nunique())}


def select_latest_aqua_class_rows(aqua_df):
    if aqua_df.empty:
        return aqua_df.copy()
    flags = AQUA_CLASS_SELECTION_FLAGS
    missing = [column for column in flags if column not in aqua_df.columns]
    if missing:
        raise ValueError(
            'AQUA Class results lack authoritative validity/latest flags. '
            f'Refetch Class data with these columns: {", ".join(missing)}'
        )
    selected = aqua_df.copy()
    if 'Operation' in selected.columns:
        selected = selected.loc[pd.to_numeric(selected['Operation'], errors='coerce').eq(7820)].copy()
    for column in flags:
        affirmative = selected[column].astype(str).str.strip().str.upper().isin(['Y', 'YES', 'TRUE', '1', '1.0'])
        selected = selected.loc[affirmative].copy()
    selected = selected.dropna(subset=['Visual_ID'])
    selected['Visual_ID'] = selected['Visual_ID'].map(_normalize_visual_id)
    selected = selected.loc[selected['Visual_ID'].ne('')].copy()
    bin_columns = ['Unit_Interface_Bin', 'Unit_Functional_Bin']
    bin_values = selected[bin_columns].apply(pd.to_numeric, errors='raise')
    conflicts = bin_values.groupby(selected['Visual_ID']).nunique(dropna=False).gt(1).any(axis=1)
    if conflicts.any():
        raise ValueError(
            f'AQUA returned conflicting final-latest Class bins for {int(conflicts.sum())} VID(s). '
            'Resolve the conflicting authoritative records before reporting.'
        )
    return selected.drop_duplicates(subset=['Visual_ID'], keep='first')


def apply_aqua_class_bins_to_final_workbook(final_workbook_path, aqua_workbook_path=None, aqua_df=None, aqua_sheet_name=AQUA_CLASS_SHEET):
    final_workbook = Path(final_workbook_path)
    if not final_workbook.exists():
        raise FileNotFoundError(f"Final workbook not found: {final_workbook}")

    cleaned_df = pd.read_excel(final_workbook, sheet_name=CATTS_CLEANED_SHEET)
    if aqua_df is None:
        aqua_workbook = Path(aqua_workbook_path)
        if not aqua_workbook.exists():
            raise FileNotFoundError(f"AQUA workbook not found: {aqua_workbook}")
        sheet_candidates = [aqua_sheet_name]
        if "class_7820" not in sheet_candidates:
            sheet_candidates.append("class_7820")

        aqua_df = None
        for sheet_name in sheet_candidates:
            try:
                aqua_df = pd.read_excel(aqua_workbook, sheet_name=sheet_name)
                break
            except ValueError:
                continue

        if aqua_df is None:
            raise ValueError(
                f"AQUA workbook {aqua_workbook} does not contain any expected sheet: {sheet_candidates}"
            )
    else:
        aqua_workbook = None

    required_aqua_columns = ["Visual_ID", "Unit_Interface_Bin", "Unit_Functional_Bin"]
    missing_aqua_columns = [column for column in required_aqua_columns if column not in aqua_df.columns]
    if missing_aqua_columns:
        raise ValueError(
            "AQUA workbook is missing required columns: " + ", ".join(missing_aqua_columns)
        )

    lookup_df = (
        select_latest_aqua_class_rows(aqua_df)[required_aqua_columns]
        .dropna(subset=["Visual_ID"])
        .rename(
            columns={
                "Unit_Interface_Bin": "Class IB",
                "Unit_Functional_Bin": "Class FB",
            }
        )
    )

    merged_df = cleaned_df.drop(
        columns=[column for column in ["Class IB", "Class FB"] if column in cleaned_df.columns]
    )
    merged_df = merged_df.merge(lookup_df, how="left", left_on="VISUAL_ID", right_on="Visual_ID")
    if "Visual_ID" in merged_df.columns:
        merged_df.drop(columns=["Visual_ID"], inplace=True)

    ordered_columns = CATTS_ENRICHED_COLUMNS + ["SDHMT IB", "SDHMT FB", "Class IB", "Class FB"]
    merged_df = _align_columns(merged_df, ordered_columns)

    _replace_excel_sheet(final_workbook, aqua_sheet_name, aqua_df)
    _replace_excel_sheet(final_workbook, CATTS_CLEANED_SHEET, merged_df)
    _apply_short_wid_excel_format(final_workbook, CATTS_CLEANED_SHEET, ["FAB_WAFER"])

    summary = {
        "final_workbook": str(final_workbook),
        "aqua_sheet": aqua_sheet_name,
        "matched_visual_ids": int(merged_df["Class IB"].notna().sum()),
        "catts_cleaned_sheet": CATTS_CLEANED_SHEET,
    }
    return merged_df, summary