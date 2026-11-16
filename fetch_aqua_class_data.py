import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

from utils.aqua_client import (
    DEFAULT_AQUA_CREDENTIALS_PATH,
    DEFAULT_AQUA_OUTPUT_PATH,
    fetch_aqua_class_7820_from_workbook_vid_sheet,
    fetch_aqua_class_7820_from_vid_csv,
    load_cached_aqua_result_workbook,
    save_aqua_result_workbook,
)
from utils.ioe_workflow import AQUA_CLASS_SHEET
from utils.ioe_workflow import add_date_stamp_to_path, DEFAULT_FINAL_WORKBOOK, DEFAULT_VID_OUTPUT, VID_SHEET


STEP5_LOCK_FILENAME = "step5_aqua_fetch.lock"


def _is_pid_running(pid):
    if not pid or pid <= 0:
        return False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            capture_output=True,
            text=True,
            check=False,
        )
        return str(pid) in result.stdout
    except Exception:
        return False


def _acquire_step5_lock(lock_path):
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY

    try:
        fd = os.open(str(lock_path), flags)
    except FileExistsError:
        existing_pid = None
        try:
            existing_text = lock_path.read_text(encoding="utf-8").strip()
            if existing_text:
                existing_pid = int(existing_text.splitlines()[0].strip())
        except Exception:
            existing_pid = None

        if existing_pid and _is_pid_running(existing_pid):
            raise RuntimeError(
                f"Step5 already running under PID {existing_pid}. "
                "Please wait for it to complete before starting another run."
            )

        lock_path.unlink(missing_ok=True)
        fd = os.open(str(lock_path), flags)

    lock_body = f"{os.getpid()}\n{int(time.time())}\n"
    with os.fdopen(fd, "w", encoding="utf-8") as handle:
        handle.write(lock_body)


def _release_step5_lock(lock_path):
    lock_path.unlink(missing_ok=True)


def _cleanup_stale_aqua_temp_dirs(output_path):
    output_dir = Path(output_path).resolve().parent
    removed = []
    for temp_dir in output_dir.glob("aqua_class_*"):
        if temp_dir.is_dir():
            shutil.rmtree(temp_dir, ignore_errors=True)
            removed.append(str(temp_dir))
    return removed


def _show_step5_popup(success, title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if success:
            messagebox.showinfo(title, message, parent=root)
        else:
            messagebox.showerror(title, message, parent=root)
        root.destroy()
    except Exception as exc:
        # Keep a console fallback for headless environments.
        print(f"Unable to show Step5 popup: {exc}")


def main():
    parser = argparse.ArgumentParser(
        description="Fetch AQUA Class 7820 data from either outputs/Final.xlsx VID sheet or an explicit VID.csv"
    )
    parser.add_argument(
        "--vid-csv",
        default=None,
        help="Optional path to a VID.csv file containing true VID values; if omitted, the script uses outputs/Final.xlsx",
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_FINAL_WORKBOOK),
        help="Path to the workflow workbook containing the VID sheet",
    )
    parser.add_argument(
        "--vid-sheet",
        default=VID_SHEET,
        help="Workbook sheet name containing the true VID list",
    )
    parser.add_argument(
        "--credentials",
        default=str(DEFAULT_AQUA_CREDENTIALS_PATH),
        help="Path to the AQUA API Key YAML file",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_AQUA_OUTPUT_PATH),
        help="Path to the final AQUA Class workbook",
    )
    parser.add_argument(
        "--vid-limit",
        type=int,
        default=None,
        help="Optional limit for the number of true VIDs sent to AQUA",
    )
    args = parser.parse_args()
    args.output = str(add_date_stamp_to_path(args.output))

    lock_path = Path(args.output).resolve().parent / STEP5_LOCK_FILENAME
    try:
        _acquire_step5_lock(lock_path)
    except RuntimeError as exc:
        lock_error = str(exc)
        print(lock_error)
        _show_step5_popup(
            success=False,
            title="Step 5 AQUA Fetch Failed",
            message=lock_error,
        )
        raise SystemExit(2) from exc

    removed_temp_dirs = _cleanup_stale_aqua_temp_dirs(args.output)
    if removed_temp_dirs:
        print(f"Removed {len(removed_temp_dirs)} stale Step5 temp folder(s) before fetch.")

    try:
        vid_csv_path = Path(args.vid_csv) if args.vid_csv else None
        workbook_path = Path(args.workbook)

        if vid_csv_path is not None:
            result = fetch_aqua_class_7820_from_vid_csv(
                vid_csv_path=vid_csv_path,
                credentials_path=args.credentials,
                output_path=args.output,
                vid_limit=args.vid_limit,
            )
            result["source"] = "aqua"
            source_desc = str(vid_csv_path)
        else:
            if workbook_path.exists():
                result = load_cached_aqua_result_workbook(
                    final_workbook_path=workbook_path,
                    cache_workbook_path=args.output,
                    vid_sheet_name=args.vid_sheet,
                    cache_sheet_name=AQUA_CLASS_SHEET,
                    vid_limit=args.vid_limit,
                )
                if result is None:
                    result = fetch_aqua_class_7820_from_workbook_vid_sheet(
                        final_workbook_path=workbook_path,
                        vid_sheet_name=args.vid_sheet,
                        credentials_path=args.credentials,
                        vid_limit=args.vid_limit,
                    )
                    result["source"] = "aqua"
                else:
                    result["source"] = "cache"

                source_desc = f"{workbook_path} [{args.vid_sheet}]"
            else:
                raise FileNotFoundError(
                    f"Neither VID.csv nor workbook source is available. Looked for workbook: {workbook_path}"
                )

        result["output_path"] = save_aqua_result_workbook(
            result["result_df"],
            output_path=args.output,
            sheet_name=AQUA_CLASS_SHEET,
        )
        output_desc = result["output_path"]

        success_message = (
            f"AQUA Class 7820 fetch completed for {result['vid_count']} VIDs. "
            f"Rows written: {result['row_count']}. "
            f"Source type: {result.get('source', 'aqua')}. "
            f"Source: {source_desc}. "
            f"Workbook: {output_desc}"
        )
        print(success_message)
        _show_step5_popup(
            success=True,
            title="Step 5 AQUA Fetch Complete",
            message=success_message,
        )
    except Exception as exc:
        failure_message = f"Step 5 AQUA fetch failed: {exc}"
        print(failure_message)
        _show_step5_popup(
            success=False,
            title="Step 5 AQUA Fetch Failed",
            message=failure_message,
        )
        raise
    finally:
        _release_step5_lock(lock_path)


if __name__ == "__main__":
    main()