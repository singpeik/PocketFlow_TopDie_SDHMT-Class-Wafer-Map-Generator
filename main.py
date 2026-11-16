from flow import create_ioe_flow
import argparse
from utils.aqua_client import DEFAULT_AQUA_CREDENTIALS_PATH, DEFAULT_AQUA_OUTPUT_PATH
from utils.ioe_workflow import (
    add_date_stamp_to_path,
    DEFAULT_ALL_PRODUCT_MAPS_OUTPUT,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_FINAL_WORKBOOK,
    DEFAULT_IOE_WAFER_MAP_OUTPUT,
    DEFAULT_IOE_WAFER_MAP_SOURCE,
    DEFAULT_GPU_WAFER_MAP_SOURCE,
    DEFAULT_GPU_WAFER_MAP_OUTPUT,
    DEFAULT_VID_OUTPUT,
    DEFAULT_WLA_SDHMT_OUTPUT,
)


def _show_step_popup(success, title, message):
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
        print(f"Unable to show workflow popup: {exc}")

def main():
    parser = argparse.ArgumentParser(description="Wafer Test Result Analyzer - SDHMT/Class (IOE/GPU)")
    parser.add_argument(
        "--ioe-map-source",
        default=str(DEFAULT_IOE_WAFER_MAP_SOURCE),
        help="Path to the IOE wafer map Excel source file used to build the master IOE map",
    )
    parser.add_argument(
        "--ioe-map-output",
        default=str(DEFAULT_IOE_WAFER_MAP_OUTPUT),
        help="Path to the generated IOE wafer map JSON file",
    )
    parser.add_argument(
        "--gpu-map-source",
        default=str(DEFAULT_GPU_WAFER_MAP_SOURCE),
        help="Path to the GPU reference wafer map Excel file",
    )
    parser.add_argument(
        "--gpu-map-output",
        default=str(DEFAULT_GPU_WAFER_MAP_OUTPUT),
        help="Path to the GPU wafer geometry JSON file",
    )
    parser.add_argument(
        "--all-maps-output",
        default=str(DEFAULT_ALL_PRODUCT_MAPS_OUTPUT),
        help="Path to the wafer map library JSON file",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Directory for filtered CATTS outputs",
    )
    parser.add_argument(
        "--config-path",
        default="config.yaml",
        help="Path to the XEUS datasource configuration for Step 4",
    )
    parser.add_argument(
        "--query-path",
        default="query.sql",
        help="Path to the SQL file used for Step 4 SDHMT fetch",
    )
    parser.add_argument(
        "--final-workbook",
        default=str(DEFAULT_FINAL_WORKBOOK),
        help="Path to the workflow workbook, defaulting to outputs/Final.xlsx",
    )
    parser.add_argument(
        "--vid-output",
        default=str(DEFAULT_VID_OUTPUT),
        help="Legacy standalone VID.csv path; the workflow now writes the VID tab into outputs/Final.xlsx",
    )
    parser.add_argument(
        "--wla-sdhmt-output",
        default=str(DEFAULT_WLA_SDHMT_OUTPUT),
        help="Legacy standalone Step 4 workbook path; the workflow now writes WLA lots SDHMT into outputs/Final.xlsx",
    )
    parser.add_argument(
        "--aqua-credentials-path",
        default=str(DEFAULT_AQUA_CREDENTIALS_PATH),
        help="Path to the local AQUA credentials YAML file used for Step 5",
    )
    parser.add_argument(
        "--aqua-output",
        default=str(DEFAULT_AQUA_OUTPUT_PATH),
        help="Legacy standalone Step 5 workbook path; the workflow now writes AQUA class_7820 into outputs/Final.xlsx",
    )
    parser.add_argument(
        "--class-vid-limit",
        type=int,
        default=None,
        help="Optional limit for the number of true VID values sent to Step 5",
    )
    parser.add_argument(
        "--mrb-ppt-output",
        default=None,
        help="Optional output path for the Step 6 wafer test result PPT",
    )
    parser.add_argument(
        "--stop-after-step",
        type=int,
        choices=[4, 5, 6],
        default=6,
        help="Optionally stop the wafer workflow after Step 4, 5, or 6",
    )
    args = parser.parse_args()

    shared = {
        "ioe_wafer_map": None,
        "selected_wafer": None,
        "filtered_catts_raw_df": None,
        "filtered_catts_raw_summary": None,
        "final_cleansed_df": None,
        "wla_sdhmt_df": None,
        "step4_summary": None,
        "class_fetch_summary": None,
        "step6_summary": None,
    }

    print("Starting Wafer Test Result Analyzer (IOE/GPU)...")
    flow = create_ioe_flow(stop_after_step=args.stop_after_step)

    params = {
        "ioe_map_source": args.ioe_map_source,
        "ioe_map_output": args.ioe_map_output,
        "gpu_map_source": args.gpu_map_source,
        "gpu_map_output": args.gpu_map_output,
        "all_maps_output": args.all_maps_output,
        "output_dir": args.output_dir,
        "raw_report_initial_dir": str(DEFAULT_IOE_WAFER_MAP_SOURCE.parent),
        "config_path": args.config_path,
        "query_path": args.query_path,
        "final_workbook": str(add_date_stamp_to_path(args.final_workbook)),
        "vid_output": args.vid_output,
        "wla_sdhmt_output": args.wla_sdhmt_output,
        "aqua_credentials_path": args.aqua_credentials_path,
        "aqua_output": str(add_date_stamp_to_path(args.aqua_output)),
        "class_vid_limit": args.class_vid_limit,
        "mrb_ppt_output": args.mrb_ppt_output,
        "stop_after_step": args.stop_after_step,
    }

    flow.set_params(params)
    try:
        flow.run(shared)
    except Exception as exc:
        if not shared.get("step4_summary"):
            _show_step_popup(
                success=False,
                title="Step 1-4 Failed",
                message=(
                    "Step 1-4 failed before completion.\n"
                    "Please check terminal logs for details.\n\n"
                    f"Error: {exc}"
                ),
            )
        else:
            _show_step_popup(
                success=False,
                title="Workflow Failed",
                message=(
                    "Workflow failed after Step 1-4 completion.\n"
                    "Please check terminal logs for details.\n\n"
                    f"Error: {exc}"
                ),
            )
        raise

    summary = shared.get("filtered_catts_raw_summary")
    step4_summary = shared.get("step4_summary")
    class_summary = shared.get("class_fetch_summary")
    step6_summary = shared.get("step6_summary")
    if summary and step4_summary:
        class_status = class_summary.get("status") if isinstance(class_summary, dict) else "not-run"
        class_step_succeeded = class_status in {"completed", "skipped_no_true_vid"}
        mrb_output = step6_summary.get("output_ppt") if isinstance(step6_summary, dict) else "not-generated"
        wafer_label = ", ".join(summary.get("lot7_wid3_values", []))
        stop_suffix = ""
        if args.stop_after_step == 4:
            stop_suffix = " Workflow stopped after Step 4 by request."
        elif args.stop_after_step == 5:
            stop_suffix = " Workflow stopped after Step 5 by request."
        elif not class_step_succeeded:
            stop_suffix = " Workflow stopped after Step 5 failure; Step 6 skipped."

        completion_prefix = "Workflow finished successfully"
        if not class_step_succeeded and args.stop_after_step == 6:
            completion_prefix = "Workflow finished with Step 5 failure"

        print(
            f"{completion_prefix} for {wafer_label}. "
            f"Final workbook saved to {summary['final_workbook']}, "
            f"SDHMT sheet: {step4_summary['wla_sdhmt_sheet']}, "
            f"AQUA Class fetch status: {class_status}, "
            f"Test result PPT: {mrb_output}."
            f"{stop_suffix}"
        )

        return shared.get("final_cleansed_df")

    print("Workflow finished without producing a filtered CATTS raw report.")
    if args.stop_after_step == 4:
        _show_step_popup(
            success=False,
            title="Step 1-4 Failed",
            message=(
                "Step 1-4 did not complete successfully.\n"
                "Required Step 1-4 outputs were not produced.\n"
                "Please review the terminal logs and retry."
            ),
        )
    return None

if __name__ == "__main__":
    main()
