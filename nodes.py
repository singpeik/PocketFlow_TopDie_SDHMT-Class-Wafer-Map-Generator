from pocketflow import Node
from utils.file_io import read_yaml_config, save_results
from utils.pyuber_client import execute_query
from utils.data_wrangling import wrangle_results
from utils.data_processing import analyze_data
from utils.reporting import generate_ioe_mrb_summary_ppt, generate_lot_level_rollup, generate_wafer_maps_and_ppt
from utils.ioe_workflow import (
    AQUA_CLASS_SHEET,
    VID_SHEET,
    add_date_stamp_to_path,
    apply_aqua_class_bins_to_final_workbook,
    build_ioe_wafer_map,
    cleanup_ioe_intermediate_files,
    collect_ioe_user_inputs,
    parse_catts_raw_report,
    fetch_sdhmt_for_wla_lots,
)
from utils.aqua_client import (
    DEFAULT_AQUA_CREDENTIALS_PATH,
    DEFAULT_AQUA_OUTPUT_PATH,
    fetch_aqua_class_7820_from_vid_csv,
    fetch_aqua_class_7820_from_workbook_vid_sheet,
    load_cached_aqua_result_workbook,
    save_aqua_result_workbook,
)
from generate_reject_summary import generate_reject_ppt
import os
import pandas as pd

class FetchDataNode(Node):
    def prep(self, shared):
        # Get config path from params or default to 'config.yaml'
        config_path = self.params.get("config_path", "config.yaml")
        # Get filters
        filters = self.params.get("filters", {})
        
        if "filters" not in self.params:
             if self.params.get("filter_mode") == "lot":
                  filters = {
                       "lots": self.params.get("filter_data"), 
                       "product": "%8P%", 
                       "time": {"mode": "days", "days": 600}
                  }
        
        return config_path, filters

    def exec(self, prep_res):
        config_path, filters = prep_res
        print(f"Loading configuration from {config_path}...")
        config = read_yaml_config(config_path)

        if "data_source" not in config:
             raise ValueError("Missing required config key: data_source")
        
        if "output_path" not in config:
            config["output_path"] = "SDHMT die level data.csv"
        if "output_format" not in config:
            config["output_format"] = "csv"
            
        query = ""
        if "sql_file" in config:
            sql_file_path = config["sql_file"]
            if not os.path.exists(sql_file_path):
                 raise FileNotFoundError(f"SQL file specified in config not found: {sql_file_path}")
            with open(sql_file_path, "r") as f:
                query = f.read()
            print(f"Loaded SQL query from file: {sql_file_path}")
        elif "sql_query" in config:
            query = config["sql_query"]
        else:
             raise ValueError("Config must contain either 'sql_query' or 'sql_file'.")
        
        lot_condition = "1=1"
        lots = filters.get("lots", ["W%"])
        if isinstance(lots, str): lots = [lots]
        
        processed_lots = []
        for lot in lots:
            lot = lot.strip()
            if not lot: continue
            if '%' in lot:
                processed_lots.append(f"v0.lot LIKE '{lot}'")
            elif len(lot) in [7, 8]:
                processed_lots.append(f"v0.lot LIKE '{lot}%'")
            else:
                processed_lots.append(f"v0.lot LIKE '{lot}'")
        
        if processed_lots:
            lot_condition = "(" + " OR ".join(processed_lots) + ")"
        
        prod_condition = "1=1"
        prods = filters.get("products") 
        if not prods:
             single_prod = filters.get("product", "%8P%")
             if single_prod:
                 prods = [single_prod]
             else:
                 prods = ["%8P%"] 

        processed_prods = []
        for prod in prods:
            prod = prod.strip()
            if not prod: continue
            if '%' in prod:
                processed_prods.append(f"v0.devrevstep LIKE '{prod}'")
            else:
                processed_prods.append(f"v0.devrevstep = '{prod}'")
        
        if processed_prods:
            prod_condition = "(" + " OR ".join(processed_prods) + ")"

        time_condition = "1=1"
        time_cfg = filters.get("time", {})
        t_mode = time_cfg.get("mode", "days")
        
        clauses = []
        if lot_condition and lot_condition != "1=1": clauses.append(lot_condition)
        if prod_condition and prod_condition != "1=1": clauses.append(prod_condition)
        
        if t_mode == "days":
             days = time_cfg.get("days", 600)
             clauses.append(f"v0.test_end_date_time >= TRUNC(SYSDATE) - {days}")
        elif t_mode == "range":
             start = time_cfg.get("start")
             end = time_cfg.get("end")
             if start and end:
                 clauses.append(f"v0.test_end_date_time BETWEEN TO_DATE('{start}', 'MM-DD-YYYY') AND TO_DATE('{end}', 'MM-DD-YYYY') + 1")
        
        if not clauses:
             full_condition = "1=1"
        else:
             full_condition = " AND ".join(clauses)
             
        if "{FILTER_CONDITION}" in query:
            query = query.replace("{FILTER_CONDITION}", full_condition)
            print("Injected Combined Filters")
        elif "{LOT_ID}" in query:
             print("Warning: Using legacy {LOT_ID} replacement.")
             pass
             
        print(f"DEBUG: Full Query to execute:\n{query[:100]}...")
        
        datasource = config["data_source"]
        
        all_results = []
        if isinstance(datasource, list):
            for ds in datasource:
                print(f"Querying datasource: {ds}")
                results = execute_query(query, ds)
                all_results.extend(results)
        else:
            print(f"Querying datasource: {datasource}")
            all_results = execute_query(query, datasource)
            
        df = wrangle_results(all_results)
        
        return df, config, filters

    def post(self, shared, prep_res, exec_res):
        df, config, filters = exec_res
        shared["df"] = df
        shared["config"] = config
        shared["filters"] = filters
        return "default"

class SaveOutputNode(Node):
    def prep(self, shared):
        config = shared.get("config")
        df = shared.get("df")
        return df, config["output_path"], config["output_format"]

    def exec(self, inputs):
        df, output_path, output_format = inputs
        save_results(df, output_path, output_format)
        return output_path

    def post(self, shared, prep_res, exec_res):
        print(f"Flow completed. Output saved to {exec_res}")
        return "default"

class WaferMapCheckNode(Node):
    def prep(self, shared):
        return shared.get("df")

    def exec(self, df):
        if df is None or df.empty: return True
        
        map_json_path = "domain-knowledge/wafer-map/all_product_maps.json"
        try:
            import json
            import tkinter as tk
            from tkinter import filedialog, messagebox

            with open(map_json_path, 'r', encoding='utf-8') as f:
                maps = json.load(f)
            
            if 'DEVREVSTEP' not in df.columns:
                 return True
            
            bom_codes = df['DEVREVSTEP'].dropna().unique()
            
            updated = False
            for bom in bom_codes:
                bom_str = str(bom)
                # Ignore empty strings or missing mappings
                if bom_str.strip() and bom_str not in maps:
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    
                    msg = (f"No wafer map found in library for product BOM: {bom_str}.\n\n"
                           "Please upload the example-map file for this specific BOM code product.\n"
                           "Instructions to download:\n"
                           "1. Go to CATTS\n"
                           "2. Extended report\n"
                           "3. Die index reports\n"
                           "4. input lot8 to download example-map excel sheet.")
                    messagebox.showinfo("Missing Wafer Map", msg, parent=root)
                    
                    file_path = filedialog.askopenfilename(
                        parent=root,
                        title=f"Select example-map excel file for {bom_str}",
                        filetypes=[("Excel files", "*.xlsx *.xls")]
                    )
                    
                    if file_path:
                        try:
                            ex_df = pd.read_excel(file_path)
                            if 'X_LOCATION' in ex_df.columns and 'Y_LOCATION' in ex_df.columns:
                                coords = ex_df[['X_LOCATION', 'Y_LOCATION']].dropna().drop_duplicates()
                                active_dies = coords.astype(int).values.tolist()
                                
                                maps[bom_str] = {"active_dies": active_dies}
                                updated = True
                                messagebox.showinfo("Success", f"Map for {bom_str} added successfully with {len(active_dies)} dies!", parent=root)
                            else:
                                messagebox.showerror("Error", f"Uploaded file lacks X_LOCATION or Y_LOCATION.", parent=root)
                        except Exception as e:
                            messagebox.showerror("Error", f"Failed to read/process excel file:\n{e}", parent=root)
                    else:
                        messagebox.showwarning("Warning", "No file selected. Map generation for this product might result in empty backgrounds.", parent=root)
                    
                    root.destroy()
            
            if updated:
                 with open(map_json_path, 'w', encoding='utf-8') as f:
                      json.dump(maps, f, indent=4)
                      
        except Exception as e:
            print(f"Error checking wafer maps: {e}")
        
        return True

    def post(self, shared, prep_res, exec_res):
        return "default"

class DataProcessingNode(Node):
    def prep(self, shared):
        return shared.get("df")

    def exec(self, df):
        if df is None or df.empty:
            print("No data to process for pivot table.")
            return None

        print("DataProcessingNode: Starting analysis...")
        try:
            result = analyze_data(df)

            if result is not None:
                output_file = "wafer level rollup.xlsx"

                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    if isinstance(result, dict):
                        for sheet_name, sheet_df in result.items():        
                             sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        result.to_excel(writer, sheet_name="Analysis", index=False)

                print(f"Analysis results saved to {output_file}")
                return result
            else:
                 print("Analysis returned no results.")
                 return None

        except Exception as e:
            print(f"Error during data processing: {e}")
            import traceback
            traceback.print_exc()
            return None

    def post(self, shared, prep_res, exec_res):
        if exec_res is not None:
             shared["pivot_df"] = exec_res
        return "default"

class LotLevelRollupNode(Node):
    def prep(self, shared):
        return shared.get("pivot_df")
    
    def exec(self, pivot_df):
        if pivot_df is None:
            print("No pivot_df found to run lot level rollup.")
            return None
        generate_lot_level_rollup("wafer level rollup.xlsx", "lot level rollup.xlsx")
        return "lot level rollup.xlsx"
    
    def post(self, shared, prep_res, exec_res):
        if exec_res:
            print(f"Lot level rollup completed. Output saved to {exec_res}")
        return "default"

class ReportGenerationNode(Node):
    def prep(self, shared):
        return shared.get("filters", {})
    
    def exec(self, filters):
        die_csv = "SDHMT die level data.csv"
        wafer_xlsx = "wafer level rollup.xlsx"
        lot_xlsx = "lot level rollup.xlsx"
        map_json = "domain-knowledge/wafer-map/all_product_maps.json"
        
        from utils.reporting import generate_wafer_maps_and_ppt
        generate_wafer_maps_and_ppt(die_csv, lot_xlsx, wafer_xlsx, map_json, "SDHMT_Report.pptx", filters=filters)
        
        return "SDHMT_Report.pptx"
    
    def post(self, shared, prep_res, exec_res):
        print(f"Report Generation completed. Output saved to {exec_res}")
        return "default"

class GenerateRejectSummaryNode(Node):
    def prep(self, shared):
        return None
    
    def exec(self, _):
        die_csv = "SDHMT die level data.csv"
        map_json = "domain-knowledge/wafer-map/all_product_maps.json"
        from generate_reject_summary import generate_reject_ppt
        generate_reject_ppt(die_csv, map_json, "SDHMT reject bin summary.pptx")
        return "SDHMT reject bin summary.pptx"
    
    def post(self, shared, prep_res, exec_res):
        print(f"Reject Binary Summary generated. Output saved to {exec_res}")
        return "default"


class GenerateIOEWaferMapNode(Node):
    def prep(self, shared):
        product = (shared.get("selected_wafer") or {}).get("product", "IOE")
        prefix = "gpu" if product == "GPU" else "ioe"
        print(f"[Step 2/6] Preparing {product} wafer geometry...")
        return {
            "source_path": self.params.get(f"{prefix}_map_source"),
            "output_path": self.params.get(f"{prefix}_map_output"),
            "all_maps_path": self.params.get("all_maps_output"),
            "map_name": product,
        }

    def exec(self, prep_res):
        result = build_ioe_wafer_map(
            source_path=prep_res.get("source_path"),
            output_path=prep_res.get("output_path"),
            all_maps_path=prep_res.get("all_maps_path"),
            map_name=prep_res["map_name"],
        )
        if result.get("build_mode") == "reused_json":
            print(
                f"Reused existing {prep_res['map_name']} wafer map JSON at {prep_res.get('output_path')} with "
                f"{result['active_die_count']} active die coordinates across {result['wafer_count']} wafers."
            )
        else:
            print(
                f"Generated {prep_res['map_name']} wafer map from {result['source_file']} with "
                f"{result['active_die_count']} active die coordinates across {result['wafer_count']} wafers."
            )
        return result

    def post(self, shared, prep_res, exec_res):
        shared["ioe_wafer_map"] = exec_res
        shared["wafer_map_path"] = prep_res.get("output_path") or (
            "domain-knowledge/wafer-map/gpu_wafer_map.json" if prep_res["map_name"] == "GPU"
            else "domain-knowledge/wafer-map/ioe_wafer_map.json"
        )
        return "default"


class CollectIOEInputsNode(Node):
    def prep(self, shared):
        print("[Step 1/6] Waiting for product, wafer selection, and CATTS raw report input...")
        return self.params.get("raw_report_initial_dir")

    def exec(self, initial_dir):
        result = collect_ioe_user_inputs(initial_dir=initial_dir)
        print(
            f"Selected {result['wafer_count']} {result.get('product', 'IOE')} wafer(s): "
            f"{', '.join(result['lot7_wid3_values'])} using raw report {result['raw_report_path']}"
        )
        return result

    def post(self, shared, prep_res, exec_res):
        shared["selected_wafer"] = exec_res
        return "default"


class FilterCATTSRawReportNode(Node):
    def prep(self, shared):
        print("[Step 3/6] Filtering CATTS raw report...")
        selection = shared.get("selected_wafer") or {}
        return {
            "raw_report_path": selection.get("raw_report_path"),
            "lot7": selection.get("lot7"),
            "wid3": selection.get("wid3"),
            "lot7_wid3_values": selection.get("lot7_wid3_values"),
            "output_dir": self.params.get("output_dir"),
            "final_workbook": self.params.get("final_workbook"),
        }

    def exec(self, prep_res):
        filtered_df, summary = parse_catts_raw_report(
            raw_report_path=prep_res.get("raw_report_path"),
            lot7=prep_res.get("lot7"),
            wid3=prep_res.get("wid3"),
            output_dir=prep_res.get("output_dir"),
            lot7_wid3_values=prep_res.get("lot7_wid3_values"),
            final_workbook_path=prep_res.get("final_workbook"),
        )
        print(
            f"Filtered CATTS raw report down to {summary['filtered_rows']} rows for "
            f"{summary['selected_wafer_count']} wafer(s)."
        )
        return filtered_df, summary

    def post(self, shared, prep_res, exec_res):
        filtered_df, summary = exec_res
        shared["filtered_catts_raw_df"] = filtered_df
        shared["filtered_catts_raw_summary"] = summary
        cleanup_paths = cleanup_ioe_intermediate_files(summary)
        shared["ioe_filter_cleanup_paths"] = cleanup_paths

        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showinfo(
                "Wafer Raw Report Ready",
                (
                    f"Completed filtering for {summary['selected_wafer_count']} wafer(s).\n"
                    f"Selections: {', '.join(summary['lot7_wid3_values'])}\n\n"
                    f"Filtered rows: {summary['filtered_rows']}\n"
                    f"Fake VID rows: {summary['fake_vid_rows']}\n"
                    f"True VID rows: {summary['true_vid_rows']}\n\n"
                    f"VID tab: {summary['vid_sheet_name']}\n"
                    f"Wafer/WLA cross-reference tab: {summary['vid_to_ioe_wla_sheet']} ({summary['vid_to_ioe_wla_rows']} rows)\n"
                    f"VID to WLA tab: {summary['vid_to_wla_sheet']} ({summary['vid_to_wla_rows']} rows)\n"
                    f"CATTS rows with WLA lookup: {summary['wla_lookup_matches']}\n"
                    f"Final workbook: {summary['final_workbook']}\n"
                    "Temporary filtered CATTS extract files were cleaned up automatically."
                ),
                parent=root,
            )
            root.destroy()
        except Exception as exc:
            print(f"Unable to show completion popup: {exc}")

        if cleanup_paths:
            print("Removed Step 3 intermediate files:")
            for cleanup_path in cleanup_paths:
                print(f"- {cleanup_path}")

        return "default"


class FetchSDHMTDataNode(Node):
    def prep(self, shared):
        print("[Step 4/6] Fetching SDHMT data via PyUber...")
        return {
            "summary": shared.get("filtered_catts_raw_summary") or {},
            "config_path": self.params.get("config_path", "config.yaml"),
            "query_path": self.params.get("query_path", "query.sql"),
            "wla_sdhmt_output": self.params.get("wla_sdhmt_output"),
        }

    def exec(self, prep_res):
        cleaned_df, sdhmt_df, step4_summary = fetch_sdhmt_for_wla_lots(
            summary=prep_res["summary"],
            config_path=prep_res["config_path"],
            query_path=prep_res["query_path"],
            output_path=prep_res.get("wla_sdhmt_output"),
        )
        print(
            f"Fetched {step4_summary['sdhmt_rows']} SDHMT rows and matched "
            f"{step4_summary['matched_wla_die_indexes']} CATTS WLA die index values."
        )
        return cleaned_df, sdhmt_df, step4_summary

    def post(self, shared, prep_res, exec_res):
        cleaned_df, sdhmt_df, step4_summary = exec_res
        shared["final_cleansed_df"] = cleaned_df
        shared["wla_sdhmt_df"] = sdhmt_df
        shared["step4_summary"] = step4_summary
        filter_summary = shared.get("filtered_catts_raw_summary") or {}
        selected_wafers = ", ".join(filter_summary.get("lot7_wid3_values", []))
        selected_wla_lots = filter_summary.get("wla_lot7_count", 0)
        filtered_rows = filter_summary.get("filtered_rows", 0)
        true_vid_rows = filter_summary.get("true_vid_rows", 0)
        fake_vid_rows = filter_summary.get("fake_vid_rows", 0)

        next_step_msg = "Next: Step 5 AQUA Class fetch will start automatically."
        if self.params.get("stop_after_step") == 4:
            next_step_msg = "Run stopped at Step 4 by request."

        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showinfo(
                "Step 1-4 Passed",
                (
                    f"Selected wafer(s): {selected_wafers}\n"
                    f"Filtered CATTS rows: {filtered_rows}\n"
                    f"True VID rows: {true_vid_rows}\n"
                    f"Fake VID rows: {fake_vid_rows}\n"
                    f"Derived WLA lot7% count: {selected_wla_lots}\n"
                    f"Final workbook: {step4_summary['final_workbook']}\n"
                    f"Updated sheet: {step4_summary['wla_sdhmt_sheet']}\n"
                    f"SDHMT rows: {step4_summary['sdhmt_rows']}\n"
                    f"Matched CATTS WLA die index rows: {step4_summary['matched_wla_die_indexes']}\n\n"
                    f"{next_step_msg}"
                ),
                parent=root,
            )
            root.destroy()
        except Exception as exc:
            print(f"Unable to show SDHMT completion popup: {exc}")

        return "default"


class FetchClassDataNode(Node):
    def prep(self, shared):
        print("[Step 5/6] Fetching AQUA Class 7820 data...")
        summary = shared.get("filtered_catts_raw_summary") or {}
        return {
            "vid_sheet": summary.get("vid_sheet") or VID_SHEET,
            "final_workbook": summary.get("final_workbook"),
            "aqua_credentials_path": self.params.get("aqua_credentials_path", str(DEFAULT_AQUA_CREDENTIALS_PATH)),
            "aqua_output": self.params.get("aqua_output", str(DEFAULT_AQUA_OUTPUT_PATH)),
            "class_vid_limit": self.params.get("class_vid_limit"),
        }

    def exec(self, prep_res):
        try:
            vid_df = pd.read_excel(prep_res["final_workbook"], sheet_name=prep_res["vid_sheet"])
            if "VISUAL_ID" not in vid_df.columns:
                raise ValueError(
                    f"Workbook sheet {prep_res['vid_sheet']} must contain a VISUAL_ID column."
                )

            true_vids = [
                value.strip()
                for value in vid_df["VISUAL_ID"].dropna().astype(str)
                if value.strip() and "_" not in value.strip()
            ]
            if not true_vids:
                empty_aqua_df = pd.DataFrame(
                    columns=["Visual_ID", "Unit_Interface_Bin", "Unit_Functional_Bin"]
                )
                _, workbook_summary = apply_aqua_class_bins_to_final_workbook(
                    final_workbook_path=prep_res["final_workbook"],
                    aqua_df=empty_aqua_df,
                    aqua_sheet_name=AQUA_CLASS_SHEET,
                )
                result = {
                    "status": "skipped_no_true_vid",
                    "vid_sheet": prep_res["vid_sheet"],
                    "final_workbook": workbook_summary["final_workbook"],
                    "aqua_sheet": workbook_summary["aqua_sheet"],
                    "vid_count": 0,
                    "row_count": 0,
                    "matched_visual_ids": 0,
                    "reason": "No true VID values were found; AQUA Class fetch was skipped.",
                }
                print(result["reason"] + " Continuing to Step 6 with blank Class bins.")
                return result

            result = load_cached_aqua_result_workbook(
                final_workbook_path=prep_res["final_workbook"],
                cache_workbook_path=prep_res["aqua_output"],
                vid_sheet_name=prep_res["vid_sheet"],
                cache_sheet_name=AQUA_CLASS_SHEET,
                vid_limit=prep_res.get("class_vid_limit"),
            )
            if result is None:
                result = fetch_aqua_class_7820_from_workbook_vid_sheet(
                    final_workbook_path=prep_res["final_workbook"],
                    vid_sheet_name=prep_res["vid_sheet"],
                    credentials_path=prep_res["aqua_credentials_path"],
                    vid_limit=prep_res.get("class_vid_limit"),
                )
                result["source"] = "aqua"

            result["aqua_output"] = save_aqua_result_workbook(
                result["result_df"],
                output_path=prep_res["aqua_output"],
                sheet_name=AQUA_CLASS_SHEET,
            )
            _, workbook_summary = apply_aqua_class_bins_to_final_workbook(
                final_workbook_path=prep_res["final_workbook"],
                aqua_df=result["result_df"],
                aqua_sheet_name=AQUA_CLASS_SHEET,
            )
            result["final_workbook"] = workbook_summary["final_workbook"]
            result["matched_visual_ids"] = workbook_summary["matched_visual_ids"]
            result["aqua_sheet"] = workbook_summary["aqua_sheet"]
            print(
                f"Prepared AQUA Class 7820 data for {result['vid_count']} true VID values from {result.get('source', 'aqua')} into {result['final_workbook']} sheet {result['aqua_sheet']}."
            )
            return result
        except Exception as exc:
            result = {
                "status": "failed",
                "vid_sheet": prep_res.get("vid_sheet"),
                "final_workbook": prep_res.get("final_workbook"),
                "aqua_credentials_path": prep_res.get("aqua_credentials_path"),
                "error": str(exc),
            }
            print(f"AQUA Class data fetch failed: {exc}")
            return result

    def post(self, shared, prep_res, exec_res):
        shared["class_fetch_summary"] = exec_res
        status = exec_res.get("status")

        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            if status == "completed":
                message = (
                    f"Step 5 Passed: AQUA Class 7820 fetch complete.\n\n"
                    f"Source: {exec_res.get('source', 'aqua')}\n"
                    f"VIDs fetched: {exec_res['vid_count']}\n"
                    f"Rows fetched: {exec_res['row_count']}\n"
                    f"Matched CATTS Visual_ID rows: {exec_res['matched_visual_ids']}\n"
                    f"AQUA sheet: {exec_res['aqua_sheet']}\n"
                    f"Standalone workbook: {exec_res['aqua_output']}\n"
                    f"Final workbook: {exec_res['final_workbook']}\n\n"
                    "Next: Step 6 will start automatically."
                )
                messagebox.showinfo("Step 5 Passed", message, parent=root)
            elif status == "skipped_no_true_vid":
                message = (
                    "Step 5 skipped: no true VIDs were found in the CATTS report.\n\n"
                    "AQUA Class data will be blank. Step 6 report generation will continue.\n"
                    f"Final workbook: {exec_res.get('final_workbook')}"
                )
                messagebox.showinfo("Step 5 Skipped", message, parent=root)
            else:
                message = (
                    "Step 5 Failed: AQUA Class data fetch did not complete.\n\n"
                    f"VID sheet: {exec_res.get('vid_sheet')}\n"
                    f"Final workbook: {exec_res.get('final_workbook')}\n"
                    f"Credentials: {exec_res.get('aqua_credentials_path')}\n"
                    f"Error: {exec_res.get('error')}"
                )
                messagebox.showerror("Step 5 Failed", message, parent=root)
            root.destroy()
        except Exception as exc:
            print(f"Unable to show Class fetch popup: {exc}")

        if status not in {"completed", "skipped_no_true_vid"}:
            print("Stopping flow after Step 5 failure; skipping Step 6 report generation.")
            return "stop"

        return "default"


class GenerateIOEMRBReportNode(Node):
    def prep(self, shared):
        product = (shared.get("selected_wafer") or {}).get("product", "IOE")
        print(f"[Step 6/6] Generating {product} interest-bin summary PPT...")
        summary = shared.get("filtered_catts_raw_summary") or {}
        output_dir = self.params.get("output_dir", ".")
        default_ppt_output = add_date_stamp_to_path(
            os.path.join(output_dir, f"{product} MRB map summary.pptx"),
            separator="",
        )
        requested_ppt_output = self.params.get("mrb_ppt_output")
        if requested_ppt_output:
            requested_ppt_output = add_date_stamp_to_path(requested_ppt_output)

        final_workbook = summary.get("final_workbook") or self.params.get("final_workbook")
        if not final_workbook:
            raise ValueError(
                "Step 6 requires a final workbook path. "
                "Run Step 1-4 and Step 5 first, then retry Step 6."
            )

        cleaned_df = pd.read_excel(final_workbook, sheet_name="CATTS-cleaned")
        required_columns = ["Class IB", "Class FB"]
        missing_columns = [column for column in required_columns if column not in cleaned_df.columns]
        if missing_columns:
            raise RuntimeError(
                "Step 6 blocked: Class columns are missing in CATTS-cleaned. "
                "Please complete Step 5 first so Class results are written into the final workbook. "
                f"Missing columns: {', '.join(missing_columns)}"
            )

        class_populated_rows = int(cleaned_df[required_columns].notna().any(axis=1).sum())
        class_fetch_summary = shared.get("class_fetch_summary") or {}
        skipped_without_vids = class_fetch_summary.get("status") == "skipped_no_true_vid"
        if class_populated_rows == 0 and not skipped_without_vids:
            raise RuntimeError(
                "Step 6 blocked: Class columns exist but contain no populated rows. "
                "Please confirm Step 5 completed successfully and wrote data into the final workbook."
            )

        return {
            "final_workbook": final_workbook,
            "ioe_map_path": shared.get("wafer_map_path") or self.params.get(
                "gpu_map_output" if product == "GPU" else "ioe_map_output",
                f"domain-knowledge/wafer-map/{product.lower()}_wafer_map.json",
            ),
            "output_dir": output_dir,
            "mrb_ppt_output": requested_ppt_output or default_ppt_output,
        }

    def exec(self, prep_res):
        result = generate_ioe_mrb_summary_ppt(
            final_workbook_path=prep_res["final_workbook"],
            ioe_map_path=prep_res["ioe_map_path"],
            output_ppt_path=prep_res["mrb_ppt_output"],
            output_dir=prep_res["output_dir"],
        )
        print(
            f"Generated Step 6 wafer test result PPT for {result['wafer_count']} wafer(s) at {result['output_ppt']}."
        )
        return result

    def post(self, shared, prep_res, exec_res):
        shared["step6_summary"] = exec_res
        cleanup_paths = cleanup_ioe_intermediate_files(
            output_dir=prep_res.get("output_dir"),
        )
        shared["ioe_cleanup_paths"] = cleanup_paths

        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showinfo(
                "Wafer Test Result Report Complete",
                (
                    f"Test result PPT: {exec_res['output_ppt']}\n"
                    f"Wafer count: {exec_res['wafer_count']}\n"
                    "Interest: all IB15 and IB36 results\n"
                    "IB15 FB1590-1599: up triangle; other/no FB: down triangle\n"
                    "IB36: diamond\n"
                    "SDHMT: red; Class: purple\n"
                    "Whole-wafer and outermost edge-die failures reported\n"
                    f"Intermediate files removed: {len(cleanup_paths)}"
                ),
                parent=root,
            )
            root.destroy()
        except Exception as exc:
            print(f"Unable to show Step 6 completion popup: {exc}")

        if cleanup_paths:
            print("Removed wafer workflow intermediate files:")
            for cleanup_path in cleanup_paths:
                print(f"- {cleanup_path}")

        return "default"
