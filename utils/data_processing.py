

import pandas as pd
import os
import yaml

def generate_pivot_table(df: pd.DataFrame, settings: dict = None) -> pd.DataFrame:
    """
    Generates a pivot table based on settings.
    Default settings mirror the previous hardcoded behavior.
    """
    if settings is None:
        settings = {
            "index": ["LOT", "WAFER_SCRIBE"],
            "columns": ["INTERFACE_BIN"],
            "values": "PROGRAM_NAME",
            "aggfunc": "count"
        }

    index_cols = settings.get("index", ["LOT", "WAFER_SCRIBE"])
    columns = settings.get("columns", ["INTERFACE_BIN"])
    values = settings.get("values", "PROGRAM_NAME")
    aggfunc = settings.get("aggfunc", "count")
    fill_value = settings.get("fill_value", 0)

    # Validate columns exist
    required_cols = []
    if isinstance(index_cols, list): required_cols.extend(index_cols)
    else: required_cols.append(index_cols)
    
    if isinstance(columns, list): required_cols.extend(columns)
    else: required_cols.append(columns)
    
    if isinstance(values, list): required_cols.extend(values)
    else: required_cols.append(values)
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
         print(f"Warning: Missing columns for pivot table: {missing_cols}")
         # Attempt to proceed? Or raise error? Raised error previously.
         raise ValueError(f"Missing columns in DataFrame: {missing_cols}")

    print(f"Generating pivot table with settings: Index={index_cols}, Columns={columns}, Values={values}, AggFunc={aggfunc}")

    pivot_table = pd.pivot_table(
        df,
        index=index_cols,
        columns=columns,
        values=values,
        aggfunc=aggfunc,
        fill_value=fill_value
    )
    
    return pivot_table

def analyze_data(df: pd.DataFrame, config_path: str = "data_process_details.yaml"):
    """
    Analyzes data to produce IB and FB pivot tables.
    Returns: Dict[str, pd.DataFrame] containing the sheet names and dataframes.
    """
    # Create IB Pivot
    # Index: LOT, WAFER_SCRIBE
    # Columns: INTERFACE_BIN
    # Values: PROGRAM_NAME (count)
    if "INTERFACE_BIN" not in df.columns or "FUNCTIONAL_BIN" not in df.columns:
        print("Required bin columns missing for analysis.")
        return None
        
    # --- IB Pivot ---
    ib_pivot = pd.pivot_table(
        df,
        index=["LOT", "WAFER_SCRIBE", "DEVREVSTEP"],
        columns=["INTERFACE_BIN"],
        values="PROGRAM_NAME",
        aggfunc="count",
        fill_value=0
    ).reset_index()  # Flatten index to avoid merged cells
    
    # Rename columns (add 'IB' prefix) 
    # Must preserve Index columns first
    new_cols = []
    for col in ib_pivot.columns:
        if col in ["LOT", "WAFER_SCRIBE", "DEVREVSTEP"]:
            new_cols.append(col)
        else:
            new_cols.append(f"IB{col}")
    ib_pivot.columns = new_cols
    
    # --- FB Pivot ---
    fb_pivot = pd.pivot_table(
        df,
        index=["LOT", "WAFER_SCRIBE", "DEVREVSTEP"],
        columns=["FUNCTIONAL_BIN"],
        values="PROGRAM_NAME",
        aggfunc="count",
        fill_value=0
    ).reset_index()  # Flatten index to avoid merged cells
    
    # Rename columns (add 'FB' prefix)
    new_cols = []
    for col in fb_pivot.columns:
        if col in ["LOT", "WAFER_SCRIBE", "DEVREVSTEP"]:
            new_cols.append(col)
        else:
            new_cols.append(f"FB{col}")
    fb_pivot.columns = new_cols
    
    # --- Combined Pivot (IB + FB) ---
    # Merge on keys to keep rows aligned
    combined_pivot = pd.merge(ib_pivot, fb_pivot, on=["LOT", "WAFER_SCRIBE", "DEVREVSTEP"], how="outer").fillna(0)
    
    return {
        "IB pivot": ib_pivot, 
        "FB pivot": fb_pivot,
        "IB+FB pivot": combined_pivot
    }

if __name__ == "__main__":
    # Test logic
    input_file = "SDHMT die level data.csv"
    output_file = "wafer level rollup.xlsx"
    
    if os.path.exists(input_file):
        print(f"Reading data from {input_file}...")
        try:
            df = pd.read_csv(input_file)
            results = analyze_data(df)
            
            if results:
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for sheet_name, sheet_df in results.items():
                        sheet_df.to_excel(writer, sheet_name=sheet_name)
                print(f"Results saved to {output_file}")
        except Exception as e:
            print(f"Error: {e}")
            print(f"Pivot table saved to {output_file}")
            
        except Exception as e:
            print(f"Error processing data: {e}")
    else:
        print(f"File {input_file} not found. Please run the main flow to generate data first.")