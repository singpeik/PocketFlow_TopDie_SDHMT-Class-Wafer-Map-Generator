import os
import json
import tempfile
import pandas as pd

# Ensure pythonnet loads .NET 8 with BinaryFormatter enabled if PyUber needs CoreCLR
try:
    import pythonnet
    _runtime_cfg = {
        "runtimeOptions": {
            "tfm": "net8.0",
            "framework": {
                "name": "Microsoft.NETCore.App",
                "version": "8.0.31",
            },
            "configProperties": {
                "System.Runtime.Serialization.EnableUnsafeBinaryFormatterSerialization": True
            },
        }
    }
    _tmp_cfg = os.path.join(tempfile.gettempdir(), "pyuber.net8.runtimeconfig.json")
    with open(_tmp_cfg, "w", encoding="utf-8") as _f:
        json.dump(_runtime_cfg, _f)
    try:
        pythonnet.load("coreclr", runtime_config=_tmp_cfg)
    except Exception:
        pass
except Exception:
    pass

try:
    import PyUber
except ImportError:
    PyUber = None
    print("Warning: PyUber not found. Running in mock mode.")

def execute_query(query: str, datasource: str) -> list:
    """
    Executes a SQL query against the specified PyUber datasource.
    Returns a list of dictionaries (records).
    """
    if PyUber is None:
        # Mock behavior for development/testing without PyUber installed
        print(f"[MOCK] Connecting to {datasource}")
        print(f"[MOCK] Executing: {query}")
        return [
            {
                "lot": "W521101A",
                "operation": "197254",
                "devrevstep": "D8P1JAFH0",
                "test_end_date": "2024-05-01 10:00:00",
                "session_latest_flag": "Y",
                "device_latest_flag": "Y",
                "program_name": "MOCK_PROGRAM",
                "wafer_id": "102",
                "sort_lot": "W521101A",
                "wafer_scribe": "W521101A0102",
                "sort_wafer_id": "102",
                "sort_x": 0,
                "sort_y": 1,
                "interface_bin": 1,
                "functional_bin": 1,
            },
            {
                "lot": "W521101A",
                "operation": "197254",
                "devrevstep": "D8P1JAFH0",
                "test_end_date": "2024-05-01 10:00:00",
                "session_latest_flag": "Y",
                "device_latest_flag": "Y",
                "program_name": "MOCK_PROGRAM",
                "wafer_id": "102",
                "sort_lot": "W521101A",
                "wafer_scribe": "W521101A0102",
                "sort_wafer_id": "102",
                "sort_x": 1,
                "sort_y": 2,
                "interface_bin": 2,
                "functional_bin": 2,
            },
        ]

    try:
        print(f"Connecting to {datasource}...")
        conn = PyUber.connect(datasource=datasource)
        cursor = conn.cursor()
        
        print(f"Executing query...")
        cursor.execute(query)
        
        # Fetch all results
        # Assuming cursor.description is available for column names
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        
        results = []
        for row in rows:
            results.append(dict(zip(columns, row)))
            
        conn.close()
        return results
        
    except Exception as e:
        print(f"Error executing query: {e}")
        raise e
