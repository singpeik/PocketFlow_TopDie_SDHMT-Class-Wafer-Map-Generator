import yaml
import pandas as pd
import os
import json

def read_yaml_config(filepath: str) -> dict:
    """
    Reads a YAML configuration file.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Config file not found: {filepath}")
    
    with open(filepath, 'r') as f:
        config = yaml.safe_load(f)
    
    return config

def save_results(data, filepath: str, format: str = 'csv'):
    """
    Saves data to a file in the specified format.
    Data is expected to be a list of dictionaries or a pandas DataFrame.
    """
    if data is None:
        print("No data to save.")
        return

    if isinstance(data, pd.DataFrame):
        if data.empty:
            print("No data to save.")
            return
        df = data
    elif isinstance(data, list):
        if not data:
            print("No data to save.")
            return
        df = pd.DataFrame(data)
    else:
        raise ValueError("Data must be a list of dictionaries or a pandas DataFrame.")

    # Create directory if it doesn't exist
    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or '.', exist_ok=True)

    if format.lower() == 'csv':
        df.to_csv(filepath, index=False)
    elif format.lower() == 'json':
        df.to_json(filepath, orient='records', indent=4)
    else:
        raise ValueError(f"Unsupported format: {format}")
    
    print(f"Results saved to {filepath}")
