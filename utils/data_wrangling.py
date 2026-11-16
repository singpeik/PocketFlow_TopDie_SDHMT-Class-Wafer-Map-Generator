import pandas as pd
from typing import List, Dict, Any

def wrangle_results(results: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Converts the raw list of dictionaries from the query results into a pandas DataFrame.
    Performs initial data cleaning and type conversion if necessary.
    """
    if not results:
        return pd.DataFrame()
    
    df = pd.DataFrame(results)
    
    # Ensure column names are uppercase to match downstream expectations
    df.columns = [str(c).upper() for c in df.columns]
    
    if 'DEVREVSTEP' in df.columns:
        # Extract starting with '8' and the next 3 characters (e.g., 8PXJAVFA -> 8PXJ)
        extracted = df['DEVREVSTEP'].str.extract(r'(8.{3})')[0]
        # Fallback to the original value if it doesn't match the pattern
        df['DEVREVSTEP'] = extracted.fillna(df['DEVREVSTEP'])
        
    return df
