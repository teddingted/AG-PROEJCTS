
import pandas as pd
import numpy as np
import re
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def parse_tag(tag):
    """
    Parses a tag name like 'ERS_51DPIC111_BSP' into components.
    Custom logic based on observed patterns.
    """
    if tag == 'localtime':
        return {'FullTag': tag, 'Type': 'Time', 'System': 'Meta', 'Equipment': 'Time', 'Attribute': 'Time'}
    
    parts = tag.split('_')
    
    # Default structure
    info = {
        'FullTag': tag,
        'System': parts[0] if len(parts) > 0 else 'Unknown',
        'Unit': 'Unknown',
        'Equipment': 'Unknown',
        'Attribute': 'Unknown'
    }
    
    # Try to extract Unit (e.g., 51 from ERS_51...)
    # Pattern: Prefix_UnitEquipment_Attribute
    
    if len(parts) >= 2:
        # Check if the second part starts with digits (Unit + Eq)
        match = re.match(r'^(\d+)([A-Z0-9]+)$', parts[1])
        if match:
            info['Unit'] = match.group(1)
            info['Equipment'] = match.group(2)
        else:
            info['Equipment'] = parts[1]
            
        if len(parts) >= 3:
            info['Attribute'] = '_'.join(parts[2:]) # Join the rest as attribute
    
    return info

def classify_signal_type(series):
    """
    Heuristics to determine if a signal is Analog or Digital.
    """
    unique_counts = series.nunique()
    if unique_counts <= 2:
        return 'Digital' # Likely State (0/1)
    elif unique_counts < 10 and series.dtype != 'float64':
        return 'Discrete' # Likely Mode (1, 2, 3...)
    else:
        return 'Analog' # Continuous variable

def profile_data():
    print(f"Loading {FILE_PATH}...")
    # Load first few rows to infer types safely, then full load
    df = pd.read_csv(FILE_PATH)
    
    print(f"Data Loaded: {df.shape[0]} rows, {df.shape[1]} columns")
    
    tag_catalog = []
    
    print("Profiling columns...")
    for col in df.columns:
        if col == 'localtime':
            continue
            
        # Parse Tag Structure
        meta = parse_tag(col)
        
        # Stats
        series = df[col]
        sig_type = classify_signal_type(series)
        
        meta['SignalType'] = sig_type
        meta['Min'] = series.min()
        meta['Max'] = series.max()
        meta['Mean'] = series.mean()
        meta['StdDev'] = series.std()
        meta['UniqueValues'] = series.nunique()
        meta['Zeros'] = (series == 0).sum()
        
        tag_catalog.append(meta)
        
    catalog_df = pd.DataFrame(tag_catalog)
    
    # Save Catalog
    catalog_path = os.path.join(OUTPUT_DIR, "signal_catalog.csv")
    catalog_df.to_csv(catalog_path, index=False)
    print(f"Signal Catalog saved to {catalog_path}")
    
    # Summary
    print("\n--- Summary ---")
    print(catalog_df['SignalType'].value_counts())
    print("\nTop 5 Units:")
    print(catalog_df['Unit'].value_counts().head())
    
    return catalog_df

if __name__ == "__main__":
    profile_data()
