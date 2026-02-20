
import pandas as pd
import numpy as np
# import seaborn as sns # Skipped to avoid dependency issues
# import matplotlib.pyplot as plt # Skipped to avoid dependency issues
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"
CATALOG_PATH = os.path.join(OUTPUT_DIR, "signal_catalog.csv")

def analyze_advanced():
    print("Loading data...")
    df = pd.read_csv(FILE_PATH)
    catalog = pd.read_csv(CATALOG_PATH)
    
    # 1. Segment Data by Unit
    units = catalog['Unit'].unique()
    print(f"Found Units: {units}")
    
    # Ensure Unit column is treated consistently (as string or int)
    catalog['Unit'] = catalog['Unit'].astype(str)
    
    # 2. Operating Mode Detection (State Machine Inference)
    # Look for digital signals that hold steady state for long periods (low switch count)
    # but still have variance (not constantly 0 or 1).
    print("\n--- Identifying Potential Mode Selectors ---")
    digital_cols = catalog[catalog['SignalType'] == 'Digital']['FullTag'].tolist()
    
    potential_modes = []
    for col in digital_cols:
        if col not in df.columns: continue
        
        switches = (df[col].diff() != 0).sum()
        # If switches are rare (e.g., < 20 in 43000 rows) but > 0
        if 0 < switches < 50:
            potential_modes.append(col)
            
    print(f"Potential Operating Mode Signals: {len(potential_modes)}")
    # Print top 10 candidates
    print(potential_modes[:10]) 
    
    # 3. Correlation Analysis (Focusing on Analog Signals in Unit 51 - Main Plant)
    print("\n--- Correlation Analysis (Unit 51) ---")
    unit51_analog = catalog[(catalog['Unit'] == '51') & (catalog['SignalType'] == 'Analog')]['FullTag'].tolist()
    # Filter out columns not in df
    unit51_analog = [c for c in unit51_analog if c in df.columns]
    print(f"DEBUG: Found {len(unit51_analog)} analog signals in Unit 51")
    
    # Remove constant columns (std dev near 0) to avoid NaN correlations
    # Also handle numeric conversion just in case
    df_51 = df[unit51_analog].apply(pd.to_numeric, errors='coerce')
    df_51 = df_51.loc[:, df_51.std() > 0.01]
    
    print(f"DEBUG: df_51 shape for correlation: {df_51.shape}")
    
    corr_matrix = df_51.corr()
    
    # Find strong correlations (> 0.8 or < -0.8)
    strong_pairs = []
    for i in range(len(corr_matrix.columns)):
        for j in range(i+1, len(corr_matrix.columns)):
            val = corr_matrix.iloc[i, j]
            if abs(val) > 0.80:
                strong_pairs.append((corr_matrix.columns[i], corr_matrix.columns[j], val))

                
    print(f"Found {len(strong_pairs)} strong correlation pairs (>0.80)")
    
    # Save Strong Pairs
    pairs_df = pd.DataFrame(strong_pairs, columns=['Signal_A', 'Signal_B', 'Correlation'])
    pairs_df.to_csv(os.path.join(OUTPUT_DIR, "strong_correlations_unit51.csv"), index=False)
    
    # 4. KPI Identification (Variance Analysis)
    # Signals with high variance might be key control variables or unstable loops
    print("\n--- High Variance Signals (Possible KPIs or Instability) ---")
    variance = df_51.var().sort_values(ascending=False)
    print(variance.head(10))
    
    return pairs_df

if __name__ == "__main__":
    analyze_advanced()
