
import pandas as pd
import numpy as np
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"
CATALOG_PATH = os.path.join(OUTPUT_DIR, "signal_catalog.csv")

def generate_timeline():
    print("Loading data...")
    df = pd.read_csv(FILE_PATH)
    catalog = pd.read_csv(CATALOG_PATH)
    
    # 1. Select Active Digital Signals (Event Markers)
    # We want signals that DID change state, but not TOO often (checking for bouncing/noise)
    print("Scanning for event markers...")
    
    events = []
    
    cols_to_check = catalog[catalog['SignalType'] == 'Digital']['FullTag'].tolist()
    
    # Also include the explicit 'MODE' signal if it exists
    if 'ERS_MODE_Y' in df.columns:
        cols_to_check.append('ERS_MODE_Y')
        
    cols_to_check = [c for c in cols_to_check if c in df.columns]
    
    for col in cols_to_check:
        # Detect changes: shift and compare
        changes = df[col].diff().fillna(0)
        change_indices = changes[changes != 0].index.tolist()
        
        # Filter noisy signals (more than 50 changes in 24h might be noise or high-freq control)
        if 0 < len(change_indices) < 50:
            for idx in change_indices:
                timestamp = df.loc[idx, 'localtime']
                new_val = df.loc[idx, col]
                prev_val = df.loc[idx-1, col]
                
                events.append({
                    'Timestamp': timestamp,
                    'TimeIndex': idx,
                    'Tag': col,
                    'Event': f"Changed from {prev_val} to {new_val}",
                    'From': prev_val,
                    'To': new_val
                })
                
    # 2. Sort and Aggregate events
    events_df = pd.DataFrame(events)
    
    if events_df.empty:
        print("No significant events found! The system seems completely static.")
        return
        
    events_df['Timestamp'] = pd.to_datetime(events_df['Timestamp'])
    events_df = events_df.sort_values(by='Timestamp')
    
    # 3. Create Narrative Timeline
    # Group events that happen within the same minute to define "Macro Events"
    print("\n--- Operational Timeline ---")
    
    events_df['TimeGroup'] = events_df['Timestamp'].dt.floor('5min') # Group by 5 minute windows
    
    grouped = events_df.groupby('TimeGroup')
    
    timeline_summary = []
    
    for time_bucket, group in grouped:
        print(f"\n[ Time: {time_bucket} ]")
        summary_line = f"At {time_bucket}, {len(group)} signals changed state."
        print(summary_line)
        timeline_summary.append(summary_line)
        
        # List specific changes
        for _, row in group.iterrows():
            detail = f"  - {row['Tag']}: {row['From']} -> {row['To']}"
            print(detail)
            timeline_summary.append(detail)
            
    # Save textual timeline
    with open(os.path.join(OUTPUT_DIR, "operational_timeline.txt"), "w") as f:
        f.write("\n".join(timeline_summary))
        
    print(f"\nTimeline saved to {os.path.join(OUTPUT_DIR, 'operational_timeline.txt')}")

if __name__ == "__main__":
    generate_timeline()
