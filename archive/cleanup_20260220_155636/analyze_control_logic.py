
# -*- coding: utf-8 -*-
import sys
import os

# UTF-8 Fix for Windows cp949
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Set style
plt.style.use('seaborn-v0_8-darkgrid')

def load_data():
    print("Loading data...")
    cols = [
        'localtime',
        # Unit 51 Press Control
        'ERS_51DPIC111_BSP', 'ERS_51DPIC111_CTRNOUT', 'ERS_51PDIT0111_Y',
        # Unit 81 Split Range (Section 3.1.2.3 in FDS)
        'ERS_81PIC0003_CTRNOUT', # OP
        'ERS_81PIC0003_BSP',     # SP
        'ERS_51PIT011A_Y',       # PV (Confirmed in FDS)
        'ERS_81XV0002_3_Y',      # Split Range Outcome 1 (Valve Open/Close)
        'ERS_81PCV0003_OUTPOS',  # Split Range Outcome 2
        'ERS_81PCV0004_OUTPOS',  # Split Range Outcome 3
        # Mode
        'ERS_MODE_Y'
    ]
    # Filter only existing columns
    df_head = pd.read_csv(FILE_PATH, nrows=5)
    cols = [c for c in cols if c in df_head.columns]
    
    df = pd.read_csv(FILE_PATH, usecols=cols)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    return df

def analyze_split_range(df):
    """
    Validates FDS Sec 3.1.2.3: 81PIC0003 Split Range Control.
    OP 0-50%: 81XV0002 / 81PCV0003
    OP 50-100%: 81PCV0004
    """
    print("\n--- Analyzing Split Range Logic (81PIC0003) ---")
    
    if 'ERS_81PIC0003_CTRNOUT' not in df.columns:
        print("Controller Signal missing!")
        return

    # Create a plot to show the correlation
    # We want to see if PCV0004 moves when OP > 50
    # And if PCV0003 moves when OP < 50
    
    fig, ax1 = plt.subplots(figsize=(10, 6))
    
    # Plot OP (Control Output)
    ax1.plot(df.index, df['ERS_81PIC0003_CTRNOUT'], color='black', label='Controller OP (81PIC003)', linewidth=1)
    ax1.set_ylabel('Controller Output (%)')
    ax1.set_ylim(0, 100)
    
    ax2 = ax1.twinx()
    
    # Plot Valves
    colors = {'ERS_81PCV0003_OUTPOS': 'blue', 'ERS_81PCV0004_OUTPOS': 'red'}
    for col, color in colors.items():
        if col in df.columns:
            ax2.plot(df.index, df[col], color=color, alpha=0.6, label=f'Valve Pos ({col})')
            
    ax2.set_ylabel('Valve Position (%)')
    ax2.set_ylim(0, 100)
    
    # Combine legends
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
    
    plt.title('Validation of Split Range Control (FDS 3.1.2.3)')
    
    save_path = os.path.join(OUTPUT_DIR, "split_range_validation.png")
    plt.savefig(save_path, dpi=150)
    print(f"Saved {save_path}")
    
    # QUANTITATIVE CHECK
    # Check correlation between OP and Valves in specific ranges
    op = df['ERS_81PIC0003_CTRNOUT']
    
    # Low Range (0-50): Should correlate with PCV0003
    low_range_mask = (op > 0) & (op < 50)
    if 'ERS_81PCV0003_OUTPOS' in df.columns and low_range_mask.sum() > 100:
        corr_low = op[low_range_mask].corr(df.loc[low_range_mask, 'ERS_81PCV0003_OUTPOS'])
        print(f"Correlation OP vs PCV0003 (Low Range): {corr_low:.4f}")
        
    # High Range (50-100): Should correlate with PCV0004
    high_range_mask = (op > 50)
    if 'ERS_81PCV0004_OUTPOS' in df.columns and high_range_mask.sum() > 100:
        corr_high = op[high_range_mask].corr(df.loc[high_range_mask, 'ERS_81PCV0004_OUTPOS'])
        print(f"Correlation OP vs PCV0004 (High Range): {corr_high:.4f}")

def analyze_startup_ramp(df):
    """
    Deep dive into the 15:35 Unit 71 startup event.
    """
    print("\n--- Analyzing Startup Ramp (15:35) ---")
    start_window = "2024-08-28 15:30:00"
    end_window = "2024-08-28 15:50:00"
    
    df_slice = df.loc[start_window:end_window]
    
    if df_slice.empty:
        print("No data for startup window")
        return
        
    # Plot Setpoint vs PV for 51DPIC111 (Main Pressure Loop)
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # SP
    if 'ERS_51DPIC111_BSP' in df.columns:
        ax.plot(df_slice.index, df_slice['ERS_51DPIC111_BSP'], 'r--', label='Setpoint (SP)', drawstyle='steps-post')
        
    # PV (Assuming PDIT0111_Y is the PV based on tag logic)
    if 'ERS_51PDIT0111_Y' in df.columns:
        ax.plot(df_slice.index, df_slice['ERS_51PDIT0111_Y'], 'g-', label='Pressure (PV)')
        
    # Mode
    if 'ERS_MODE_Y' in df.columns:
        ax2 = ax.twinx()
        ax2.plot(df_slice.index, df_slice['ERS_MODE_Y'], 'k:', label='System Mode', alpha=0.5)
        ax2.set_ylabel('Mode')
        
    plt.title('Pressure Ramp-up during Startup (15:35)')
    ax.legend(loc='upper left')
    
    save_path = os.path.join(OUTPUT_DIR, "startup_ramp_detail.png")
    plt.savefig(save_path, dpi=150)
    print(f"Saved {save_path}")

if __name__ == "__main__":
    df = load_data()
    analyze_split_range(df)
    analyze_startup_ramp(df)
