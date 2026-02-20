
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

# Set style
plt.style.use('seaborn-v0_8')

def load_data():
    print("Loading data...")
    # Load specific columns to save memory/time
    cols = [
        'localtime',
        'ERS_MODE_Y',             # General Mode
        'ERS_71TIC0001_CTRNOUT',  # Unit 71 Active (if > 0)
        'ERS_51DPIC111_BSP',      # Unit 51 Press Mode (Low/High)
        'ERS_81XV0001_10_Y',      # Unit 81 Valve config
        'ERS_51FY0001_Y'          # Main Flow (for background context)
    ]
    # Check what actually exists
    # We'll read first to check columns roughly or just read all and filter
    # Reading all is safer given size is small enough
    df = pd.read_csv(FILE_PATH)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    return df

def create_gantt_chart(df):
    print("Creating Operational Gantt Chart...")
    
    # 1. Define Boolean States
    # Resample to 1 min to make plotting lighter
    df_res = df.resample('1min').mean()
    
    # State Logic
    # Unit 71 Active: Control Output > 1%
    s_unit71 = df_res['ERS_71TIC0001_CTRNOUT'] > 1.0
    
    # High Pressure Mode: Setpoint > 50 (It jumped from 12 to 100)
    s_press_high = df_res['ERS_51DPIC111_BSP'] > 50.0
    
    # Unit 81 Active: Valve Open (Logic might be inverted, checking timeline: 0->1 at 22:15, let's assume 1 is Active/Open)
    # Actually identifying "Config A" vs "Config B"
    s_unit81_monitor = df_res['ERS_81XV0001_10_Y'] > 0.5
    
    # Main Compressor Load: Normalized Flow
    s_load = df_res['ERS_51FY0001_Y'] / df_res['ERS_51FY0001_Y'].max()
    
    # 2. Plotting (Strip Chart)
    fig, ax = plt.subplots(figsize=(16, 8))
    
    # Strip 1: System Mode (Background)
    # Plot raw mode as a step line/shaded area?
    # Let's simple use bars
    
    # Configuration for Strips
    strips = [
        {'data': s_unit71, 'label': 'Unit 71 (Aux) Active', 'color': '#2ecc71', 'y_pos': 3},
        {'data': s_press_high, 'label': 'Unit 51 High Pressure Mode', 'color': '#e74c3c', 'y_pos': 2},
        {'data': s_unit81_monitor, 'label': 'Unit 81 Valve Config', 'color': '#f39c12', 'y_pos': 1}
    ]
    
    # Draw bars
    for strip in strips:
        # Fill between where condition is met
        # Create a boolean series
        ax.fill_between(df_res.index, 
                        strip['y_pos'] - 0.4, 
                        strip['y_pos'] + 0.4, 
                        where=strip['data'], 
                        color=strip['color'], 
                        alpha=0.8, 
                        label=strip['label'],
                        step='mid')
                        
    # Add Main Load as a line overlay on top? Or separate? 
    # Let's put it on a secondary axis or just below?
    # Let's add it as a "Load Intensity" area at the bottom (y=0)
    ax.fill_between(df_res.index, 0, s_load * 0.8, color='#3498db', alpha=0.4, label='Main Compressor Load')
    ax.text(df_res.index[0], 0.5, "Main Comp Load", color='#3498db', fontweight='bold')

    # Formatting
    ax.set_yticks([0.4, 1, 2, 3])
    ax.set_yticklabels(['', 'Unit 81 Config', 'High Press Mode', 'Unit 71 Active'])
    ax.set_ylim(0, 4)
    
    # X Axis
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    plt.xlabel('Time (Hours)')
    
    plt.title('Daily Operational Log (Gantt View)', fontsize=18)
    plt.grid(True, axis='x', linestyle='--', alpha=0.5)
    
    # Annotate Events
    events = [
        ('2024-08-28 15:35', 'Unit 71 Start'),
        ('2024-08-28 21:55', 'System Reset')
    ]
    for time_str, label in events:
        dt = pd.to_datetime(time_str)
        plt.axvline(dt, color='black', linestyle=':', linewidth=2)
        plt.text(dt, 3.6, f" {label}", rotation=0, fontweight='bold', bbox=dict(facecolor='white', alpha=0.7))

    plt.tight_layout()
    
    save_path = os.path.join(OUTPUT_DIR, "operational_gantt.png")
    plt.savefig(save_path, dpi=150)
    print(f"Saved {save_path}")

if __name__ == "__main__":
    df = load_data()
    create_gantt_chart(df)
