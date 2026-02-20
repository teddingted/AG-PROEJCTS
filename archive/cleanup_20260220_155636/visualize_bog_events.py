
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
import seaborn as sns
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"
CATALOG_PATH = os.path.join(OUTPUT_DIR, "signal_catalog.csv")

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_context("talk")

def load_data():
    print("Loading data...")
    df = pd.read_csv(FILE_PATH)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    
    catalog = pd.read_csv(CATALOG_PATH)
    return df, catalog

def plot_global_volatility(df, catalog):
    print("Generating Global Volatility Plot...")
    
    # 1. Select Key Analog Signals (Unit 51 & 81)
    # Filter for signals that actually vary (avoid flat lines)
    analogs = catalog[(catalog['SignalType'] == 'Analog') & (catalog['StdDev'] > 0.1)]['FullTag'].tolist()
    analogs = [c for c in analogs if c in df.columns]
    
    print(f"Calculating volatility for {len(analogs)} signals...")
    
    # 2. Normalize Data (Z-Score) so we can aggregate them
    # We use a rolling window to calculate "activity"
    # Volatility = Rolling Std Dev of normalized data
    
    # Processing in chunks to save memory if needed, but 40k rows is fine.
    # Normalize: (Val - Mean) / Std
    df_norm = (df[analogs] - df[analogs].mean()) / df[analogs].std()
    
    # Rolling S.D. over 1 minute window
    volatility = df_norm.rolling(window='1min').std()
    
    # Mean Volatility across all signals = System "Temperature"
    system_index = volatility.mean(axis=1)
    
    # 3. Plot
    plt.figure(figsize=(15, 6))
    plt.plot(system_index.index, system_index, color='#e74c3c', linewidth=1.5, label='System Volatility Index')
    
    # Add Markers for known events
    events = [
        ('2024-08-28 15:35', 'Unit 71 Startup'),
        ('2024-08-28 18:40', 'Mode Switch'),
        ('2024-08-28 21:55', 'System Reset')
    ]
    
    for time_str, label in events:
        dt = pd.to_datetime(time_str)
        plt.axvline(dt, color='#2c3e50', linestyle='--', alpha=0.7)
        plt.text(dt, system_index.max()*0.9, f" {label}", rotation=90, verticalalignment='top')
        
    plt.title('Global System Volatility (Dynamic Activity)', fontsize=16)
    plt.ylabel('Avg Std Dev (Normalized)')
    plt.xlabel('Time')
    plt.legend()
    plt.tight_layout()
    
    save_path = os.path.join(OUTPUT_DIR, "global_volatility.png")
    plt.savefig(save_path, dpi=150)
    print(f"Saved {save_path}")

def plot_event_snapshot(df, event_time_str, label, window_min=20):
    print(f"Generating snapshot for {label} around {event_time_str}...")
    
    center_time = pd.to_datetime(event_time_str)
    start_time = center_time - pd.Timedelta(minutes=window_min/2)
    end_time = center_time + pd.Timedelta(minutes=window_min/2)
    
    # Slice data
    mask = (df.index >= start_time) & (df.index <= end_time)
    df_slice = df.loc[mask]
    
    if df_slice.empty:
        print("Empty slice!")
        return

    # Select interesting signals for this specific view
    # 1. Flow (Load)
    flows = ['ERS_51FY0001_Y', 'ERS_51FY0001_YE', 'ERS_51FY0002_Y']
    flows = [f for f in flows if f in df.columns]
    
    # 2. Pressure Setpoints & Values
    pressures = ['ERS_51DPIC111_BSP', 'ERS_51DPIC111_CTRNOUT', 'ERS_51DPIC111_Y', # Unit 51 Press
                 'ERS_81PIC0002_BSP', 'ERS_81PIC0002_CTRNOUT'] # Unit 81 Press
    # Note: DPIC111_Y might not exist, check catalog. Usually _Y is PV.
    # Looking at catalog, ERS_51DPIC111_BSP is Digital (Setpoint?), ERS_51DPIC111_CTRNOUT is Analog.
    # Let's find the PV for DPIC111. Usually PDIT...
    # ERS_51PDIT0111_Y corresponds to DPIC111 loop.
    
    pressures = [p for p in pressures if p in df.columns]
    
    # 3. Mode / Valves
    valves = ['ERS_51L64_Y', 'ERS_81XV0001_2_Y', 'ERS_MODE_Y']
    valves = [v for v in valves if v in df.columns]
    
    # Create Subplots
    fig, axes = plt.subplots(3, 1, figsize=(12, 12), sharex=True)
    
    # Plot 1: Flows (Load)
    for col in flows:
        axes[0].plot(df_slice.index, df_slice[col], label=col)
    axes[0].set_title(f"{label}: Flow & Load Analysis")
    axes[0].set_ylabel('Flow / Current')
    axes[0].legend(loc='upper right', fontsize='small')
    
    # Plot 2: Pressure Controllers
    for col in pressures:
        # Scale if necessary, or plot on secondary y? keeping simple for now
        axes[1].plot(df_slice.index, df_slice[col], label=col)
    axes[1].set_title(f"{label}: Pressure Controls")
    axes[1].set_ylabel('Pressure / %')
    axes[1].legend(loc='upper right', fontsize='small')
    
    # Plot 3: Discrete States
    for i, col in enumerate(valves):
        # Offset them so they don't overlap
        axes[2].plot(df_slice.index, df_slice[col] + (i*1.2), label=col, drawstyle='steps-post')
    axes[2].set_title(f"{label}: Logic / Valves")
    axes[2].set_yticks([])
    axes[2].legend(loc='upper right', fontsize='small')
    
    plt.xlabel('Time')
    plt.tight_layout()
    
    filename = f"event_{label.lower().replace(' ', '_')}.png"
    save_path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(save_path, dpi=150)
    print(f"Saved {save_path}")

def main():
    df, catalog = load_data()
    plot_global_volatility(df, catalog)
    
    # Plot specific events
    plot_event_snapshot(df, '2024-08-28 15:35:00', 'Unit 71 Startup', window_min=20)
    plot_event_snapshot(df, '2024-08-28 21:55:00', 'System Reset', window_min=20)
    
    # Specific zoom for the 22:00 reset which was complex
    plot_event_snapshot(df, '2024-08-28 22:00:00', 'Shutdown Sequence', window_min=60)

if __name__ == "__main__":
    main()
