
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

plt.style.use('seaborn-v0_8-darkgrid')

def load_full_data():
    """Load complete dataset for deep analysis"""
    print("Loading full dataset...")
    df = pd.read_csv(FILE_PATH)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    print(f"Loaded {len(df)} rows, {len(df.columns)} columns")
    return df

def analyze_control_performance(df):
    """
    Calculate Control Loop Performance Metrics:
    - MSE (Mean Squared Error between SP and PV)
    - Overshoot during ramp events
    - Settling time
    """
    print("\n=== CONTROL PERFORMANCE ANALYSIS ===")
    
    # Define key control loops from FDS
    loops = {
        '71FIC0001 (BOG Flow)': {
            'SP': 'ERS_71FIC0001_BSP',
            'PV': 'ERS_72FI0001_Y',  # Actual flow measurement
            'OP': 'ERS_71FCV0001_OUTPOS'
        },
        '81TIC0001 (LD Inlet Temp)': {
            'SP': 'ERS_81TIC0001_BSP',
            'PV': 'ERS_72TE0001_Y',
            'OP': 'ERS_81TCV0002_OUTPOS'  # Split range, using first valve
        },
        '81PIC0003 (N2 Suction Press)': {
            'SP': 'ERS_81PIC0003_BSP',
            'PV': 'ERS_51PIT011A_Y',
            'OP': 'ERS_81PIC0003_CTRNOUT'
        }
    }
    
    results = []
    
    for loop_name, tags in loops.items():
        # Check if tags exist
        sp_col = tags['SP']
        pv_col = tags['PV']
        
        if sp_col not in df.columns or pv_col not in df.columns:
            print(f"  {loop_name}: MISSING TAGS")
            continue
            
        # Calculate Error
        error = df[sp_col] - df[pv_col]
        
        # Metrics
        mse = (error ** 2).mean()
        mae = error.abs().mean()
        max_error = error.abs().max()
        
        # Identify when controller is active (SP != 0 and changing)
        active_mask = (df[sp_col] > 0.1) & (df[sp_col].diff().abs() < 10)  # Not ramping wildly
        
        if active_mask.sum() > 100:
            mse_active = (error[active_mask] ** 2).mean()
            mae_active = error[active_mask].abs().mean()
        else:
            mse_active = np.nan
            mae_active = np.nan
        
        results.append({
            'Loop': loop_name,
            'MSE (All)': mse,
            'MAE (All)': mae,
            'MSE (Active)': mse_active,
            'MAE (Active)': mae_active,
            'Max Error': max_error
        })
        
        print(f"  {loop_name}:")
        print(f"    MSE: {mse:.4f}, MAE: {mae:.4f}, Max Error: {max_error:.2f}")
    
    # Save results
    results_df = pd.DataFrame(results)
    results_df.to_csv(os.path.join(OUTPUT_DIR, "control_performance.csv"), index=False)
    print(f"  Saved control_performance.csv")
    
    return results_df

def analyze_valve_health(df):
    """
    Valve Stiction & Deadband Analysis:
    - Plot OP vs Valve Position to detect hysteresis
    - Calculate valve response delay
    """
    print("\n=== VALVE HEALTH DIAGNOSTICS ===")
    
    # Focus on critical valves
    valves = {
        '71FCV0001 (BOG Flow Valve)': {
            'CMD': 'ERS_71FIC0001_CTRNOUT',  # Controller Output
            'POS': 'ERS_71FCV0001_OUTPOS'
        },
        '81TCV0002 (Temp Control)': {
            'CMD': 'ERS_81TIC0001_CTRNOUT',
            'POS': 'ERS_81TCV0002_OUTPOS'
        }
    }
    
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    for idx, (valve_name, tags) in enumerate(valves.items()):
        cmd_col = tags['CMD']
        pos_col = tags['POS']
        
        if cmd_col not in df.columns or pos_col not in df.columns:
            print(f"  {valve_name}: MISSING")
            continue
        
        # Subsample for clarity (every 100th point)
        df_sub = df[[cmd_col, pos_col]].dropna()[::100]
        
        axes[idx].scatter(df_sub[cmd_col], df_sub[pos_col], s=1, alpha=0.3)
        axes[idx].plot([0, 100], [0, 100], 'r--', label='Ideal (1:1)')
        axes[idx].set_xlabel('Command (%)')
        axes[idx].set_ylabel('Actual Position (%)')
        axes[idx].set_title(f'{valve_name}\nStiction Check')
        axes[idx].legend()
        axes[idx].grid(True)
        
        # Calculate linearity (R²)
        valid_mask = (df_sub[cmd_col] > 0) & (df_sub[pos_col] > 0)
        if valid_mask.sum() > 10:
            corr = df_sub.loc[valid_mask, cmd_col].corr(df_sub.loc[valid_mask, pos_col])
            print(f"  {valve_name}: Correlation = {corr:.4f}")
    
    plt.tight_layout()
    save_path = os.path.join(OUTPUT_DIR, "valve_health.png")
    plt.savefig(save_path, dpi=150)
    print(f"  Saved valve_health.png")

def trace_causal_chain(df):
    """
    Trace Temperature → Pressure → Flow causality during startup (15:35)
    """
    print("\n=== CAUSAL CHAIN TRACKING (15:35 Startup) ===")
    
    start = "2024-08-28 15:30:00"
    end = "2024-08-28 15:50:00"
    
    df_event = df.loc[start:end]
    
    # Key variables
    vars_of_interest = {
        'Temp (LDC Inlet)': 'ERS_72TE0001_Y',
        'Press (N2 Suction)': 'ERS_51PIT011A_Y',
        'Flow (BOG)': 'ERS_72FI0001_Y',
        'Mode': 'ERS_MODE_Y'
    }
    
    # Filter existing columns
    plot_cols = {k: v for k, v in vars_of_interest.items() if v in df.columns}
    
    if len(plot_cols) < 2:
        print("  Insufficient data for causal analysis")
        return
    
    # Normalize for comparison (Z-score within event window)
    df_norm = pd.DataFrame()
    for label, col in plot_cols.items():
        if df_event[col].std() > 0.01:
            df_norm[label] = (df_event[col] - df_event[col].mean()) / df_event[col].std()
        else:
            df_norm[label] = df_event[col]  # Constant signal
    
    df_norm.index = df_event.index
    
    # Plot
    fig, ax = plt.subplots(figsize=(12, 6))
    
    for label in df_norm.columns:
        ax.plot(df_norm.index, df_norm[label], label=label, linewidth=1.5)
    
    ax.set_title('Causal Chain During Startup (Normalized)')
    ax.set_ylabel('Normalized Value (Z-score)')
    ax.legend(loc='upper left')
    ax.grid(True)
    
    plt.tight_layout()
    save_path = os.path.join(OUTPUT_DIR, "causal_chain_startup.png")
    plt.savefig(save_path, dpi=150)
    print(f"  Saved causal_chain_startup.png")
    
    # Calculate Lead-Lag (Cross-correlation)
    print("\n  Lead-Lag Analysis (찾는 중: 어떤 변수가 먼저 움직였는가?)")
    
    if 'Flow (BOG)' in df_norm.columns and 'Press (N2 Suction)' in df_norm.columns:
        # Cross-correlation between Flow and Pressure
        flow = df_norm['Flow (BOG)'].dropna()
        press = df_norm['Press (N2 Suction)'].dropna()
        
        # Align
        common_idx = flow.index.intersection(press.index)
        if len(common_idx) > 100:
            corr_at_lag = []
            lags = range(-60, 61)  # ±60 seconds
            
            for lag in lags:
                if lag >= 0:
                    corr = flow.iloc[lag:].corr(press.iloc[:len(press)-lag] if lag > 0 else press)
                else:
                    corr = flow.iloc[:len(flow)+lag].corr(press.iloc[-lag:])
                corr_at_lag.append(corr)
            
            max_corr_idx = np.argmax(np.abs(corr_at_lag))
            best_lag = lags[max_corr_idx]
            
            print(f"    Flow vs Pressure: Best Lag = {best_lag}s (Corr={corr_at_lag[max_corr_idx]:.3f})")
            if best_lag > 0:
                print(f"      → Flow leads Pressure by {best_lag} seconds")
            elif best_lag < 0:
                print(f"      → Pressure leads Flow by {-best_lag} seconds")
            else:
                print(f"      → Simultaneous")

def main():
    df = load_full_data()
    
    # 1. Control Performance
    analyze_control_performance(df)
    
    # 2. Valve Health
    analyze_valve_health(df)
    
    # 3. Causal Chain
    trace_causal_chain(df)
    
    print("\n=== ANALYSIS COMPLETE ===")

if __name__ == "__main__":
    main()
