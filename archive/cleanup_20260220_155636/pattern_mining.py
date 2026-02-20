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
CSV_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

plt.style.use('seaborn-v0_8-darkgrid')

def load_full_dataset():
    """Load complete dataset for comprehensive mining"""
    print("Loading full dataset for pattern mining...")
    df = pd.read_csv(CSV_PATH)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    print(f"Loaded {len(df)} samples, {len(df.columns)} columns")
    return df

def scan_alarm_trip_signals(df):
    """
    Search for alarm and trip signals that may indicate failure precursors
    """
    print("\n=== SCANNING ALARM/TRIP SIGNALS ===")
    
    # Identify alarm-related columns
    alarm_cols = [c for c in df.columns if any(x in c.upper() for x in ['ALM', 'ALARM', 'TRIP', 'FAIL', 'WARN', 'HH', 'LL', 'HHLL'])]
    
    print(f"Found {len(alarm_cols)} potential alarm/trip signals")
    
    if len(alarm_cols) == 0:
        print("No explicit alarm signals found in tag names")
        return None
    
    # Analyze alarm activation patterns
    alarm_df = df[alarm_cols]
    
    # Find which alarms were active
    alarm_summary = []
    for col in alarm_cols:
        if alarm_df[col].dtype in [np.float64, np.int64]:
            # Digital alarms (0/1)
            activations = (alarm_df[col] > 0).sum()
            if activations > 0:
                first_activation = alarm_df[alarm_df[col] > 0].index[0]
                last_activation = alarm_df[alarm_df[col] > 0].index[-1]
                alarm_summary.append({
                    'Signal': col,
                    'Activations': activations,
                    'First': first_activation,
                    'Last': last_activation,
                    'Duration_min': (last_activation - first_activation).total_seconds() / 60
                })
    
    if alarm_summary:
        alarm_summary_df = pd.DataFrame(alarm_summary)
        alarm_summary_df = alarm_summary_df.sort_values('Activations', ascending=False)
        
        print(f"\nTop 10 Most Frequent Alarms:")
        print(alarm_summary_df.head(10).to_string(index=False))
        
        # Save
        alarm_summary_df.to_csv(os.path.join(OUTPUT_DIR, "alarm_analysis.csv"), index=False)
        return alarm_summary_df
    else:
        print("No alarm activations detected")
        return None

def calculate_energy_metrics(df):
    """
    Calculate Coefficient of Performance (COP) and energy efficiency
    """
    print("\n=== ENERGY EFFICIENCY ANALYSIS ===")
    
    # Key energy-related signals
    power_signals = [c for c in df.columns if 'KW' in c.upper() or 'POWER' in c.upper() or 'CURRENT' in c.upper()]
    flow_signals = [c for c in df.columns if 'FI' in c and '_Y' in c]
    temp_signals = [c for c in df.columns if 'TE' in c or 'TI' in c]
    
    print(f"Found {len(power_signals)} power-related signals")
    print(f"Found {len(flow_signals)} flow signals")
    print(f"Found {len(temp_signals)} temperature signals")
    
    # Estimate cooling capacity (simplified)
    # Q = m * Cp * dT
    
    results = {
        'avg_power_signals': len(power_signals),
        'avg_flow_signals': len(flow_signals),
        'avg_temp_range': 0
    }
    
    if len(temp_signals) > 0:
        # Calculate temperature span (max - min across all temp sensors)
        temp_df = df[temp_signals].dropna()
        if len(temp_df) > 0:
            temp_range = temp_df.max().max() - temp_df.min().min()
            results['avg_temp_range'] = temp_range
            print(f"Temperature Operating Range: {temp_range:.1f}°C")
    
    return results

def analyze_heat_exchanger_performance(df):
    """
    Calculate heat exchanger effectiveness and UA value
    """
    print("\n=== HEAT EXCHANGER PERFORMANCE ===")
    
    # Identify inlet/outlet temperature pairs
    # Pattern: Unit_HX_inlet, Unit_HX_outlet
    
    # For BOG heat exchanger (Unit 72)
    hx_temps = {
        'BOG_Inlet': 'ERS_72TE0001_Y',
        'BOG_Outlet': None,  # Need to find
        'N2_Inlet': 'ERS_81TE0005_Y',
        'N2_Outlet': 'ERS_81TE0006_Y'
    }
    
    # Check which exist
    hx_data = {}
    for label, tag in hx_temps.items():
        if tag and tag in df.columns:
            hx_data[label] = df[tag]
            print(f"[FOUND] {label}: {tag}")
        else:
            print(f"[MISSING] {label}")
    
    if len(hx_data) >= 2:
        # Calculate temperature difference
        if 'N2_Inlet' in hx_data and 'N2_Outlet' in hx_data:
            n2_delta = hx_data['N2_Inlet'] - hx_data['N2_Outlet']
            avg_delta = n2_delta.mean()
            print(f"\nN2 Side Temperature Drop: {avg_delta:.2f}°C (avg)")
            print(f"  Min: {n2_delta.min():.2f}°C, Max: {n2_delta.max():.2f}°C")
            
            # Visualize
            fig, ax = plt.subplots(figsize=(12, 4))
            time_axis = [(t - df.index[0]).total_seconds()/3600 for t in n2_delta.index]
            ax.plot(time_axis, n2_delta, 'b-', linewidth=0.5, alpha=0.7)
            ax.axhline(avg_delta, color='r', linestyle='--', label=f'Avg: {avg_delta:.1f}°C')
            ax.set_xlabel('Time (hours)')
            ax.set_ylabel('N2 Temperature Drop (°C)')
            ax.set_title('Heat Exchanger Performance (N2 Side)')
            ax.legend()
            ax.grid(True, alpha=0.3)
            
            save_path = os.path.join(OUTPUT_DIR, "hx_performance.png")
            plt.savefig(save_path, dpi=150, bbox_inches='tight')
            print(f"Saved {save_path}")
            
            return {'N2_avg_delta': avg_delta}
    
    return None

def analyze_valve_cycling(df):
    """
    Track valve movement frequency to identify wear patterns
    """
    print("\n=== VALVE CYCLING ANALYSIS ===")
    
    # Find all valve position signals
    valve_signals = [c for c in df.columns if 'OUTPOS' in c or ('CV' in c and '_Y' in c)]
    
    print(f"Found {len(valve_signals)} valve position signals")
    
    valve_stats = []
    
    for valve in valve_signals[:10]:  # Analyze top 10 for performance
        pos = df[valve]
        
        # Calculate movement (derivative)
        movement = pos.diff().abs()
        
        # Count significant movements (>1% change)
        significant_moves = (movement > 1.0).sum()
        
        # Total travel distance
        total_travel = movement.sum()
        
        # Range of operation
        op_range = pos.max() - pos.min()
        
        valve_stats.append({
            'Valve': valve,
            'Movements': significant_moves,
            'Total_Travel_%': total_travel,
            'Operating_Range_%': op_range,
            'Avg_Position_%': pos.mean()
        })
    
    valve_stats_df = pd.DataFrame(valve_stats)
    valve_stats_df = valve_stats_df.sort_values('Movements', ascending=False)
    
    print("\nTop Valve Activity:")
    print(valve_stats_df.to_string(index=False))
    
    # Save
    valve_stats_df.to_csv(os.path.join(OUTPUT_DIR, "valve_cycling_analysis.csv"), index=False)
    
    return valve_stats_df

def find_time_patterns(df):
    """
    Identify hourly operational patterns
    """
    print("\n=== TIME-OF-DAY PATTERN ANALYSIS ===")
    
    # Add hour column
    df['hour'] = df.index.hour
    
    # Select a representative signal (Mode)
    if 'ERS_MODE_Y' in df.columns:
        mode_by_hour = df.groupby('hour')['ERS_MODE_Y'].agg(['mean', 'std', 'count'])
        
        print("\nMode by Hour of Day:")
        print(mode_by_hour)
        
        # Visualize
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.bar(mode_by_hour.index, mode_by_hour['mean'], alpha=0.7, color='steelblue')
        ax.errorbar(mode_by_hour.index, mode_by_hour['mean'], yerr=mode_by_hour['std'], 
                   fmt='none', ecolor='black', capsize=3)
        ax.set_xlabel('Hour of Day')
        ax.set_ylabel('Average System Mode')
        ax.set_title('Daily Operational Pattern (Mode by Hour)')
        ax.set_xticks(range(0, 24))
        ax.grid(True, alpha=0.3)
        
        save_path = os.path.join(OUTPUT_DIR, "hourly_pattern.png")
        plt.savefig(save_path, dpi=150, bbox_inches='tight')
        print(f"Saved {save_path}")
        
        return mode_by_hour
    
    return None

def generate_correlation_heatmap(df):
    """
    Create advanced correlation heatmap for key signals
    """
    print("\n=== GENERATING CORRELATION HEATMAP ===")
    
    # Select key analog signals (top 20 by variance)
    analog_cols = df.select_dtypes(include=[np.float64, np.int64]).columns
    
    # Calculate variance
    variances = df[analog_cols].var().sort_values(ascending=False)
    top_signals = variances.head(20).index.tolist()
    
    print(f"Selected top 20 high-variance signals for correlation")
    
    # Compute correlation matrix
    corr_matrix = df[top_signals].corr()
    
    # Visualize
    fig, ax = plt.subplots(figsize=(14, 12))
    sns.heatmap(corr_matrix, annot=False, cmap='coolwarm', center=0, 
                square=True, linewidths=0.5, cbar_kws={"shrink": 0.8}, ax=ax)
    ax.set_title('Advanced Correlation Heatmap (Top 20 Signals)', fontsize=14, fontweight='bold')
    
    save_path = os.path.join(OUTPUT_DIR, "advanced_correlation_heatmap.png")
    plt.savefig(save_path, dpi=200, bbox_inches='tight')
    print(f"Saved {save_path}")
    
    # Find strongest unexpected correlations
    # Mask diagonal and lower triangle
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
    corr_matrix_masked = corr_matrix.where(~mask)
    
    # Find top correlations
    corr_flat = corr_matrix_masked.unstack().dropna()
    top_corr = corr_flat.abs().sort_values(ascending=False).head(10)
    
    print("\nTop 10 Unexpected Correlations:")
    for (sig1, sig2), corr_val in top_corr.items():
        print(f"  {sig1} <-> {sig2}: {corr_val:.3f}")
    
    return corr_matrix

def main():
    print("=== DEEP PATTERN MINING & INSIGHT EXPANSION ===\n")
    
    df = load_full_dataset()
    
    # 1. Alarm/Trip Analysis
    alarm_data = scan_alarm_trip_signals(df)
    
    # 2. Energy Efficiency
    energy_metrics = calculate_energy_metrics(df)
    
    # 3. Heat Exchanger Performance
    hx_performance = analyze_heat_exchanger_performance(df)
    
    # 4. Valve Wear Patterns
    valve_cycling = analyze_valve_cycling(df)
    
    # 5. Time-of-Day Patterns
    hourly_patterns = find_time_patterns(df)
    
    # 6. Advanced Correlation
    corr_heatmap = generate_correlation_heatmap(df)
    
    print("\n=== PATTERN MINING COMPLETE ===")
    print("Generated new insights and visualizations")

if __name__ == "__main__":
    main()
