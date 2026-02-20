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
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np

# Configuration
CSV_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
SEQUENCE_CSV = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output\fds_sequence_validation.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

plt.style.use('seaborn-v0_8-darkgrid')

def load_sequence_timeline():
    """Load detected FDS sequence steps"""
    print("Loading FDS sequence timeline...")
    df = pd.read_csv(SEQUENCE_CSV)
    df['Timestamp'] = pd.to_datetime(df['Timestamp'])
    print(f"Loaded {len(df)} detected steps")
    return df

def load_signal_data():
    """Load key signals for workflow visualization"""
    print("Loading signal data...")
    
    key_signals = [
        'localtime',
        'ERS_MODE_Y',
        'ERS_81PIC0003_CTRNOUT',
        'ERS_81PIC0003_BSP',
        'ERS_51PIT011A_Y',
        'ERS_51SCV1_OUTPOS',
        'ERS_51SCV2_OUTPOS',
        'ERS_51FRIC10_CTRNOUT',
        'ERS_51TIC021_CTRNOUT',
        'ERS_71FIC0001_CTRNOUT',
        'ERS_81TE0006_Y'
    ]
    
    # Load only these columns
    df_sample = pd.read_csv(CSV_PATH, nrows=1)
    existing = [c for c in key_signals if c in df_sample.columns]
    
    df = pd.read_csv(CSV_PATH, usecols=existing)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    
    print(f"Loaded {len(df)} samples with {len(existing)} signals")
    return df

def create_workflow_gantt(sequence_df, signal_df):
    """
    Create Gantt-style timeline showing FDS workflow execution
    """
    print("\nCreating FDS Workflow Timeline...")
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 10), height_ratios=[1, 3])
    
    # Define FDS sequence steps (from FDS Sec 3.1.3.3)
    fds_steps = {
        'Step01': {'name': 'Initialize Controllers', 'duration': 150, 'color': '#3498db'},
        'Step02': {'name': 'Start N2 Compander', 'duration': 300, 'color': '#e74c3c'},
        'Step03': {'name': 'Enable Anti-Surge', 'duration': 10, 'color': '#f39c12'},
        'Step04': {'name': 'Close Bypass Valves', 'duration': 100, 'color': '#9b59b6'},
        'Step05': {'name': 'Ramp FRIC10', 'duration': 240, 'color': '#1abc9c'},
        'Step06': {'name': 'Fine Ramp FRIC10', 'duration': 10, 'color': '#16a085'},
        'Step07': {'name': 'Enable TIC021', 'duration': None, 'color': '#27ae60'},
        'Step08': {'name': 'Enable FRIC41', 'duration': None, 'color': '#2ecc71'}
    }
    
    # Top panel: FDS Expected Timeline
    ax1.set_title('FDS Expected Sequence (Cool-down)', fontsize=14, fontweight='bold')
    ax1.set_xlim(0, 800)
    ax1.set_ylim(0, 1)
    
    cumulative_time = 0
    for step_id, info in fds_steps.items():
        if info['duration']:
            rect = FancyBboxPatch(
                (cumulative_time, 0.3), info['duration'], 0.4,
                boxstyle="round,pad=0.01", 
                edgecolor='black', 
                facecolor=info['color'],
                alpha=0.7,
                linewidth=2
            )
            ax1.add_patch(rect)
            
            # Add text label
            ax1.text(
                cumulative_time + info['duration']/2, 0.5,
                f"{step_id}\n{info['name']}\n({info['duration']}s)",
                ha='center', va='center',
                fontsize=9, fontweight='bold',
                color='white'
            )
            
            cumulative_time += info['duration']
    
    ax1.set_xlabel('Time (seconds from start)', fontsize=11)
    ax1.set_yticks([])
    ax1.grid(True, axis='x', alpha=0.3)
    
    # Bottom panel: Actual Execution Timeline
    ax2.set_title('Actual Execution (Data-Driven)', fontsize=14, fontweight='bold')
    
    if len(sequence_df) > 0:
        start_time = sequence_df['Timestamp'].min()
        
        for idx, row in sequence_df.iterrows():
            step_id = row['Step']
            timestamp = row['Timestamp']
            
            # Calculate relative time from start
            rel_time = (timestamp - start_time).total_seconds()
            
            # Get step info
            step_info = fds_steps.get(step_id, {'name': 'Unknown', 'color': 'gray'})
            
            # Draw vertical bar
            ax2.axvline(rel_time, color=step_info['color'], linewidth=3, alpha=0.7, linestyle='--')
            
            # Add label
            ax2.text(
                rel_time, 0.95,
                f"{step_id}\n{timestamp.strftime('%H:%M:%S')}",
                ha='center', va='top',
                fontsize=9,
                bbox=dict(boxstyle='round', facecolor=step_info['color'], alpha=0.5)
            )
    
    # Overlay signal behavior
    if 'ERS_MODE_Y' in signal_df.columns:
        mode_signal = signal_df['ERS_MODE_Y']
        start_time = sequence_df['Timestamp'].min()
        end_time = start_time + pd.Timedelta(seconds=800)
        
        mode_slice = mode_signal.loc[start_time:end_time]
        
        if len(mode_slice) > 0:
            time_axis = [(t - start_time).total_seconds() for t in mode_slice.index]
            
            ax2_twin = ax2.twinx()
            ax2_twin.plot(time_axis, mode_slice.values, 'k-', linewidth=2, label='System Mode')
            ax2_twin.set_ylabel('System Mode', fontsize=11)
            ax2_twin.set_ylim(-0.5, 3.5)
            ax2_twin.legend(loc='upper right')
    
    ax2.set_xlabel('Time (seconds from start)', fontsize=11)
    ax2.set_xlim(0, 800)
    ax2.set_ylim(0, 1)
    ax2.set_yticks([])
    ax2.grid(True, axis='x', alpha=0.3)
    
    plt.tight_layout()
    save_path = os.path.join(OUTPUT_DIR, "fds_workflow_timeline.png")
    plt.savefig(save_path, dpi=200, bbox_inches='tight')
    print(f"Saved {save_path}")
    
    return save_path

def create_signal_progression(sequence_df, signal_df):
    """
    Create multi-panel plot showing how each signal changes during workflow
    """
    print("\nCreating Signal Progression Chart...")
    
    if len(sequence_df) == 0:
        print("No sequence steps detected, skipping signal progression")
        return None
    
    start_time = sequence_df['Timestamp'].min()
    end_time = start_time + pd.Timedelta(minutes=20)
    
    data_slice = signal_df.loc[start_time:end_time]
    
    # Define signal groups to plot
    signal_groups = [
        {
            'title': 'Pressure Control (81PIC0003)',
            'signals': [
                ('ERS_81PIC0003_BSP', 'Setpoint (SP)', 'r--'),
                ('ERS_51PIT011A_Y', 'Process Value (PV)', 'b-'),
                ('ERS_81PIC0003_CTRNOUT', 'Controller Output (OP)', 'g-')
            ],
            'ylabel': 'Pressure / %'
        },
        {
            'title': 'Bypass Valves (Step04)',
            'signals': [
                ('ERS_51SCV1_OUTPOS', 'SCV1 Position', 'purple'),
                ('ERS_51SCV2_OUTPOS', 'SCV2 Position', 'orange')
            ],
            'ylabel': 'Valve Position (%)'
        },
        {
            'title': 'Flow & Temperature Controllers',
            'signals': [
                ('ERS_51FRIC10_CTRNOUT', 'FRIC10 (Step05)', 'teal'),
                ('ERS_51TIC021_CTRNOUT', 'TIC021 (Step07)', 'brown'),
                ('ERS_71FIC0001_CTRNOUT', 'FIC0001 (BOG)', 'navy')
            ],
            'ylabel': 'Controller Output (%)'
        }
    ]
    
    num_panels = len(signal_groups)
    fig, axes = plt.subplots(num_panels, 1, figsize=(14, 4*num_panels), sharex=True)
    
    if num_panels == 1:
        axes = [axes]
    
    for idx, group in enumerate(signal_groups):
        ax = axes[idx]
        ax.set_title(group['title'], fontsize=12, fontweight='bold')
        
        for sig_name, label, style in group['signals']:
            if sig_name in data_slice.columns:
                time_axis = [(t - start_time).total_seconds()/60 for t in data_slice.index]
                ax.plot(time_axis, data_slice[sig_name], style, label=label, linewidth=1.5)
        
        # Add vertical lines for sequence steps
        for _, row in sequence_df.iterrows():
            step_time = (row['Timestamp'] - start_time).total_seconds() / 60
            ax.axvline(step_time, color='red', linestyle=':', alpha=0.5, linewidth=1)
            ax.text(step_time, ax.get_ylim()[1]*0.95, row['Step'], 
                   rotation=90, va='top', fontsize=8, color='red')
        
        ax.set_ylabel(group['ylabel'], fontsize=10)
        ax.legend(loc='best', fontsize=9)
        ax.grid(True, alpha=0.3)
    
    axes[-1].set_xlabel('Time from Start (minutes)', fontsize=11)
    
    plt.tight_layout()
    save_path = os.path.join(OUTPUT_DIR, "fds_signal_progression.png")
    plt.savefig(save_path, dpi=200, bbox_inches='tight')
    print(f"Saved {save_path}")
    
    return save_path

def main():
    print("=== FDS WORKFLOW VISUALIZATION ===\n")
    
    # Load data
    sequence_df = load_sequence_timeline()
    signal_df = load_signal_data()
    
    # Create visualizations
    gantt_path = create_workflow_gantt(sequence_df, signal_df)
    progression_path = create_signal_progression(sequence_df, signal_df)
    
    print("\n=== VISUALIZATION COMPLETE ===")
    print(f"Generated workflows:")
    print(f"  1. Timeline: {gantt_path}")
    print(f"  2. Signal Progression: {progression_path}")

if __name__ == "__main__":
    main()
