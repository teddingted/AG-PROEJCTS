
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
import os

# Configuration
FILE_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

# FDS Start Sequence (Sec 3.1.3.3)
# Step01: Set initial conditions
# Step02: Start N2 Compander motor
# Step03: Enable Anti-Surge controllers (51XIC1/2)
# Step04: Close bypass valves (51SCV1/2) with ramp
# Step05: Ramp FRIC10 and TIC041
# Step06: Further ramp FRIC10
# Step07: Enable TIC021 (Temp control)
# Step08: Enable FRIC41

def load_startup_window():
    """Load data around 15:35 startup event"""
    print("Loading startup window data...")
    
    # Tags referenced in FDS Sequence
    seq_tags = [
        'localtime',
        # Step01 tags
        'ERS_51HIC41B_CTRNOUT',  # Should go to 0%
        'ERS_51HIC41A_CTRNOUT',  # Should go to 100%
        'ERS_72XV0004_3_Y',      # Should OPEN
        'ERS_71FIC0001_CTRNOUT', # TRACK mode (0%)
        'ERS_81PIC0003_CTRNOUT', # TRACK mode (50%)
        # Step02 tags
        'ERS_51SET142_Y',        # Full speed indicator
        'ERS_51TIC041_CTRNOUT',  # Should go to AUTO with SP=1.0
        # Step04 tags  
        'ERS_51SCV1_OUTPOS',     # Close with 1%/sec ramp
        'ERS_51SCV2_OUTPOS',     # Close with 1%/sec ramp
        # Step05-08 tags
        'ERS_51FRIC10_CTRNOUT',  # Ramp control
        'ERS_51TIC021_CTRNOUT',  # Enable at Step07
        'ERS_81TE0006_Y',        # Temperature check for Step08
        # Mode
        'ERS_MODE_Y'
    ]
    
    # Check which tags exist
    df_sample = pd.read_csv(FILE_PATH, nrows=5)
    existing_tags = [t for t in seq_tags if t in df_sample.columns]
    print(f"Found {len(existing_tags)}/{len(seq_tags)} sequence tags in data")
    
    # Load full dataset with only these columns
    df = pd.read_csv(FILE_PATH, usecols=existing_tags)
    df['localtime'] = pd.to_datetime(df['localtime'])
    df = df.set_index('localtime')
    
    # Focus on startup window
    start_time = "2024-08-28 15:30:00"
    end_time = "2024-08-28 16:00:00"
    
    df_window = df.loc[start_time:end_time]
    print(f"Loaded {len(df_window)} samples from {start_time} to {end_time}")
    
    return df_window, existing_tags

def detect_sequence_steps(df):
    """
    Detect when each FDS sequence step occurred based on signal patterns
    """
    print("\n=== DETECTING FDS SEQUENCE STEPS ===\n")
    
    events = []
    
    # STEP 01: Initialization (Controllers set to TRACK mode, valves positioned)
    # Signature: 81PIC0003 goes to ~50%, 71FIC0001 goes to 0%
    if 'ERS_81PIC0003_CTRNOUT' in df.columns:
        pic003 = df['ERS_81PIC0003_CTRNOUT']
        # Find when it first reaches 50% (±5%)
        step01_mask = (pic003 > 45) & (pic003 < 55)
        if step01_mask.any():
            step01_time = df[step01_mask].index[0]
            events.append({
                'Step': 'Step01',
                'Description': 'Initialize controllers to TRACK mode',
                'Timestamp': step01_time,
                'Evidence': f'81PIC0003 = {pic003.loc[step01_time]:.1f}%'
            })
            print(f"[OK] Step01 detected at {step01_time}")
            print(f"  Action: Set initial conditions (81PIC0003={pic003.loc[step01_time]:.1f}%)")
    
    # STEP 02: Start Compander Motor
    # Signature: Motor starts running (difficult to detect without motor status signal)
    # Proxy: Look for Step03 activation shortly after
    
    # STEP 03: Enable Anti-Surge Controllers (51XIC1/2)
    # Signature: XIC controllers switch from TRACK to AUTO
    # Proxy: Not directly visible, but can infer from timing
    
    # STEP 04: Close Bypass Valves (51SCV1/2)
    # Signature: SCV1/SCV2 positions ramp down from 100% to <5%
    if 'ERS_51SCV1_OUTPOS' in df.columns:
        scv1 = df['ERS_51SCV1_OUTPOS']
        # Find when SCV1 crosses below 50% (midpoint of ramp)
        scv1_falling = (scv1.shift(1) > 50) & (scv1 < 50)
        if scv1_falling.any():
            step04_time = df[scv1_falling].index[0]
            events.append({
                'Step': 'Step04',
                'Description': 'Close bypass valves (51SCV1/2) with ramp',
                'Timestamp': step04_time,
                'Evidence': f'51SCV1 crossing 50% at {scv1.loc[step04_time]:.1f}%'
            })
            print(f"[OK] Step04 detected at {step04_time}")
            print(f"  Action: Bypass valve ramping closed (51SCV1={scv1.loc[step04_time]:.1f}%)")
    
    # STEP 05: Ramp FRIC10 and TIC041
    # Signature: FRIC10 starts ramping from 0% towards 5%
    if 'ERS_51FRIC10_CTRNOUT' in df.columns:
        fric10 = df['ERS_51FRIC10_CTRNOUT']
        # Find when FRIC10 first rises above 1%
        fric10_rising = (fric10.shift(1) < 1) & (fric10 > 1)
        if fric10_rising.any():
            step05_time = df[fric10_rising].index[0]
            events.append({
                'Step': 'Step05',
                'Description': 'Ramp FRIC10 (0% → 10%/min)',
                'Timestamp': step05_time,
                'Evidence': f'FRIC10 rising from {fric10.shift(1).loc[step05_time]:.1f}% to {fric10.loc[step05_time]:.1f}%'
            })
            print(f"[OK] Step05 detected at {step05_time}")
            print(f"  Action: Flow ramp initiated (FRIC10={fric10.loc[step05_time]:.1f}%)")
    
    # STEP 06: Continue FRIC10 ramp (slower rate: 1.5%/min)
    # (Hard to distinguish from Step05 without rate calculation)
    
    # STEP 07: Enable TIC021 (Temperature controller)
    # Signature: TIC021 switches to AUTO mode with SP=-130°C
    if 'ERS_51TIC021_CTRNOUT' in df.columns:
        tic021 = df['ERS_51TIC021_CTRNOUT']
        # Find when TIC021 becomes active (moves away from initial 5%)
        tic021_active = (tic021.diff().abs() > 1)  # Significant change
        if tic021_active.any():
            step07_time = df[tic021_active].index[0]
            events.append({
                'Step': 'Step07',
                'Description': 'Enable TIC021 (Temp controller to AUTO)',
                'Timestamp': step07_time,
                'Evidence': f'TIC021 activated at {tic021.loc[step07_time]:.1f}%'
            })
            print(f"[OK] Step07 detected at {step07_time}")
            print(f"  Action: Temperature control enabled (TIC021={tic021.loc[step07_time]:.1f}%)")
    
    # STEP 08: Enable FRIC41 (Final flow controller)
    # Signature: FRIC41 goes to AUTO with SP=1.0
    # Condition: 81TE0006 <= -129°C
    if 'ERS_81TE0006_Y' in df.columns:
        te0006 = df['ERS_81TE0006_Y']
        # Find when temperature drops below -129°C
        temp_ready = te0006 < -129
        if temp_ready.any():
            step08_time = df[temp_ready].index[0]
            events.append({
                'Step': 'Step08',
                'Description': 'Enable FRIC41 (Final flow control)',
                'Timestamp': step08_time,
                'Evidence': f'81TE0006 = {te0006.loc[step08_time]:.1f}°C (< -129°C threshold)'
            })
            print(f"[OK] Step08 detected at {step08_time}")
            print(f"  Action: Final flow controller enabled (Temp={te0006.loc[step08_time]:.1f}°C)")
    
    # Create timeline DataFrame
    if events:
        timeline_df = pd.DataFrame(events)
        timeline_df = timeline_df.sort_values('Timestamp')
        
        # Calculate time deltas between steps
        timeline_df['Delta (s)'] = timeline_df['Timestamp'].diff().dt.total_seconds()
        
        print("\n=== SEQUENCE TIMELINE ===")
        print(timeline_df.to_string(index=False))
        
        # Save
        save_path = os.path.join(OUTPUT_DIR, "fds_sequence_validation.csv")
        timeline_df.to_csv(save_path, index=False)
        print(f"\nSaved timeline to {save_path}")
        
        return timeline_df
    else:
        print("⚠️ No sequence steps detected (missing tags)")
        return None

def analyze_mode_transitions(df):
    """
    Track ERS_MODE_Y changes during startup
    """
    print("\n=== MODE TRANSITIONS DURING STARTUP ===\n")
    
    if 'ERS_MODE_Y' not in df.columns:
        print("⚠️ MODE signal not found")
        return
    
    mode = df['ERS_MODE_Y']
    mode_changes = mode[mode.diff() != 0]
    
    for timestamp, value in mode_changes.items():
        print(f"{timestamp}: MODE changed to {int(value)}")

def main():
    df, tags = load_startup_window()
    
    # Detect FDS sequence steps
    timeline = detect_sequence_steps(df)
    
    # Track mode changes
    analyze_mode_transitions(df)
    
    print("\n=== ANALYSIS COMPLETE ===")
    print(f"Found {len(tags)} relevant tags in data")
    if timeline is not None:
        print(f"Detected {len(timeline)} sequence steps")
        total_duration = (timeline['Timestamp'].max() - timeline['Timestamp'].min()).total_seconds()
        print(f"Total sequence duration: {total_duration:.0f} seconds ({total_duration/60:.1f} minutes)")

if __name__ == "__main__":
    main()
