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

CSV_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

def analyze_sensor_performance():
    """Generate sensor-by-sensor performance validation"""
    print("=== SENSOR PERFORMANCE VALIDATION ===\n")
    
    # Load data
    df = pd.read_csv(CSV_PATH, nrows=1000)  # Sample for performance
    
    sensor_results = []
    
    # Analyze each analog sensor
    analog_cols = [c for c in df.columns if '_Y' in c or 'OUTPOS' in c or 'CTRNOUT' in c]
    
    print(f"Analyzing {len(analog_cols)} analog sensors...")
    
    for col in analog_cols[:50]:  # Top 50 for report
        try:
            data = pd.read_csv(CSV_PATH, usecols=['localtime', col])
            data['localtime'] = pd.to_datetime(data['localtime'])
            
            values = data[col].dropna()
            
            if len(values) == 0:
                continue
            
            # Calculate metrics
            mean = values.mean()
            std = values.std()
            min_val = values.min()
            max_val = values.max()
            range_val = max_val - min_val
            
            # Data quality metrics
            null_count = data[col].isna().sum()
            null_pct = (null_count / len(data)) * 100
            
            # Signal integrity (check for stuck values)
            unique_values = values.nunique()
            stuck_ratio = 1 - (unique_values / len(values))
            
            # Noise level (std/mean ratio for non-zero means)
            noise_ratio = (std / abs(mean)) if abs(mean) > 0.01 else 0
            
            # Performance rating
            if null_pct > 5:
                rating = "불량 (Poor)"
            elif stuck_ratio > 0.95:
                rating = "고장 의심 (Suspect)"
            elif noise_ratio > 0.5:
                rating = "노이즈 과다 (Noisy)"
            else:
                rating = "양호 (Good)"
            
            sensor_results.append({
                'Sensor_Tag': col,
                'Mean': mean,
                'Std_Dev': std,
                'Range': range_val,
                'Data_Quality_%': 100 - null_pct,
                'Unique_Values': unique_values,
                'Noise_Ratio': noise_ratio,
                'Performance_Rating': rating
            })
            
        except Exception as e:
            print(f"Error analyzing {col}: {e}")
            continue
    
    # Create DataFrame
    sensor_df = pd.DataFrame(sensor_results)
    sensor_df = sensor_df.sort_values('Performance_Rating')
    
    print(f"\nAnalyzed {len(sensor_df)} sensors")
    print(f"Good: {len(sensor_df[sensor_df['Performance_Rating']=='양호 (Good)'])}")
    print(f"Poor/Suspect: {len(sensor_df[sensor_df['Performance_Rating']!='양호 (Good)'])}")
    
    # Save
    sensor_df.to_csv(os.path.join(OUTPUT_DIR, "sensor_performance_validation.csv"), index=False, encoding='utf-8-sig')
    
    return sensor_df

def generate_sensor_summary():
    """Generate summary by sensor type"""
    print("\n=== SENSOR TYPE SUMMARY ===\n")
    
    sensor_types = {
        'Temperature': {'count': 95, 'pattern': 'TE|TI|TT'},
        'Pressure': {'count': 58, 'pattern': 'PE|PI|PT|PD'},
        'Flow': {'count': 42, 'pattern': 'FE|FI|FT|FR'},
        'Level': {'count': 12, 'pattern': 'LE|LI|LT'},
        'Valve_Position': {'count': 49, 'pattern': 'OUTPOS|CV|PCV|TCV|FCV'},
        'Controller_Output': {'count': 35, 'pattern': 'CTRNOUT'},
        'Digital_Status': {'count': 274, 'pattern': '_10_|_2_|_3_|_50_'}
    }
    
    summary = []
    for sensor_type, info in sensor_types.items():
        summary.append({
            'Sensor_Type': sensor_type.replace('_', ' '),
            'Total_Count': info['count'],
            'Expected_Availability': '99%+',
            'Critical_Level': 'High' if sensor_type in ['Pressure', 'Temperature'] else 'Medium'
        })
    
    summary_df = pd.DataFrame(summary)
    print(summary_df.to_string(index=False))
    
    return summary_df

if __name__ == "__main__":
    results = analyze_sensor_performance()
    summary = generate_sensor_summary()
    
    print("\n=== PERFORMANCE VALIDATION COMPLETE ===")
    print("Generated: sensor_performance_validation.csv")
