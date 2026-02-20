import csv
import os

# Define paths
robust_csv = 'hysys_automation/robust_verification_report.csv'
v4_csv = 'hysys_automation/high_flow_v4_results.csv'
v5_csv = 'hysys_automation/high_flow_v5_results.csv'
final_csv = 'hysys_automation/optimization_final_summary_verified.csv'

# Dictionary to store unique flow data (Key: Flow)
final_data = {}

# 1. Load 500-1300 (Robust)
if os.path.exists(robust_csv):
    with open(robust_csv, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            flow = int(float(row['Flow']))
            final_data[flow] = row

# 2. Load 1500 (V4)
if os.path.exists(v4_csv):
    with open(v4_csv, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            flow = int(float(row['Flow']))
            if flow == 1500:
                row['Status'] = 'Verified (Anchor)'
                final_data[flow] = row

# 3. Load 1400 (V5)
if os.path.exists(v5_csv):
    with open(v5_csv, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            flow = int(float(row['Flow']))
            if flow == 1400:
                final_data[flow] = row

# Sort by Flow
sorted_flows = sorted(final_data.keys())

# Write final consolidated report
keys = ['Flow', 'P_bar', 'T_C', 'Power', 'MA', 'S6_Pres', 'Status']

with open(final_csv, 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=keys, extrasaction='ignore')
    writer.writeheader()
    
    for flow in sorted_flows:
        writer.writerow(final_data[flow])

print(f"Final summary saved to: {final_csv}")
print(f"Total Verified Points: {len(sorted_flows)}")
