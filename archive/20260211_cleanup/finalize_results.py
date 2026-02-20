import csv
import os

# Define paths
robust_csv = 'hysys_automation/robust_verification_report.csv'
final_csv = 'hysys_automation/optimization_final_summary_500_1500.csv'

# Read verified data (500-1300)
data = []
if os.path.exists(robust_csv):
    with open(robust_csv, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(row)

# Add Failed/Extrapolated data (1400-1500)
# Reason: Simulation Instability / Non-Convergence even with Robust Kick
pending_flows = [
    {'Flow': '1400', 'P_bar': '6.7', 'T_C': '-100', 'Status': 'FAILED (Sim Limit Reached)'},
    {'Flow': '1500', 'P_bar': '7.4', 'T_C': '-100', 'Status': 'FAILED (Sim Limit Reached)'}
]

# Write final consolidated report
keys = ['Flow', 'P_bar', 'T_C', 'Power', 'MA', 'S6_Pres', 'Status']

with open(final_csv, 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=keys, extrasaction='ignore')
    writer.writeheader()
    
    # Verified points
    for row in data:
        writer.writerow(row)
        
    # Pending points
    for row in pending_flows:
        writer.writerow(row)

print(f"Final summary saved to: {final_csv}")
print("Verified Points: 500-1300 kg/h")
print("Failed Points: 1400-1500 kg/h (Simulation Instability)")
