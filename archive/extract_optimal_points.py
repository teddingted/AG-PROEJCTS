import csv
import os

# Read full dataset
csv_path = 'hysys_automation/optimization_comprehensive_500_1500.csv'
data = list(csv.DictReader(open(csv_path)))

# Group by flow
flows = {}
for row in data:
    if row.get('Status') == 'OK':
        flow = row['Flow_Set']
        if flow not in flows:
            flows[flow] = []
        flows[flow].append(row)

# Extract optimal point for each flow
optimal_points = []
for flow in sorted(flows.keys(), key=float):
    points = flows[flow]
    
    # Find point with minimum power and MA > 0.5
    valid_points = [p for p in points if float(p.get('MA', -999)) > 0.5]
    
    if not valid_points:
        print(f"Warning: No valid MA points for {flow} kg/h")
        continue
    
    # Sort by Power
    best = min(valid_points, key=lambda p: float(p.get('Power', 99999)))
    
    optimal_points.append({
        'MassFlow': flow,
        'P_bar': round(float(best['S1_Pres']), 2),
        'T_C': round(float(best['T_Adj4_Set']), 1),
        'Power_kW': round(float(best['Power']), 2),
        'MA_C': round(float(best['MA']), 2),
        'S6_Pres_bar': round(float(best['S6_Pres']), 2)
    })

# Extrapolate for 1400 and 1500 kg/h
# Linear fit from 1200, 1300
if len(optimal_points) >= 2:
    p1200 = next((p for p in optimal_points if float(p['MassFlow']) == 1200), None)
    p1300 = next((p for p in optimal_points if float(p['MassFlow']) == 1300), None)
    
    if p1200 and p1300:
        # Slope calculation
        dp_dflow = (p1300['P_bar'] - p1200['P_bar']) / 100.0
        dt_dflow = (p1300['T_C'] - p1200['T_C']) / 100.0
        
        for flow in [1400, 1500]:
            delta = flow - 1300
            optimal_points.append({
                'MassFlow': str(flow),
                'P_bar': round(p1300['P_bar'] + dp_dflow * delta, 2),
                'T_C': round(p1300['T_C'] + dt_dflow * delta, 1),
                'Power_kW': 0.0,  # Unknown
                'MA_C': 0.0,      # Unknown
                'S6_Pres_bar': 0.0,  # Unknown
                'Note': 'EXTRAPOLATED'
            })

# Write to CSV
output_path = 'hysys_automation/optimization_final_result.csv'
with open(output_path, 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=['MassFlow', 'P_bar', 'T_C', 'Power_kW', 'MA_C', 'S6_Pres_bar', 'Note'])
    writer.writeheader()
    for p in optimal_points:
        if 'Note' not in p:
            p['Note'] = 'VERIFIED'
        writer.writerow(p)

print(f"[OK] Extracted {len(optimal_points)} optimal points")
print(f"[OK] Saved to: {output_path}")
print("\n=== Optimal Operating Points ===")
for p in optimal_points:
    note = p.get('Note', '')
    if note == 'EXTRAPOLATED':
        print(f"{p['MassFlow']:>5} kg/h: P={p['P_bar']:>4.1f} bar, T={p['T_C']:>5.0f}°C  [{note}]")
    else:
        print(f"{p['MassFlow']:>5} kg/h: P={p['P_bar']:>4.1f} bar, T={p['T_C']:>5.0f}°C, Power={p['Power_kW']:>6.1f} kW, MA={p['MA_C']:>4.2f}°C")
