import csv
import sys

def analyze_csv(filename):
    print(f"Analyzing {filename}...")
    try:
        with open(filename, 'r') as f:
            # Manually define headers matching hysys_optimizer_comprehensive.py
            headers = [
                'Flow_Set', 'P_Set', 'T_Adj4_Set', 'Time', 
                'S1_Pres', 'S1_Temp', 'S1_Flow', 'S10_Flow', 'S6_Pres', 
                'Power', 'MA', 'UA', 'LMTD', 'Status'
            ]
            reader = csv.DictReader(f, fieldnames=headers)
            data = list(reader)
    except FileNotFoundError:
        print("File not found.")
        return

    # Filter for 500-700 range
    target_flows = ['500', '600', '700']
    subset = [r for r in data if r['Flow_Set'] in target_flows]
    
    print(f"Total Rows for 500-600 kg/h: {len(subset)}")
    
    # Analysis per Flow
    for flow in target_flows:
        flow_data = [r for r in subset if r['Flow_Set'] == flow]
        if not flow_data:
            print(f"\nFlow {flow}: No data.")
            continue
            
        valid_pts = [r for r in flow_data if r['Status'] == 'OK']
        
        print(f"\nFlow {flow} kg/h:")
        print(f"  Total Attempts: {len(flow_data)}")
        print(f"  Valid Points:   {len(valid_pts)}")
        
        if valid_pts:
            pressures = sorted(set([float(r['P_Set']) for r in valid_pts]))
            temps = sorted(set([float(r['T_Adj4_Set']) for r in valid_pts]))
            ma_vals = [float(r['MA']) for r in valid_pts]
            pwr_vals = [float(r['Power']) for r in valid_pts]
            
            print(f"  Valid Pressures: {min(pressures)} - {max(pressures)} bar")
            print(f"  Valid Temps:     {min(temps)} - {max(temps)} C")
            print(f"  Min Approach:    {min(ma_vals):.2f} - {max(ma_vals):.2f} C")
            print(f"  Power Range:     {min(pwr_vals):.1f} - {max(pwr_vals):.1f} kW")
            
            # Best Point (Min Power)
            best_pt = min(valid_pts, key=lambda x: float(x['Power']))
            print(f"  BEST POINT: P={best_pt['P_Set']} bar, T={best_pt['T_Adj4_Set']} C, Power={best_pt['Power']} kW, MA={best_pt['MA']} C")

if __name__ == "__main__":
    analyze_csv("hysys_automation/optimization_comprehensive_500_1500.csv")
