import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def wait_for_solver(case, timeout=60):
    """Wait for solver to finish"""
    start = time.time()
    while True:
        try:
            if not case.Solver.IsSolving:
                break
        except:
            pass
        
        if time.time() - start > timeout:
            print(f"    [WARNING] Solver timeout after {timeout}s")
            break
        
        time.sleep(0.3)
    
    time.sleep(2.0)  # Stabilization
    return True

def reset_all_adjusts(case):
    """Reset all Adjust blocks"""
    print("  Resetting all Adjust blocks...")
    for adj_name in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            adj = case.Flowsheet.Operations.Item(adj_name)
            adj.Reset()
            print(f"    [OK] {adj_name}")
        except:
            print(f"    [SKIP] {adj_name}")

def get_snapshot(case):
    """Get current state snapshot"""
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        p7 = s7.Pressure.Value  # kPa
        min_app = lng.MinApproach.Value  # C
        power = ss.Cell("C8").CellValue  # kW
        
        valid = abs(min_app) < 10000 and abs(p7) < 10000
        
        return {
            'p7': p7,
            'min_app': min_app,
            'power': power,
            'valid': valid
        }
    except:
        return {
            'p7': -32767,
            'min_app': -32767,
            'power': 999999,
            'valid': False
        }

def test_condition(case, flow_kgh, p_bar, t_c):
    """Test a specific P, T condition at given flow"""
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # Set conditions
    s10.MassFlow.Value = flow_kgh / 3600.0
    s1.Pressure.Value = p_bar * 100.0  # Convert to kPa
    adj4.TargetValue.Value = t_c
    
    wait_for_solver(case, timeout=45)
    reset_all_adjusts(case)
    wait_for_solver(case, timeout=45)
    
    snap = get_snapshot(case)
    
    return snap

def main():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    if os.path.exists(file_path):
        case = app.SimulationCases.Open(file_path)
        case.Visible = True
    else:
        case = app.ActiveDocument
    
    print("="*80)
    print("600 kg/h OPTIMIZATION - Pressure Sweep 3.0 to 3.8 bar")
    print("="*80)
    
    # Open CSV
    csv_file = os.path.join(os.getcwd(), folder, "600_sweep_results.csv")
    f = open(csv_file, 'w', newline='')
    writer = csv.writer(f)
    writer.writerow(["P1_bar", "T_ADJ4_C", "MinApproach_C", "P7_bar", "Power_kW", "Status"])
    
    flow = 600
    
    # Pressure sweep: 3.0 to 3.8 bar in 0.1 bar steps
    pressures = [round(p, 1) for p in [3.0 + i*0.1 for i in range(9)]]  # 3.0, 3.1, ..., 3.8
    
    # Temperature candidates to try
    temperatures = [-111.0, -110.0, -109.0, -108.0, -107.0, -106.0, -105.0, -104.0, -103.0, -102.0]
    
    best_result = None
    best_score = 1e20
    
    for p_bar in pressures:
        print(f"\n{'='*80}")
        print(f"Testing P = {p_bar} bar")
        print('='*80)
        
        for t_c in temperatures:
            print(f"  P={p_bar:.1f}b, T={t_c:.1f}C...", end=" ")
            
            snap = test_condition(case, flow, p_bar, t_c)
            
            if snap['valid']:
                app_val = snap['min_app']
                p7_val = snap['p7'] / 100.0
                pwr_val = snap['power']
                
                # Check constraints
                app_ok = 2.0 <= app_val <= 2.5
                p7_ok = p7_val <= 36.0
                
                if app_ok and p7_ok:
                    status = "SUCCESS"
                    score = pwr_val
                    print(f"OK App={app_val:.2f}, P7={p7_val:.2f}b, Pwr={pwr_val:.1f} kW")
                elif app_val < 2.0:
                    status = "App<2.0"
                    score = 1e6 + (2.0 - app_val) * 10000
                    print(f"! App={app_val:.2f} (low), P7={p7_val:.2f}b, Pwr={pwr_val:.1f}")
                elif app_val > 2.5:
                    status = "App>2.5"
                    score = 1e6 + (app_val - 2.5) * 10000
                    print(f"! App={app_val:.2f} (high), P7={p7_val:.2f}b, Pwr={pwr_val:.1f}")
                else:
                    status = "P7>36"
                    score = 1e6 + (p7_val - 36.0) * 10000
                    print(f"! App={app_val:.2f}, P7={p7_val:.2f}b (high), Pwr={pwr_val:.1f}")
                
                # Save to CSV
                writer.writerow([p_bar, t_c, round(app_val, 2), round(p7_val, 2), round(pwr_val, 2), status])
                f.flush()
                
                # Track best
                if score < best_score:
                    best_score = score
                    best_result = (p_bar, t_c, snap, status)
            else:
                print(f"X Invalid/Unconverged")
                writer.writerow([p_bar, t_c, snap['min_app'], snap['p7']/100.0, snap['power'], "INVALID"])
                f.flush()
    
    f.close()
    
    print("\n" + "="*80)
    print("BEST RESULT")
    print("="*80)
    if best_result:
        p, t, snap, status = best_result
        print(f"P1 = {p:.1f} bar")
        print(f"T_ADJ4 = {t:.1f} C")
        print(f"Min Approach = {snap['min_app']:.2f} C")
        print(f"P7 = {snap['p7']/100:.2f} bar")
        print(f"Power = {snap['power']:.2f} kW")
        print(f"Status = {status}")
    else:
        print("No valid solution found!")
    
    print(f"\nResults saved to: {csv_file}")

if __name__ == "__main__":
    main()
