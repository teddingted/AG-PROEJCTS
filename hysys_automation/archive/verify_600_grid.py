import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def reset_all_adjusts(case):
    """Resets all Adjust blocks (ADJ-1 through ADJ-4)."""
    adj_names = ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]
    
    for name in adj_names:
        try:
            adj = case.Flowsheet.Operations.Item(name)
            adj.Reset()
        except:
            pass
    
    return True

def wait_for_solver(case, timeout=60):
    """Waits for solver to complete."""
    start = time.time()
    while True:
        try:
            if not case.Solver.IsSolving:
                break
        except:
            pass
        
        if time.time() - start > timeout:
            break
        
        time.sleep(0.2)
    
    # Stabilization delay
    time.sleep(1.5)
    return True

def get_snapshot(case):
    """Capture all relevant metrics."""
    valid = True
    note = "OK"
    
    # P7 Constraint
    p7 = 99999.0
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        p7 = s7.Pressure.Value  # kPa
        if p7 > 3600.0:
            valid = False
            note = f"P7_High"
    except:
        valid = False
        note = "P7_Fail"
        
    # Min Approach
    min_app = -999.0
    try:
        lng = case.Flowsheet.Operations.Item("LNG-100")
        min_app = lng.MinApproach.Value
        if min_app < 0.0:
            valid = False
            note = f"Cross"
    except:
        valid = False
        note = "App_Fail"

    # Power
    power = 999999.0
    try:
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        val = ss.Cell("C8").CellValue
        if val is not None:
            power = float(val)
    except:
        pass
        
    return {
        'p7': p7,
        'min_app': min_app,
        'power': power,
        'valid': valid,
        'note': note
    }

def main():
    """
    Grid search for 600 kg/h case.
    
    Pressure: 3.0 to 3.8 bar (0.1 bar steps) = 9 points
    Temperature: -112 to -100°C (2°C steps) = 7 points
    Total: 63 combinations
    """
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    print("Connecting to HYSYS...")
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
    
    # Get references
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # Set flow to 600 kg/h
    target_flow = 600.0
    print(f"\n{'='*80}")
    print(f"GRID SEARCH: 600 kg/h, P1: 3.0-3.8 bar, T: -112 to -100°C")
    print(f"Target: Min Approach = 2.0°C")
    print('='*80)
    
    print(f"\nSetting flow to {target_flow} kg/h...")
    s10.MassFlow.Value = target_flow / 3600.0
    wait_for_solver(case, timeout=45)
    reset_all_adjusts(case)
    wait_for_solver(case, timeout=45)
    
    # Open CSV for results
    csv_file = os.path.join(os.getcwd(), folder, "verification_600_grid.csv")
    f = open(csv_file, 'w', newline='')
    writer = csv.writer(f)
    writer.writerow(["P1_bar", "T_ADJ4_C", "MinApproach_C", "P7_bar", "Power_kW", 
                     "App_Error", "Valid", "Note"])
    
    # Define grid
    pressures = [round(3.0 + i * 0.1, 1) for i in range(9)]  # 3.0, 3.1, ..., 3.8
    temperatures = [-112, -110, -108, -106, -104, -102, -100]
    
    target_app = 2.0
    
    print(f"\nGrid: {len(pressures)} pressures × {len(temperatures)} temperatures = {len(pressures)*len(temperatures)} points")
    print(f"Pressures: {pressures}")
    print(f"Temperatures: {temperatures}\n")
    
    best_result = None
    best_error = 999.0
    total_tests = len(pressures) * len(temperatures)
    test_count = 0
    
    for p_bar in pressures:
        p_kpa = p_bar * 100.0
        
        print(f"\n{'='*80}")
        print(f"P1 = {p_bar} bar")
        print('='*80)
        
        for t_c in temperatures:
            test_count += 1
            
            try:
                # Set conditions
                s1.Pressure.Value = p_kpa
                adj4.TargetValue.Value = t_c
                wait_for_solver(case, timeout=30)
                
                # Get snapshot
                snap = get_snapshot(case)
                
                # Calculate error from target
                app_error = abs(snap['min_app'] - target_app)
                
                # Display
                status = ""
                if snap['valid'] and app_error < best_error:
                    best_error = app_error
                    best_result = {
                        'p_bar': p_bar,
                        't': t_c,
                        'snap': snap
                    }
                    status = " ← NEW BEST"
                
                print(f"  [{test_count:2d}/{total_tests}] T={t_c:4.0f}°C → "
                      f"App={snap['min_app']:5.2f}°C (err={app_error:.3f}) "
                      f"P7={snap['p7']/100:5.2f}bar Pwr={snap['power']:6.1f}kW "
                      f"{snap['note']}{status}")
                
                # Write to CSV
                row = [
                    p_bar,
                    t_c,
                    round(snap['min_app'], 2),
                    round(snap['p7']/100.0, 2),
                    round(snap['power'], 2),
                    round(app_error, 3),
                    snap['valid'],
                    snap['note']
                ]
                writer.writerow(row)
                f.flush()
                
            except Exception as e:
                print(f"  [{test_count:2d}/{total_tests}] T={t_c:4.0f}°C → ERROR: {e}")
                writer.writerow([p_bar, t_c, 0, 0, 0, 999, False, "ERROR"])
                f.flush()
    
    # Summary
    print(f"\n{'='*80}")
    print("GRID SEARCH COMPLETE")
    print('='*80)
    
    if best_result:
        br = best_result
        print(f"\nBest configuration found:")
        print(f"  P1 = {br['p_bar']} bar")
        print(f"  T_ADJ4 = {br['t']:.1f}°C")
        print(f"  Min Approach = {br['snap']['min_app']:.2f}°C (target: 2.00°C)")
        print(f"  Error from target = {best_error:.3f}°C")
        print(f"  P7 = {br['snap']['p7']/100:.2f} bar")
        print(f"  Power = {br['snap']['power']:.2f} kW")
        print(f"  Status: {br['snap']['note']}")
    else:
        print("\nNo valid configuration found!")
    
    print(f"\nResults saved to: {csv_file}")
    
    f.close()

if __name__ == "__main__":
    main()
