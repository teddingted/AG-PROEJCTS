import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def wait_solver_quick(case, timeout=20):
    """Quick solver wait"""
    start = time.time()
    while time.time() - start < timeout:
        try:
            if not case.Solver.IsSolving:
                time.sleep(0.5)
                return True
        except:
            pass
        time.sleep(0.15)
    return False

def reset_adjusts(case):
    """Reset all Adjust blocks"""
    for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            case.Flowsheet.Operations.Item(adj).Reset()
        except:
            pass

def check_temp_cross(case):
    """Check for temperature cross"""
    try:
        lng = case.Flowsheet.Operations.Item("LNG-100")
        app = lng.MinApproach.Value
        return app > 0
    except:
        return False

def get_state(case):
    """Get full state"""
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        app = lng.MinApproach.Value
        p7 = s7.Pressure.Value / 100.0
        pwr = ss.Cell("C8").CellValue
        
        valid = abs(app) < 1000 and abs(p7) < 1000 and app > 0
        return {'app': app, 'p7': p7, 'pwr': pwr, 'valid': valid}
    except:
        return {'app': -9999, 'p7': -9999, 'pwr': 99999, 'valid': False}

def recover_s1(s1, p_kpa):
    """Recover Stream 1 if abnormal"""
    try:
        t = s1.Temperature.Value
        if not (0 <= t <= 50):
            s1.Temperature.Value = 40.0
            s1.Pressure.Value = p_kpa
            return True
    except:
        pass
    return False

def find_valid_temp_range(case, s1, s10, adj4, flow, p_bar, prev_good_t=None):
    """Find temperature range without crossing using 5°C steps"""
    print(f"  [SCAN] Finding valid temp range...")
    
    p_kpa = p_bar * 100.0
    
    # Start from previous good temperature or -120
    start_t = prev_good_t if prev_good_t else -120
    
    # Scan from -120 to -90 in 5°C steps
    scan_temps = list(range(-120, -85, 5))  # -120, -115, -110, -105, -100, -95, -90
    
    valid_temps = []
    
    for t in scan_temps:
        # Reset to previous converged state
        if prev_good_t:
            s10.MassFlow.Value = flow / 3600.0
            s1.Pressure.Value = p_kpa
            adj4.TargetValue.Value = prev_good_t
            wait_solver_quick(case, 15)
        
        # Set new temperature
        adj4.TargetValue.Value = t
        wait_solver_quick(case, 15)
        
        # Check for cross WITHOUT Adjust
        if check_temp_cross(case):
            valid_temps.append(t)
            print(f"    T={t}°C: OK (no cross)")
        else:
            print(f"    T={t}°C: CROSS")
    
    if valid_temps:
        # Found valid region, expand around it with 1°C steps
        center = valid_temps[len(valid_temps)//2]
        expanded = list(range(center-7, center+8))  # ±7°C around center
        print(f"  [EXPAND] Valid region found, expanding around T={center}°C")
        return expanded
    else:
        print(f"  [WARN] No valid temps found in scan, using full range")
        return list(range(-120, -89))

def optimize_flow_v3(case, flow, p_center, best_prev=None):
    """Optimized with intelligent temp range finding"""
    print(f"\n{'='*60}")
    print(f"Flow={flow} kg/h, P={p_center:.1f}bar")
    
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # Pressure range
    if flow == 500 or flow == 1500:
        pressures = [p_center]
        print(f"  [FIXED P={p_center:.1f}]")
    else:
        pressures = [round(p_center + i*0.1, 1) for i in range(-4, 5)]
        print(f"  [VAR P={p_center-0.4:.1f}-{p_center+0.4:.1f}]")
    
    best = None
    best_score = 1e20
    total_tested = 0
    
    for p in pressures:
        # Find valid temperature range for this pressure
        prev_t = best_prev['t'] if best_prev else None
        temp_range = find_valid_temp_range(case, s1, s10, adj4, flow, p, prev_t)
        
        print(f"  [P={p:.1f}] Testing {len(temp_range)} temps...")
        
        for t in temp_range:
            # Reset to previous good state
            if best_prev:
                s10.MassFlow.Value = flow / 3600.0
                s1.Pressure.Value = best_prev['p'] * 100.0
                adj4.TargetValue.Value = best_prev['t']
                wait_solver_quick(case, 15)
            
            # Set new conditions
            p_kpa = p * 100.0
            s10.MassFlow.Value = flow / 3600.0
            s1.Pressure.Value = p_kpa
            adj4.TargetValue.Value = t
            
            wait_solver_quick(case, 20)
            
            # Check cross before Adjust
            if not check_temp_cross(case):
                continue
            
            # No cross, proceed with Adjust
            reset_adjusts(case)
            wait_solver_quick(case, 20)
            
            if recover_s1(s1, p_kpa):
                wait_solver_quick(case, 15)
                reset_adjusts(case)
                wait_solver_quick(case, 15)
            
            state = get_state(case)
            total_tested += 1
            
            if state['valid']:
                app_ok = 2.0 <= state['app'] <= 2.5
                p7_ok = state['p7'] <= 36.0
                
                if app_ok and p7_ok:
                    score = state['pwr']
                    if score < best_score:
                        best_score = score
                        best = {'flow': flow, 'p': p, 't': t,
                               'app': state['app'], 'p7': state['p7'], 'pwr': state['pwr']}
                        print(f"    [BEST] T={t:.0f} App={state['app']:.2f} Pwr={state['pwr']:.0f}")
    
    # Fallback
    if not best and best_prev:
        print(f"  [FALLBACK] Using previous result")
        best = best_prev.copy()
        best['flow'] = flow
    
    if best:
        print(f"  OK: P={best['p']:.1f} T={best['t']:.0f} App={best['app']:.2f} Pwr={best['pwr']:.0f} ({total_tested} tests)")
    else:
        print(f"  FAIL ({total_tested} tests)")
    
    return best

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
    
    p_init = {
        500: 3.0, 600: 3.5, 700: 4.1, 800: 4.3, 900: 4.8, 1000: 5.2,
        1100: 5.7, 1200: 6.1, 1300: 6.5, 1400: 6.9, 1500: 7.4
    }
    
    flows = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
    results = {}
    best_prev = None
    
    print("\n" + "="*60)
    print("SMART OPTIMIZATION WITH RANGE FINDING")
    print("="*60)
    
    for iteration in range(1, 4):
        print(f"\n{'#'*60}")
        print(f"# ITERATION {iteration}/3")
        print(f"{'#'*60}")
        
        for flow in flows:
            # ALWAYS use Pre-set Table for pressure center
            p_center = p_init[flow]
            
            # For iteration 1, use wide scan or default center
            # For subsequent iterations, we can use result as center but re-scan to be safe
            if flow in results and results[flow] and iteration > 1:
                best_prev = results[flow]
            else:
                best_prev = None  # Force fresh scan from default
            
            result = optimize_flow_v3(case, flow, p_center, best_prev)
            
            if result:
                results[flow] = result
    
    # Save
    csv_file = os.path.join(os.getcwd(), folder, "optimization_final.csv")
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["MassFlow_10_kg_h", "P1_bar", "ADJ4_Target_C",
                        "MinApproach_C", "P7_bar", "Power_kW", "Status"])
        
        for flow in flows:
            if flow in results and results[flow]:
                r = results[flow]
                writer.writerow([flow, round(r['p'], 1), round(r['t'], 1),
                               round(r['app'], 2), round(r['p7'], 2),
                               round(r['pwr'], 2), "OK"])
            else:
                writer.writerow([flow, 0, 0, 0, 0, 0, "FAILED"])
    
    print(f"\n{'='*60}")
    print(f"COMPLETE: {csv_file}")
    print('='*60)
    
    print("\nRESULTS:")
    for flow in flows:
        if flow in results and results[flow]:
            r = results[flow]
            print(f"{flow:4d}: P={r['p']:.1f} T={r['t']:.0f} App={r['app']:.2f} Pwr={r['pwr']:.0f}")

if __name__ == "__main__":
    main()
