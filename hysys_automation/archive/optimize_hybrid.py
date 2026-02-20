import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def wait_solver_quick(case, timeout=20):
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
    for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            case.Flowsheet.Operations.Item(adj).Reset()
        except:
            pass

def check_smart_range(case):
    """Check if Min Approach is in target range (2.0 - 3.0)"""
    try:
        app = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
        # Strict check for target range
        if 2.0 <= app <= 3.0:
            return 2  # Perfect target
        elif app > 0:
            return 1  # Valid but not target
        else:
            return 0  # Cross
    except:
        return 0

def get_state(case):
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
    try:
        t = s1.Temperature.Value
        if not (0 <= t <= 50):
            s1.Temperature.Value = 40.0
            s1.Pressure.Value = p_kpa
            return True
    except:
        pass
    return False

def scan_valid_range(case, s1, s10, adj4, flow, p_bar, start_t=-90):
    """Ultra-fast scan looking for 2.0-3.0 range"""
    print(f"  [SCAN] Fast checking valid range (target 2.0-3.0)...")
    p_kpa = p_bar * 100.0
    
    # 1. Reset to clean state
    s10.MassFlow.Value = flow / 3600.0
    s1.Pressure.Value = p_kpa
    adj4.TargetValue.Value = start_t
    wait_solver_quick(case, 10) # Faster wait
    reset_adjusts(case)
    wait_solver_quick(case, 10)
    
    # 2. Scan downwards
    scan_temps = range(int(start_t), -121, -5) # 5 deg steps
    valid_points = []
    
    for t in scan_temps:
        adj4.TargetValue.Value = t
        wait_solver_quick(case, 5) # Ultra fast check (5s)
        
        status = check_smart_range(case)
        
        if status == 2: # Perfect range (2.0-3.0)
            print(f"    T={t}: TARGET HIT (App ~ 2-3) -> Stop Scan")
            # Found ideal point, search strictly around here
            return range(t+2, t-3, -1) # e.g. T to T-2
            
        elif status == 1: # Positive but not target
            valid_points.append(t)
            # print(f"    T={t}: OK (>0)")
        else:
            # print(f"    T={t}: Cross")
            if len(valid_points) > 0:
                break # Stop if we hit cross after valid points
    
    if not valid_points:
        return []
    
    # If no perfect target found, use best valid range
    min_t = min(valid_points)
    max_t = max(valid_points)
    print(f"  [RANGE] checking {max_t} to {min_t}")
    return range(max_t, min_t - 6, -1)

def optimize_flow_hybrid(case, flow, p_center, best_prev=None):
    print(f"\n{'='*60}")
    print(f"Flow={flow} kg/h, P={p_center:.1f}bar")
    
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    if flow == 500 or flow == 1500:
        pressures = [p_center]
        print(f"  [FIXED P]")
    else:
        pressures = [round(p_center + i*0.1, 1) for i in range(-4, 5)]
        print(f"  [VAR P: {pressures[0]} to {pressures[-1]}]")
    
    best = None
    best_score = 1e20
    
    for p in pressures:
        # 1. Scan for valid temperature range
        start_t = best_prev['t'] if best_prev else -90
        # Ensure start_t is reasonable
        if start_t < -115: start_t = -110 
        if start_t > -90: start_t = -90
            
        temp_range = scan_valid_range(case, s1, s10, adj4, flow, p, start_t)
        
        if not temp_range:
            print(f"  [SKIP] P={p}: No valid temp range found")
            continue
            
        print(f"  [OPT] Testing {len(temp_range)} temps for P={p}...")
        
        for t in temp_range:
            p_kpa = p * 100.0
            
            # Set conditions
            adj4.TargetValue.Value = t
            s1.Pressure.Value = p_kpa # Ensure pressure set
            wait_solver_quick(case, 15)
            
            # Check cross BEFORE Adjust
            status = check_smart_range(case)
            if status == 0: # Cross
                continue
                
            # No cross, run Adjust
            reset_adjusts(case)
            wait_solver_quick(case, 20)
            
            # Recover S1 if needed
            if recover_s1(s1, p_kpa):
                wait_solver_quick(case, 15)
                reset_adjusts(case)
                wait_solver_quick(case, 15)
                
            state = get_state(case)
            
            if state['valid']:
                app_ok = 2.0 <= state['app'] <= 2.5
                p7_ok = state['p7'] <= 36.0
                
                if app_ok and p7_ok:
                    score = state['pwr']
                    if score < best_score:
                        best_score = score
                        best = {'flow': flow, 'p': p, 't': t,
                               'app': state['app'], 'p7': state['p7'], 'pwr': state['pwr']}
                        print(f"    [BEST] P={p} T={t} App={state['app']:.2f} Pwr={state['pwr']:.1f}")
                elif abs(state['app'] - 2.25) > 5.0:
                    pass # Skip if far off
                    
    # Fallback
    if not best and best_prev:
        print(f"  [FALLBACK] Using previous result")
        best = best_prev.copy()
        
    if best:
        print(f"  RESULT: P={best['p']} T={best['t']} App={best['app']:.2f} Pwr={best['pwr']:.1f}")
    else:
        print(f"  FAILED")
        
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
    
    print("\n" + "="*60)
    print("HYBRID OPTIMIZATION: SCAN -> OPTIMIZE")
    print("="*60)
    
    for iteration in range(1, 4): # 3 Iterations
        print(f"\nIT {iteration}/3")
        for flow in flows:
            p_center = p_init[flow]
            best_prev = results.get(flow)
            
            res = optimize_flow_hybrid(case, flow, p_center, best_prev)
            if res:
                results[flow] = res
                
    # Save
    csv_file = os.path.join(os.getcwd(), folder, "optimization_hybrid.csv")
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["MassFlow", "P1", "T_ADJ4", "MinApp", "P7", "Power", "Status"])
        for flow in flows:
            if flow in results:
                r = results[flow]
                writer.writerow([flow, r['p'], r['t'], round(r['app'],2), round(r['p7'],2), round(r['pwr'],2), "OK"])
            else:
                writer.writerow([flow, 0, 0, 0, 0, 0, "FAIL"])
                
    print(f"\nSaved to {csv_file}")

if __name__ == "__main__":
    main()
