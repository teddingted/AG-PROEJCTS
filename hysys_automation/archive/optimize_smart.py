import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def wait_solver_fast(case, timeout=30):
    """Fast convergence check"""
    start = time.time()
    while time.time() - start < timeout:
        try:
            if not case.Solver.IsSolving:
                time.sleep(0.8)
                return True
        except:
            pass
        time.sleep(0.2)
    return False

def reset_adjusts(case):
    """Reset all Adjust blocks"""
    for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            case.Flowsheet.Operations.Item(adj).Reset()
        except:
            pass

def recover_s1_temp(s1, p_kpa):
    """Recover Stream 1 if temperature is abnormal"""
    try:
        t = s1.Temperature.Value
        if not (0 <= t <= 50):
            s1.Temperature.Value = 40.0
            s1.Pressure.Value = p_kpa
            return True
    except:
        pass
    return False

def get_state(case):
    """Get current state"""
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        app = lng.MinApproach.Value
        p7 = s7.Pressure.Value / 100.0
        pwr = ss.Cell("C8").CellValue
        
        valid = abs(app) < 1000 and abs(p7) < 1000 and app > 0  # app > 0 to avoid crossing
        return {'app': app, 'p7': p7, 'pwr': pwr, 'valid': valid}
    except:
        return {'app': -9999, 'p7': -9999, 'pwr': 99999, 'valid': False}

def test_point(case, s1, s10, adj4, flow, p_bar, t_c):
    """Test single point with recovery"""
    p_kpa = p_bar * 100.0
    
    s10.MassFlow.Value = flow / 3600.0
    s1.Pressure.Value = p_kpa
    adj4.TargetValue.Value = t_c
    
    wait_solver_fast(case, 25)
    reset_adjusts(case)
    wait_solver_fast(case, 25)
    
    if recover_s1_temp(s1, p_kpa):
        wait_solver_fast(case, 20)
        reset_adjusts(case)
        wait_solver_fast(case, 20)
    
    return get_state(case)

def optimize_flow_smart(case, flow, p_center, best_prev=None):
    """Smart optimization with adaptive temperature range"""
    print(f"\n{'='*60}")
    print(f"Flow = {flow} kg/h, P = {p_center:.1f} bar")
    
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # Pressure range
    if flow == 500 or flow == 1500:
        pressures = [p_center]
        print(f"  [FIXED P]")
    else:
        pressures = [round(p_center + i*0.1, 1) for i in range(-4, 5)]
        print(f"  [VAR P: {p_center-0.4:.1f}-{p_center+0.4:.1f}]")
    
    # Adaptive temperature range based on previous result
    if best_prev and 'app' in best_prev:
        # Narrow range around previous optimal
        t_center = best_prev['t']
        t_range = list(range(int(t_center)-5, int(t_center)+6, 1))  # ±5°C
        print(f"  [ADAPTIVE T: {t_center-5:.0f} to {t_center+5:.0f}]")
    else:
        # Wide initial range
        t_range = list(range(-115, -100))  # -115 to -100
        print(f"  [INITIAL T: -115 to -100]")
    
    best = None
    best_score = 1e20
    tested = 0
    
    for p in pressures:
        for t in t_range:
            tested += 1
            state = test_point(case, s1, s10, adj4, flow, p, t)
            
            if state['valid']:  # Already filters out crossing (app > 0)
                app_ok = 2.0 <= state['app'] <= 2.5
                p7_ok = state['p7'] <= 36.0
                
                if app_ok and p7_ok:
                    score = state['pwr']
                    if score < best_score:
                        best_score = score
                        best = {'flow': flow, 'p': p, 't': t, 
                               'app': state['app'], 'p7': state['p7'], 'pwr': state['pwr']}
                        print(f"  [BEST] P={p:.1f} T={t:.0f} App={state['app']:.2f} Pwr={state['pwr']:.0f}")
    
    # Fallback
    if not best and best_prev:
        print(f"  [FALLBACK] P={best_prev['p']:.1f} T={best_prev['t']:.0f}")
        p_kpa = best_prev['p'] * 100.0
        recover_s1_temp(s1, p_kpa)
        s10.MassFlow.Value = flow / 3600.0
        s1.Pressure.Value = p_kpa
        adj4.TargetValue.Value = best_prev['t']
        wait_solver_fast(case, 25)
        reset_adjusts(case)
        wait_solver_fast(case, 25)
        best = best_prev.copy()
        best['flow'] = flow
    
    if best:
        print(f"  RESULT: P={best['p']:.1f} T={best['t']:.0f} App={best['app']:.2f} Pwr={best['pwr']:.0f} ({tested} tests)")
    else:
        print(f"  FAILED ({tested} tests)")
    
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
    
    # Known good starting points
    p_init = {
        500: 3.0, 600: 3.5, 700: 4.1, 800: 4.3, 900: 4.8, 1000: 5.2,
        1100: 5.7, 1200: 6.1, 1300: 6.5, 1400: 6.9, 1500: 7.4
    }
    
    flows = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
    results = {}
    best_prev = None
    
    print("\n" + "="*60)
    print("SMART PARAMETRIC OPTIMIZATION")
    print("="*60)
    
    # Run 3 iterations with adaptive refinement
    for iteration in range(1, 4):
        print(f"\n{'#'*60}")
        print(f"# ITERATION {iteration}/3")
        print(f"{'#'*60}")
        
        for flow in flows:
            if flow in results and results[flow]:
                p_center = results[flow]['p']
                best_prev = results[flow]
            else:
                p_center = p_init[flow]
            
            result = optimize_flow_smart(case, flow, p_center, best_prev)
            
            if result:
                results[flow] = result
                best_prev = result
    
    # Save results
    csv_file = os.path.join(os.getcwd(), folder, "optimization_smart_results.csv")
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
    print("COMPLETE")
    print(f"Results: {csv_file}")
    print('='*60)
    
    print("\nSUMMARY:")
    for flow in flows:
        if flow in results and results[flow]:
            r = results[flow]
            print(f"{flow:4d} kg/h: P={r['p']:.1f}b T={r['t']:.0f}C App={r['app']:.2f}C Pwr={r['pwr']:.0f}kW")
        else:
            print(f"{flow:4d} kg/h: FAILED")

if __name__ == "__main__":
    main()
