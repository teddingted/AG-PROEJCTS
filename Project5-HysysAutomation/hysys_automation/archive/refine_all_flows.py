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
            break
        
        time.sleep(0.3)
    
    time.sleep(1.5)
    return True

def reset_all_adjusts(case):
    """Reset all Adjust blocks"""
    for adj_name in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            adj = case.Flowsheet.Operations.Item(adj_name)
            adj.Reset()
        except:
            pass

def get_snapshot(case):
    """Get current state snapshot"""
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        p7 = s7.Pressure.Value
        min_app = lng.MinApproach.Value
        power = ss.Cell("C8").CellValue
        
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
    """Test a specific P, T condition"""
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    s10.MassFlow.Value = flow_kgh / 3600.0
    s1.Pressure.Value = p_bar * 100.0
    adj4.TargetValue.Value = t_c
    
    wait_for_solver(case, timeout=45)
    reset_all_adjusts(case)
    wait_for_solver(case, timeout=45)
    
    # Check if Stream 1 temperature is abnormal (outside 0-50°C indicates convergence issue)
    try:
        s1_temp = s1.Temperature.Value
        if not (0 <= s1_temp <= 50):
            print(f"    [ABNORMAL] S1 temp={s1_temp:.1f}C (outside 0-50C), recovering...")
            # Recovery: Set S1 to safe temperature and re-converge
            s1.Temperature.Value = 40.0  # Safe temperature within range
            s1.Pressure.Value = p_bar * 100.0  # Keep pressure
            wait_for_solver(case, timeout=45)
            reset_all_adjusts(case)
            wait_for_solver(case, timeout=45)
    except:
        pass
    
    return get_snapshot(case)

def optimize_flow(case, flow_kgh, p_center, t_range, iteration, fallback_result=None):
    """Optimize single flow rate with ±0.4 bar pressure range"""
    
    print(f"\n{'='*80}")
    print(f"ITERATION {iteration}: Flow = {flow_kgh} kg/h, P_center = {p_center:.1f} bar")
    print('='*80)
    
    # Pressure range: ±0.4 bar in 0.1 bar steps
    p_min = p_center - 0.4
    p_max = p_center + 0.4
    pressures = [round(p_min + i*0.1, 1) for i in range(9)]  # 9 points
    
    # Temperature range
    temperatures = list(range(t_range[0], t_range[1]+1, 1))
    
    best_result = None
    best_score = 1e20
    
    for p_bar in pressures:
        for t_c in temperatures:
            snap = test_condition(case, flow_kgh, p_bar, t_c)
            
            if snap['valid']:
                app_val = snap['min_app']
                p7_val = snap['p7'] / 100.0
                pwr_val = snap['power']
                
                # Check constraints
                app_ok = 2.0 <= app_val <= 2.5
                p7_ok = p7_val <= 36.0
                
                if app_ok and p7_ok:
                    score = pwr_val
                    if score < best_score:
                        best_score = score
                        best_result = {
                            'flow': flow_kgh,
                            'p': p_bar,
                            't': t_c,
                            'app': app_val,
                            'p7': p7_val,
                            'power': pwr_val,
                            'status': 'SUCCESS'
                        }
                        print(f"  [NEW BEST] P={p_bar:.1f}b T={t_c:.1f}C -> App={app_val:.2f}, Pwr={pwr_val:.1f} kW")
    
    if best_result:
        print(f"\n  BEST for {flow_kgh} kg/h:")
        print(f"    P={best_result['p']:.1f} bar, T={best_result['t']:.1f} C")
        print(f"    App={best_result['app']:.2f} C, P7={best_result['p7']:.2f} bar, Power={best_result['power']:.1f} kW")
    else:
        # FALLBACK: Use previous successful result if available
        if fallback_result:
            print(f"\n  No new solution found, using FALLBACK from previous iteration")
            print(f"    Applying fallback: P={fallback_result['p']:.1f} bar, T={fallback_result['t']:.1f} C")
            
            # Apply fallback conditions to recover convergence
            try:
                s1 = case.Flowsheet.MaterialStreams.Item("1")
                s10 = case.Flowsheet.MaterialStreams.Item("10")
                adj4 = case.Flowsheet.Operations.Item("ADJ-4")
                
                # Check if Stream 1 is abnormal
                s1_temp = s1.Temperature.Value
                if not (0 <= s1_temp <= 50):
                    print(f"    [RECOVERY] S1 abnormal at {s1_temp:.1f}C (outside 0-50C), resetting to 40C")
                    s1.Temperature.Value = 40.0
                    wait_for_solver(case, timeout=30)
                
                # Apply fallback P and T
                s10.MassFlow.Value = flow_kgh / 3600.0
                s1.Pressure.Value = fallback_result['p'] * 100.0
                adj4.TargetValue.Value = fallback_result['t']
                wait_for_solver(case, timeout=45)
                reset_all_adjusts(case)
                wait_for_solver(case, timeout=45)
                
                print(f"    Fallback applied successfully")
            except Exception as e:
                print(f"    [ERROR] Fallback application failed: {e}")
            
            best_result = fallback_result.copy()
            best_result['status'] = 'FALLBACK'
        else:
            print(f"\n  No valid solution found for {flow_kgh} kg/h and no fallback available")
    
    return best_result

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
    
    # Initial conditions from previous optimization
    initial_conditions = {
        500: {'p': 3.0, 't_range': (-115, -107)},
        600: {'p': 3.5, 't_range': (-112, -104)},
        700: {'p': 4.1, 't_range': (-107, -99)},
        800: {'p': 4.3, 't_range': (-109, -101)},
        900: {'p': 4.8, 't_range': (-109, -101)},
        1000: {'p': 5.2, 't_range': (-108, -100)},
        1100: {'p': 5.7, 't_range': (-106, -98)},
        1200: {'p': 6.1, 't_range': (-106, -98)},
        1300: {'p': 6.5, 't_range': (-103, -95)},
        1400: {'p': 6.9, 't_range': (-104, -96)},
        1500: {'p': 7.4, 't_range': (-101, -93)}
    }
    
    flows = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
    
    # Storage for results across iterations
    all_results = {flow: [] for flow in flows}
    
    # Run 3 iterations
    for iteration in range(1, 4):
        print(f"\n{'#'*80}")
        print(f"# ITERATION {iteration}/3")
        print(f"{'#'*80}")
        
        for flow in flows:
            # Determine fallback (previous iteration's result or known good value)
            if iteration == 1:
                # First iteration: use initial conditions, no fallback yet
                p_center = initial_conditions[flow]['p']
                t_range = initial_conditions[flow]['t_range']
                fallback = None
            else:
                # Subsequent iterations: use best result from previous iteration as fallback
                prev_best = all_results[flow][-1]
                if prev_best:
                    p_center = prev_best['p']
                    t_range = (prev_best['t'] - 4, prev_best['t'] + 4)
                    fallback = prev_best
                else:
                    # No previous result, use initial
                    p_center = initial_conditions[flow]['p']
                    t_range = initial_conditions[flow]['t_range']
                    fallback = None
            
            result = optimize_flow(case, flow, p_center, t_range, iteration, fallback)
            all_results[flow].append(result)
    
    # Save final results
    print(f"\n{'='*80}")
    print("FINAL OPTIMIZED RESULTS")
    print('='*80)
    
    csv_file = os.path.join(os.getcwd(), folder, "optimization_results_refined.csv")
    with open(csv_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["MassFlow_10_kg_h", "P1_bar", "ADJ4_Target_C", "MinApproach_C", "P7_bar", "Power_kW", "Note"])
        
        for flow in flows:
            # Get best result across all iterations
            valid_results = [r for r in all_results[flow] if r is not None]
            if valid_results:
                best = min(valid_results, key=lambda x: x['power'])
                writer.writerow([
                    flow,
                    round(best['p'], 1),
                    round(best['t'], 1),
                    round(best['app'], 2),
                    round(best['p7'], 2),
                    round(best['power'], 2),
                    best['status']
                ])
                print(f"{flow} kg/h: P={best['p']:.1f}b, T={best['t']:.1f}C, App={best['app']:.2f}C, Power={best['power']:.1f} kW")
            else:
                writer.writerow([flow, 0, 0, 0, 0, 0, "FAILED"])
                print(f"{flow} kg/h: FAILED")
    
    print(f"\nResults saved to: {csv_file}")

if __name__ == "__main__":
    main()
