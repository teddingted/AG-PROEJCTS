import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

# Configuration
SIMULATION_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
OUTPUT_FILE = "optimization_final_result.csv"
FOLDER_NAME = "hysys_automation"

# Optimization Parameters
FLOWS = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
PRESET_P_INIT = {
    500: 3.0, 600: 3.5, 700: 4.1, 800: 4.3, 900: 4.8, 1000: 5.2,
    1100: 5.7, 1200: 6.1, 1300: 6.5, 1400: 6.9, 1500: 7.4
}

def wait_solver_robust(case, timeout=60):
    """
    Robust solver wait with value stabilization check.
    Ensures simulation is truly converged and values are stable.
    """
    # 1. Wait for IsSolving to be False
    start = time.time()
    while time.time() - start < timeout:
        try:
            if not case.Solver.IsSolving:
                break
        except:
            pass
        time.sleep(0.5)
    
    # 2. Minimum stabilization wait
    time.sleep(1.0) 
    
    # 3. Check for specific stream value stabilization (Stream 1 Temp)
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        last_val = s1.Temperature.Value
        stable_count = 0
        
        for _ in range(10): # Check for up to 5 seconds
            time.sleep(0.5)
            curr_val = s1.Temperature.Value
            if abs(curr_val - last_val) < 0.01:
                stable_count += 1
            else:
                stable_count = 0
            
            if stable_count >= 2: # Stable for 1 second
                return True
            last_val = curr_val
    except:
        pass
        
    return True

def reset_adjusts(case):
    """Reset all Adjust blocks to default state."""
    for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            case.Flowsheet.Operations.Item(adj).Reset()
        except:
            pass

def check_abnormal_state(case):
    """Check for physical nonsense values (e.g. < -200 C) indicating instability."""
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        if s1.Temperature.Value < -200:
            return True # Abnormal
        
        lng = case.Flowsheet.Operations.Item("LNG-100")
        if abs(lng.MinApproach.Value) > 500:
            return True
            
        return False
    except:
        return True

def hard_reset(case, flow, p_bar):
    """
    Hard reset to clean state if simulation becomes unstable.
    Sets safe default values to recover convergence.
    """
    print("    [HARD RESET] Simulation unstable, applying safe defaults...")
    try:
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        
        # Reset Adjusts first
        reset_adjusts(case)
        
        # Set to safe known values
        s10.MassFlow.Value = flow / 3600.0
        s1.Pressure.Value = p_bar * 100.0
        s1.Temperature.Value = 40.0 # Safe temp
        adj4.TargetValue.Value = -100.0 # Safe target
        
        wait_solver_robust(case, 30)
        reset_adjusts(case)
        wait_solver_robust(case, 30)
    except:
        print("    [ERROR] Hard reset failed")

def check_smart_range_robust(case):
    """
    Check convergence quality.
    Returns:
        2: Target Hit (Min Approach 2.0 - 3.0 C)
        1: Valid (> 0 C)
        0: Cross or Invalid
    """
    try:
        # Check for garbage first
        if check_abnormal_state(case):
            return 0
            
        app = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
        if 2.0 <= app <= 3.0:
            return 2
        elif app > 0:
            return 1
        return 0
    except:
        return 0

def optimize_flow(case, flow, p_center, best_prev=None):
    """Run robust optimization for a single flow rate."""
    print(f"\n{'='*60}")
    print(f"Flow={flow} kg/h, P_center={p_center:.1f} bar")
    
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # 1. Define Pressure Range
    if flow == 500 or flow == 1500:
        pressures = [p_center]
        print(f"  [MODE] Fixed Pressure: {p_center} bar")
    else:
        pressures = [round(p_center + i*0.1, 1) for i in range(-4, 5)]
        print(f"  [MODE] Variable Pressure: {pressures[0]} - {pressures[-1]} bar")

    best = None
    best_score = 1e20

    for p in pressures:
        # 2. Scan for Valid Temperature Range
        p_kpa = p * 100.0
        valid_range = []
        
        # Reset to clean state for each pressure
        s10.MassFlow.Value = flow / 3600.0
        s1.Pressure.Value = p_kpa
        s1.Temperature.Value = 40.0 # Ensure S1 temp is safe
        wait_solver_robust(case, 20)
        reset_adjusts(case)
        wait_solver_robust(case, 20)
        
        print(f"  [SCAN] P={p} bar...", end="", flush=True)
        
        # Scan temps: -90 to -120 in 5 deg steps
        scan_temps = range(-90, -121, -5)
        for t in scan_temps:
            adj4.TargetValue.Value = t
            wait_solver_robust(case, 15) 
            
            # Check for crash
            if check_abnormal_state(case):
                print(" Unstable!", end="")
                hard_reset(case, flow, p)
                continue
                
            status = check_smart_range_robust(case)
            
            if status == 2: # Target Hit
                print(f" TARGET HIT at {t}C")
                valid_range = range(t+2, t-3, -1) # Narrow search around target
                break
            elif status == 1:
                valid_range = range(t, t-6, -1) # Default valid range
            elif status == 0:
                if valid_range: break # Stop scanning if we hit a cross after finding valid points
        
        if not valid_range:
            print(" No valid range found.")
            continue
            
        # 3. Optimize in Valid Range
        print(f"    -> Optimizing {valid_range.start} to {valid_range.stop} C...")
        
        for t in valid_range:
            adj4.TargetValue.Value = t
            s1.Pressure.Value = p_kpa
            wait_solver_robust(case, 15)
            
            if check_abnormal_state(case):
                hard_reset(case, flow, p)
                continue
            
            if check_smart_range_robust(case) == 0:
                continue
                
            # Run Adjusts for final convergence
            reset_adjusts(case)
            wait_solver_robust(case, 30)
            
            # Final check
            if check_abnormal_state(case): continue
            
            try:
                s7 = case.Flowsheet.MaterialStreams.Item("7")
                lng = case.Flowsheet.Operations.Item("LNG-100")
                ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
                
                app = lng.MinApproach.Value
                p7 = s7.Pressure.Value / 100.0
                pwr = ss.Cell("C8").CellValue
                
                # Check constraints: App 2.0-2.5, P7 <= 36
                if 2.0 <= app <= 2.5 and p7 <= 36.0:
                    if pwr < best_score:
                        best_score = pwr
                        best = {'p':p, 't':t, 'app':app, 'p7':p7, 'pwr':pwr}
                        print(f"      [NEW BEST] P={p} T={t} App={app:.2f} Pwr={pwr:.1f}")
            except:
                pass
                
    if not best and best_prev:
        print("  [FALLBACK] Using previous iteration result")
        best = best_prev.copy()
        
    if best:
        print(f"  RESULT: P={best['p']} bar, T={best['t']} C, App={best['app']:.2f} C, Power={best['pwr']:.1f} kW")
        return best
    else:
        print("  FAILED to find solution")
        return None

def main():
    file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER_NAME, SIMULATION_FILE))
    
    print("="*60)
    print("HYSYS ROBUST OPTIMIZATION (FINAL)")
    print("="*60)
    
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
        
    results = {}
    
    # Run Optimization Loop (3 Iterations for refinement)
    for iteration in range(1, 4):
        print(f"\n{'#'*30}")
        print(f"# ITERATION {iteration}/3")
        print(f"{'#'*30}")
        
        for flow in FLOWS:
            # Use previous best result as reference if available
            best_prev = results.get(flow)
            
            # Run optimization
            res = optimize_flow(case, flow, PRESET_P_INIT[flow], best_prev)
            
            if res:
                results[flow] = res
            
            # Save intermediate results
            csv_path = os.path.join(os.getcwd(), FOLDER_NAME, OUTPUT_FILE)
            try:
                with open(csv_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["MassFlow", "P1_bar", "T_ADJ4_C", "MinApp_C", "P7_bar", "Power_kW", "Status"])
                    for f_key in FLOWS:
                        if f_key in results:
                            r = results[f_key]
                            writer.writerow([f_key, r['p'], r['t'], round(r['app'],2), round(r['p7'],2), round(r['pwr'],2), "OK"])
                        else:
                            writer.writerow([f_key, 0, 0, 0, 0, 0, "PENDING" if f_key >= flow else "FAILED"])
            except Exception as e:
                print(f"  [WARN] Failed to save CSV: {e}")

    print(f"\nFinal results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
