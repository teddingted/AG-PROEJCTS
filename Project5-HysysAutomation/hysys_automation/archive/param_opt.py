import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup, read_data_from_case

# --- Logic Functions ---


def reset_all_adjusts(case):
    """
    Resets all Adjust blocks (ADJ-1 through ADJ-4).
    Critical for convergence when operating conditions change.
    """
    adj_names = ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]
    print("  Resetting all Adjust blocks...")
    
    for name in adj_names:
        try:
            adj = case.Flowsheet.Operations.Item(name)
            adj.Reset()
            print(f"    [OK] {name}")
        except Exception as e:
            print(f"    [FAIL] {name}: {e}")
    
    return True

def check_adjusts_and_fix(case):
    """
    Checks if global solver is stuck or if critical adjusts are failing.
    """
    # Heuristic: Global Solver
    try:
        if not case.Solver.CanSolve:
            print("    > Global Solver Stopped (CanSolve=False). Restarting...")
            case.Solver.CanSolve = True
            return True
    except: pass
    
    # We could iterate adjusts and check warnings, but let's rely on Valid check later
    return False

def wait_for_solver(case, timeout=60, stabilization_time=1.5):
    """
    Waits for solver to complete AND verifies value stabilization.
    
    HYSYS convergence involves:
    1. Solver reports IsSolving=False (calculations stopped)
    2. Recycle blocks converge
    3. Adjust blocks converge
    4. Values stabilize
    
    This function ensures ALL conditions are met.
    Optimized for speed while maintaining reliability.
    """
    max_retries = 2  # Reduced from 3
    for attempt in range(max_retries):
        # Step 1: Wait for solver to report idle
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
            
            time.sleep(0.2)
        
        # Step 2: Stabilization delay (reduced for speed)
        time.sleep(stabilization_time)
        
        # Step 3: Verify values are stable (faster sampling)
        try:
            # First sample
            s7_1 = case.Flowsheet.MaterialStreams.Item("7").Pressure.Value
            lng_1 = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
            
            # Wait shorter interval
            time.sleep(0.5)  # Reduced from 1.0s
            
            # Second sample
            s7_2 = case.Flowsheet.MaterialStreams.Item("7").Pressure.Value
            lng_2 = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
            
            # Check if values changed significantly
            p7_change = abs(s7_2 - s7_1)
            app_change = abs(lng_2 - lng_1)
            
            # Slightly relaxed tolerance for speed
            if p7_change > 2.0 or app_change > 0.15:  # Was 1.0 kPa, 0.1 C
                print(f"    [INFO] Values changing (P7:{p7_change:.2f} kPa, App:{app_change:.2f} C), waiting...")
                time.sleep(stabilization_time)
                continue
                
        except:
            pass
        
        # Step 4: Check if solver needs restart
        if check_adjusts_and_fix(case):
            continue
        
        return True
    
    return False

def get_snapshot(case):
    """
    Capture metrics.
    """
    valid = True
    note = "OK"
    
    p7 = 99999.0
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        p7 = s7.Pressure.Value # kPa
        if p7 > 3600.0 + 1.0: # 36 bar limit
            valid = False
            note = f"P7 High ({p7/100:.2f} bar)"
    except: valid=False; note="P7 Fail"
        
    min_app = -999.0
    try:
        lng = case.Flowsheet.Operations.Item("LNG-100")
        min_app = lng.MinApproach.Value
        if min_app < 0.0:
            valid = False
            note = f"Cross ({min_app:.2f} C)"
    except: valid=False; note="App Fail"

    power = 999999.0
    try:
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        val = ss.Cell("C8").CellValue
        if val is None: val = 999999.0
        power = float(val)
    except: pass
        
    return {
        'p7': p7,
        'min_app': min_app,
        'power': power,
        'valid': valid,
        'note': note
    }

def get_initial_pressure(flow_kgh):
    """
    Pre-set table for initial pressure based on reliq flow.
    
    Fixed endpoints:
    - 500 kg/h → 3.0 bara (300 kPa)
    - 1500 kg/h → 7.4 bara (740 kPa)
    
    Intermediate values: Linear interpolation with gradual pressure increase
    Rounded to 0.1 bar precision (10 kPa units) for exact values like 3.0, 3.1, 3.2...
    """
    if flow_kgh <= 500:
        p_kpa = 300.0  # 3.0 bar
    elif flow_kgh >= 1500:
        p_kpa = 740.0  # 7.4 bar
    else:
        # Linear interpolation
        p_kpa = 300.0 + (flow_kgh - 500.0) * (740.0 - 300.0) / (1500.0 - 500.0)
    
    # Round to nearest 10 kPa (0.1 bar precision)
    # This ensures values like 300, 310, 320... (3.0, 3.1, 3.2 bar)
    p_kpa = round(p_kpa / 10.0) * 10.0
    
    return p_kpa

def run_optimization_for_flow(case, flow_kgh, start_t, writer, best_success=None):
    """
    Optimizes P and T for a specific Flow10.
    
    Pressure strategy:
    - 500 kg/h, 1500 kg/h: FIXED at pre-set table value
    - Others: ±0.2 bar (±20 kPa) range from pre-set table
    
    Fallback: If convergence fails, try starting from best_success (P, T)
    """
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        
        print(f"\n{'='*80}")
        print(f"OPTIMIZING: Stream 10 Flow = {flow_kgh} kg/h")
        print('='*80)
        
        # Set Flow
        print(f"Setting flow to {flow_kgh} kg/h...")
        s10.MassFlow.Value = flow_kgh / 3600.0
        wait_for_solver(case, timeout=45)
        
        # CRITICAL: Reset all Adjust blocks after flow change
        reset_all_adjusts(case)
        wait_for_solver(case, timeout=45)
        
        # Determine pressure range
        p_preset = get_initial_pressure(flow_kgh)
        
        # Fixed pressure for endpoints, variable for others
        if flow_kgh == 500 or flow_kgh == 1500:
            p_range = (p_preset, p_preset)  # Fixed
            allow_p_variation = False
            print(f"\nFixed pressure: P={p_preset/100:.1f} bar (endpoint)")
        elif flow_kgh == 600:
            # Special case for 600 kg/h which struggled
            p_min = p_preset - 40.0  # -0.4 bar
            p_max = p_preset + 40.0  # +0.4 bar
            p_range = (p_min, p_max)
            allow_p_variation = True
            print(f"\nPressure range: P={p_min/100:.1f}-{p_max/100:.1f} bar (preset ±0.4 for 600kg/h)")
        else:
            p_min = p_preset - 20.0  # -0.2 bar
            p_max = p_preset + 20.0  # +0.2 bar
            p_range = (p_min, p_max)
            allow_p_variation = True
            print(f"\nPressure range: P={p_min/100:.1f}-{p_max/100:.1f} bar (preset ±0.2)")
        
        t_curr = start_t
        print(f"Starting temperature: T={t_curr:.1f} C")
        
        # Try initial setup with preset pressure
        p_curr = p_preset
        s1.Pressure.Value = p_curr
        adj4.TargetValue.Value = t_curr
        wait_for_solver(case, timeout=45)
        
        # Check initial state
        snap = get_snapshot(case)
        print(f"Initial state: App={snap['min_app']:.2f}, P7={snap['p7']/100:.2f}b, Pwr={snap['power']:.1f}")
        
        # If initial convergence failed and we have a fallback, try it
        if (not snap['valid'] or snap['min_app'] < -100) and best_success:
            print(f"Initial convergence failed, trying fallback from best success case...")
            p_fallback, t_fallback = best_success
            
            # Constrain fallback pressure to allowed range
            if allow_p_variation:
                p_fallback = max(p_range[0], min(p_range[1], p_fallback))
            else:
                p_fallback = p_preset  # Keep fixed
            
            s1.Pressure.Value = p_fallback
            adj4.TargetValue.Value = t_fallback
            wait_for_solver(case, timeout=45)
            reset_all_adjusts(case)
            wait_for_solver(case, timeout=45)
            
            snap = get_snapshot(case)
            print(f"After fallback: App={snap['min_app']:.2f}, P7={snap['p7']/100:.2f}b, Pwr={snap['power']:.1f}")
            
            if snap['valid'] and snap['min_app'] > -100:
                p_curr = p_fallback
                t_curr = t_fallback
                print(f"Fallback successful! Starting from P={p_curr/100:.1f}b, T={t_curr:.1f}C")
        
        # Start optimization
        best_snap = snap if (snap['valid'] and 2.0 <= snap['min_app'] <= 2.5 and snap['p7'] <= 3600) else None
        best_state = (p_curr, t_curr)
        best_score = 1e20
        
        # Scoring function
        def score(s):
            # Hard constraints
            if not s['valid']:
                pen = 1e9
                if "P7" in s['note']: pen += (s['p7'] - 3600)*1000
                if "Cross" in s['note']: pen += abs(s['min_app'])*10000
                return pen
            
            # P7 constraint (CRITICAL)
            if s['p7'] > 3600.0:
                return 1e8 + (s['p7'] - 3600.0)*10000
            
            # Min Approach constraint
            ma = s['min_app']
            if ma > 2.5:
                return 1e6 + (ma - 2.5)*10000
            if ma < 2.0:
                return 1e6 + (2.0 - ma)*10000
            
            # Minimize power within valid range
            return s['power']
        
        # Update best if initial is good
        sc = score(snap)
        if sc < best_score:
            best_score = sc
            best_snap = snap
            best_state = (p_curr, t_curr)
        
        # Grid Search with temperature AND optionally pressure
        max_steps = 50
        
        # Define neighbors based on whether pressure can vary
        if allow_p_variation:
            neighbors = [(10.0, 0.0), (-10.0, 0.0), (0.0, 1.0), (0.0, -1.0)]  # P±0.1bar, T±1C
            print(f"\nOptimizing P (±0.2bar from {p_preset/100:.1f}) and T (max {max_steps} steps)...")
        else:
            neighbors = [(0.0, 1.0), (0.0, -1.0)]  # T only
            print(f"\nOptimizing T only at fixed P={p_preset/100:.1f}bar (max {max_steps} steps)...")
        
        for step in range(max_steps):
            sc = score(snap)
            
            # Update best if improved
            if sc < best_score:
                best_score = sc
                best_snap = snap
                best_state = (p_curr, t_curr)
                status = "[NEW BEST]"
            else:
                status = ""
            
            # Diagnostic output every 5 steps or if best
            if step % 5 == 0 or status:
                print(f"  Step {step:2d}: P={p_curr/100:.1f}b T={t_curr:.1f}C | "
                      f"App={snap['min_app']:.2f} P7={snap['p7']/100:.2f}b Pwr={snap['power']:.1f} | "
                      f"Score={sc:.0f} {status}")
            
            # Early exit if we found excellent solution
            if (snap['valid'] and 2.0 <= snap['min_app'] <= 2.5 and 
                snap['p7'] <= 3600 and snap['power'] < best_score * 1.01):
                if step > 10:  # Give it at least 10 steps
                    print(f"  Excellent solution found, stopping early.")
                    break
            
            # Explore neighbors (P and/or T)
            candidates = []
            for dp, dt in neighbors:
                pn = p_curr + dp
                tn = t_curr + dt
                
                # Check pressure bounds
                if pn < p_range[0] or pn > p_range[1]:
                    continue
                
                # Temperature bounds
                if tn < -115.0 or tn > -90.0:
                    continue
                
                try:
                    # Set new P and T
                    if abs(dp) > 0.1:  # Pressure changed
                        s1.Pressure.Value = pn
                    adj4.TargetValue.Value = tn
                    wait_for_solver(case, timeout=20)
                    
                    sn = get_snapshot(case)
                    
                    # If invalid, try reset
                    if not sn['valid']:
                        reset_all_adjusts(case)
                        wait_for_solver(case, timeout=20)
                        sn = get_snapshot(case)
                    
                    candidates.append((pn, tn, score(sn), sn))
                except:
                    continue
                
                # Revert to current state
                s1.Pressure.Value = p_curr
                adj4.TargetValue.Value = t_curr
                wait_for_solver(case, timeout=20)
  # Reduced from 30
            
            if not candidates:
                print("  No valid neighbors, stopping.")
                break
            
            # Pick best neighbor
            candidates.sort(key=lambda x: x[2])  # Sort by score
            winner = candidates[0]
            
            if winner[2] < sc:
                p_curr, t_curr = winner[0], winner[1]
                snap = winner[3]
                # Apply winner
                s1.Pressure.Value = p_curr
                adj4.TargetValue.Value = t_curr
                wait_for_solver(case, timeout=20)
            else:
                print("  Local optimum reached.")
                break
        
        # Record best result
        print(f"\n{'='*80}")
        if best_snap and best_snap['valid'] and 2.0 <= best_snap['min_app'] <= 2.5 and best_snap['p7'] <= 3600:
            print(f"[SUCCESS] Valid optimum found:")
            print(f"  P1 = {best_state[0]/100:.1f} bar")
            print(f"  T_ADJ4 = {best_state[1]:.1f} C")
            print(f"  Min Approach = {best_snap['min_app']:.2f} C [2.0-2.5 OK]")
            print(f"  P7 = {best_snap['p7']/100:.2f} bar [<=36 OK]")
            print(f"  Power = {best_snap['power']:.2f} kW")
        else:
            print(f"[PARTIAL] Best solution (may violate constraints):")
            if best_snap:
                print(f"  P1 = {best_state[0]/100:.1f} bar")
                print(f"  T_ADJ4 = {best_state[1]:.1f} C")
                print(f"  Min Approach = {best_snap['min_app']:.2f} C")
                print(f"  P7 = {best_snap['p7']/100:.2f} bar")
                print(f"  Power = {best_snap['power']:.2f} kW")
        print('='*80)
        
        # Write to CSV
        if best_snap:
            row = [
                flow_kgh,
                round(best_state[0]/100.0, 1),  # P1 with 0.1 bar precision
                round(best_state[1], 1),         # T with 0.1 C precision
                round(best_snap['min_app'], 2),
                round(best_snap['p7']/100.0, 2),
                round(best_snap['power'], 2),
                best_snap['note']
            ]
            writer.writerow(row)
            
            # Return temperature and success state
            if best_snap['valid'] and 2.0 <= best_snap['min_app'] <= 2.5:
                return best_state[1], best_state  # (T, (P, T))
            else:
                return best_state[1], None  # Partial success, no fallback update
        else:
            # Fallback
            writer.writerow([flow_kgh, 0, 0, 0, 0, 0, "FAILED"])
            return start_t, None
            
    except Exception as e:
        print(f"ERROR in optimization: {e}")
        import traceback
        traceback.print_exc()
        return start_t, None


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
    
    # Open CSV
    csv_file = os.path.join(os.getcwd(), folder, "optimization_results.csv")
    f = open(csv_file, 'w', newline='')
    writer = csv.writer(f)
    writer.writerow(["MassFlow_10_kg_h", "P1_bar", "ADJ4_Target_C", "MinApproach_C", "P7_bar", "Power_kW", "Note"])
    
    # Step 1: Set and verify 500 kg/h optimal state (from previous optimization)
    print("\n" + "="*80)
    print("STEP 1: Setting 500 kg/h to optimal state")
    print("="*80)
    
    # Hardcoded optimal values from previous successful run
    OPTIMAL_500 = {
        'flow': 500,
        'p1': 3.0,      # bar
        't_adj4': -111.0,  # C
        'expected_app': 2.37,  # C
        'expected_p7': 13.59,  # bar
        'expected_power': 669.2  # kW
    }
    
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        print(f"\nSetting optimal conditions:")
        print(f"  Flow: {OPTIMAL_500['flow']} kg/h")
        print(f"  P1: {OPTIMAL_500['p1']} bar")
        print(f"  T_ADJ4: {OPTIMAL_500['t_adj4']} C")
        
        # Set conditions
        s10.MassFlow.Value = OPTIMAL_500['flow'] / 3600.0
        wait_for_solver(case, timeout=45)
        
        reset_all_adjusts(case)
        wait_for_solver(case, timeout=45)
        
        s1.Pressure.Value = OPTIMAL_500['p1'] * 100.0  # Convert to kPa
        adj4.TargetValue.Value = OPTIMAL_500['t_adj4']
        wait_for_solver(case, timeout=60)
        
        # Verify convergence
        print("\nVerifying convergence...")
        flow10 = s10.MassFlow.Value * 3600.0
        p1 = s1.Pressure.Value / 100.0
        p7 = s7.Pressure.Value / 100.0
        t_adj4 = adj4.TargetValue.Value
        min_app = lng.MinApproach.Value
        power = ss.Cell("C8").CellValue
        
        print(f"\nActual state:")
        print(f"  Flow: {flow10:.1f} kg/h")
        print(f"  P1: {p1:.1f} bar")
        print(f"  T_ADJ4: {t_adj4:.1f} C")
        print(f"  Min Approach: {min_app:.2f} C (expected: {OPTIMAL_500['expected_app']:.2f})")
        print(f"  P7: {p7:.2f} bar (expected: {OPTIMAL_500['expected_p7']:.2f})")
        print(f"  Power: {power:.2f} kW (expected: {OPTIMAL_500['expected_power']:.2f})")
        
        # Check if converged properly
        app_ok = abs(min_app - OPTIMAL_500['expected_app']) < 0.5
        p7_ok = abs(p7 - OPTIMAL_500['expected_p7']) < 2.0
        
        if app_ok and p7_ok:
            print("  ✓ Convergence verified!")
            status = "User Set (Verified)"
        else:
            print("  ⚠ Values differ from expected - may need adjustment")
            status = "User Set (Check)"
        
        row_500 = [
            500,
            round(p1, 1),
            round(t_adj4, 1),
            round(min_app, 2),
            round(p7, 2),
            round(power, 2),
            status
        ]
        writer.writerow(row_500)
        f.flush()
        
        # Use this as initial best_success and warm start temperature
        curr_t = t_adj4
        best_success = (OPTIMAL_500['p1'] * 100.0, OPTIMAL_500['t_adj4'])  # (P_kPa, T_C)
        
    except Exception as e:
        print(f"ERROR setting 500 kg/h state: {e}")
        import traceback
        traceback.print_exc()
        curr_t = -98.0
        best_success = None
    
    # Step 2: Optimize specific flow rates
    print("\n" + "="*80)
    print("STEP 2: Optimizing 600 kg/h (Refining)")
    print("="*80)
    
    # Targeting 600 kg/h based on previous results
    flows = [600]
    
    # best_success already set from 500 kg/h in Step 1
    # For 600, we can also hint from the previous run (P=3.6, T=-104.0) if we wanted,
    # but starting from the 500 fallback is also safe. 
    # Let's try to start near the previous result to save time.
    curr_t = -104.0 
    
    if best_success:
        print(f"\nStarting with fallback from 500 kg/h: P={best_success[0]/100:.1f}b, T={best_success[1]:.1f}C\n")
    
    for flow in flows:
        # We want to allow slightly more pressure variation for 600 kg/h to find the sweet spot
        # Original logic allows +-0.2. Let's see if that's enough. 
        # Previous result: P=3.6, T=-104.0, App=1.05. 
        # Needs Higher App -> Likely needs Lower P or Higher T?
        # Actually, higher App usually means moving away from pinch.
        
        result_t, success_state = run_optimization_for_flow(case, flow, curr_t, writer, best_success)
        curr_t = result_t
        
        # Update best_success if this case succeeded
        if success_state:
            best_success = success_state
            print(f"  → Best success updated: P={success_state[0]/100:.1f}b, T={success_state[1]:.1f}C")
        
        f.flush()
        
    f.close()
    print("\nParametric Optimization Complete.")

if __name__ == "__main__":
    main()
