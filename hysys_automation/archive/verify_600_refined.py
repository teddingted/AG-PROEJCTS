import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def reset_all_adjusts(case):
    """Resets all Adjust blocks (ADJ-1 through ADJ-4)."""
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

def wait_for_solver(case, timeout=60, stabilization_time=2.0):
    """Waits for solver to complete AND verifies value stabilization."""
    max_retries = 2
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
        
        # Step 2: Stabilization delay
        time.sleep(stabilization_time)
        
        # Step 3: Verify values are stable
        try:
            # First sample
            s7_1 = case.Flowsheet.MaterialStreams.Item("7").Pressure.Value
            lng_1 = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
            
            # Wait
            time.sleep(0.8)
            
            # Second sample
            s7_2 = case.Flowsheet.MaterialStreams.Item("7").Pressure.Value
            lng_2 = case.Flowsheet.Operations.Item("LNG-100").MinApproach.Value
            
            # Check if values changed significantly
            p7_change = abs(s7_2 - s7_1)
            app_change = abs(lng_2 - lng_1)
            
            if p7_change > 2.0 or app_change > 0.15:
                print(f"    [INFO] Values changing (P7:{p7_change:.2f} kPa, App:{app_change:.2f} C), waiting...")
                time.sleep(stabilization_time)
                continue
                
        except:
            pass
        
        # Step 4: Check if solver needs restart
        try:
            if not case.Solver.CanSolve:
                print("    > Global Solver Stopped. Restarting...")
                case.Solver.CanSolve = True
                continue
        except:
            pass
        
        return True
    
    return False

def get_snapshot(case):
    """Capture all relevant metrics."""
    valid = True
    note = "OK"
    
    # P7 Constraint
    p7 = 99999.0
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        p7 = s7.Pressure.Value  # kPa
        if p7 > 3600.0 + 1.0:  # 36 bar limit
            valid = False
            note = f"P7 High ({p7/100:.2f} bar)"
    except:
        valid = False
        note = "P7 Fail"
        
    # Min Approach
    min_app = -999.0
    try:
        lng = case.Flowsheet.Operations.Item("LNG-100")
        min_app = lng.MinApproach.Value
        if min_app < 0.0:
            valid = False
            note = f"Cross ({min_app:.2f} C)"
    except:
        valid = False
        note = "App Fail"

    # Power
    power = 999999.0
    try:
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        val = ss.Cell("C8").CellValue
        if val is None:
            val = 999999.0
        power = float(val)
    except:
        pass
    
    # P1 (for reference)
    p1 = 0.0
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        p1 = s1.Pressure.Value
    except:
        pass
    
    # T_ADJ4 (for reference)
    t_adj4 = 0.0
    try:
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        t_adj4 = adj4.TargetValue.Value
    except:
        pass
        
    return {
        'p1': p1,
        'p7': p7,
        'min_app': min_app,
        'power': power,
        't_adj4': t_adj4,
        'valid': valid,
        'note': note
    }

def main():
    """
    Refined verification for 600 kg/h case.
    
    Goal: Find optimal P1 in range 3.0-3.8 bara that achieves Min Approach closest to 2.0°C
    Strategy: Grid search with 0.1 bar steps, optimizing temperature at each pressure
    """
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
    
    # Get references
    s1 = case.Flowsheet.MaterialStreams.Item("1")
    s10 = case.Flowsheet.MaterialStreams.Item("10")
    adj4 = case.Flowsheet.Operations.Item("ADJ-4")
    
    # Set flow to 600 kg/h
    target_flow = 600.0
    print(f"\n{'='*80}")
    print(f"REFINED VERIFICATION: 600 kg/h, P1 range 3.0-3.8 bara")
    print(f"Target: Min Approach = 2.0°C")
    print('='*80)
    
    print(f"\nSetting flow to {target_flow} kg/h...")
    s10.MassFlow.Value = target_flow / 3600.0
    wait_for_solver(case, timeout=45)
    reset_all_adjusts(case)
    wait_for_solver(case, timeout=45)
    
    # Open CSV for results
    csv_file = os.path.join(os.getcwd(), folder, "verification_600_refined.csv")
    f = open(csv_file, 'w', newline='')
    writer = csv.writer(f)
    writer.writerow(["P1_bar", "T_ADJ4_C", "MinApproach_C", "P7_bar", "Power_kW", 
                     "App_Error", "Valid", "Note"])
    
    # Pressure range: 3.0 to 3.8 bar in 0.1 bar steps
    p_min = 3.0
    p_max = 3.8
    p_step = 0.1
    
    # Generate pressure points
    pressures = []
    p = p_min
    while p <= p_max + 0.01:  # Small epsilon for floating point
        pressures.append(round(p, 1))
        p += p_step
    
    print(f"\nTesting {len(pressures)} pressure points: {pressures[0]} to {pressures[-1]} bar")
    print(f"For each pressure, will optimize temperature to target Min Approach = 2.0°C\n")
    
    best_result = None
    best_error = 999.0
    
    for p_bar in pressures:
        p_kpa = p_bar * 100.0
        
        print(f"\n{'-'*80}")
        print(f"Testing P1 = {p_bar} bar")
        print('-'*80)
        
        # Set pressure
        s1.Pressure.Value = p_kpa
        
        # Initial temperature guess (interpolate from known points)
        # At 500 kg/h: P=3.0, T=-111
        # At 600 kg/h (previous): P=3.7, T=-105
        # Linear interpolation
        if p_bar <= 3.0:
            t_start = -111.0
        elif p_bar >= 3.7:
            t_start = -105.0
        else:
            # Interpolate
            t_start = -111.0 + (p_bar - 3.0) * (-105.0 - (-111.0)) / (3.7 - 3.0)
        
        t_start = round(t_start, 1)
        
        print(f"Starting temperature: {t_start}°C")
        adj4.TargetValue.Value = t_start
        wait_for_solver(case, timeout=45)
        
        # Reset adjusts for clean start
        reset_all_adjusts(case)
        wait_for_solver(case, timeout=45)
        
        # Get initial snapshot
        snap = get_snapshot(case)
        print(f"Initial: App={snap['min_app']:.2f}°C, P7={snap['p7']/100:.2f}bar, Pwr={snap['power']:.1f}kW")
        
        # Local optimization on temperature to hit Min Approach = 2.0°C
        # Use simple gradient descent / binary search approach
        t_curr = t_start
        best_snap = snap
        best_t = t_curr
        
        max_iterations = 20
        target_app = 2.0
        tolerance = 0.05  # Accept ±0.05°C from target
        
        for iteration in range(max_iterations):
            snap = get_snapshot(case)
            
            if not snap['valid']:
                print(f"  Iter {iteration}: Invalid state ({snap['note']})")
                # Try to recover
                reset_all_adjusts(case)
                wait_for_solver(case, timeout=30)
                snap = get_snapshot(case)
                if not snap['valid']:
                    break
            
            app_error = abs(snap['min_app'] - target_app)
            
            # Update best if closer to target
            if app_error < abs(best_snap['min_app'] - target_app):
                best_snap = snap
                best_t = t_curr
            
            print(f"  Iter {iteration}: T={t_curr:.1f}°C → App={snap['min_app']:.2f}°C (error={app_error:.3f}°C)")
            
            # Check convergence
            if app_error < tolerance:
                print(f"  ✓ Target reached! App={snap['min_app']:.2f}°C")
                break
            
            # Determine direction
            # If Min Approach > target: need to decrease approach
            #   → Increase pressure OR decrease temperature
            # If Min Approach < target: need to increase approach  
            #   → Decrease pressure OR increase temperature
            # Since we're fixing pressure, adjust temperature only
            
            if snap['min_app'] > target_app:
                # Too high, need to reduce approach
                # Typically: lower temperature increases approach (moves away from pinch)
                # But this is system-dependent. Let's try decreasing T
                t_next = t_curr - 1.0
            else:
                # Too low, need to increase approach
                t_next = t_curr + 1.0
            
            # Bounds check
            if t_next < -115.0 or t_next > -90.0:
                print(f"  Temperature limit reached")
                break
            
            # Apply new temperature
            adj4.TargetValue.Value = t_next
            wait_for_solver(case, timeout=30)
            
            # Check if we're oscillating or stuck
            if iteration > 5:
                # If error isn't improving, try smaller steps
                if app_error > 0.5:
                    # Try half step
                    if snap['min_app'] > target_app:
                        t_next = t_curr - 0.5
                    else:
                        t_next = t_curr + 0.5
                    adj4.TargetValue.Value = t_next
                    wait_for_solver(case, timeout=30)
            
            t_curr = t_next
        
        # Record best result for this pressure
        final_error = abs(best_snap['min_app'] - target_app)
        
        print(f"\nBest for P={p_bar}bar: T={best_t:.1f}°C, App={best_snap['min_app']:.2f}°C, Error={final_error:.3f}°C")
        
        # Write to CSV
        row = [
            p_bar,
            round(best_t, 1),
            round(best_snap['min_app'], 2),
            round(best_snap['p7']/100.0, 2),
            round(best_snap['power'], 2),
            round(final_error, 3),
            best_snap['valid'],
            best_snap['note']
        ]
        writer.writerow(row)
        f.flush()
        
        # Track overall best
        if best_snap['valid'] and final_error < best_error:
            best_error = final_error
            best_result = {
                'p_bar': p_bar,
                't': best_t,
                'snap': best_snap
            }
    
    # Summary
    print(f"\n{'='*80}")
    print("VERIFICATION COMPLETE")
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
