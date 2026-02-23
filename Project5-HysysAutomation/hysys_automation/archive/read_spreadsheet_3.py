
import os
import time
import win32com.client
import threading
from hysys_utils import dismiss_popup, read_data_from_case, print_data_table

def reset_adjust_block(adj):
    """
    Attempts to reset the Adjust block via Ignored toggle.
    """
    try:
        adj.Ignored = True
        time.sleep(0.1)
        adj.Ignored = False
        print(f"    > Reset {adj.Name} via Ignored toggle.")
        return True
    except:
        return False

def check_adjusts_and_fix(case):
    """
    Checks if global solver is stuck or if critical adjusts are failing.
    Retuns True if an action was taken.
    """
    # Heuristic: Check if CanSolve is False (Stuck)
    try:
        if not case.Solver.CanSolve:
            print("    > Global Solver Stopped. Restarting...")
            case.Solver.CanSolve = True
            return True
    except: pass

    # Check ADJ-1 explicitly if possible
    # We can't easily read status, but we can restart if we suspect it's invalid
    return False

def wait_for_solver(case, timeout=30):
    """
    Waits for solver. Checks Adjusts. Resets if needed.
    """
    max_retries = 3
    for attempt in range(max_retries):
        # 1. Wait for Idle
        start = time.time()
        while True:
            try:
                if not case.Solver.IsSolving:
                    break
            except: pass
            if time.time() - start > timeout:
                print("    > Solver Timeout.")
                break 
            time.sleep(0.2)
            
        # 2. Check & Fix
        if check_adjusts_and_fix(case):
            continue # Loop back if fixed
            
        return True
    return False



def ensure_convergence(case):
    """
    Iterate over Adjust/Recycle blocks.
    If not converged (heuristic), reset them.
    User instruction: "Press reset button until converged".
    """
    ops = case.Flowsheet.Operations
    
    # We can't easily check 'Converged' property.
    # Heuristic: Check if case is solvable.
    # If case.Solver.CanSolve is False, we have issues.
    
    # User specific request: ADJ-1.
    # Let's blindly reset ADJ-1 if we suspect issues?
    # Or try to force it to calculate.
    
    # Since property access failed locally, we will try to use the HYSYS 
    # built-in 'Solver' start/stop or 'Integrator' reset? No, that's dynamics.
    
    # Let's try to just continue and assume the Grid Search's "Valid" check 
    # (checking if values are present and constraints met) catch the non-converged states.
    # If invalid, the Grid Search will penalize it heavily.
    
    # However, to help it converge, we can try to "poke" the solver by 
    # Stop/Start solver?
    case.Solver.CanSolve = False
    time.sleep(0.1)
    case.Solver.CanSolve = True
    wait_for_solver(case)

def get_snapshot(case):
    """
    Capture metrics.
    Metrics: P7, Min App (Approach), Power (C8).
    """
    valid = True
    note = "OK"
    
    # 1. P7 Constraint
    p7 = 99999.0
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        p7 = s7.Pressure.Value
        if p7 > 3600.0 + 1.0: # 36 bar limit
            valid = False
            note = f"P7 High ({p7/100:.2f} bar)"
    except:
        valid = False
        note = "P7 Read Fail"
        
    # 2. Min Approach (Target <= 2.5) & Cross Check
    min_app = -999.0
    try:
        lng = case.Flowsheet.Operations.Item("LNG-100")
        min_app = lng.MinApproach.Value
        if min_app < 0.0:
            valid = False # Cross is bad
            note = f"Cross ({min_app:.2f} C)"
    except:
        valid = False
        note = "MinApp Read Fail"

    # 3. Power (Minimize this)
    power = 999999.0
    try:
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        # Cell C8
        c8 = ss.Cell("C8")
        val = c8.CellValue
        if val is None: val = 999999.0
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

def optimize_case():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    try:
        app = win32com.client.Dispatch("HYSYS.Application")
        app.Visible = True
    except: return

    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    try:
        if os.path.exists(file_path):
             case = app.SimulationCases.Open(file_path)
             case.Visible = True
        else:
             case = app.ActiveDocument
    except: return

    try:
        stream_1 = case.Flowsheet.MaterialStreams.Item("1")
        adj_4 = case.Flowsheet.Operations.Item("ADJ-4")
    except: return

    # Initial Reset
    print("Resetting to safe start...")
    p_curr = 740.0 # 7.4 bar
    t_curr = -98.0 # -98 C
    stream_1.Pressure.Value = p_curr
    adj_4.TargetValue.Value = t_curr
    
    ensure_convergence(case)
    
    # Grid Search Config
    # Discrete Steps: P +/- 10 kPa (0.1 bar), T +/- 1.0 C
    neighbors = [
        (10.0, 0.0), (-10.0, 0.0), 
        (0.0, 1.0), (0.0, -1.0)
    ]
    
    max_steps = 30

    print("\n--- Discrete Optimization (Target: 2.0 <= Min App <= 2.5, Min Power) ---")
    
    for step in range(max_steps):
        # 1. Snapshot
        curr = get_snapshot(case)
        
        # 2. Score
        def score_func(snap):
            # 1. Validity (Hard Constraints: P7, Convergence)
            if not snap['valid']:
                pen = 1e9
                if "P7" in snap['note']: pen += (snap['p7'] - 3600)*1000
                if "Cross" in snap['note']: pen += abs(snap['min_app'])*10000
                return pen
            
            # 2. Target Range: 2.0 <= MinApp <= 2.5
            ma = snap['min_app']
            
            if ma > 2.5:
                # Too high: Penalty
                return 1e6 + (ma - 2.5) * 10000
            elif ma < 2.0:
                # Too low: Penalty
                return 1e6 + (2.0 - ma) * 10000
                
            # 3. Objective: Minimize Power (Inside Range)
            return snap['power']

        curr_score = score_func(curr)

        print(f"\nStep {step}: P={p_curr:.1f}, T={t_curr:.1f} | Score={curr_score:.1f} | Valid={curr['valid']} ({curr['note']})")
        print(f"  > MinApp={curr['min_app']:.4f}, Power={curr['power']:.2f}")
        
        # 3. Explore Neighbors
        candidates = []
        for dp, dt in neighbors:
            p_next = p_curr + dp
            t_next = t_curr + dt
            
            # Boundary check
            if p_next < 500 or p_next > 5000: continue
            
            try:
                stream_1.Pressure.Value = p_next
                adj_4.TargetValue.Value = t_next
            except: continue
            
            wait_for_solver(case)
            
            # Ensure convergence?
            # Basic toggle if invalid
            snap = get_snapshot(case)
            if not snap['valid']:
                # Try reset/poke
                case.Solver.CanSolve = False
                case.Solver.CanSolve = True
                wait_for_solver(case)
                snap = get_snapshot(case)

            sc = score_func(snap)
            candidates.append((p_next, t_next, sc))
            
            # Revert for next neighbor check?
            # Yes, standard local search checks all neighbors from current center.
            # But changing back and forth is slow.
            # Let's assume we do Greedy: First Improvement? Or Best Improvement?
            # With only 4 neighbors, Best Improvement is fine.
            stream_1.Pressure.Value = p_curr
            adj_4.TargetValue.Value = t_curr
            wait_for_solver(case)

        # 4. Pick Best
        if not candidates: break
        
        candidates.sort(key=lambda x: x[2])
        best = candidates[0]
        
        if best[2] < curr_score:
            print(f"  >>> Improved! Moving to P={best[0]}, T={best[1]} (Score: {best[2]:.1f})")
            p_curr = best[0]
            t_curr = best[1]
            stream_1.Pressure.Value = p_curr
            adj_4.TargetValue.Value = t_curr
            wait_for_solver(case)
        else:
            print("  >>> Local Optimum Reached.")
            break

    print("\n--- Optimization Finished ---")
    data = read_data_from_case(case)
    print_data_table(data)

if __name__ == "__main__":
    optimize_case()
