import os
import time
import win32com.client
import threading
from hysys_utils import dismiss_popup, read_data_from_case, print_data_table

def wait_for_solver(case, timeout=60):
    """Wait for solver to finish"""
    start = time.time()
    while True:
        try:
            if not case.Solver.IsSolving:
                return True
        except: pass
        if time.time() - start > timeout:
            print("    > Solver Timeout.")
            return False
        time.sleep(0.5)

def reset_all_adjusts(case):
    """Reset all non-ignored Adjust blocks"""
    adj_names = ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]
    
    print("\n=== Resetting All Adjust Blocks ===")
    for name in adj_names:
        try:
            adj = case.Flowsheet.Operations.Item(name)
            adj.Reset()
            print(f"  [OK] Reset {name}")
        except Exception as e:
            print(f"  [FAIL] Failed to reset {name}: {e}")
    
    # Wait for solver after resets
    print("\nWaiting for solver after resets...")
    wait_for_solver(case)
    print("Solver idle.\n")

def get_snapshot(case):
    """Get current simulation state"""
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        p1 = s1.Pressure.Value / 100.0  # bar
        p7 = s7.Pressure.Value / 100.0  # bar
        flow10 = s10.MassFlow.Value * 3600.0  # kg/h
        min_app = lng.MinApproach.Value
        power = ss.Cell("C8").CellValue
        
        print(f"Stream 1 P: {p1:.2f} bar")
        print(f"Stream 7 P: {p7:.2f} bar")
        print(f"Stream 10 Flow: {flow10:.1f} kg/h")
        print(f"Min Approach: {min_app:.2f} °C")
        print(f"Power (C8): {power:.2f} kW")
        
        return {
            'p1': p1,
            'p7': p7,
            'flow10': flow10,
            'min_app': min_app,
            'power': power,
            'valid': min_app > -100  # Simple validity check
        }
    except Exception as e:
        print(f"Error reading snapshot: {e}")
        return None

def test_convergence():
    """Test convergence at a single operating point"""
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    # Connect to HYSYS
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
    print("CONVERGENCE TEST - Debugging Script")
    print("="*80)
    
    # Step 1: Check current state
    print("\n--- Step 1: Current State (User Reset) ---")
    snap = get_snapshot(case)
    
    # Step 2: Reset all Adjust blocks
    reset_all_adjusts(case)
    
    print("\n--- Step 2: State After Adjust Resets ---")
    snap = get_snapshot(case)
    
    # Step 3: Try setting a known flow condition (500 kg/h)
    print("\n--- Step 3: Setting Flow to 500 kg/h ---")
    try:
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        s10.MassFlow.Value = 500.0 / 3600.0  # kg/s
        print("Flow set to 500 kg/h")
    except Exception as e:
        print(f"Error setting flow: {e}")
    
    print("\nWaiting for solver...")
    wait_for_solver(case, timeout=120)  # Longer timeout
    
    print("\n--- Step 4: State After Flow Change ---")
    snap = get_snapshot(case)
    
    # Step 4: If still invalid, try setting pressure to 3 bar
    if snap and not snap['valid']:
        print("\n!!! Simulation not converged, trying pressure adjustment...")
        print("\n--- Step 5: Setting Pressure to 3 bar ---")
        try:
            s1 = case.Flowsheet.MaterialStreams.Item("1")
            s1.Pressure.Value = 300.0  # kPa
            print("Pressure set to 3 bar")
        except Exception as e:
            print(f"Error setting pressure: {e}")
        
        print("\nWaiting for solver...")
        wait_for_solver(case, timeout=120)
        
        # Reset adjusts again
        reset_all_adjusts(case)
        
        print("\n--- Step 6: State After Pressure Change + Reset ---")
        snap = get_snapshot(case)
    
    # Final check
    print("\n" + "="*80)
    print("FINAL DIAGNOSIS")
    print("="*80)
    if snap and snap['valid']:
        print("[SUCCESS] Simulation CONVERGED")
        print(f"  Min Approach: {snap['min_app']:.2f} °C")
        print(f"  Power: {snap['power']:.2f} kW")
    else:
        print("[FAILED] Simulation FAILED TO CONVERGE")
        print("\nPossible issues:")
        print("  1. Flow change requires manual intervention in HYSYS")
        print("  2. ADJ blocks need specific sequence of operations")
        print("  3. Initial guess values incompatible with new flow")
        print("  4. Additional constraints/specifications active")
    
    print("\nFull spreadsheet data:")
    data = read_data_from_case(case)
    print_data_table(data)

if __name__ == "__main__":
    test_convergence()
