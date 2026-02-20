import sys
import os
import time

# Add current directory to path so we can import modules
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def check_adjust_status(mgr, adjust_name):
    """Reads status of an Adjust operation"""
    try:
        # Access the Adjust Object directly via internal method or new register
        # Since 'control.vol_flow' is registered to ADJ-1 TargetValue, we can access the object parent
        # But better to get it fresh
        fs = mgr.case.Flowsheet
        adj = fs.Operations.Item(adjust_name)
        
        # Read key properties
        target = adj.TargetValue.Value
        error = adj.Error.Value
        # adjusted_val = adj.AdjustedValue.Value # Might not be direct property
        
        # Iterations or Status?
        # status = adj.Status # Might be an integer
        
        print(f"  [{adjust_name} STATUS]")
        print(f"    Target: {target:.4f}")
        print(f"    Error:  {error:.4f}")
        
        if abs(error) < 1e-4:
            print(f"    State:  CONVERGED")
        else:
            print(f"    State:  DRIFTING / SOLVING")
            
        return dict(target=target, error=error)
        
    except Exception as e:
        print(f"  [CHECK FAIL] Could not read {adjust_name}: {e}")
        return None

def set_with_check():
    print("="*60)
    print(" SEQUENTIAL UPDATE WITH ADJUST CHECK ")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        mgr = HysysNodeManager(app)
        
        # 1. Set Temperature
        print("\n[STEP 1] Setting Temperature to 40.0 C...")
        if mgr.write('inlet.temperature', 40.0):
            print("  > Set Temp: DONE")
        else:
            print("  > Set Temp: FAILED (Likely Calculated)")
            
        print("  > Checking ADJ-1...")
        time.sleep(2.0)
        check_adjust_status(mgr, "ADJ-1")
        
        # 2. Set Mass Flow
        print("\n[STEP 2] Setting Mass Flow to 20,000 kg/h...")
        if mgr.write('control.s1_mass_flow', 20000.0):
            print("  > Set Flow: DONE")
        else:
            print("  > Set Flow: FAILED")
            
        print("  > Checking ADJ-1...")
        time.sleep(2.0)
        status = check_adjust_status(mgr, "ADJ-1")
        
        # 3. Final Observation
        print("\n[FINAL OBSERVATION]")
        for i in range(5):
            t = mgr.read('inlet.temperature')
            mf = mgr.read('control.s1_mass_flow')
            print(f"  State: T={t:.2f} C, Flow={mf:.1f} kg/h")
            time.sleep(1.0)
            
    except Exception as e:
        print(f"\n[ERROR] {e}")

if __name__ == "__main__":
    set_with_check()
