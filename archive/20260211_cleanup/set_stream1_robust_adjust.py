import sys
import os
import time
import win32com.client

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def run_robust_update():
    print("="*60)
    print(" ROBUST UPDATE STREAM 1 (WITH ADJ-1 RESET) ")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        mgr = HysysNodeManager(app)
        
        # 1. Set Inputs
        print("\n[STEP 1] Setting Inputs...")
        mgr.write('inlet.temperature', 40.0, verify=False)
        mgr.write('control.s1_mass_flow', 20000.0)
        
        # 2. Check and Reset ADJ-1
        print("\n[STEP 2] Checking ADJ-1 Status...")
        try:
            fs = mgr.case.Flowsheet
            adj = fs.Operations.Item("ADJ-1")
            
            # Use 'IsIgnored' as discovered in scan
            is_ignored = adj.IsIgnored
            print(f"  > ADJ-1 IsIgnored: {is_ignored}")
            
            if not is_ignored:
                print("  > Adjust is ACTIVE. Triggering RESET to force convergence...")
                adj.Reset()
                print("  > Reset Triggered. Waiting for Solver (10s)...")
                time.sleep(10.0)
            else:
                print("  > Adjust is IGNORED. Skipping Reset.")
                
        except Exception as e:
            print(f"  > [ADJ CHECK FAIL] {e}")

        # 3. Final Observation
        print("\n[FINAL STATE]")
        t = mgr.read('inlet.temperature')
        mf = mgr.read('control.s1_mass_flow')
        p = mgr.read('inlet.pressure')
        
        print(f"  Temp: {t:.2f} C")
        print(f"  Flow: {mf:.1f} kg/h")
        print(f"  Pres: {p:.4f} bar")
        
        # Check if Temp is closer to 40?
        if abs(t - 40.0) < 1.0:
            print("  > SUCCESS: Temperature moved to target.")
        else:
            print("  > RESULT: Temperature still drifting (likely constrained).")

    except Exception as e:
        print(f"\n[ERROR] {e}")

if __name__ == "__main__":
    run_robust_update()
