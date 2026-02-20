import sys
import os
import time
import win32com.client

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def run_monitor_reset():
    print("="*60)
    print(" MONITORING STREAM 1 AFTER RESET ")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        mgr = HysysNodeManager(app)
        
        # 1. Set Inputs
        print("\n[STEP 1] Setting Inputs...")
        # Force write check false to speed up, we rely on reset
        mgr.write('inlet.temperature', 40.0, verify=False) 
        mgr.write('control.s1_mass_flow', 20000.0, verify=False)
        
        # 2. Reset ADJ-1
        print("\n[STEP 2] Resetting ADJ-1...")
        fs = mgr.case.Flowsheet
        adj = fs.Operations.Item("ADJ-1")
        if not adj.IsIgnored:
            adj.Reset()
            print("  > Reset Triggered.")
        else:
            print("  > ADJ-1 is Ignored. Skipping Reset.")
            
        # 3. Monitor Loop
        print("\n[STEP 3] Monitoring Logic (20 samples)...")
        print(f"{'Time':<5} | {'Temp (C)':<10} | {'Flow (kg/h)':<15} | {'Pres (bar)':<10}")
        print("-" * 50)
        
        for i in range(20):
            time.sleep(1.0)
            t = mgr.read('inlet.temperature')
            mf = mgr.read('control.s1_mass_flow')
            p = mgr.read('inlet.pressure')
            
            # Formatting
            t_str = f"{t:.2f}" if t is not None else "None"
            mf_str = f"{mf:.1f}" if mf is not None else "None"
            p_str = f"{p:.4f}" if p is not None else "None"
            
            print(f"{i+1:<5} | {t_str:<10} | {mf_str:<15} | {p_str:<10}")

    except Exception as e:
        print(f"\n[ERROR] {e}")

if __name__ == "__main__":
    run_monitor_reset()
