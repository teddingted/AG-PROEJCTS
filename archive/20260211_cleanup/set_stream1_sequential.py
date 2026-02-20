import sys
import os
import time
from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def set_sequential():
    print("="*60)
    print(" SEQUENTIAL UPDATE: STREAM 1 ")
    print(" Target: Temp=40C -> (Wait) -> Flow=20000kg/h")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        mgr = HysysNodeManager(app)
        
        # 1. Initial State
        print("\n[INITIAL STATE]")
        p_init = mgr.read('inlet.pressure')
        print(f"  Pressure: {p_init} bar")
        
        # 2. Set Temperature
        print("\n[STEP 1] Setting Temperature to 40.0 C...")
        if mgr.write('inlet.temperature', 40.0):
            print("  > Success.")
        else:
            print("  > Failed.")
            
        print("  > Waiting 5 seconds...")
        time.sleep(5.0)
        
        # 3. Set Mass Flow
        print("\n[STEP 2] Setting Mass Flow to 20,000 kg/h...")
        if mgr.write('control.s1_mass_flow', 20000.0):
            print("  > Success.")
        else:
            print("  > Failed.")
            
        print("  > Waiting 5 seconds...")
        time.sleep(5.0)

        # 4. Final Observation
        print("\n[OBSERVING FINAL STATE (10s)]...")
        for i in range(10):
            t = mgr.read('inlet.temperature')
            mf = mgr.read('control.s1_mass_flow')
            p = mgr.read('inlet.pressure')
            pwr = mgr.read('result.compressor_power')
            
            print(f"  {i+1}s | T: {t:.2f} C | Flow: {mf:.1f} kg/h | P: {p:.4f} bar | Pwr: {pwr:.1f} kW")
            time.sleep(1.0)
            
    except Exception as e:
        print(f"\n[ERROR] {e}")

if __name__ == "__main__":
    set_sequential()
