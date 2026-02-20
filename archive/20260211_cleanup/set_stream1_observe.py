import sys
import os
import time
from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def set_and_observe():
    print("="*50)
    print(" SETTING STREAM 1 & OBSERVING ")
    print(" Target: Temp = 40.0 C, MassFlow = 20,000 kg/h")
    print("="*50)
    
    app = connect_hysys()
    if not app: return

    try:
        mgr = HysysNodeManager(app)
        
        # 1. Write Values
        print("\n[WRITING INPUTS]...")
        
        # Note: In the node manager, 'inlet.temperature' is Stream 1 Temp
        # 'control.s1_mass_flow' is Stream 1 Mass Flow (registered as extra node)
        
        # Set Temperature
        if mgr.write('inlet.temperature', 40.0):
            print("  > Set Temperature: OK")
        else:
            print("  > Set Temperature: FAILED")
            
        # Set Mass Flow
        # Use existing node 'control.s1_mass_flow' which maps to Stream 1 MassFlow
        if mgr.write('control.s1_mass_flow', 20000.0):
             print("  > Set Mass Flow: OK")
        else:
             print("  > Set Mass Flow: FAILED")

        # 2. Observe Stability
        print("\n[OBSERVING RESPONSE (10s)]...")
        for i in range(10):
            time.sleep(1.0)
            
            # Read feedback
            t = mgr.read('inlet.temperature')
            mf_s10 = mgr.read('inlet.mass_flow') # Stream 10 (Main Feed)
            mf_s1 = mgr.read('control.s1_mass_flow') # Stream 1 (N2)
            pwr = mgr.read('result.compressor_power')
            
            print(f"  T+ {i+1}s | S1 Temp: {t:.2f} C | S1 Flow: {mf_s1:.1f} kg/h | Power: {pwr:.1f} kW")
            
    except Exception as e:
        print(f"\n[ERROR] {e}")

if __name__ == "__main__":
    set_and_observe()
