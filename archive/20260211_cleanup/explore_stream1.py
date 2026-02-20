import sys
import os
import time

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def explore_stream_1():
    print("="*40)
    print(" EXPLORING STREAM 1 ")
    print("="*40)
    
    app = connect_hysys()
    if not app:
        print("Failed to connect to HYSYS.")
        return

    try:
        # Attach to active document
        mgr = HysysNodeManager(app)
        
        # Read keys mapped to Stream 1
        # Note: 'inlet.mass_flow' is Stream 10 in the manager, so we use 'control.s1_mass_flow' for Stream 1
        keys = {
            'Temperature': 'inlet.temperature',
            'Pressure': 'inlet.pressure',
            'Mass Flow': 'control.s1_mass_flow'
        }
        
        print(f"\n[Reading Stream 1 Data]...")
        
        for name, key in keys.items():
            val = mgr.read(key)
            if val is not None:
                unit = mgr.nodes[key].unit
                print(f"  > {name}: {val:.4f} {unit}")
            else:
                print(f"  > {name}: [READ ERROR]")

    except Exception as e:
        print(f"\n[ERROR] {e}")
    finally:
        # We don't dispose/close to keep it open for user
        print("\nDone.")

if __name__ == "__main__":
    explore_stream_1()
