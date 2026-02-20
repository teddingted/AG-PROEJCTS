import csv
import os
import time
import threading
from hysys_utils import dismiss_popup
from hysys_node_manager import HysysNodeManager, connect_hysys

# --- Configuration ---
folder = "hysys_automation"
out_file = "optimization_robust_result.csv"
flows = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]

preset_p = {
    500: 3.0, 600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 7.2, 1500: 7.5
}

def run_optimization():
    print("="*60)
    print("HYSYS ROBUST OPTIMIZER (MANAGER BASED)")
    print("="*60)

    # 1. Start Popup Handler
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()

    # 2. Connect via Robust Manager
    app = connect_hysys()
    if not app:
        print("CRITICAL: Could not connect to HYSYS.")
        return

    try:
        mgr = HysysNodeManager(app)
        
        # Init CSV
        csv_path = os.path.join(folder, out_file)
        if not os.path.exists(csv_path):
            with open(csv_path, 'w', newline='') as f:
                csv.writer(f).writerow(["MassFlow", "P_bar", "T_C", "MinApp_C", "Power_kW"])

        for flow in flows:
            print(f"\nProcessing Flow: {flow} kg/h")
            # Logic similar to previous optimizer but using mgr methods
            # ... (Full logic would go here, simplified for demo) ...
            
            # Example usage:
            if not mgr.write('inlet.mass_flow', flow):
                print("Failed to set flow.")
                continue
                
            # Read back
            temp = mgr.read('inlet.temperature')
            print(f" Current T: {temp}")

    except Exception as e:
        print(f"Optimization Error: {e}")
    finally:
        if 'mgr' in locals():
            mgr.dispose()

if __name__ == "__main__":
    run_optimization()
