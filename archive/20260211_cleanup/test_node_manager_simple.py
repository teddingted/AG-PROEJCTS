import sys
import os
import time

# Add current directory to path so we can import modules
sys.path.append(os.getcwd())

print("STARTING TEST...", flush=True)

try:
    from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys
    print("Module imported successfully.", flush=True)
except Exception as e:
    print(f"Import failed: {e}", flush=True)
    sys.exit(1)

print("Connecting...", flush=True)
app = connect_hysys()
if app:
    print("Connected!", flush=True)
    try:
        # Try to init manager. If no active doc, this might fail unless we pass a path.
        # But we want to test 'active doc' attachment first.
        try:
            mgr = HysysNodeManager(app)
            print("Manager initialized (Active Document).", flush=True)
        except RuntimeError:
            print("No Active Document found. Attempting to open default file...", flush=True)
            # Define default file path
            sim_file = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
            full_path = os.path.abspath(os.path.join("hysys_automation", sim_file))
            
            if os.path.exists(full_path):
                mgr = HysysNodeManager(app, full_path)
                print(f"Manager initialized with file: {sim_file}", flush=True)
            else:
                print(f"File not found: {full_path}", flush=True)
                raise RuntimeError("No file to open.")

        val = mgr.read('inlet.temperature')
        print(f"Read Temperature: {val}", flush=True)
        
        # mgr.dispose() 
        # COMMENTED OUT DISPOSE to keep HYSYS open for user inspection
        print("Test Complete. HYSYS should remain open.", flush=True)
        
    except Exception as e:
        print(f"Manager Error: {e}", flush=True)
else:
    print("Failed to connect.", flush=True)
