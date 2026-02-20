import sys
import os
from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

def test_set_flow():
    print("TEST: Setting Mass Flow to 500...")
    app = connect_hysys()
    mgr = HysysNodeManager(app)
    
    try:
        mgr.write('inlet.mass_flow', 500.0)
        print("SUCCESS: Set Flow to 500.")
    except Exception as e:
        print(f"FAILED: {e}")

if __name__ == "__main__":
    test_set_flow()
