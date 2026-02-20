import os
import win32com.client
import threading
from hysys_utils import dismiss_popup
import time

def inspect_adj():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    
    # Open
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    try:
        if os.path.exists(file_path):
             case = app.SimulationCases.Open(file_path)
        else:
             case = app.ActiveDocument
        case.Visible = True
    except: return

    try:
        adj1 = case.Flowsheet.Operations.Item("ADJ-1")
        print(f"ADJ-1 Found: {adj1.Name}")
        
        # Check Converged
        try: print(f"  Converged (Prop): {adj1.Converged}")
        except Exception as e: print(f"  Converged Err: {e}")
        
        # Check Reset
        try: 
            # Check if Reset method exists (hard to check in Python without calling, but can try calling it safely?)
            # Usually Reset methods don't return value.
            # We won't call it blindly yet, just check other status.
            pass
        except: pass
        
        # Check Ignored
        print(f"  Ignored: {adj1.Ignored}")
        
        # Check Recycle
        try:
            rcY = case.Flowsheet.Operations.Item("RCY-1") # Guessing name or looking for one
            print(f"RCY-1 Found: {rcY.Name}")
            print(f"  Recycle Converged: {rcY.Converged}")
        except:
            print("No RCY-1 found or error.")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_adj()
