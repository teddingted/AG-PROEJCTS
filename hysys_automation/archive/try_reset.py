import os
import win32com.client
import threading
from hysys_utils import dismiss_popup

def try_reset_methods():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    try:
        if os.path.exists(file_path):
             case = app.SimulationCases.Open(file_path)
             case.Visible = True
        else:
             case = app.ActiveDocument
    except: return

    try:
        adj = case.Flowsheet.Operations.Item("ADJ-1")
        print(f"Target: {adj.Name}")
        
        # Method 1: direct Reset()
        print("Attempting adj.Reset()...")
        try:
            adj.Reset()
            print("  > Success (Method call)")
        except Exception as e:
            print(f"  > Failed: {e}")

        # Method 2: Reset variable
        print("Attempting adj.Reset = 1...")
        try:
            adj.Reset = True # or 1
            print("  > Success (Property set)")
        except Exception as e:
            print(f"  > Failed: {e}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    try_reset_methods()
