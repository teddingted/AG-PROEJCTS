import os
import win32com.client
from hysys_utils import dismiss_popup
import threading
import time

def inspect_status():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    
    case = None
    try:
        if os.path.exists(file_path):
            t = threading.Thread(target=dismiss_popup)
            t.daemon = True
            t.start()
            case = app.SimulationCases.Open(file_path)
            case.Visible = True
        else:
            case = app.ActiveDocument
    except:
        case = app.ActiveDocument

    print("--- Operation Status Inspection ---")
    ops = case.Flowsheet.Operations

    for op in ops:
        name = op.Name
        if name.startswith("ADJ") or name == "LNG-100":
             print(f"\nOp: {name}")
             
             # Ignored?
             try: print(f"  Ignored: {op.Ignored}")
             except Exception as e: print(f"  Ignored Err: {e}")
             
             # Adjust specific
             if name.startswith("ADJ"):
                 try:
                     # Attempt to read 'Converged' variable if accessible
                     # Or check iteration count
                     pass
                 except: pass

             # Check ErrorStatus
             try:
                 # Some objects have ErrorStatus
                 print(f"  ErrorStatus: {op.ErrorStatus}")
             except: pass

    print("\n--- Stream 7 Inspection ---")
    try:
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        print(f"Stream 7 Pressure: {s7.Pressure.Value} kPa")
    except Exception as e:
        print(f"Stream 7 Error: {e}")

if __name__ == "__main__":
    inspect_status()
