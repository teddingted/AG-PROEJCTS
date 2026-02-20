import os
import win32com.client
import threading
from hysys_utils import dismiss_popup

def inspect_s10():
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
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        print(f"Stream 10 Found: {s10.Name}")
        print(f"  MassFlow (Value): {s10.MassFlow.Value} (Internal)")
        # Usually internal is kg/s?
        # Let's try setting it to 500 kg/h = 500/3600 kg/s = 0.138
        
        # Check if we can set it
        # s10.MassFlow.Value = 0.1388
    except Exception as e:
        print(f"Error accessing Stream 10: {e}")

if __name__ == "__main__":
    inspect_s10()
