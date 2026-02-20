import os
import win32com.client
import threading
from hysys_utils import dismiss_popup

def inspect_adj_variables():
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
        else:
             case = app.ActiveDocument
    except: return

    try:
        adj = case.Flowsheet.Operations.Item("ADJ-1")
        print(f"Inspecting {adj.Name}...")
        
        # Try listing variables
        # Usually op.VariableNames or we iterate UserVariables
        # Or look at specific known properties for Adjust
        
        # 1. Check UserVariables (often where buttons live in older versions)
        print("User Variables:")
        try:
            for i in range(adj.UserVariables.Count):
                var = adj.UserVariables.Item(i)
                print(f"  {var.Name} = {var.Value}")
        except: print("  (None or Error)")
        
        # 2. Check direct properties via __dir__() or standard attributes if possible
        # Adjust blocks often have: 'Start', 'Continue', 'Reset'
        # Let's try to access them and print status
        potentials = ["Reset", "ResetCalculation", "Start", "Calculate"]
        print("\nPotential Properties:")
        for p in potentials:
            try:
                # Try getting the variable object
                v = adj.Variable(p)
                print(f"  Found Variable: {p}, Value: {v.Value}")
            except:
                pass
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_adj_variables()
