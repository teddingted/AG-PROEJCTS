import sys
import os
import win32com.client

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import connect_hysys

def scan_adj():
    print("="*60)
    print(" SCANNING ADJ-1 PROPERTIES ")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        # Direct COM access
        case = app.ActiveDocument
        fs = case.Flowsheet
        adj = fs.Operations.Item("ADJ-1")
        
        print(f"\nObject: {adj.Name}")
        print(f"Type: {adj.TypeName}")
        
        # 1. Try Standard Properties
        props = ["Ignored", "TargetValue", "AdjustedValue", "MeasuredValue", "Error", "Tolerance", "Status", "MaximumIterations", "SignalError"]
        
        print("\n[Standard Property Check]")
        for p in props:
            try:
                val = getattr(adj, p)
                # If it's an object with .Value, try to read it
                if hasattr(val, 'Value'):
                    v_str = f"{val.Value}"
                    print(f"  > {p}: {v_str} (Has .Value)")
                else:
                    print(f"  > {p}: {val} (Direct)")
            except Exception as e:
                print(f"  > {p}: [ACCESS FAILED] {e}")

        # 2. Try Reset Method
        print("\n[Method Check]")
        if hasattr(adj, 'Reset'):
            print("  > Reset(): Available")
        else:
            print("  > Reset(): Not Found")
            
        if hasattr(adj, 'Calculate'):
            print("  > Calculate(): Available")

    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")

if __name__ == "__main__":
    scan_adj()
