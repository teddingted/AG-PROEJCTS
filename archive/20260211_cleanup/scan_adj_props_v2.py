import sys
import os
import win32com.client

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import connect_hysys

def scan_adj_v2():
    print("="*60)
    print(" SCANNING ADJ-1 PROPERTIES (DEEP) ")
    print("="*60)
    
    app = connect_hysys()
    if not app: return

    try:
        case = app.ActiveDocument
        fs = case.Flowsheet
        adj = fs.Operations.Item("ADJ-1")
        
        # 1. Alternative Status Flags
        flags = ["Active", "IsIgnored", "Calculating", "Solved", "UnderSpecified"]
        print("\n[Status Flags]")
        for f in flags:
            try:
                val = getattr(adj, f)
                print(f"  > {f}: {val}")
            except:
                pass # print(f"  > {f}: [FAIL]")

        # 2. Variable Access
        print("\n[Variables]")
        try:
            mv = adj.MeasuredVariable
            print(f"  > MeasuredVariable: {mv.Name} (Type: {mv.TypeName})")
            if hasattr(mv, 'Value'): print(f"    Value: {mv.Value}")
            elif hasattr(mv, 'Variable'): print(f"    VarValue: {mv.Variable.Value}")
        except Exception as e:
            print(f"  > MeasuredVariable: [FAIL] {e}")

        try:
            av = adj.AdjustedVariable
            print(f"  > AdjustedVariable: {av.Name} (Type: {av.TypeName})")
        except Exception as e:
            print(f"  > AdjustedVariable: [FAIL] {e}")

        # 3. Known HYSYS Generic Get/Set
        # Sometimes direct access fails but generic GetVariable works if wrapped
        # But HYSYS automation usually uses properties.
        
        # 4. Check "Ignored" again with proper Error handling
        # Verify if it's truly missing or just failed for some reason
        try:
            ign = adj.Ignored
            print(f"  > Ignored: {ign}")
        except:
             print("  > Ignored: [FAIL] Again")

    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")

if __name__ == "__main__":
    scan_adj_v2()
