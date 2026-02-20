import os, time
import win32com.client

"""
HYSYS SNAPSHOT ANALYZER V2
- Reads MA correctly
- No UserValue error
"""

def analyze_state():
    try:
        try:
            app = win32com.client.GetActiveObject("HYSYS.Application")
        except:
            app = win32com.client.Dispatch("HYSYS.Application")

        case = app.ActiveDocument
        if not case: return

        fs = case.Flowsheet
        solver = case.Solver
        
        print("\n=== HYSYS STATE SNAPSHOT V2 ===")
        # Streams
        s1 = fs.MaterialStreams.Item("1")
        s10 = fs.MaterialStreams.Item("10")
        
        print(f"Flow (S10): {s10.MassFlow.Value * 3600.0:.2f} kg/h")
        print(f"Pres (S1): {s1.Pressure.Value / 100.0:.2f} bar")
        
        # Operations
        adj4 = fs.Operations.Item("ADJ-4")
        lng = fs.Operations.Item("LNG-100")
        comp = fs.Operations.Item("K-100")
        sprdsht = fs.Operations.Item("SPRDSHT-1")
        
        print(f"Target Temp (ADJ-4): {adj4.TargetValue.Value:.2f} C")
        print(f"Min Approach (LNG): {lng.MinApproach.Value:.4f} C")
        
        try:
            p = sprdsht.Cell('C8').CellValue
            print(f"Power (C8): {p:.2f} kW")
        except:
             print(f"Power (K-100): {comp.EnergyValue:.2f} kW")

        print("\n============================")
        
    except Exception as e:
        print(f"ERROR: {str(e)}")

if __name__ == "__main__":
    analyze_state()
