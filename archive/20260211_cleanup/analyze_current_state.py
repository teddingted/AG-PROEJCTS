import os, time
import win32com.client

"""
HYSYS SNAPSHOT ANALYZER
Reads current state of active HYSYS case to reverse-engineer successful convergence.
"""

def analyze_state():
    try:
        try:
            app = win32com.client.GetActiveObject("HYSYS.Application")
        except:
            print("ERROR: Could not attach to active HYSYS. Is it running?")
            return

        case = app.ActiveDocument
        if not case:
            print("ERROR: No active case.")
            return

        fs = case.Flowsheet
        solver = case.Solver
        
        print("\n=== HYSYS STATE SNAPSHOT ===")
        print(f"Case: {case.Name}")
        print(f"Solver Status: {'SOLVING' if solver.IsSolving else 'IDLE'}")
        print(f"Can Solve: {solver.CanSolve}")
        
        # Streams
        s1 = fs.MaterialStreams.Item("1")
        s6 = fs.MaterialStreams.Item("6")
        s10 = fs.MaterialStreams.Item("10")
        
        print(f"\n[Stream 10 - Feed]")
        print(f"  Flow: {s10.MassFlow.Value * 3600.0:.2f} kg/h")
        print(f"  Temp: {s10.Temperature.Value:.2f} C")
        print(f"  Pres: {s10.Pressure.Value / 100.0:.2f} bar")

        print(f"\n[Stream 1 - N2 Loop High]")
        print(f"  Flow: {s1.MassFlow.Value * 3600.0:.2f} kg/h")
        print(f"  Temp: {s1.Temperature.Value:.2f} C")
        print(f"  Pres: {s1.Pressure.Value / 100.0:.2f} bar")
        
        print(f"\n[Stream 6 - N2 Loop Low]")
        print(f"  Pres: {s6.Pressure.Value / 100.0:.2f} bar")
        
        # Operations
        adj4 = fs.Operations.Item("ADJ-4")
        lng = fs.Operations.Item("LNG-100")
        comp = fs.Operations.Item("K-100")
        sprdsht = fs.Operations.Item("SPRDSHT-1")
        
        print(f"\n[ADJ-4 Target]")
        print(f"  TargetValue: {adj4.TargetValue.Value:.2f}")
        print(f"  UserValue: {adj4.UserValue}")
        print(f"  CalculatedValue: {adj4.CalculatedValue}")
        
        print(f"\n[LNG-100 Exchanger]")
        print(f"  MinApproach: {lng.MinApproach.Value:.4f} C")
        print(f"  LMTD: {lng.LMTD.Value:.2f} C")
        print(f"  UA: {lng.UA.Value:.2f}")
        
        print(f"\n[K-100 Compressor]")
        print(f"  Power: {comp.EnergyValue:.2f} kW")
        print(f"  Efficiency: {comp.AdiabaticEfficiency.Value:.2f} %")
        
        print(f"\n[Spreadsheet Power]")
        try:
            print(f"  Cell C8: {sprdsht.Cell('C8').CellValue:.2f} kW")
        except:
            print("  Cell C8: Error reading")
            
        print("\n============================")
        
    except Exception as e:
        print(f"CRITICAL ERROR: {str(e)}")

if __name__ == "__main__":
    analyze_state()
