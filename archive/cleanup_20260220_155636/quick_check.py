import win32com.client
import time

print("--- HYSYS CONNECTION CHECK ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    print("Connected to HYSYS Application")
    print(f"Version: {app.Version}")
    print(f"Active Case: {app.ActiveDocument.Title.Value}")
    
    fs = app.ActiveDocument.Flowsheet
    print("Flowsheet accessed.")
    
    # Test Node Access
    s10 = fs.MaterialStreams.Item("10")
    print(f"Stream 10 Mass Flow: {s10.MassFlow.Value * 3600} kg/h")
    
    adj1 = fs.Operations.Item("ADJ-1")
    print(f"ADJ-1 Target: {adj1.TargetValue.Value}")
    
    print("--- SUCCESS ---")
except Exception as e:
    print(f"--- FAILURE: {e} ---")
