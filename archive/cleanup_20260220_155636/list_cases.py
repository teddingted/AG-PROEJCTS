import win32com.client

print("--- LIST HYSYS CASES ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    print(f"HYSYS Version: {app.Version}")
    
    print(f"Open Cases Count: {app.SimulationCases.Count}")
    
    for i in range(app.SimulationCases.Count):
        # Index is 0-based or 1-based? Usually 0-based in Python for COM collection iteration?
        # But safest is for-loop over collection
        pass

    for c in app.SimulationCases:
        print(f" - {c.Title.Value} (Path: {c.FullPath})")
        
    if app.ActiveDocument:
        print(f"Active Document: {app.ActiveDocument.Title.Value}")
    else:
        print("Active Document is NONE")

except Exception as e:
    print(f"Error: {e}")
