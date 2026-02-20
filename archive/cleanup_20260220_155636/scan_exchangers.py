import win32com.client

print("--- EXCHANGER SCAN ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    fs = app.ActiveDocument.Flowsheet
    
    for op in fs.Operations:
        if "Exchanger" in op.TypeName or "LNG" in op.TypeName:
            print(f"Name: {op.name}, Type: {op.TypeName}")
            try:
                print(f"  MinApproach: {op.MinApproach.Value}")
            except:
                print("  (Cannot read MinApproach)")

except Exception as e:
    print(f"Error: {e}")
