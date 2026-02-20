import win32com.client

print("--- HYSYS ITEM CHECK ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    fs = app.ActiveDocument.Flowsheet
    print(f"Flowsheet: {fs.Name}")
    
    # Try finding LNG-100
    try:
        lng = fs.Operations.Item("LNG-100")
        print(f"Found LNG-100 (Type: {lng.TypeName})")
        print(f"MinApproach: {lng.MinApproach.Value}")
    except Exception as e:
        print(f"Could not find LNG-100: {e}")
        
    # List all Operations
    print("\nList of Operations:")
    for op in fs.Operations:
        print(f" - {op.name}")
        
except Exception as e:
    print(f"Error: {e}")
