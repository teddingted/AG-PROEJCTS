import win32com.client

print("--- ALL OPS ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    fs = app.ActiveDocument.Flowsheet
    
    count = 0
    for op in fs.Operations:
        if count < 10: # Sample 10 for debug
            print(f"Name: {op.name}, Type: {op.TypeName}")
        count += 1
    
    # Try finding LNG-100 specifically by name in collection
    try:
        lng = fs.Operations.Item("LNG-100")
        print(f"LNG-100 FOUND! Type: {lng.TypeName}")
    except:
        print("LNG-100 NOT FOUND in Operations collection.")

except Exception as e:
    print(f"Error: {e}")
