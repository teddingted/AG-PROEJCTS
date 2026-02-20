import win32com.client

print("--- STREAM CONNECTION CHECK ---")
try:
    app = win32com.client.GetActiveObject("HYSYS.Application")
    fs = app.ActiveDocument.Flowsheet
    
    s6 = fs.MaterialStreams.Item("6")
    print(f"Stream 6 found.")
    
    # Check downstream operations
    print("Stream 6 connects to:")
    for op in s6.DownstreamOperations:
        print(f" - {op.name} ({op.TypeName})")
        
    s7 = fs.MaterialStreams.Item("7")
    print(f"Stream 7 found.")
    print("Stream 7 connects from:")
    for op in s7.UpstreamOperations:
        print(f" - {op.name} ({op.TypeName})")

except Exception as e:
    print(f"Error: {e}")
