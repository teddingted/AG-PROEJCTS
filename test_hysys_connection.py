import win32com.client
import time

"""
HYSYS 연결 테스트 스크립트
"""

print("=" * 60)
print("HYSYS CONNECTION TEST")
print("=" * 60)

try:
    print("\n[1] Attempting to connect to HYSYS...")
    try:
        app = win32com.client.GetActiveObject("HYSYS.Application")
        print("    ✓ Connected via GetActiveObject")
    except:
        print("    GetActiveObject failed. Trying Dispatch...")
        app = win32com.client.Dispatch("HYSYS.Application")
        print("    ✓ Connected via Dispatch")
    
    print(f"\n[2] HYSYS Version: {app.Version}")
    
    print("\n[3] Accessing Active Case...")
    case = app.ActiveDocument
    print(f"    ✓ Case Name: {case.Title.Value}")
    
    print("\n[4] Accessing Flowsheet...")
    fs = case.Flowsheet
    print("    ✓ Flowsheet accessed")
    
    print("\n[5] Testing Node Access...")
    
    # Test Stream 1
    try:
        s1 = fs.MaterialStreams.Item("1")
        print(f"    ✓ Stream 1: P={s1.Pressure.Value/100:.2f} bar, T={s1.Temperature.Value:.2f} °C")
    except Exception as e:
        print(f"    ✗ Stream 1 Error: {e}")
    
    # Test Stream 10
    try:
        s10 = fs.MaterialStreams.Item("10")
        print(f"    ✓ Stream 10: MassFlow={s10.MassFlow.Value*3600:.2f} kg/h")
    except Exception as e:
        print(f"    ✗ Stream 10 Error: {e}")
    
    # Test ADJ-1
    try:
        adj1 = fs.Operations.Item("ADJ-1")
        target_val = adj1.TargetValue.Value
        print(f"    ✓ ADJ-1: Current Target={target_val}")
        
        # Test writing
        print("\n[6] Testing ADJ-1 Write Access...")
        test_val = 3600.0
        adj1.TargetValue.Value = test_val
        time.sleep(2)
        new_val = adj1.TargetValue.Value
        print(f"    Set: {test_val}, Read: {new_val}")
        if abs(new_val - test_val) < 1:
            print("    ✓ Write successful")
        else:
            print("    ✗ Write failed (value didn't stick)")
            
    except Exception as e:
        print(f"    ✗ ADJ-1 Error: {e}")
    
    # Test ADJ-4
    try:
        adj4 = fs.Operations.Item("ADJ-4")
        print(f"    ✓ ADJ-4: Target={adj4.TargetValue.Value}")
    except Exception as e:
        print(f"    ✗ ADJ-4 Error: {e}")
    
    print("\n" + "=" * 60)
    print("TEST COMPLETED")
    print("=" * 60)
    
except Exception as e:
    print(f"\n✗ CRITICAL ERROR: {e}")
    import traceback
    traceback.print_exc()
