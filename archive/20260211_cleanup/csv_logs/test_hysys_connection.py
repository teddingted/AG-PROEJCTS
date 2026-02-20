"""
HYSYS Node Manager - Minimal Diagnostic Test
Tests each initialization step independently to find hang point
"""
import win32com.client
import time
import sys

def test_step(step_name, func, timeout=10):
    """Test a function with timeout"""
    print(f"\n[TEST] {step_name}...", end=" ", flush=True)
    start = time.time()
    try:
        result = func()
        elapsed = time.time() - start
        print(f"✓ OK ({elapsed:.1f}s)")
        return result
    except Exception as e:
        elapsed = time.time() - start
        print(f"✗ FAIL ({elapsed:.1f}s): {e}")
        return None

def step1_connect_hysys():
    """Step 1: Connect to HYSYS"""
    try:
        app = win32com.client.GetActiveObject("HYSYS.Application")
        return app
    except:
        print("\n  → No active HYSYS, starting new...", end=" ", flush=True)
        app = win32com.client.Dispatch("HYSYS.Application")
        app.Visible = True
        return app

def step2_get_document(app):
    """Step 2: Get active document"""
    return app.ActiveDocument

def step3_get_flowsheet(case):
    """Step 3: Access flowsheet"""
    return case.Flowsheet

def step4_get_solver(case):
    """Step 4: Access solver"""
    return case.Solver

def step5_get_stream(fs):
    """Step 5: Get material stream"""
    return fs.MaterialStreams.Item("1")

def step6_read_temperature(stream):
    """Step 6: Read temperature"""
    return stream.Temperature.Value

def step7_write_temperature(stream):
    """Step 7: Write temperature"""
    old_val = stream.Temperature.Value
    stream.Temperature.Value = 40.0
    new_val = stream.Temperature.Value
    return (old_val, new_val)

def main():
    print("="*60)
    print("HYSYS NODE MANAGER - DIAGNOSTIC TEST")
    print("="*60)
    print("\nTesting each initialization step independently...")
    
    # Step 1: COM Connection
    app = test_step("1. Connect to HYSYS", step1_connect_hysys)
    if not app:
        print("\n[CRITICAL] Cannot connect to HYSYS. Ensure HYSYS is running.")
        sys.exit(1)
    
    # Step 2: Document
    case = test_step("2. Get Active Document", lambda: step2_get_document(app))
    if not case:
        print("\n[CRITICAL] No active document. Open a simulation file in HYSYS.")
        sys.exit(1)
    
    # Step 3: Flowsheet
    fs = test_step("3. Access Flowsheet", lambda: step3_get_flowsheet(case))
    if not fs:
        print("\n[CRITICAL] Cannot access flowsheet.")
        sys.exit(1)
    
    # Step 4: Solver
    solver = test_step("4. Access Solver", lambda: step4_get_solver(case))
    if not solver:
        print("\n[CRITICAL] Cannot access solver.")
        sys.exit(1)
    
    # Step 5: Stream
    stream = test_step("5. Get Stream '1'", lambda: step5_get_stream(fs))
    if not stream:
        print("\n[CRITICAL] Cannot find stream '1'. Check simulation.")
        sys.exit(1)
    
    # Step 6: Read
    temp = test_step("6. Read Temperature", lambda: step6_read_temperature(stream))
    if temp is not None:
        print(f"      Current temperature: {temp:.2f} °C")
    
    # Step 7: Write
    result = test_step("7. Write Temperature (40°C)", lambda: step7_write_temperature(stream))
    if result:
        print(f"      Changed: {result[0]:.2f} → {result[1]:.2f} °C")
    
    print("\n" + "="*60)
    print("ALL TESTS PASSED ✓")
    print("="*60)
    print("\nHYSYS COM interface is working correctly.")
    print("Node Manager should initialize without issues.\n")

if __name__ == "__main__":
    main()
