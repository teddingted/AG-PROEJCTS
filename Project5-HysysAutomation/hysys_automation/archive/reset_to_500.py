import os
import time
import win32com.client
import threading
from hysys_utils import dismiss_popup

def wait_for_solver(case, timeout=60):
    """Wait for solver to finish"""
    start = time.time()
    while True:
        try:
            if not case.Solver.IsSolving:
                break
        except:
            pass
        
        if time.time() - start > timeout:
            break
        
        time.sleep(0.3)
    
    time.sleep(1.5)
    return True

def reset_all_adjusts(case):
    """Reset all Adjust blocks"""
    for adj_name in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
        try:
            adj = case.Flowsheet.Operations.Item(adj_name)
            adj.Reset()
        except:
            pass

def main():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    if os.path.exists(file_path):
        case = app.SimulationCases.Open(file_path)
        case.Visible = True
    else:
        case = app.ActiveDocument
    
    print("="*80)
    print("RESETTING TO 500 kg/h BASELINE")
    print("="*80)
    
    # 500 kg/h optimal conditions
    OPTIMAL_500 = {
        'flow': 500,
        'p1': 3.0,      # bar
        't_adj4': -111.0  # C
    }
    
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        print(f"\nSetting 500 kg/h baseline:")
        print(f"  Flow: {OPTIMAL_500['flow']} kg/h")
        print(f"  P1: {OPTIMAL_500['p1']} bar")
        print(f"  T_ADJ4: {OPTIMAL_500['t_adj4']} C")
        
        # Set conditions
        s10.MassFlow.Value = OPTIMAL_500['flow'] / 3600.0
        wait_for_solver(case, timeout=45)
        
        reset_all_adjusts(case)
        wait_for_solver(case, timeout=45)
        
        s1.Pressure.Value = OPTIMAL_500['p1'] * 100.0  # Convert to kPa
        adj4.TargetValue.Value = OPTIMAL_500['t_adj4']
        wait_for_solver(case, timeout=60)
        
        # Verify
        print("\nVerifying state...")
        flow10 = s10.MassFlow.Value * 3600.0
        p1 = s1.Pressure.Value / 100.0
        p7 = s7.Pressure.Value / 100.0
        t_adj4 = adj4.TargetValue.Value
        min_app = lng.MinApproach.Value
        power = ss.Cell("C8").CellValue
        
        print(f"\nActual state:")
        print(f"  Flow: {flow10:.1f} kg/h")
        print(f"  P1: {p1:.1f} bar")
        print(f"  T_ADJ4: {t_adj4:.1f} C")
        print(f"  Min Approach: {min_app:.2f} C")
        print(f"  P7: {p7:.2f} bar")
        print(f"  Power: {power:.2f} kW")
        
        print("\n" + "="*80)
        print("RESET COMPLETE - Ready for next steps")
        print("="*80)
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
