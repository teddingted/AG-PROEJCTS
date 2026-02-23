import os
import time
import win32com.client
import threading
from hysys_utils import dismiss_popup

def read_current_state_500(case):
    """
    Read the current state at 500 kg/h (manually set by user).
    Returns a dict with the state information.
    """
    try:
        s1 = case.Flowsheet.MaterialStreams.Item("1")
        s7 = case.Flowsheet.MaterialStreams.Item("7")
        s10 = case.Flowsheet.MaterialStreams.Item("10")
        adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        lng = case.Flowsheet.Operations.Item("LNG-100")
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        flow10 = s10.MassFlow.Value * 3600.0  # kg/h
        p1 = s1.Pressure.Value / 100.0  # bar
        p7 = s7.Pressure.Value / 100.0  # bar
        t_adj4 = adj4.TargetValue.Value  # C
        min_app = lng.MinApproach.Value  # C
        power = ss.Cell("C8").CellValue  # kW
        
        return {
            'flow': round(flow10, 1),
            'p1': round(p1, 1),
            't_adj4': round(t_adj4, 1),
            'min_app': round(min_app, 2),
            'p7': round(p7, 2),
            'power': round(power, 2),
            'note': 'User Set'
        }
    except Exception as e:
        print(f"Error reading 500 kg/h state: {e}")
        return None

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
    print("READING 500 kg/h STATE (User Configured)")
    print("="*80)
    
    state_500 = read_current_state_500(case)
    
    if state_500:
        print(f"\n500 kg/h Configuration:")
        print(f"  Flow: {state_500['flow']} kg/h")
        print(f"  P1: {state_500['p1']} bar")
        print(f"  T_ADJ4: {state_500['t_adj4']} C")
        print(f"  Min Approach: {state_500['min_app']} C")
        print(f"  P7: {state_500['p7']} bar")
        print(f"  Power: {state_500['power']} kW")
        print(f"  Note: {state_500['note']}")
        print("\nThis configuration will be saved to CSV.")
    else:
        print("\nFailed to read 500 kg/h state. Please check HYSYS.")
        return
    
    print("\nPress Enter to confirm and continue with 600-1500 kg/h optimization...")
    input()

if __name__ == "__main__":
    main()
