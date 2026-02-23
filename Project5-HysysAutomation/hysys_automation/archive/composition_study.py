import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

def set_composition(case, stream_name, propene_frac):
    """
    Sets the composition of 1_propene.
    Assumes components are [..., Ethylene, Propene, ...].
    We need to find the indices for Propene and Ethylene.
    """
    try:
        stream = case.Flowsheet.MaterialStreams.Item(stream_name)
        comp_names = case.Flowsheet.FluidPackage.Components.Names
        
        # Create a zero array
        new_comps = [0.0] * len(comp_names)
        
        # Find indices
        idx_propene = -1
        idx_ethylene = -1
        
        for i, name in enumerate(comp_names):
            if "Propene" in name:
                idx_propene = i
            elif "Ethylene" in name:
                idx_ethylene = i
                
        if idx_propene == -1 or idx_ethylene == -1:
            print(f"Error: Could not find Propene or Ethylene components. Found: {comp_names}")
            return False
            
        # Set values
        # Ethylene gets the remainder
        ethylene_frac = 1.0 - propene_frac
        
        new_comps[idx_propene] = propene_frac
        new_comps[idx_ethylene] = ethylene_frac
        
        stream.ComponentMassFraction.Values = new_comps
        return True
    except Exception as e:
        print(f"Error setting composition: {e}")
        return False

def wait_for_solver(case, timeout=30):
    start = time.time()
    while True:
        if not case.Solver.IsSolving:
            return True
        if time.time() - start > timeout:
            return False
        time.sleep(0.5)

def get_target_temp(case):
    try:
        # F_ethylene
        s = case.Flowsheet.MaterialStreams.Item("F_ethylene")
        return s.Temperature.Value
    except:
        return -999.0

def run_control_loop(case, target_t=-31.0, tolerance=0.5):
    """
    Adjusts 1_propene MassFlow to hit target_t at F_ethylene.
    Simple Proportional-like logic or Bisection.
    Since real physics is involved, T decreases as Flow usually increases (cooling).
    Let's check the relationship.
    """
    s_propene = case.Flowsheet.MaterialStreams.Item("1_propene")
    
    max_iter = 20
    
    # Heuristic:
    # If T > Target (too hot) -> Need more cooling -> Increase Flow
    # If T < Target (too cold) -> Need less cooling -> Decrease Flow
    
    for i in range(max_iter):
        wait_for_solver(case)
        curr_t = get_target_temp(case)
        curr_flow = s_propene.MassFlow.Value * 3600 # kg/h
        
        delta = curr_t - target_t
        
        print(f"    Iter {i}: T={curr_t:.2f} C, Flow={curr_flow:.1f} kg/h, Delta={delta:.2f}")
        
        if abs(delta) < tolerance:
            return True, curr_t, curr_flow
            
        # Adjustment steps
        # Gain factor: 100 kg/h per 1 degree error? Let's try adaptive.
        # If delta is large (>5), use larger step.
        
        step = delta * 50.0 # 1C difference -> 50 kg/h change
        
        # Limit step size
        step = max(-500, min(500, step))
        
        new_flow = curr_flow + step
        if new_flow < 100: new_flow = 100 # Minimum flow
        
        s_propene.MassFlow.Value = new_flow / 3600.0
        
        # Allow solver to work
        time.sleep(1.0)
        wait_for_solver(case)
        
    return False, curr_t, curr_flow

def main():
    filename = "ETHANE_CHS_MIMIC_REV1_FULL_LOOP.hsc"
    folder = "hysys_automation"
    
    # Path handling
    if folder in os.getcwd():
        file_path = os.path.abspath(filename)
    else:
        file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))
        
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    
    # Popup handler
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    
    if os.path.exists(file_path):
        case = app.SimulationCases.Open(file_path)
        case.Visible = True
    else:
        case = app.ActiveDocument
        
    # Prepare CSV
    csv_file = "composition_study.csv"
    mode = 'w'
    if os.path.exists(csv_file): mode = 'w' # Overwrite for this run
    
    with open(csv_file, mode, newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Propene_Frac", "Ethylene_Frac", "Flow_1_propene_kgh", "Pressure_1_propene_kPa", "F_ethylene_T_C", "Power_kW", "Converged"])
        
        # Composition loop: 1.0 down to 0.5 in 0.05 steps
        propene_fracs = [1.0, 0.95, 0.90, 0.85, 0.80, 0.75, 0.70, 0.65, 0.60, 0.55, 0.50]
        
        for p_frac in propene_fracs:
            print(f"\n--- Testing Propene Fraction: {p_frac:.2f} ---")
            
            # 1. Set Composition
            if not set_composition(case, "1_propene", p_frac):
                print("Skipping due to error setting composition.")
                continue
                
            wait_for_solver(case)
            
            # 2. Control Loop to hit -31C
            converged, final_t, final_flow = run_control_loop(case, target_t=-31.0)
            
            # 3. Get other stats
            try:
                s_propene = case.Flowsheet.MaterialStreams.Item("1_propene")
                p_val = s_propene.Pressure.Value # kPa
                
                # Power? Assuming from spreadsheet similar to other scripts, or just checking 'Q-100' or similar if known.
                # checking param_opt.py: val = ss.Cell("C8").CellValue
                power = 0.0
                try:
                    ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
                    val = ss.Cell("C8").CellValue
                    if val: power = float(val)
                except:
                    pass
                    
            except:
                p_val = 0
                power = 0
                
            writer.writerow([p_frac, 1.0-p_frac, final_flow, p_val, final_t, power, converged])
            f.flush()
            
            if not converged:
                print("Warning: Did not fully converge to target temperature.")

    print("\nStudy Complete.")

if __name__ == "__main__":
    main()
