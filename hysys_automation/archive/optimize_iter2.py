import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

# --- Configuration ---
SIMULATION_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER_NAME = "hysys_automation"
PREV_RESULT_FILE = "optimization_refined.csv"
NEW_RESULT_FILE = "optimization_iter2_highflow.csv"

# Flows (Targeted)
FLOWS = [1400, 1500]
PRESET_P_INIT = {
    1400: 7.2, 1500: 7.5
}

class HysysController:
    def __init__(self, app, case):
        self.app = app
        self.case = case
        self.s1 = case.Flowsheet.MaterialStreams.Item("1")
        self.s10 = case.Flowsheet.MaterialStreams.Item("10")
        self.s7 = case.Flowsheet.MaterialStreams.Item("7")
        self.adj4 = case.Flowsheet.Operations.Item("ADJ-4")
        self.lng = case.Flowsheet.Operations.Item("LNG-100")
        self.spreadsheet = case.Flowsheet.Operations.Item("SPRDSHT-1")

    def wait_solver_optimized(self, timeout=30, stability_duration=1.5):
        time.sleep(0.5) 
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if not self.case.Solver.IsSolving: break
            except: pass
            time.sleep(0.5)
            
        stable_start = None
        last_val = -9999
        while time.time() - start_time < timeout:
            try:
                curr_val = self.s1.Temperature.Value
                if curr_val < -200: return False
                if abs(curr_val - last_val) < 0.01:
                    if stable_start is None: stable_start = time.time()
                    elif time.time() - stable_start >= stability_duration: return True
                else:
                    stable_start = None
                last_val = curr_val
                time.sleep(0.5)
            except: return False
        return True

    def reset_adjusts(self):
        for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
            try: self.case.Flowsheet.Operations.Item(adj).Reset()
            except: pass

    def check_state(self):
        try:
            val = self.s1.Temperature.Value
            if val < -200: return 0 # Crash -> Treat as Cross to reset? No, return 0 is fine.
            app = self.lng.MinApproach.Value
            if abs(app) > 500: return 0
            
            if 0.5 <= app <= 3.5: return 2
            elif app > 0: return 1 # Wide
            else: return 0 # Cross/Tight (App < 0.5 or Negative optimization value)
        except: return 0

    def hard_reset(self, flow, p_bar):
        print("    [RESET] Unstable...")
        try:
            self.case.Solver.CanSolve = False
            self.reset_adjusts()
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0 
            self.adj4.TargetValue.Value = -90.0
            self.case.Solver.CanSolve = True
            self.wait_solver_optimized(20)
            self.reset_adjusts()
            self.wait_solver_optimized(20)
        except: pass

    def get_metrics(self):
        try:
            return {
                't': self.adj4.TargetValue.Value,
                'app': self.lng.MinApproach.Value,
                'p7': self.s7.Pressure.Value / 100.0,
                'pwr': self.spreadsheet.Cell("C8").CellValue
            }
        except: return None

def optimize_iter2(ctrl, flow):
    print(f"\n{'='*60}")
    print(f"Flow={flow} kg/h [DEEP_SCAN 1C]")
    
    # DEEP_SCAN Strategy
    center_p = PRESET_P_INIT.get(flow, 7.5)
    # Scan wider pressure range: Center ± 0.6
    pressures = [round(center_p + i*0.1, 1) for i in range(-6, 7)]
    scan_step = 1.0
    start_t = -90
    end_t = -130 # Extended for safety
    
    print(f"  Pressure Scope: {pressures}")
    print(f"  Temp Scope: {start_t} -> {end_t} (Step {scan_step})")
    
    best = None
    best_score = 1e20
    
    for p in pressures:
        p_kpa = p * 100.0
        
        # Init P
        ctrl.s10.MassFlow.Value = flow / 3600.0
        ctrl.s1.Pressure.Value = p_kpa
        ctrl.adj4.TargetValue.Value = start_t
        ctrl.wait_solver_optimized(10)
        ctrl.reset_adjusts()
        ctrl.wait_solver_optimized(10)
        
        print(f"  [SCAN] P={p}...", end="", flush=True)
        
        scan_temps = range(start_t, end_t - 1, int(-scan_step))
        valid_range = None
        
        for t in scan_temps:
            ctrl.adj4.TargetValue.Value = t
            ctrl.wait_solver_optimized(15)
            status = ctrl.check_state()
            
            if status == 2: # Hit
                print(f" HIT {t}")
                if scan_step == 1.0:
                    valid_range = range(t + 1, t - 2, -1) 
                else:
                    valid_range = range(t + 2, t - 4, -1)
                break
                
            elif status == 0: # Cross / Tight -> Continue Colder
                 if ctrl.s1.Temperature.Value < -200:
                    print(" CRASH", end="")
                    ctrl.hard_reset(flow, p)
                    break
                    
            elif status == 1: # Wide -> Overshot (Too Cold -> Too Wide)
                 print(f" WIDE {t} (Backtrack)", end="")
                 valid_range = range(t + 4, t - 1, -1)
                 break
        
        if not valid_range:
            print(" No valid range")
            continue
            
        print(f"    -> Opt {valid_range.start}~{valid_range.stop}...", end="")
        for t in valid_range:
            ctrl.adj4.TargetValue.Value = t
            ctrl.wait_solver_optimized(10)
            if ctrl.check_state() == 0: continue
            
            # Final settle
            ctrl.reset_adjusts()
            ctrl.wait_solver_optimized(20)
            
            m = ctrl.get_metrics()
            if m and 2.0 <= m['app'] <= 2.5 and m['p7'] <= 36.5:
                if m['pwr'] < best_score:
                    best_score = m['pwr']
                    best = {'p':p, **m}
                    
        if best and best['p'] == p: print(f" OK (Pwr={best['pwr']:.1f})")
        else: print(" -")

    return best

def main():
    print("="*60)
    print("HYSYS ITERATION 2 (High Flow Fix)")
    print("="*60)
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()
    case = app.ActiveDocument
    ctrl = HysysController(app, case)
    
    for flow in FLOWS:
        res = optimize_iter2(ctrl, flow)
        
        if res:
            with open(os.path.join(FOLDER_NAME, NEW_RESULT_FILE), 'a', newline='') as f:
                w = csv.writer(f)
                r = res
                w.writerow([flow, r['p'], r['t'], r['app'], r['p7'], r['pwr']])

if __name__ == "__main__":
    main()
