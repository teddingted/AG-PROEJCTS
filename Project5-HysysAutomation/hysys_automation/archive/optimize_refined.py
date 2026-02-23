import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

# --- Configuration ---
SIMULATION_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER_NAME = "hysys_automation"

# Optimization Parameters
FLOWS = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
PRESET_P_INIT = {
    500: 3.0, 600: 3.5, 700: 4.1, 800: 4.3, 900: 4.8, 1000: 5.2,
    1100: 5.7, 1200: 6.1, 1300: 6.5, 1400: 6.9, 1500: 7.4
}

# Scan Settings
SCAN_STEP = 2.0  # Deg C
TARGET_APP_MIN = 0.5
TARGET_APP_MAX = 3.5
FINE_TUNE_STEP = 1.0

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
        self.last_stable_state = None

    def wait_solver_optimized(self, timeout=30, stability_duration=1.5):
        """
        Slower Robust Polling (Stability Priority):
        - Checks IsSolving every 0.5s
        - Checks Value Stability every 0.5s
        - Returns True as soon as stable for 'stability_duration'
        """
        time.sleep(0.5) # Initial buffer to let HYSYS react
        start_time = time.time()
        
        # 1. Wait for Solver Idle
        while time.time() - start_time < timeout:
            try:
                if not self.case.Solver.IsSolving:
                    break
            except: pass
            time.sleep(0.5)
            
        # 2. Monitor Stability (Slower Poll)
        stable_start = None
        last_val = -9999
        
        while time.time() - start_time < timeout:
            try:
                curr_val = self.s1.Temperature.Value
                
                # Check for physical nonsense (instability)
                if curr_val < -200:
                    return False # Failed/Unstable
                
                if abs(curr_val - last_val) < 0.01:
                    if stable_start is None:
                        stable_start = time.time()
                    elif time.time() - stable_start >= stability_duration:
                        return True # Stable
                else:
                    stable_start = None # Reset if changed
                
                last_val = curr_val
                time.sleep(0.5) # Slower poll
            except:
                return False
        
        return True # Timeout but maybe ok?

    def reset_adjusts(self):
        for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
            try:
                self.case.Flowsheet.Operations.Item(adj).Reset()
            except:
                pass

    def check_state(self):
        """
        Returns status code:
        2: Inside Target Range (0.5 - 3.5 C)
        1: Valid (> 0 C) but outside target
        0: Cross (< 0 C) or Abnormal
        """
        try:
            val = self.s1.Temperature.Value
            if val < -200: return 0
            
            app = self.lng.MinApproach.Value
            if abs(app) > 500: return 0
            
            if TARGET_APP_MIN <= app <= TARGET_APP_MAX:
                return 2
            elif app > 0:
                return 1
            else:
                return 0
        except:
            return 0

    def hard_reset(self, flow, p_bar):
        print("    [RESET] Simulation unstable, applying safe defaults...")
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
        except:
            print("    [ERROR] Reset failed")

    def get_metrics(self):
        try:
            return {
                't': self.adj4.TargetValue.Value,
                'app': self.lng.MinApproach.Value,
                'p7': self.s7.Pressure.Value / 100.0,
                'pwr': self.spreadsheet.Cell("C8").CellValue
            }
        except:
            return None

def optimize_flow_refined(ctrl, flow, p_center, best_prev=None):
    print(f"\n{'='*60}")
    print(f"Flow={flow} kg/h, P_center={p_center:.1f} bar")
    
    # Pressure strategy
    if flow == 500 or flow == 1500:
        pressures = [p_center]
    else:
        pressures = [round(p_center + i*0.1, 1) for i in range(-4, 5)] # ±0.4 bar
        
    best = None
    best_score = 1e20
    
    for p in pressures:
        p_kpa = p * 100.0
        
        # 1. Init / Clean Start for P
        ctrl.s10.MassFlow.Value = flow / 3600.0
        ctrl.s1.Pressure.Value = p_kpa
        ctrl.adj4.TargetValue.Value = -90.0 # Start safe
        ctrl.wait_solver_optimized(10)
        ctrl.reset_adjusts()
        ctrl.wait_solver_optimized(10)
        
        print(f"  [SCAN] P={p} (2C step)...", end="", flush=True)
        
        valid_range = None
        
        # 2. Scan loop (-90 down to -124 in 2C steps)
        scan_temps = range(-90, -125, -2) 
        
        prev_status = None
        
        for t in scan_temps:
            ctrl.adj4.TargetValue.Value = t
            ctrl.wait_solver_optimized(15)
            
            # Check range clamping (User request)
            if t > -90.0: continue # Should not happen with range() but safe
            if t < -120.0: break   # Strict limit
            
            status = ctrl.check_state()
            
            if status == 2: # Target Hit (0.5 ~ 3.5)
                print(f" HIT at {t}C")
                valid_range = range(t + 2, t - 3, -1) 
                break
                
            elif status == 0: # Cross / Tight
                # In this system, Colder T = Wider App.
                # So if we are Narrow/Cross, we need to go Colder (Continue).
                print(f" CROSS/TIGHT at {t}C... ", end="")
                if ctrl.s1.Temperature.Value < -200:
                   print(" CRASH!", end="")
                   ctrl.hard_reset(flow, p)
                   continue # Try next colder point? Or abort? Usually crash needs reset.
                
            elif status == 1: # Wide (> 3.5)
                # We went from Cross(0) or nothing -> Wide(1). 
                # We skipped the Target(2).
                print(f" WIDE at {t}C (Overshot)", end="")
                # Target is likely between prev (t+2) and curr (t).
                valid_range = range(t + 2, t - 1, -1)
                print(f" -> Backtracking {valid_range.start} to {valid_range.stop}")
                break
            
            prev_status = status
            
        if not valid_range:
            print(" No valid range")
            continue
            
        # 3. Optimize Fine-tune
        print(f"    -> Finetuning {valid_range.start} to {valid_range.stop}...", end="")
        
        local_best = None
        
        for t in valid_range:
            ctrl.adj4.TargetValue.Value = t
            ctrl.wait_solver_optimized(10)
            
            if ctrl.check_state() == 0: continue
            
            ctrl.reset_adjusts()
            ctrl.wait_solver_optimized(20) # Full solve
            
            m = ctrl.get_metrics()
            if m:
                # Target Verification: App 2.0 ~ 2.5
                if 2.0 <= m['app'] <= 2.5 and m['p7'] <= 36.5:
                    if m['pwr'] < best_score:
                        best_score = m['pwr']
                        best = {'p':p, **m}
                        local_best = m
                        # print(f" [NEW BEST {m['pwr']:.1f}]", end="")
                        
        if local_best:
            print(f" OK (Best: {local_best['pwr']:.1f} kW)")
        else:
            print(" No constr solution")
            
    if best:
        print(f"  RESULT: P={best['p']} T={best['t']:.1f} App={best['app']:.2f} Pwr={best['pwr']:.1f}")
        return best
    elif best_prev:
        print("  [FALLBACK] Using previous result")
        return best_prev
    return None

def main():
    print("="*60)
    print("HYSYS REFINED ROBUST OPTIMIZATION (2C STEP)")
    print("="*60)
    
    file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER_NAME, SIMULATION_FILE))
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

    ctrl = HysysController(app, case)
    results = {}
    
    for flow in FLOWS:
        best_prev = results.get(flow)
        res = optimize_flow_refined(ctrl, flow, PRESET_P_INIT[flow], best_prev)
        if res:
            results[flow] = res
            
            # Auto-save
            with open(os.path.join(FOLDER_NAME, "optimization_refined.csv"), 'a', newline='') as f:
                w = csv.writer(f)
                r = res
                w.writerow([flow, r['p'], r['t'], r['app'], r['p7'], r['pwr']])

if __name__ == "__main__":
    main()
