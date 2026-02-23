import os
import time
import win32com.client
import threading
import csv
from hysys_utils import dismiss_popup

# --- Configuration ---
SIMULATION_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER_NAME = "hysys_automation"
TARGET_APP = 2.25
APP_TOLERANCE = 0.25 # 2.0 to 2.5
MAX_STEP_T = 10.0 # Max temperature jump per step
MIN_STEP_T = 0.5
SAFE_T_MAX = -80.0
SAFE_T_MIN = -125.0

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

    def wait_stable(self, timeout=30, stability_duration=1.0):
        """Adaptive wait: returns immediately when values stabilize."""
        start_time = time.time()
        
        # 1. Wait for Solver Idle
        while time.time() - start_time < timeout:
            try:
                if not self.case.Solver.IsSolving:
                    break
            except: pass
            time.sleep(0.1)
            
        # 2. Monitor Stability
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
                time.sleep(0.1)
            except:
                return False
        
        return True # Timeout but maybe ok?

    def get_metrics(self):
        try:
            return {
                'app': self.lng.MinApproach.Value,
                'p7': self.s7.Pressure.Value / 100.0,
                'pwr': self.spreadsheet.Cell("C8").CellValue,
                't': self.adj4.TargetValue.Value
            }
        except:
            return None

    def save_state(self):
        try:
            self.last_stable_state = {
                't': self.adj4.TargetValue.Value,
                'flow': self.s10.MassFlow.Value,
                'p': self.s1.Pressure.Value
            }
        except:
            pass

    def restore_last_stable(self):
        if not self.last_stable_state:
            return False
        print("    [RESTORE] Rolling back to last stable state...")
        try:
            self.case.Solver.CanSolve = False # Pause solver to prevent fighting
            self.s10.MassFlow.Value = self.last_stable_state['flow']
            self.s1.Pressure.Value = self.last_stable_state['p']
            self.adj4.TargetValue.Value = self.last_stable_state['t']
            self.case.Solver.CanSolve = True
            self.wait_stable(20)
            return True
        except:
            return False

    def set_condition(self, t):
        try:
            self.save_state() # Save before moving
            self.adj4.TargetValue.Value = t
            is_stable = self.wait_stable(15)
            
            if not is_stable:
                print(f"    [WARN] Instability at T={t:.1f}, rolling back...")
                self.restore_last_stable()
                return False
            
            # Reset Adjusts (Important for HYSYS internal logic)
            for adj in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]:
                 self.case.Flowsheet.Operations.Item(adj).Reset()
            
            is_stable_2 = self.wait_stable(20)
            if not is_stable_2:
                 self.restore_last_stable()
                 return False
                 
            return True
        except:
            self.restore_last_stable()
            return False

class NewtonOptimizer:
    def __init__(self, controller):
        self.ctrl = controller
        
    def find_optimal_t(self, p_bar, start_t=-100.0):
        print(f"  [OPT] Newton Search P={p_bar} bar (Start T={start_t})...")
        
        # 1. Initialize
        self.ctrl.s1.Pressure.Value = p_bar * 100.0
        
        # Point A
        t1 = start_t
        if not self.ctrl.set_condition(t1): return None
        m1 = self.ctrl.get_metrics()
        if not m1: return None
        
        # Check if lucky
        if 2.0 <= m1['app'] <= 3.0:
            print(f"    HIT: T={t1:.1f} -> App={m1['app']:.2f}")
            return {'t': t1, **m1}

        # Point B (Sample gradient)
        # Determine direction: if App is small/negative, we need Higher T (Less cold) -> No, Wait.
        # Temp down -> Recovery Up? No.
        # LNG Heat Exchanger: 
        # Usually: Lower T (colder) -> Closer Approach (smaller App). 
        # So if App < 2.0 (too close), we need Higher T (warmer).
        # If App > 2.5 (too wide), we need Lower T (colder).
        
        # Let's verify direction with a small step
        step = -2.0 if m1['app'] > 2.25 else 2.0
        t2 = t1 + step
        
        if not self.ctrl.set_condition(t2): return None
        m2 = self.ctrl.get_metrics()
        if not m2: return None
        
        # Newton Loop
        for i in range(5): # Max 5 steps
            app1 = m1['app']
            app2 = m2['app']
            
            # Gradient k = d(App)/d(T)
            if abs(t2 - t1) < 1e-3: break
            k = (app2 - app1) / (t2 - t1)
            
            print(f"    Step {i}: T1={t1:.1f}({app1:.2f}) -> T2={t2:.1f}({app2:.2f}) | k={k:.3f}")
            
            # Avoid division by zero or crazy gradients
            if abs(k) < 0.01: 
                print("    [Gradient Flat] Taking default step")
                t_target = t2 - 2.0 # Default move colder
            else:
                # Predict T for Target App = 2.25
                # 2.25 = app2 + k * (t_target - t2)
                # t_target - t2 = (2.25 - app2) / k
                delta_t = (TARGET_APP - app2) / k
                
                # Clamp Delta T
                if delta_t > MAX_STEP_T: delta_t = MAX_STEP_T
                if delta_t < -MAX_STEP_T: delta_t = -MAX_STEP_T
                
                t_target = t2 + delta_t
            
            # Clamp Absolute T
            if t_target < SAFE_T_MIN: t_target = SAFE_T_MIN
            if t_target > SAFE_T_MAX: t_target = SAFE_T_MAX
            
            # Done if converged
            if 2.0 <= app2 <= 2.5 and m2['p7'] <= 36.5: # Allow slight P7 margin during search
                print(f"    CONVERGED: T={t2:.1f} -> App={app2:.2f}")
                return {'t': t2, **m2}
            
            # Move
            if abs(t_target - t2) < MIN_STEP_T:
                print("    [Step too small] Stopping")
                break
                
            prev_t2 = t2
            prev_m2 = m2
            
            if not self.ctrl.set_condition(t_target):
                print("    [Move Failed] Reducing step...")
                t_target = (t2 + t_target) / 2 # Half step
                if not self.ctrl.set_condition(t_target):
                    break
            
            # Update points for next iter
            t1, m1 = prev_t2, prev_m2
            t2 = t_target
            m2 = self.ctrl.get_metrics()
            if not m2: break
            
        return None

def main():
    print("="*60)
    print("HYSYS ADVANCED OPTIMIZATION (NEWTON-RAPHSON)")
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
    opt = NewtonOptimizer(ctrl)
    
    # Test Flow Range
    flows = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
    results = {}
    
    # Pre-determined pressure centers (from previous runs)
    p_centers = {
        500: 3.0, 600: 3.5, 700: 4.1, 800: 4.3, 900: 4.8, 
        1000: 5.2, 1100: 5.7, 1200: 6.1, 1300: 6.5, 1400: 6.9, 1500: 7.4
    }

    for flow in flows:
        print(f"\nOptimization Flow={flow} kg/h")
        
        # 1. Init Flow
        ctrl.s10.MassFlow.Value = flow / 3600.0
        
        # 2. Search Pressures (Center ± 0.2 bar)
        p_base = p_centers.get(flow, 4.0)
        pressures = [p_base]
        if flow != 500 and flow != 1500:
             pressures = [round(p_base + x, 1) for x in [0, -0.1, 0.1, -0.2, 0.2]] # Priority: Center -> Out
        
        best_res = None
        
        for p in pressures:
            # Predict Start Temp based on flow (Heuristic)
            # Low flow -> Colder (-110), High flow -> Warmer (-95)
            start_t = -110.0 + (flow - 500) * (15.0/1000.0) 
            
            res = opt.find_optimal_t(p, start_t)
            
            if res and res['pwr'] < 1e9:
                # Valid result
                if best_res is None or res['pwr'] < best_res['pwr']:
                    best_res = {'p': p, **res}
                    print(f"  [NEW BEST] P={p} T={res['t']:.1f} Pwr={res['pwr']:.1f}")
                    
        if best_res:
            results[flow] = best_res
            # Log to CSV immediately
            with open(os.path.join(FOLDER_NAME, "optimization_advanced.csv"), 'a', newline='') as f:
                w = csv.writer(f)
                r = best_res
                w.writerow([flow, r['p'], r['t'], r['app'], r['p7'], r['pwr']])
        else:
            print(f"  [FAIL] Could not optimize {flow}")

if __name__ == "__main__":
    main()
