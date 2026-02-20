import os, time, csv, threading
import win32com.client
from hysys_utils import dismiss_popup

"""
HYSYS ROBUST OPTIMIZER (FINAL)
------------------------------
Features:
1. Compact Architecture: Atomic-like updates with sequential safety fallback.
2. Anti-Lag Logic: Detects and waits out transient 'Tight' states (critical for 700 kg/h).
3. Auto-Recovery: Automatically resets simulation if divergence (< -200C) is detected.
4. Robust Stability: explicit settle buffers to ensure Adjust/Spreadsheet convergence.

Target Range: 500 - 1500 kg/h
"""

# --- Configuration ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_final_result.csv"

FLOWS = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]

# Optimally calibrated starting pressures
PRESET_P = {
    500: 3.0, 600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 7.2, 1500: 7.5 
}

class HysysOptimizer:
    def __init__(self, app_instance):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app_instance.SimulationCases.Open(file_path)
        else:
            self.case = app_instance.ActiveDocument
        self.case.Visible = True
        
        # Cache HYSYS Objects
        fs = self.case.Flowsheet
        self.solver = self.case.Solver
        self.s1 = fs.MaterialStreams.Item("1")
        self.s10 = fs.MaterialStreams.Item("10")
        self.s7 = fs.MaterialStreams.Item("7")
        self.adj4 = fs.Operations.Item("ADJ-4")
        self.lng = fs.Operations.Item("LNG-100")
        self.comp_k100 = fs.Operations.Item("K-100")
        self.cell_pwr = fs.Operations.Item("SPRDSHT-1").Cell("C8")
        self.adjs = [fs.Operations.Item(n) for n in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]]

    def is_healthy(self):
        """Returns True if simulation is converged and valid."""
        try:
            t = self.s1.Temperature.Value
            if t < -200: return False # Cryogenic Failure
            if self.comp_k100.EnergyValue < 0.1: return False # Compressor Trip
            return True
        except: return False

    def recover_state(self, flow, p_bar):
        """Hard Reset to restore simulation from a crashed state."""
        try:
            self.solver.CanSolve = False
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0 # Safe warm start
            self.adj4.TargetValue.Value = -90.0
            
            for adj in self.adjs: adj.Reset()
                
            self.solver.CanSolve = True
            # Extended wait for recovery
            self.wait_stable(20, 1.0)
            
            # Second Reset to clear logical latches
            for adj in self.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            return self.is_healthy()
        except: return False

    def wait_stable(self, timeout=30, stable_time=1.0):
        """Waits for solver idle and S1 temperature stability."""
        time.sleep(0.5)
        start = time.time()
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
            
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.s1.Temperature.Value
                if curr < -200: return False
                
                if abs(curr - last_val) < 0.01:
                    if ref_start is None: ref_start = time.time()
                    elif time.time() - ref_start >= stable_time: return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.2)
            except: return False
        return True

    def set_inputs_safe(self, flow, p_bar, t_deg):
        """Sets inputs sequentially with safety checks."""
        try:
            if not self.is_healthy():
                if not self.recover_state(flow, p_bar): return False

            self.s10.MassFlow.Value = flow / 3600.0
            
            # Gradual P change if difference is large (> 1 bar)
            if abs(self.s1.Pressure.Value - p_bar*100) > 100:
                self.s1.Pressure.Value = p_bar * 100.0
                self.wait_stable(5, 0.5)
            else:
                self.s1.Pressure.Value = p_bar * 100.0
            
            self.adj4.TargetValue.Value = t_deg
            
            if not self.wait_stable(20, 1.0): return False
            
            # CRITICAL: Settle Buffer for Laggy Adjusts (700 kg/h fix)
            time.sleep(1.5) 
            
            if not self.is_healthy(): return False
            return True
        except: return False

    def get_metrics(self):
        try:
            return {
                'app': self.lng.MinApproach.Value,
                'p7': self.s7.Pressure.Value / 100.0,
                'pwr': self.cell_pwr.CellValue
            }
        except: return None

def optimize_flow(optimizer, flow):
    print(f"\nFlow {flow} kg/h")
    
    # 1. Initialize
    center_p = PRESET_P.get(flow, 4.0)
    optimizer.recover_state(flow, center_p)
    
    # Scan Range: Center +/- 0.2 bar
    pressures = [round(center_p + i*0.1, 1) for i in range(-2, 3)]
    best = None
    
    for p in pressures:
        print(f"  P={p}: ", end="", flush=True)
        valid_range = None
        
        # 2. Temperature Scan (-90 downto -124)
        for t in range(-90, -125, -2):
            if not optimizer.set_inputs_safe(flow, p, t):
                print("!", end="") 
                optimizer.recover_state(flow, p)
                continue

            m = optimizer.get_metrics()
            if not m: continue
            
            app = m['app']
            
            # **Anti-Lag Logic**: Double check 'Tight' values
            if app < 0.5:
                time.sleep(1.5)
                m = optimizer.get_metrics()
                app = m['app']
            
            if 0.5 <= app <= 3.5:
                print(f" HIT({t})", end="")
                valid_range = range(t+2, t-5, -1)
                break
            elif app < 0.5 or app > 100: 
                print(".", end="") # Tight
            elif app > 3.5:
                print("W", end="") # Wide

        # 3. Fine Tuning
        if valid_range:
            print(f" -> Fine", end="")
            local_best = None
            for t in valid_range:
                if not optimizer.set_inputs_safe(flow, p, t): 
                   print("!", end="")
                   continue
                
                m = optimizer.get_metrics()
                # Strict constraints for final selection
                if m and 2.0 <= m['app'] <= 3.0 and m['p7'] <= 36.5:
                    if local_best is None or m['pwr'] < local_best['pwr']:
                        local_best = {'p':p, 't':t, **m}
                        best = local_best if best is None or local_best['pwr'] < best['pwr'] else best
                        print("*", end="")
            print(f" OK ({local_best['pwr']:.1f})" if local_best else " (No Cstr)")
        else:
            print(" -")

    if best:
        print(f"  BEST: {best['p']}bar, {best['t']}C, {best['pwr']:.1f}kW")
        return best
    return None

def main():
    print("="*60)
    print("HYSYS OPTIMIZER FINAL RUN")
    print("="*60)
    
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True; t.start()
    
    app = win32com.client.Dispatch("HYSYS.Application")
    optimizer = HysysOptimizer(app)
    
    # Init Result CSV
    if not os.path.exists(os.path.join(FOLDER, OUT_FILE)):
        with open(os.path.join(FOLDER, OUT_FILE), 'w', newline='') as csvf:
            csv.writer(csvf).writerow(["MassFlow", "P_bar", "T_C", "MinApp_C", "Power_kW"])

    for f in FLOWS:
        res = optimize_flow(optimizer, f)
        if res:
             with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as csvf:
                 csv.writer(csvf).writerow([f, res['p'], res['t'], res['app'], res['pwr']])

if __name__ == "__main__":
    main()
