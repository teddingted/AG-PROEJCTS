import os, time, csv, threading
import win32com.client
from hysys_utils import dismiss_popup

# --- Config ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_compact_v2.csv"
FLOWS = [600, 700, 800] 
PRESET_P = {600: 3.5, 700: 4.1, 800: 4.3}

class HysysEngine:
    def __init__(self, app_instance):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app_instance.SimulationCases.Open(file_path)
        else:
            self.case = app_instance.ActiveDocument
        self.case.Visible = True
        
        # Cache Critical Objects
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
        """Check if simulation is alive (Not diverged/crashed)"""
        try:
            t = self.s1.Temperature.Value
            if t < -200: return False # Cryogenic crash
            if self.comp_k100.EnergyValue < 0.1: return False # Unit Op dead
            return True
        except: return False

    def recover_state(self, flow, p_bar):
        """Emergency Reset Routine"""
        try:
            self.solver.CanSolve = False
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0 # Safe warm start
            self.adj4.TargetValue.Value = -90.0
            
            for adj in self.adjs: adj.Reset()
                
            self.solver.CanSolve = True
            self.wait_stable(10)
            
            # Reset Again to be sure
            for adj in self.adjs: adj.Reset()
            self.wait_stable(10)
            
            return self.is_healthy()
        except: return False

    def wait_stable(self, timeout=30, stable_time=1.0):
        time.sleep(0.5)
        start = time.time()
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.5)
            
        # Value Stability Check
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.s1.Temperature.Value
                if curr < -200: return False # Fail early
                
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
        """Sequential Safe Update with Health Check"""
        try:
            # 1. Check Health & Recover if needed BEFORE starting
            if not self.is_healthy():
                if not self.recover_state(flow, p_bar): return False

            # 2. Set Values (Sequentially to avoid shock)
            self.s10.MassFlow.Value = flow / 3600.0
            if abs(self.s1.Pressure.Value - p_bar*100) > 10:
                self.s1.Pressure.Value = p_bar * 100.0
                self.wait_stable(10) # Wait if pressure changed significantly
            
            self.adj4.TargetValue.Value = t_deg
            
            # 3. Wait & Verify
            if not self.wait_stable(20): return False
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

def optimize(eng, flow):
    print(f"\nFlow {flow} kg/h [CompactV2]")
    
    # Ensure fresh start for new flow
    eng.recover_state(flow, PRESET_P.get(flow, 4.0))
    
    center_p = PRESET_P.get(flow, 4.0)
    pressures = [round(center_p + i*0.1, 1) for i in range(-2, 3)]
    
    best = None
    
    for p in pressures:
        print(f"  P={p} bar: ", end="", flush=True)
        valid_range = None
        
        # 1. Scan (-90 to -124)
        for t in range(-90, -125, -2):
            if not eng.set_inputs_safe(flow, p, t):
                print("!", end="") # Unstable/Crash
                if not eng.recover_state(flow, p): # Force recover
                    print("X", end="")
                    break 
                continue

            m = eng.get_metrics()
            if not m: continue
            
            if 0.5 <= m['app'] <= 3.5:
                print(f" HIT({t}C)", end="")
                valid_range = range(t+2, t-5, -1) # Fine tune range
                break
            elif m['app'] < 0.5 or m['app'] > 100: # Cross/Tight
                print(".", end="")
            elif m['app'] > 3.5: # Wide
                print("W", end="") # Just logging, keep scanning colder

        # 2. Fine Tune
        if valid_range:
            print(f" -> Fine", end="")
            local_best = None
            for t in valid_range:
                if not eng.set_inputs_safe(flow, p, t): 
                   print("!", end="")
                   continue
                
                m = eng.get_metrics()
                if m and 2.0 <= m['app'] <= 2.5 and m['p7'] <= 36.5:
                    if local_best is None or m['pwr'] < local_best['pwr']:
                        local_best = {'p':p, 't':t, **m}
                        best = local_best if best is None or local_best['pwr'] < best['pwr'] else best
                        print("*", end="")
            print(f" OK" if local_best else " (No Constr)")
        else:
            print(" -")

    if best:
        print(f"  BEST: {best['p']}bar, {best['t']}C, {best['pwr']:.1f}kW")
        return best
    return None

if __name__ == "__main__":
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True; t.start()
    
    app = win32com.client.Dispatch("HYSYS.Application")
    eng = HysysEngine(app)
    
    # Run Test
    for f in FLOWS:
        res = optimize(eng, f)
        if res:
             with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as csvf:
                 csv.writer(csvf).writerow([f, res['p'], res['t'], res['app'], res['pwr']])
