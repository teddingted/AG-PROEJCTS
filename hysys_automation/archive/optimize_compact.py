import os, time, csv, threading
import win32com.client
from hysys_utils import dismiss_popup

# --- Config ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_compact.csv"
FLOWS = [600, 700, 800] # Test Range
PRESET_P = {600: 3.5, 700: 4.1, 800: 4.3} # Targets from previous run

class HysysEngine:
    def __init__(self, app_instance):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app_instance.SimulationCases.Open(file_path)
        else:
            self.case = app_instance.ActiveDocument
        self.case.Visible = True
        
        # Cache Objects (Faster Access)
        fs = self.case.Flowsheet
        self.solver = self.case.Solver
        self.s1 = fs.MaterialStreams.Item("1")
        self.s10 = fs.MaterialStreams.Item("10")
        self.s7 = fs.MaterialStreams.Item("7")
        self.adj4 = fs.Operations.Item("ADJ-4")
        self.lng = fs.Operations.Item("LNG-100")
        self.cell_pwr = fs.Operations.Item("SPRDSHT-1").Cell("C8")
        self.ops_to_reset = [fs.Operations.Item(n) for n in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]]

    def set_inputs(self, flow, p_bar, t_deg):
        """Atomic Update: Disable Solver -> Set All -> Enable Solver"""
        try:
            self.solver.CanSolve = False
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.adj4.TargetValue.Value = t_deg
            
            # Reset Adjusts while frozen
            for op in self.ops_to_reset: op.Reset()
                
            self.solver.CanSolve = True
            return self._wait_stable()
        except:
            self.solver.CanSolve = True # Safety
            return False

    def _wait_stable(self, timeout=30, stable_time=1.0):
        time.sleep(0.5) # Allow solver to kick in
        start = time.time()
        
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
            
        # Stability Check
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.s1.Temperature.Value
                if curr < -200: return False # Crash
                
                if abs(curr - last_val) < 0.01:
                    if ref_start is None: ref_start = time.time()
                    elif time.time() - ref_start >= stable_time: return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.2)
            except: return False
        return True # Time out but maybe ok

    def get_state(self):
        try:
            return {
                'app': self.lng.MinApproach.Value,
                'p7': self.s7.Pressure.Value / 100.0,
                'pwr': self.cell_pwr.CellValue
            }
        except: return None

def optimize(eng, flow):
    print(f"\nFlow {flow} kg/h [Compact]")
    center_p = PRESET_P.get(flow, 4.0)
    pressures = [round(center_p + i*0.1, 1) for i in range(-2, 3)] # Local Scan
    best = None
    
    for p in pressures:
        print(f"  P={p} bar: ", end="", flush=True)
        valid_range = None
        
        # 1. Fast Scan (2C steps)
        for t in range(-90, -125, -2):
            if not eng.set_inputs(flow, p, t): 
                print("!", end="") # Unstable
                continue
                
            st = eng.get_state()
            if not st: continue
            
            is_cross = st['app'] < 0.5 or st['app'] > 100
            is_target = 0.5 <= st['app'] <= 3.5
            is_wide = st['app'] > 3.5
            
            if is_target:
                print(f" HIT({t}C)", end="")
                valid_range = range(t+2, t-4, -1)
                break
            elif is_cross:
                print(".", end="") # Tight/Cross -> Continue Colder
            elif is_wide:
                # Wide means we overshoot coldness? Or start too cold?
                # Scan is Warmer -> Colder.
                # If we see Wide... and we haven't seen Target?
                print("W", end="")
        
        if valid_range:
            print(f" -> FineTune", end="")
            for t in valid_range:
                eng.set_inputs(flow, p, t)
                st = eng.get_state()
                if st and 2.0 <= st['app'] <= 2.5 and st['p7'] <= 36.5:
                    if best is None or st['pwr'] < best['pwr']:
                        best = {'p':p, 't':t, **st}
                        print("*", end="")
        print(" OK" if valid_range else " -")

    if best:
        print(f"  BEST: {best['p']}bar, {best['t']}C, {best['pwr']:.1f}kW")
        return best
    return None

if __name__ == "__main__":
    t = threading.Thread(target=dismiss_popup)
    t.daemon = True; t.start()
    
    app = win32com.client.Dispatch("HYSYS.Application")
    eng = HysysEngine(app)
    
    for f in FLOWS:
        res = optimize(eng, f)
        if res:
             with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as csvf:
                 csv.writer(csvf).writerow([f, res['p'], res['t'], res['app'], res['pwr']])
