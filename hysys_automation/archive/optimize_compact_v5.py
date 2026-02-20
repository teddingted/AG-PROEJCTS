import os, time, csv, threading
import win32com.client
from hysys_utils import dismiss_popup

# --- Config ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_compact_v5.csv"
FLOWS = [700, 800] 
PRESET_P = {700: 4.1, 800: 4.3}

class HysysEngine:
    def __init__(self, app_instance):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app_instance.SimulationCases.Open(file_path)
        else:
            self.case = app_instance.ActiveDocument
        self.case.Visible = True
        
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
        try:
            t = self.s1.Temperature.Value
            if t < -200: return False 
            if self.comp_k100.EnergyValue < 0.1: return False 
            return True
        except: return False

    def recover_state(self, flow, p_bar):
        try:
            self.solver.CanSolve = False
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0 
            self.adj4.TargetValue.Value = -90.0
            
            for adj in self.adjs: adj.Reset()
                
            self.solver.CanSolve = True
            self.wait_stable(20, 1.0)
            for adj in self.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            return self.is_healthy()
        except: return False

    def wait_stable(self, timeout=30, stable_time=1.0):
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
        try:
            if not self.is_healthy():
                if not self.recover_state(flow, p_bar): return False

            self.s10.MassFlow.Value = flow / 3600.0
            if abs(self.s1.Pressure.Value - p_bar*100) > 100:
                self.s1.Pressure.Value = p_bar * 100.0
                self.wait_stable(5, 0.5)
            else:
                self.s1.Pressure.Value = p_bar * 100.0
            
            self.adj4.TargetValue.Value = t_deg
            
            if not self.wait_stable(20, 1.0): return False
            time.sleep(1.5) # Extended buffer
            
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
    print(f"\nFlow {flow} kg/h [CompactV5]")
    
    eng.recover_state(flow, PRESET_P.get(flow, 4.0))
    center_p = PRESET_P.get(flow, 4.0)
    pressures = [round(center_p + i*0.1, 1) for i in range(-2, 3)]
    best = None
    
    for p in pressures:
        print(f"  P={p}: ", end="", flush=True)
        valid_range = None
        
        # 1. Scan (-90 to -124)
        for t in range(-90, -125, -2):
            if not eng.set_inputs_safe(flow, p, t):
                print("!", end="") 
                eng.recover_state(flow, p)
                continue

            m = eng.get_metrics()
            if not m: continue
            
            app = m['app']
            
            # Anti-Lag check
            if app < 0.5:
                time.sleep(2.0) # Wait and read again
                m = eng.get_metrics()
                app = m['app']
            
            if 0.5 <= app <= 3.5:
                print(f" HIT({t}C,App={app:.2f})", end="")
                valid_range = range(t+2, t-5, -1)
                break
            elif app < 0.5 or app > 100: 
                print(f".({app:.1f})", end="") # Debug Value
            elif app > 3.5:
                print("W", end="")

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
            print(f" OK ({local_best['pwr']:.1f})" if local_best else " (No Constr)")
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
    
    for f in FLOWS:
        res = optimize(eng, f)
        if res:
             with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as csvf:
                 csv.writer(csvf).writerow([f, res['p'], res['t'], res['app'], res['pwr']])
