import os, time, csv
import win32com.client

"""
SAFE HIGH FLOW OPTIMIZER (1400 & 1500 kg/h)
- Removed Popup Killer Thread (Source of Freezing)
- Simplified Solver Control (No CanSolve toggling)
- Robust Wait Logic
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"

class HighFlowOptimizerSafe:
    def __init__(self, app):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app.SimulationCases.Open(file_path)
        else:
            self.case = app.ActiveDocument
        
        fs = self.case.Flowsheet
        self.solver = self.case.Solver
        self.s1 = fs.MaterialStreams.Item("1")
        self.s10 = fs.MaterialStreams.Item("10")
        self.s6 = fs.MaterialStreams.Item("6")
        self.adj4 = fs.Operations.Item("ADJ-4")
        self.lng = fs.Operations.Item("LNG-100")
        self.comp = fs.Operations.Item("K-100")
        self.sprdsht = fs.Operations.Item("SPRDSHT-1")
        self.adjs = [fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

    def is_healthy(self):
        try:
            if self.s1.Temperature.Value < -200: return False
            if self.comp.EnergyValue < 0.1: return False
            return True
        except: return False

    def recover_state(self, flow, p_bar):
        print(f"   [RECOVERY] Hard Reset -> Flow={flow}, P={p_bar}...", flush=True)
        try:
            # Removed CanSolve toggle to avoid freezing
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0
            self.adj4.TargetValue.Value = -90.0
            
            for adj in self.adjs: 
                try: adj.Reset() 
                except: pass
            
            self.wait_stable(20, 1.0)
            
            # Double reset
            for adj in self.adjs: 
                try: adj.Reset() 
                except: pass
            
            self.wait_stable(10, 1.0)
            return self.is_healthy()
        except: return False

    def wait_stable(self, timeout=30, stable_time=1.0):
        time.sleep(0.5)
        start = time.time()
        
        # Wait for solver idle
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
        
        # Wait for temperature stability
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
            # Creep pressure if change > 1 bar
            curr_p = self.s1.Pressure.Value / 100.0
            if abs(curr_p - p_bar) > 1.0:
                step = 0.5 * (1 if p_bar > curr_p else -1)
                temp_p = curr_p
                while abs(temp_p - p_bar) > 0.5:
                    temp_p += step
                    self.s1.Pressure.Value = temp_p * 100.0
                    self.wait_stable(5, 0.5)
            
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.adj4.TargetValue.Value = t_deg
            
            if not self.wait_stable(25, 1.0): return False
            time.sleep(2.0) # Anti-lag buffer
            return self.is_healthy()
        except: return False

    def get_metrics(self):
        try:
            return {
                'MA': self.lng.MinApproach.Value,
                'S6_Pres': self.s6.Pressure.Value / 100.0,
                'Power': self.sprdsht.Cell("C8").CellValue
            }
        except: return None

    def optimize_point(self, flow, target_p):
        print(f"\n>> OPTIMIZING Flow={flow} kg/h near P={target_p} bar")
        
        # Baseline
        curr_flow = self.s10.MassFlow.Value * 3600.0
        if abs(curr_flow - flow) > 200:
             self.recover_state(flow, target_p)
        
        best = None
        
        for p in [target_p - 0.1, target_p, target_p + 0.1]:
            print(f"   Scanning P={p}: ", end="", flush=True)
            
            # Start warmer (-90) but closer to target than -80
            for t in range(-90, -116, -1):
                if not self.set_inputs_safe(flow, p, t):
                    print("!", end="") # Inputs failed
                    self.recover_state(flow, p)
                    continue
                
                m = self.get_metrics()
                if not m: continue
                
                # Loose MA check for high flow
                if m['MA'] < 0.01:
                    print(f".({m['MA']:.2f})", end="")
                    if m['MA'] < -0.5: break 
                
                if m['S6_Pres'] > 39.0: # Loose P constraint
                    print("P", end="") 
                    continue

                if m['MA'] > 5.0:
                    print("W", end="") 
                else:
                    print("*", end="") 
                    if best is None or m['Power'] < best['Power']:
                        best = {'Flow': flow, 'P_bar': p, 'T_C': t, **m}
            print("")
            
        if best:
            print(f"   BEST: P={best['P_bar']} bar, T={best['T_C']} C, Power={best['Power']:.2f} kW")
        else:
            print("   FAILED to find feasible point.")
            
        return best

    def run(self):
        print("SAFE OPTIMIZATION STARTED (Popup Killer Disabled)")
        results = []
        
        r1400 = self.optimize_point(1400, 6.7)
        if r1400: results.append(r1400)
        
        r1500 = self.optimize_point(1500, 7.4)
        if r1500: results.append(r1500)
        
        if results:
            keys = ['Flow', 'P_bar', 'T_C', 'MA', 'S6_Pres', 'Power']
            with open('hysys_automation/high_flow_safe_results.csv', 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                writer.writerows(results)
            print("\nSaved to hysys_automation/high_flow_safe_results.csv")

def main():
    try:
        app = win32com.client.GetActiveObject("HYSYS.Application")
    except:
        app = win32com.client.Dispatch("HYSYS.Application")
        
    opt = HighFlowOptimizerSafe(app)
    opt.run()

if __name__ == "__main__":
    main()
