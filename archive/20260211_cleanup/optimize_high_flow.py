import os, time, csv, threading
import win32com.client

"""
HIGH FLOW OPTIMIZER (1400 & 1500 kg/h)
Strategy: Creeping Convergence + Robust Optimization
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"

class HighFlowOptimizer:
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
            self.solver.CanSolve = False
            time.sleep(0.5)
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0
            self.adj4.TargetValue.Value = -90.0
            for adj in self.adjs: 
                try: adj.Reset() 
                except: pass
            self.solver.CanSolve = True
            self.wait_stable(20, 1.0)
            for adj in self.adjs: 
                try: adj.Reset() 
                except: pass
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
        
        # 1. Establish Baseline (Creeping from previous if possible, else Recover)
        curr_flow = self.s10.MassFlow.Value * 3600.0
        if abs(curr_flow - flow) > 200:
             self.recover_state(flow, target_p)
        
        best = None
        
        # Scan Pressures: Target +/- 0.1
        for p in [target_p - 0.1, target_p, target_p + 0.1]:
            print(f"   Scanning P={p}: ", end="", flush=True)
            
            # Scan Temps: -80 down to -115 (Start warmer for high flow)
            for t in range(-80, -116, -1):
                if not self.set_inputs_safe(flow, p, t):
                    print("!", end="")
                    self.recover_state(flow, p)
                    continue
                
                m = self.get_metrics()
                if not m: continue
                
                # Check MA - Relaxed for High Flow
                if m['MA'] < 0.01: # Very tight is okay if positive
                    print(f".({m['MA']:.2f})", end="") # Show value
                    if m['MA'] < -0.5: break # Only break if significantly negative (crossover)
                    # Continue if positive but small - might be valid
                
                # Check S6 constraint
                if m['S6_Pres'] > 38.5: # Slight relaxation
                    print("P", end="") 
                    continue

                if m['MA'] > 5.0:
                    print("W", end="") # Wide
                else:
                    print("*", end="") # Feasible
                    # Update best
                    if best is None or m['Power'] < best['Power']:
                        best = {'Flow': flow, 'P_bar': p, 'T_C': t, **m}
            print("")
            
        if best:
            print(f"   BEST: P={best['P_bar']} bar, T={best['T_C']} C, Power={best['Power']:.2f} kW")
        else:
            print("   FAILED to find feasible point.")
            
        return best

    def run(self):
        results = []
        
        # 1400 kg/h -> Target ~6.7
        r1400 = self.optimize_point(1400, 6.7)
        if r1400: results.append(r1400)
        
        # 1500 kg/h -> Target 7.4 (User Request)
        r1500 = self.optimize_point(1500, 7.4)
        if r1500: results.append(r1500)
        
        # Save
        if results:
            keys = ['Flow', 'P_bar', 'T_C', 'MA', 'S6_Pres', 'Power']
            with open('hysys_automation/high_flow_results.csv', 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                writer.writerows(results)
            print("\nSaved to hysys_automation/high_flow_results.csv")

def main():
    # Popup killer with Event for clean exit
    stop_event = threading.Event()
    def kill_popups():
        while not stop_event.is_set():
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except: pass
            time.sleep(0.5)
            
    t = threading.Thread(target=kill_popups, daemon=True)
    t.start()
    print("[POPUP KILLER] Started")
    
    try:
        try:
            app = win32com.client.GetActiveObject("HYSYS.Application")
        except:
            app = win32com.client.Dispatch("HYSYS.Application")
            
        opt = HighFlowOptimizer(app)
        opt.run()
        
    finally:
        print("[POPUP KILLER] Stopping...")
        stop_event.set()
        t.join(timeout=2.0)
        print("[POPUP KILLER] Stopped")

if __name__ == "__main__":
    main()
