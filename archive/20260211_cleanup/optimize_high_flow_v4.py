import os, time, csv, threading
import win32com.client
import win32gui, win32con

"""
HIGH FLOW OPTIMIZER V4 (SMART ANCHOR)
- Logic: Start from User's Converged Point (1500kg/h, 7.4bar, -99C)
- Strategy: Verify Anchor -> Fine Tune Downwards -> Interpolate 1400
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"

class HysysNodeManager:
    def __init__(self, app):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app.SimulationCases.Open(file_path)
        else:
            self.case = app.ActiveDocument
        
        self.fs = self.case.Flowsheet
        self.solver = self.case.Solver
        
        # Cache Nodes
        self.nodes = {
            'S1': self.fs.MaterialStreams.Item("1"),
            'S10': self.fs.MaterialStreams.Item("10"),
            'S6': self.fs.MaterialStreams.Item("6"),
            'Comp': self.fs.Operations.Item("K-100"),
            'LNG': self.fs.Operations.Item("LNG-100"),
            'Spreadsheet': self.fs.Operations.Item("SPRDSHT-1"),
            'ADJ4': self.fs.Operations.Item("ADJ-4"),
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class SmartOptimizer:
    def __init__(self, manager):
        self.mgr = manager
        self.solver = manager.solver

    def wait_stable(self, timeout=30, stable_time=1.0):
        start = time.time()
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.1)
            
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.mgr.nodes['S1'].Temperature.Value
                if curr < -200: return False
                
                if abs(curr - last_val) < 0.01:
                    if ref_start is None: ref_start = time.time()
                    elif time.time() - ref_start >= stable_time: return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.1)
            except: return False
        return True

    def get_metrics(self):
        try:
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            # Anti-lag
            if ma < 0.1:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
                
            return {
                'MA': ma,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'Status': 'OK'
            }
        except: return None

    def set_inputs_smart(self, flow, p_bar, t_deg):
        # Direct set but with stability check
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            
            if not self.wait_stable(20, 1.0):
                # Kick
                print("K", end="", flush=True)
                self.solver.CanSolve = False
                for adj in self.mgr.adjs: adj.Reset()
                time.sleep(0.5)
                self.solver.CanSolve = True
                self.wait_stable(20, 1.0)
            
            return True
        except: return False

    def optimize_from_anchor(self, flow, start_p, start_t):
        print(f"\n>> OPTIMIZING Flow={flow} kg/h (Anchor: P={start_p}, T={start_t})")
        
        best = None
        
        # 1. Verify Anchor First
        if not self.set_inputs_smart(flow, start_p, start_t):
            print("   [FAIL] Anchor point unreachable")
            return None
            
        anchor_m = self.get_metrics()
        if anchor_m:
            print(f"   Anchor: P={start_p}, T={start_t} -> MA={anchor_m['MA']:.2f}, Power={anchor_m['Power']:.2f} kW")
            best = {'Flow': flow, 'P_bar': start_p, 'T_C': start_t, **anchor_m}
        else:
            print("   [FAIL] Anchor metrics read error")
            return None

        # 2. Try to improve (Lower T = ? usually lower power? or higher?)
        # Actually lower T target (-100) means COLDER, which usually means MORE Power.
        # But we want to Minimize Power.
        # So we should try HIGHER T (-98, -97...) until MA constraint hit?
        # Wait, usually lower T_target (S4) -> S1 needs to be adjusted.
        # Let's Scan T downwards from Anchor until Tight.
        
        print("   Fine tuning T: ", end="", flush=True)
        # Scan T down (Colder)
        for t in range(int(start_t)-1, -116, -1):
            if not self.set_inputs_smart(flow, start_p, t):
                print("!", end="")
                continue
            
            m = self.get_metrics()
            if not m: continue
            
            if m['MA'] < 0.1: # Tight
                print(f".({m['MA']:.2f})", end="")
                break
            
            print("*", end="")
            if m['Power'] < best['Power']:
                best = {'Flow': flow, 'P_bar': start_p, 'T_C': t, **m}
            # If Power increases, we might stop? No, search a bit more.
            
        print(f"\n   BEST: T={best['T_C']}, Power={best['Power']:.2f} kW")
        return best

def main():
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
    
    try:
        try: app = win32com.client.GetActiveObject("HYSYS.Application")
        except: app = win32com.client.Dispatch("HYSYS.Application")
        
        mgr = HysysNodeManager(app)
        opt = SmartOptimizer(mgr)
        
        results = []
        
        # 1. 1500 kg/h (Anchor: 7.4 bar, -99 C)
        # Verify current state first
        r1500 = opt.optimize_from_anchor(1500, 7.4, -99)
        if r1500: results.append(r1500)
        
        # 2. 1400 kg/h (Interpolated Anchor: 6.7 bar, -99.5 C)
        # 1300 was -100, 1500 is -99. So 1400 might be -99.5
        r1400 = opt.optimize_from_anchor(1400, 6.7, -99) 
        if r1400: results.append(r1400)
        
        if results:
            keys = ['Flow', 'P_bar', 'T_C', 'MA', 'S6_Pres', 'Power', 'Status']
            with open('hysys_automation/high_flow_v4_results.csv', 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                writer.writerows(results)
            print("\nSaved to hysys_automation/high_flow_v4_results.csv")
            
    finally:
        stop_event.set()
        t.join(timeout=1.0)

if __name__ == "__main__":
    main()
