import os, time, csv, threading
import win32com.client
import win32gui, win32con

"""
REVERSE ROBUSTNESS VERIFIER (1300 -> 500 kg/h)
- Purpose: Check Hysteresis & Stability by approaching points from High Flow.
- Strategy: Load verified points, verify in descending order.
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
CSV_FILE = "hysys_automation/optimization_final_summary_verified.csv"

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

class ReverseVerifier:
    def __init__(self, manager):
        self.mgr = manager
        self.solver = manager.solver

    def kick_solver(self):
        try:
            self.solver.CanSolve = False
            for adj in self.mgr.adjs: adj.Reset()
            time.sleep(0.5)
            self.solver.CanSolve = True
            return True
        except: return False

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

    def set_inputs(self, flow, p_bar, t_deg):
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            
            # Robust Wait
            if not self.wait_stable(20, 1.0):
                print("K", end="", flush=True) # Kick
                self.kick_solver()
                self.wait_stable(15, 1.0)
            
            return True
        except: return False

    def get_metrics(self):
        try:
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            if ma < 0.1:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
            return {
                'MA': ma,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0
            }
        except: return None

    def run(self):
        # 1. Load Data
        data = []
        if os.path.exists(CSV_FILE):
            with open(CSV_FILE, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    data.append(row)
        else:
            print("CSV not found.")
            return

        # 2. Filter 1300 -> 500
        targets = [d for d in data if 500 <= float(d['Flow']) <= 1300]
        targets.sort(key=lambda x: float(x['Flow']), reverse=True) # DESCENDING

        print(f"Starting Reverse Verification ({len(targets)} points)...")
        print("Flow | Target P | Target T | Power | MA | Status")
        
        results = []
        for t in targets:
            flow = float(t['Flow'])
            p = float(t['P_bar'])
            temp = float(t['T_C'])
            
            print(f"{flow:.0f} | {p:.1f} | {temp:.1f} | ", end="", flush=True)
            
            if self.set_inputs(flow, p, temp):
                m = self.get_metrics()
                if m:
                    status = 'OK'
                    if m['MA'] < 0.1: status = 'Tight'
                    # Check deviation from original
                    p_diff = abs(m['Power'] - float(t['Power']))
                    if p_diff > 50: status = 'Deviated'
                    
                    print(f"{m['Power']:.2f} | {m['MA']:.2f} | {status}")
                    results.append({'Flow': flow, 'Status': status, 'Power_Rev': m['Power'], 'MA_Rev': m['MA']})
                else:
                    print("RedErr | - | Fail")
            else:
                print("SetErr | - | Fail")

        # Save Reverse Report
        with open('hysys_automation/reverse_verification_report.csv', 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['Flow', 'Status', 'Power_Rev', 'MA_Rev'])
            writer.writeheader()
            writer.writerows(results)
        print("\nReverse verification completed.")

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
        verifier = ReverseVerifier(mgr)
        verifier.run()
            
    finally:
        stop_event.set()
        t.join(timeout=1.0)

if __name__ == "__main__":
    main()
