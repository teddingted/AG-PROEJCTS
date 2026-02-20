import os, time, csv, threading
import win32com.client
import win32gui, win32con

"""
HIGH FLOW OPTIMIZER V3 (THE FINAL INTEGRATION)
- HysysNodeManager (Lightweight Access)
- HysysOptimizer (Robust Logic: Kick, Anti-Lag)
- Popup Killer (Non-Blocking Event)
- Target: 1400/1500 kg/h
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"

class HysysNodeManager:
    """Lightweight Node Manager from hysys_node_manager.py"""
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

class RobustOptimizer:
    """Robust Logic from hysys_optimizer_final.py"""
    def __init__(self, manager):
        self.mgr = manager
        self.solver = manager.solver
    
    def kick_solver(self):
        """Standard Kick: Pause -> Reset Adjusts -> Resume"""
        print("   [KICK] Resetting Adjusts...", flush=True)
        try:
            self.solver.CanSolve = False
            for adj in self.mgr.adjs: adj.Reset()
            time.sleep(0.5)
            self.solver.CanSolve = True
            return True
        except: return False

    def recover_state(self, flow, p_bar):
        """Hard Reset Strategy"""
        print(f"   [RECOVERY] Hard Reset -> Flow={flow}, P={p_bar}...", flush=True)
        try:
            self.solver.CanSolve = False
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            
            for adj in self.mgr.adjs: adj.Reset()
            
            self.solver.CanSolve = True
            if not self.wait_stable(25, 1.0): return False
            
            # Double Kick for Logic Latches
            self.kick_solver()
            if not self.wait_stable(15, 1.0): return False
            
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
                if curr < -200: return False # Cryogenic Failure
                
                if abs(curr - last_val) < 0.01:
                    if ref_start is None: ref_start = time.time()
                    elif time.time() - ref_start >= stable_time: return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.1)
            except:
                return False
        return False

    def get_metrics(self):
        try:
            # Anti-Lag: Read twice if suspicious
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            if ma < 0.1:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
                
            return {
                'MA': ma,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue
            }
        except: return None

    def set_inputs_robust(self, flow, p_bar, t_deg):
        try:
            # 1. Check Pressure Delta - Creep if large
            curr_p = self.mgr.nodes['S1'].Pressure.Value / 100.0
            if abs(curr_p - p_bar) > 1.0:
                print(f"   [CREEP] P: {curr_p:.1f} -> {p_bar:.1f}", end="")
                step = 1.0 if p_bar > curr_p else -1.0
                temp_p = curr_p
                while abs(temp_p - p_bar) > 0.8:
                    temp_p += step
                    self.mgr.nodes['S1'].Pressure.Value = temp_p * 100.0
                    self.wait_stable(3, 0.5)
                    print(".", end="", flush=True)
                print(" Done")

            # 2. Set Values
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            
            # 3. Wait & Check
            if not self.wait_stable(30, 1.0): 
                # Kick if stuck
                print("K", end="", flush=True)
                self.kick_solver()
                if not self.wait_stable(20, 1.0): return False
            
            return True
        except: return False

    def optimize_point(self, flow, target_p):
        print(f"\n>> OPTIMIZING Flow={flow} kg/h @ {target_p} bar")
        
        # Recovery Check
        curr_flow = self.mgr.nodes['S10'].MassFlow.Value * 3600.0
        if abs(curr_flow - flow) > 200:
             if not self.recover_state(flow, target_p):
                 print("   [ERROR] Initial Recovery Failed")
                 return None

        best = None
        
        # Scan Pressures (Target +/- 0.1)
        for p in [target_p]: # Single Point First for Speed
            print(f"   Scan P={p}: ", end="", flush=True)
            
            # Scan Temps (-90 to -115)
            for t in range(-90, -116, -2): # Faster Step 2 deg
                if not self.set_inputs_robust(flow, p, t):
                    print("!", end="")
                    self.recover_state(flow, p)
                    continue
                
                m = self.get_metrics()
                if not m: continue
                
                # Check Constraints
                if m['MA'] < 0.01:
                    print(f".({m['MA']:.2f})", end="")
                    if m['MA'] < -0.5: break # Crossover
                
                if m['S6_Pres'] > 39.0:
                    print("P", end="")
                elif m['MA'] > 5.0:
                    print("W", end="")
                else:
                    print("*", end="")
                    if best is None or m['Power'] < best['Power']:
                        best = {'Flow': flow, 'P_bar': p, 'T_C': t, **m}
            print("")
            
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
    print("[SYSTEM] Popup Killer Active")
    
    try:
        try: app = win32com.client.GetActiveObject("HYSYS.Application")
        except: app = win32com.client.Dispatch("HYSYS.Application")
        
        mgr = HysysNodeManager(app)
        opt = RobustOptimizer(mgr)
        
        results = []
        
        # 1400 / 6.7
        r1 = opt.optimize_point(1400, 6.7)
        if r1: results.append(r1)
        else: print("   [FAIL] 1400 kg/h failed")
        
        # 1500 / 7.4
        r2 = opt.optimize_point(1500, 7.4)
        if r2: results.append(r2)
        else: print("   [FAIL] 1500 kg/h failed")
        
        if results:
            keys = ['Flow', 'P_bar', 'T_C', 'MA', 'S6_Pres', 'Power']
            with open('hysys_automation/high_flow_v3_results.csv', 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=keys)
                writer.writeheader()
                writer.writerows(results)
            print("\nSaved to hysys_automation/high_flow_v3_results.csv")
            
            # Print Summary
            print("\n" + "="*60)
            for r in results:
                print(f"{r['Flow']}: {r['Power']:.2f} kW (P={r['P_bar']}, T={r['T_C']}, MA={r['MA']:.2f})")
            print("="*60)
            
    finally:
        stop_event.set()
        t.join(timeout=1.0)
        print("[SYSTEM] Popup Killer Stopped")

if __name__ == "__main__":
    main()
