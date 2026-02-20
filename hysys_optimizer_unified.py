import os, time, csv, threading
import win32com.client
import win32gui, win32con

"""
UNIFIED HYSYS OPTIMIZER
- Merges: Legacy Stability (Hard Reset) + Node Efficiency + Smart Anchor
- Target: 500 - 1500 kg/h
- Safety: Threaded Popup Killer (Controlled), Anti-Freeze Timeouts
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_unified_result.csv"

# Configuration
FLOWS = range(500, 1501, 100)
ANCHOR_POINTS = {
    500: {'P': 3.00, 'T': -111.0},
    1500: {'P': 7.40, 'T': -99.0}
}
PRESET_P = {
    600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 6.81
}

class HysysNodeManager:
    """Lightweight Node Manager for Efficient Access"""
    def __init__(self, app):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app.SimulationCases.Open(file_path)
        else:
            self.case = app.ActiveDocument
        self.case.Visible = True
        
        self.fs = self.case.Flowsheet
        self.solver = self.case.Solver
        
        # Cache Critical Nodes
        self.nodes = {
            'S1': self.fs.MaterialStreams.Item("1"),
            'S10': self.fs.MaterialStreams.Item("10"),
            'S7': self.fs.MaterialStreams.Item("7"),
            'S6': self.fs.MaterialStreams.Item("6"),
            'Comp': self.fs.Operations.Item("K-100"),
            'LNG': self.fs.Operations.Item("LNG-100"),
            'Spreadsheet': self.fs.Operations.Item("SPRDSHT-1"),
            'ADJ4': self.fs.Operations.Item("ADJ-4"),
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class UnifiedOptimizer:
    def __init__(self, manager):
        self.mgr = manager
        self.solver = manager.solver

    def wait_stable(self, timeout=30, stable_time=1.0):
        """Robust Stability Check"""
        time.sleep(0.5)
        start = time.time()
        # 1. Wait for Solver Idle
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
            
        # 2. Monitor S1 Temp Stability
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
                time.sleep(0.2)
            except: return False
        return True

    def recover_state(self, flow, p_bar):
        """LEGACY HARD RESET STRATEGY (The Key to Stability)"""
        try:
            self.solver.CanSolve = False
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0 # Warm Start
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            
            for adj in self.mgr.adjs: adj.Reset()
            self.solver.CanSolve = True
            
            self.wait_stable(20, 1.0)
            
            # Double Reset
            for adj in self.mgr.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            
            return True
        except: return False

    def set_inputs(self, flow, p_bar, t_deg):
        """Safe Input Setting"""
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            
            # Pressure Ramping if needed (Legacy logic)
            curr_p = self.mgr.nodes['S1'].Pressure.Value / 100.0
            if abs(curr_p - p_bar) > 1.0:
                self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
                self.wait_stable(5, 0.5)
            else:
                self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
                
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            
            if not self.wait_stable(20, 1.0): return False
            time.sleep(1.0) # Settle buffer
            return True
        except: return False

    def get_metrics(self):
        try:
            t = self.mgr.nodes['S1'].Temperature.Value
            if t < -200: return None
            
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            # Anti-lag check
            if ma < 0.5:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
                
            return {
                'MA': ma,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S7_Pres': self.mgr.nodes['S7'].Pressure.Value / 100.0
            }
        except: return None

    # --- STRATEGY 1: GRID SCAN (For 500-1300 kg/h) ---
    def strategy_grid_scan(self, flow):
        center_p = PRESET_P.get(flow, 4.0)
        print(f"   [GRID SCAN] Center P={center_p}")
        
        self.recover_state(flow, center_p)
        
        best = None
        # Scan P: Center +/- 0.1
        for p in [round(center_p + i*0.1, 1) for i in range(-1, 2)]:
            print(f"   Scan P={p}: ", end="", flush=True)
            valid_range = []
            
            # Coarse Scan Temp
            for t in range(-90, -125, -2):
                if not self.set_inputs(flow, p, t):
                    print("!", end="")
                    self.recover_state(flow, p)
                    continue
                
                m = self.get_metrics()
                if not m: continue
                
                if 0.5 <= m['MA'] <= 3.5:
                    # Found good region
                    print("H", end="")
                    valid_range = range(t+2, t-5, -1)
                    break
                elif m['MA'] < 0.5: print(".", end="") # Tight
                else: print("W", end="") # Wide
            
            # Fine Scan
            if valid_range:
                print(" -> Fine", end="")
                local_best = None
                for t in valid_range:
                    if not self.set_inputs(flow, p, t): continue
                    m = self.get_metrics()
                    if m and 2.0 <= m['MA'] <= 3.0 and m['S7_Pres'] <= 37.0: # Strict Constraints
                        if local_best is None or m['Power'] < local_best['Power']:
                            local_best = {'P':p, 'T':t, **m}
                            best = local_best if best is None or local_best['Power'] < best['Power'] else best
                            print("*", end="")
                print(f" OK ({local_best['Power']:.1f})" if local_best else " (No match)")
            else:
                print(" -")
                
        return best

    # --- STRATEGY 2: SMART ANCHOR (Data-Driven) ---
    def strategy_anchor(self, flow):
        anchor = ANCHOR_POINTS.get(flow)
        if not anchor: return None
        
        print(f"   [ANCHOR] Target P={anchor['P']}, T={anchor['T']}")
        
        # 1. Try Soft Update first (Fast)
        if self.set_inputs(flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m and m['MA'] > 0.05: # Basic sanity check
                 print(f"   Verified (Soft): Power={m['Power']:.1f}, MA={m['MA']:.2f}")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        # 2. Fallback to Hard Reset (Robust)
        print("   [SOFT FAIL] Triggering Hard Reset...")
        self.recover_state(flow, anchor['P']) 
        
        # Verify Anchor after Reset
        if not self.set_inputs(flow, anchor['P'], anchor['T']):
            print("   [FAIL] Anchor Unreachable")
            return None
            
        m = self.get_metrics()
        if m:
            print(f"   Verified (Hard): Power={m['Power']:.1f}, MA={m['MA']:.2f}")
            return {'P': anchor['P'], 'T': anchor['T'], **m}
        return None

    def run(self):
        results = []
        
        # CSV Init
        keys = ['Flow', 'P_bar', 'T_C', 'MA', 'Power', 'S6_Pres', 'Status']
        if not os.path.exists('hysys_automation/' + OUT_FILE):
             with open('hysys_automation/' + OUT_FILE, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()
        
        for flow in FLOWS:
            print(f"\n>> PROCESSING FLOW: {flow} kg/h", flush=True)
            
            res = None
            # Data-Driven Strategy Dispatch
            if flow in ANCHOR_POINTS:
                res = self.strategy_anchor(flow)
            else:
                res = self.strategy_grid_scan(flow)
            
            if res:
                row = {'Flow': flow, 'P_bar': res['P'], 'T_C': res['T'], 
                       'MA': res['MA'], 'Power': res['Power'], 'S6_Pres': res.get('S6_Pres', 0),
                       'Status': 'Unified Optimized'}
                results.append(row)
                
                with open('hysys_automation/' + OUT_FILE, 'a', newline='') as f:
                    csv.DictWriter(f, fieldnames=keys).writerow(row)
                print(f"   Saved {flow} kg/h: {res['Power']:.2f} kW")
            else:
                print(f"   [FAIL] No optimal point found for {flow}")

def main():
    # POPUP KILLER (Non-Blocking Event)
    stop_event = threading.Event()
    def kill_popups():
        while not stop_event.is_set():
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except: pass
            time.sleep(1.0) # Non-aggressive sleep
            
    t = threading.Thread(target=kill_popups, daemon=True)
    t.start()
    print("[SYSTEM] Unified Optimizer Started (Popup Killer Active)")
    
    try:
        try: app = win32com.client.GetActiveObject("HYSYS.Application")
        except: app = win32com.client.Dispatch("HYSYS.Application")
        
        mgr = HysysNodeManager(app)
        opt = UnifiedOptimizer(mgr)
        opt.run()
        
    finally:
        stop_event.set()
        t.join(timeout=1.0)
        print("[SYSTEM] Stopped")

if __name__ == "__main__":
    main()
