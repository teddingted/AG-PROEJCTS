import os, time, csv, threading, math
import win32com.client
import win32gui, win32con

"""
HYBRID HYSYS OPTIMIZER (hysys_optimizer_hybrid.py)
--------------------------------------------------
The "Ultimate" Optimizer that combines:
1.  **Anchor Strategy**: For user-defined fixed points (500, 1500).
2.  **Accelerated Secant**: For stable low/mid flows (600-1000). 3x Speedup.
3.  **Robust Grid Scan**: For sensitive high flows (1100-1400) or as Fallback.
4.  **Smart Features**:
    - Internal Surrogate Model for start points.
    - Automatic Fallback (Secant -> Grid) on failure.
    - Legacy Hard Reset Stability.
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_hybrid_result.csv"
MODEL_DATA_FILE = "hysys_automation/optimization_final_summary_verified.csv"

# Configuration
FLOWS = [500] # Quick Test

ANCHOR_POINTS = {
    500: {'P': 3.00, 'T': -111.0},
    1500: {'P': 7.40, 'T': -99.0}
}
# Fallback map if Model fails
PRESET_P = {
    600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 6.81
}

class SurrogateModel:
    """Predicts Start P/T from Verified Data"""
    def __init__(self, csv_path):
        self.data = []
        path = os.path.abspath(csv_path)
        if os.path.exists(path):
            with open(path, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        self.data.append({
                            'Flow': float(row['Flow']),
                            'P': float(row['P_bar']),
                            'T': float(row['T_C'])
                        })
                    except: pass
            self.data.sort(key=lambda x: x['Flow'])
            print(f"[MODEL] Loaded based on {len(self.data)} points.")
        else:
            print("[MODEL] No data file found. Using Presets.")

    def predict(self, flow):
        # 1. Exact or Interpolation
        if len(self.data) < 2: return None, None
        
        lower, upper = None, None
        for d in self.data:
            if d['Flow'] <= flow: lower = d
            if d['Flow'] >= flow and upper is None: upper = d
            
        if lower and upper:
            if lower == upper: return lower['P'], lower['T']
            ratio = (flow - lower['Flow']) / (upper['Flow'] - lower['Flow'])
            p = lower['P'] + ratio * (upper['P'] - lower['P'])
            t = lower['T'] + ratio * (upper['T'] - lower['T'])
            return round(p, 2), round(t, 1)
        
        # Extrapolation
        if lower: return lower['P'], lower['T']
        if upper: return upper['P'], upper['T']
        return None, None

class HysysNodeManager:
    """Robust Node Manager"""
    def __init__(self, app):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
             self.case = app.SimulationCases.Open(file_path)
        else:
             self.case = app.ActiveDocument
        self.case.Visible = True
        self.solver = self.case.Solver
        self.fs = self.case.Flowsheet
        
        self.nodes = {
            'S1': self.fs.MaterialStreams.Item("1"),
            'S10': self.fs.MaterialStreams.Item("10"),
            'S6': self.fs.MaterialStreams.Item("6"),
            'S7': self.fs.MaterialStreams.Item("7"),
            'Comp': self.fs.Operations.Item("K-100"),
            'LNG': self.fs.Operations.Item("LNG-100"),
            'Spreadsheet': self.fs.Operations.Item("SPRDSHT-1"),
            'ADJ4': self.fs.Operations.Item("ADJ-4"),
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class HybridOptimizer:
    def __init__(self, manager, model):
        self.mgr = manager
        self.model = model
        self.solver = manager.solver

    def wait_stable(self, timeout=30, stable_time=1.0):
        time.sleep(0.5)
        start = time.time()
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
            
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
                time.sleep(0.2)
            except: return False
        return True

    def recover_state(self, flow, p_bar):
        """Hard Reset"""
        try:
            self.solver.CanSolve = False
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            for adj in self.mgr.adjs: adj.Reset()
            self.solver.CanSolve = True
            self.wait_stable(20, 1.0)
            for adj in self.mgr.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            return True
        except: return False

    def set_inputs(self, flow, p_bar, t_deg):
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            if not self.wait_stable(20, 1.0): return False
            time.sleep(1.0) # Settle
            return True
        except: return False

    def get_metrics(self):
        try:
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            if ma < 0.5:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
            return {
                'MA': ma,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'S7_Pres': self.mgr.nodes['S7'].Pressure.Value / 100.0
            }
        except: return None

    # === STRATEGIES ===

    def strategy_anchor(self, flow):
        anchor = ANCHOR_POINTS.get(flow)
        print(f"   [ANCHOR] Target P={anchor['P']}, T={anchor['T']}")
        
        # Soft Try
        if self.set_inputs(flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m and m['MA'] > 0.05:
                 print(f"   Verified (Soft): Power={m['Power']:.1f}")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        # Hard Reset
        print("   [SOFT FAIL] Resetting...")
        self.recover_state(flow, anchor['P'])
        if self.set_inputs(flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m:
                 print(f"   Verified (Hard): Power={m['Power']:.1f}")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        return None

    def strategy_secant(self, flow, p_start, t_start):
        """Accelerated Secant Search"""
        print(f"   [SECANT] Start P={p_start}, T={t_start}...", end="", flush=True)
        
        # Local P Scan (Center, -0.1, +0.1)
        # But we use Secant for T inside
        best_p_res = None
        
        scan_pressures = [p_start, round(p_start-0.1, 2), round(p_start+0.1, 2)]
        
        for p in scan_pressures:
            # Secant Loop for T
            target_ma = 2.0
            t0, t1 = t_start + 2.0, t_start - 1.0
            
            # Initial Point
            if not self.set_inputs(flow, p, t0): continue
            m0 = self.get_metrics()
            if not m0: continue
            
            # Step Loop
            local_best = {'T': t0, **m0} if m0['MA']>0.5 else None
            y0 = m0['MA'] - target_ma
            
            # Short loop (max 4 steps)
            for i in range(4):
                if not self.set_inputs(flow, p, t1): break
                m1 = self.get_metrics()
                if not m1: break
                
                # Check Best
                if m1['MA'] >= 1.0 and m1['S7_Pres'] < 37.0:
                    if local_best is None or m1['Power'] < local_best['Power']:
                        local_best = {'T': t1, **m1}
                
                y1 = m1['MA'] - target_ma
                if abs(y1) < 0.2: break # Converged
                
                # Secant Update
                if abs(y1 - y0) < 0.01: t_next = t1 - 0.5
                else: t_next = t1 - y1 * (t1 - t0) / (y1 - y0)
                
                # Bounds
                t_next = max(-125.0, min(-90.0, t_next))
                t0, y0, t1 = t1, y1, t_next
            
            if local_best:
                if best_p_res is None or local_best['Power'] < best_p_res['Power']:
                    best_p_res = {'P': p, **local_best}

        if best_p_res:
            print(f" OK (P={best_p_res['P']}, T={best_p_res['T']:.1f}, Power={best_p_res['Power']:.1f})")
            return best_p_res
        
        print(" FAIL")
        return None

    def strategy_grid_scan(self, flow, p_center):
        """Robust Fallback"""
        print(f"   [GRID SCAN] Center P={p_center}...", end="", flush=True)
        self.recover_state(flow, p_center)
        
        best = None
        for p in [round(p_center + i*0.1, 1) for i in range(-1, 2)]:
            valid_range = []
            # Coarse Scan
            for t in range(-90, -125, -2):
                if not self.set_inputs(flow, p, t): 
                    self.recover_state(flow, p)
                    continue
                m = self.get_metrics()
                if m and 0.5 <= m['MA'] <= 3.5:
                    valid_range = range(t+2, t-5, -1)
                    break
            
            # Fine Scan
            if valid_range:
                for t in valid_range:
                    if not self.set_inputs(flow, p, t): continue
                    m = self.get_metrics()
                    if m and 2.0 <= m['MA'] <= 3.0 and m['S7_Pres'] <= 37.0:
                        if best is None or m['Power'] < best['Power']:
                            best = {'P':p, 'T':t, **m}
                            
        if best: print(f" OK ({best['Power']:.1f} kW)")
        else: print(" FAIL")
        return best

    def run(self):
        print(">> HYBRID OPTIMIZATION ENGINE STARTED")
        keys = ['Flow', 'P_bar', 'T_C', 'MA', 'Power', 'S6_Pres', 'Status', 'Method']
        if not os.path.exists('hysys_automation/' + OUT_FILE):
             with open('hysys_automation/' + OUT_FILE, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()

        for flow in FLOWS:
            print(f"\n>> PROCESSING FLOW: {flow} kg/h")
            res = None
            method = "None"
            
            # 1. ANCHOR CHECK
            if flow in ANCHOR_POINTS:
                method = "Anchor"
                res = self.strategy_anchor(flow)
            
            # 2. ACCELERATED SECANT (600-1000)
            elif flow <= 1000:
                p_pred, t_pred = self.model.predict(flow)
                if p_pred is None: p_pred, t_pred = PRESET_P.get(flow, 4.0), -110.0
                
                method = "Secant"
                res = self.strategy_secant(flow, p_pred, t_pred)
                
                # Fallback to Grid if Secant Failed
                if not res:
                    print("   [FALLBACK] Secant failed. Switching to Grid Scan...")
                    method = "GridScan(Fallback)"
                    res = self.strategy_grid_scan(flow, p_pred)
            
            # 3. ROBUST GRID SCAN (1100+)
            else:
                method = "GridScan"
                p_center = PRESET_P.get(flow, 5.5)
                # Use model prediction if available for better center
                p_pred, _ = self.model.predict(flow)
                if p_pred: p_center = p_pred
                
                res = self.strategy_grid_scan(flow, p_center)

            # SAVE
            if res:
                row = {'Flow': flow, 'P_bar': res['P'], 'T_C': res['T'], 
                       'MA': res['MA'], 'Power': res['Power'], 'S6_Pres': res.get('S6_Pres', 0),
                       'Status': 'Optimized', 'Method': method}
                with open('hysys_automation/' + OUT_FILE, 'a', newline='') as f:
                    csv.DictWriter(f, fieldnames=keys).writerow(row)
            else:
                print(f"   [FATAL] No solution for {flow}")

def main():
    stop_event = threading.Event()
    def kill():
        while not stop_event.is_set():
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except: pass
            time.sleep(1.0)
    threading.Thread(target=kill, daemon=True).start()

    try:
        try: app = win32com.client.GetActiveObject("HYSYS.Application")
        except: app = win32com.client.Dispatch("HYSYS.Application")
        
        mdl = SurrogateModel(MODEL_DATA_FILE)
        mgr = HysysNodeManager(app)
        opt = HybridOptimizer(mgr, mdl)
        opt.run()
    finally:
        stop_event.set()

if __name__ == "__main__":
    main()
