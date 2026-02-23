import os, time, csv, threading, math
import win32com.client
import win32gui, win32con

"""
MULTIDIMENSIONAL HYBRID OPTIMIZER (hysys_optimizer_multidim.py)
---------------------------------------------------------------
OPTIMIZATION TARGETS:
1. Mass Flow (S10): 500 - 1500 kg/h
2. Volume Flow (ADJ-1 Target): 3500 - 3800 m3/h (Step 50)
3. Pressure (S1): Optimized
4. Temperature (ADJ-4 Target): Optimized

ARCHITECTURE:
- Hybrid Strategy (Anchor / Secant / Grid)
- Multi-Loop Execution
- Robust State Recovery (Manages ADJ-1 Target explicitly with Unit Conversion)
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_multidim_result.csv"
MODEL_DATA_FILE = "hysys_automation/optimization_final_summary_verified.csv"

# Configuration
FLOWS = range(500, 701, 100)
VOL_FLOWS = range(3500, 3801, 50)

# Anchor Points (Calibrated for Vol=3669)
ANCHOR_POINTS = {
    500: {'P': 3.00, 'T': -111.0},
    1500: {'P': 7.40, 'T': -99.0}
}
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
        if lower: return lower['P'], lower['T']
        if upper: return upper['P'], upper['T']
        return None, None

class HysysNodeManager:
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
            'ADJ1': self.fs.Operations.Item("ADJ-1"), # Added for Multidim
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class MultidimOptimizer:
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

    def recover_state(self, flow, vol_flow, p_bar):
        """Hard Reset - Manages Volume Flow Correctly"""
        try:
            self.solver.CanSolve = False
            
            # CORE SETTINGS
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            
            # UNIT CONVERSION: m3/h -> m3/s
            vf_s = vol_flow / 3600.0
            print(f"   [DEBUG_RESET] Setting VolFlow: {vol_flow} m3/h -> {vf_s:.4f} m3/s")
            self.mgr.nodes['ADJ1'].TargetValue.Value = vf_s
            
            # SAFE CONDITION
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0 
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            
            for adj in self.mgr.adjs: adj.Reset()
            self.solver.CanSolve = True
            
            self.wait_stable(20, 1.0)
            
            # Double Reset
            for adj in self.mgr.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            
            return True
        except: return False

    def set_inputs(self, flow, vol_flow, p_bar, t_deg):
        try:
            # Although Flow/Vol are usually set, re-asserting them ensures consistency
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            
            # UNIT CONVERSION: m3/h -> m3/s
            vf_s = vol_flow / 3600.0
            # print(f"   [DEBUG_SET] Setting VolFlow: {vol_flow} m3/h -> {vf_s:.4f} m3/s")
            self.mgr.nodes['ADJ1'].TargetValue.Value = vf_s
            
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            
            if not self.wait_stable(20, 1.0): return False
            time.sleep(1.0)
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

    def strategy_anchor(self, flow, vol_flow):
        anchor = ANCHOR_POINTS.get(flow)
        print(f"   [ANCHOR] Target P={anchor['P']}, T={anchor['T']}")
        
        # 1. Soft Try
        if self.set_inputs(flow, vol_flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m and m['MA'] > 0.05:
                 print(f"   Verified (Soft): Power={m['Power']:.1f}")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        # 2. Hard Reset & Retry
        print("   [SOFT FAIL] Resetting...")
        self.recover_state(flow, vol_flow, anchor['P'])
        if self.set_inputs(flow, vol_flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m: 
                 print(f"   Verified (Hard): Power={m['Power']:.1f}")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        # 3. Fallback to Grid (New Feature)
        print("   [ANCHOR FAIL] Switching to Grid Scan...")
        return self.strategy_grid_scan(flow, vol_flow, anchor['P'])

    def strategy_secant(self, flow, vol_flow, p_start, t_start):
        print(f"   [SECANT] Start P={p_start}, T={t_start}...", end="", flush=True)
        best_p_res = None
        
        for p in [p_start, round(p_start-0.1, 2), round(p_start+0.1, 2)]:
            target_ma = 2.0
            t0, t1 = t_start + 2.0, t_start - 1.0
            
            if not self.set_inputs(flow, vol_flow, p, t0): continue
            m0 = self.get_metrics()
            if not m0: continue
            
            local_best = {'T': t0, **m0} if m0['MA']>0.5 else None
            y0 = m0['MA'] - target_ma
            
            for i in range(4):
                if not self.set_inputs(flow, vol_flow, p, t1): break
                m1 = self.get_metrics()
                if not m1: break
                
                if m1['MA'] >= 1.0 and m1['S7_Pres'] < 37.0:
                    if local_best is None or m1['Power'] < local_best['Power']:
                        local_best = {'T': t1, **m1}
                
                y1 = m1['MA'] - target_ma
                if abs(y1) < 0.2: break 
                
                if abs(y1 - y0) < 0.01: t_next = t1 - 0.5
                else: t_next = t1 - y1 * (t1 - t0) / (y1 - y0)
                
                t_next = max(-125.0, min(-90.0, t_next))
                t0, y0, t1 = t1, y1, t_next
            
            if local_best:
                if best_p_res is None or local_best['Power'] < best_p_res['Power']:
                    best_p_res = {'P': p, **local_best}

        if best_p_res:
            print(f" OK ({best_p_res['Power']:.1f} kW)")
            return best_p_res
        
        print(" FAIL")
        return None

    def strategy_grid_scan(self, flow, vol_flow, p_center):
        print(f"   [GRID SCAN] Center P={p_center}...", end="", flush=True)
        self.recover_state(flow, vol_flow, p_center)
        best = None
        for p in [round(p_center + i*0.1, 1) for i in range(-1, 2)]:
            valid_range = []
            for t in range(-90, -125, -2):
                if not self.set_inputs(flow, vol_flow, p, t): 
                    self.recover_state(flow, vol_flow, p)
                    continue
                m = self.get_metrics()
                if m and 0.5 <= m['MA'] <= 3.5:
                    valid_range = range(t+2, t-5, -1)
                    break
            
            if valid_range:
                for t in valid_range:
                    if not self.set_inputs(flow, vol_flow, p, t): continue
                    m = self.get_metrics()
                    if m and 2.0 <= m['MA'] <= 3.0 and m['S7_Pres'] <= 37.0:
                        if best is None or m['Power'] < best['Power']:
                            best = {'P':p, 'T':t, **m}
        if best: print(f" OK ({best['Power']:.1f} kW)")
        else: print(" FAIL")
        return best

    def run(self):
        print(">> MULTIDIM HYBRID OPTIMIZER STARTED")
        keys = ['Flow', 'VolFlow', 'P_bar', 'T_C', 'MA', 'Power', 'S6_Pres', 'Status', 'Method']
        if not os.path.exists('hysys_automation/' + OUT_FILE):
             with open('hysys_automation/' + OUT_FILE, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()

        for flow in FLOWS:
            print(f"\n>>>> MASS FLOW: {flow} kg/h")
            
            for vol in VOL_FLOWS:
                print(f"   >> VolFlow: {vol} m3/h")
                
                res = None
                method = "None"
                
                # DISPATCH
                if flow in ANCHOR_POINTS:
                    method = "Anchor"
                    res = self.strategy_anchor(flow, vol)
                    
                elif flow <= 1000:
                    p_pred, t_pred = self.model.predict(flow)
                    if p_pred is None: p_pred, t_pred = PRESET_P.get(flow, 4.0), -110.0
                    
                    method = "Secant"
                    res = self.strategy_secant(flow, vol, p_pred, t_pred)
                    if not res:
                        method = "GridScan(Fallback)"
                        res = self.strategy_grid_scan(flow, vol, p_pred)
                else:
                    method = "GridScan"
                    p_center = PRESET_P.get(flow, 5.5)
                    p_pred, _ = self.model.predict(flow)
                    if p_pred: p_center = p_pred
                    res = self.strategy_grid_scan(flow, vol, p_center)

                # SAVE
                if res:
                    row = {'Flow': flow, 'VolFlow': vol, 'P_bar': res['P'], 'T_C': res['T'], 
                           'MA': res['MA'], 'Power': res['Power'], 'S6_Pres': res.get('S6_Pres', 0),
                           'Status': 'Optimized', 'Method': method}
                    with open('hysys_automation/' + OUT_FILE, 'a', newline='') as f:
                        csv.DictWriter(f, fieldnames=keys).writerow(row)
                else:
                    print(f"   [no solution]")

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
        opt = MultidimOptimizer(mgr, mdl)
        opt.run()
    finally:
        stop_event.set()

if __name__ == "__main__":
    main()
