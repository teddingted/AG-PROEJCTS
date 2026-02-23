import os, time, csv, threading
import win32com.client
import win32gui, win32con

"""
2D HYBRID OPTIMIZER (hysys_optimizer_2d.py)
-------------------------------------------
Dimensions:
1. Mass Flow (S10): 500-1500 kg/h (100 step)
2. Volume Flow (ADJ-1): 3500-3800 m3/h (50 step) -> m3/s conversion required
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_2d_result.csv"
MODEL_DATA_FILE = "hysys_automation/optimization_final_summary_verified.csv"

# Configuration
FLOWS = range(500, 1501, 100)
VOL_FLOWS = range(3400, 3801, 50)

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
            print(f"[MODEL] Loaded {len(self.data)} points.")

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
        sim_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        print(f"[INFO] Opening Case: {sim_path}")
        
        try:
            # Try to activate if already open, or open if not
            self.case = app.SimulationCases.Open(sim_path)
        except Exception as e:
            print(f"[WARN] Open failed ({e}). Falling back to finding in open cases...")
            self.case = None
            for c in app.SimulationCases:
                if SIM_FILE in c.FullPath:
                    self.case = c
                    break
            if not self.case:
                 # Last resort: active document
                 self.case = app.ActiveDocument

        if self.case:
            print(f"[INFO] Active Case: {self.case.Title.Value}")
            self.case.Visible = True
        else:
            raise Exception("Could not find or open target simulation case.")
        self.solver = self.case.Solver
        self.fs = self.case.Flowsheet
        
        self.nodes = {
            'S1': self.fs.MaterialStreams.Item("1"),
            'S10': self.fs.MaterialStreams.Item("10"),
            'S6': self.fs.MaterialStreams.Item("6"),
            'S7': self.fs.MaterialStreams.Item("7"),
            'LNG': self.fs.Operations.Item("LNG-100"),
            'Spreadsheet': self.fs.Operations.Item("SPRDSHT-1"),
            'ADJ4': self.fs.Operations.Item("ADJ-4"),
            'ADJ1': self.fs.Operations.Item("ADJ-1"),
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class Optimizer2D:
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
        """Hard Reset with Volume Flow"""
        try:
            self.solver.CanSolve = False
            
            # Set Mass Flow (kg/h -> kg/s)
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            
            # Set Volume Flow (m3/h -> m3/s)
            vol_s = vol_flow / 3600.0
            print(f"   [RESET] Vol={vol_flow} m3/h -> {vol_s:.4f} m3/s", flush=True)
            self.mgr.nodes['ADJ1'].TargetValue.Value = vol_s
            
            # Safe Condition
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

    def set_inputs(self, flow, vol_flow, p_bar, t_deg):
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            
            vol_s = vol_flow / 3600.0
            self.mgr.nodes['ADJ1'].TargetValue.Value = vol_s
            
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
            
            # Basic Metrics
            base_metrics = {
                'MA': ma,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'S7_Pres': self.mgr.nodes['S7'].Pressure.Value / 100.0
            }
            
            # Extended Metrics from Spreadsheet A12-J12
            ss = self.mgr.nodes['Spreadsheet']
            extended = {}
            
            # Retry loop for Spreadsheet Read (Handling Transient COM Errors)
            for attempt in range(3):
                try:
                    extended = {
                        'SS_MassFlow': ss.Cell("A12").CellValue,
                        'SS_VolFlow': ss.Cell("B12").CellValue,
                        'SS_SuctionP': ss.Cell("C12").CellValue,
                        'SS_DischargeP': ss.Cell("D12").CellValue,
                        'SS_Ratio': ss.Cell("E12").CellValue,
                        'SS_Duty_kW': ss.Cell("F12").CellValue,
                        'SS_LMTD': ss.Cell("G12").CellValue,
                        'SS_MinAppr': ss.Cell("H12").CellValue,
                        'SS_UA_kWC': ss.Cell("I12").CellValue / 1000.0,
                        'SS_ExpInletT': ss.Cell("J12").CellValue
                    }
                    base_metrics.update(extended)
                    break # Success
                except Exception as e:
                    if attempt < 2: 
                        time.sleep(0.5) # Wait and retry
                    else:
                        print(f"   [WARN] Failed to read Spreadsheet A12-J12 after 3 attempts: {e}")
                
            return base_metrics
        except: return None

    def strategy_anchor(self, flow, vol_flow):
        anchor = ANCHOR_POINTS.get(flow)
        print(f"   [ANCHOR] P={anchor['P']}, T={anchor['T']}")
        
        if self.set_inputs(flow, vol_flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m and m['MA'] > 0.05:
                 print(f"   OK: {m['Power']:.1f} kW")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        print("   [RESET]...")
        self.recover_state(flow, vol_flow, anchor['P'])
        if self.set_inputs(flow, vol_flow, anchor['P'], anchor['T']):
             m = self.get_metrics()
             if m:
                 print(f"   OK: {m['Power']:.1f} kW")
                 return {'P': anchor['P'], 'T': anchor['T'], **m}
        
        print("   [FALLBACK -> GRID]")
        return self.strategy_grid_scan(flow, vol_flow, anchor['P'])

    def strategy_secant(self, flow, vol_flow, p_start, t_start):
        print(f"   [SECANT] P={p_start}, T={t_start}...", end="", flush=True)
        best = None
        
        for p in [p_start, round(p_start-0.1, 2), round(p_start+0.1, 2)]:
            target_ma, t0, t1 = 2.0, t_start + 2.0, t_start - 1.0
            
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
                
                t_next = t1 - 0.5 if abs(y1 - y0) < 0.01 else t1 - y1 * (t1 - t0) / (y1 - y0)
                t_next = max(-125.0, min(-90.0, t_next))
                t0, y0, t1 = t1, y1, t_next
            
            if local_best and (best is None or local_best['Power'] < best['Power']):
                best = {'P': p, **local_best}

        if best:
            print(f" OK ({best['Power']:.1f} kW)")
            return best
        print(" FAIL")
        return None

    def strategy_grid_scan(self, flow, vol_flow, p_center):
        print(f"   [GRID] P={p_center}...", end="", flush=True)
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
                            best = {'P': p, 'T': t, **m}
        
        if best: print(f" OK ({best['Power']:.1f} kW)")
        else: print(" FAIL")
        return best

    def get_completed_points(self):
        done = set()
        path = 'hysys_automation/' + OUT_FILE
        if os.path.exists(path):
            with open(path, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        f_val = int(float(row['Flow']))
                        v_val = int(float(row['VolFlow']))
                        done.add((f_val, v_val))
                    except: pass
        print(f"[RESUME] Found {len(done)} completed points.")
        return done

    def run(self):
        global OUT_FILE
        print("="*60)
        print(">> 2D HYBRID OPTIMIZER STARTED (Mass x Volume) + Extended Data")
        print("="*60)
        
        # Define Columns
        keys = ['Flow', 'VolFlow', 'P_bar', 'T_C', 'MA', 'Power', 'S6_Pres', 'Method']
        # Add Extended Keys
        extended_keys = ['SS_MassFlow', 'SS_VolFlow', 'SS_SuctionP', 'SS_DischargeP', 
                         'SS_Ratio', 'SS_Duty_kW', 'SS_LMTD', 'SS_MinAppr', 'SS_UA_kWC', 'SS_ExpInletT']
        keys.extend(extended_keys)
        
        # Use new filename for extended data
        OUT_FILE = "optimization_2d_extended.csv"
        
        if not os.path.exists('hysys_automation/' + OUT_FILE):
             with open('hysys_automation/' + OUT_FILE, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()

        done_points = self.get_completed_points()

        for flow in FLOWS:
            print(f"\n{'='*60}\nMASS FLOW: {flow} kg/h\n{'='*60}")
            
            for vol in VOL_FLOWS:
                if (flow, vol) in done_points:
                    print(f"  Vol: {vol} m3/h [SKIPPING - ALREADY DONE]")
                    continue

                print(f"  Vol: {vol} m3/h")
                res, method = None, "None"
                
                if flow in ANCHOR_POINTS:
                    method = "Anchor"
                    res = self.strategy_anchor(flow, vol)
                elif flow <= 1000:
                    p_pred, t_pred = self.model.predict(flow)
                    if p_pred is None: p_pred, t_pred = PRESET_P.get(flow, 4.0), -110.0
                    method = "Secant"
                    res = self.strategy_secant(flow, vol, p_pred, t_pred)
                    if not res:
                        method = "Grid(FB)"
                        res = self.strategy_grid_scan(flow, vol, p_pred)
                else:
                    method = "Grid"
                    p_center = PRESET_P.get(flow, 5.5)
                    p_pred, _ = self.model.predict(flow)
                    if p_pred: p_center = p_pred
                    res = self.strategy_grid_scan(flow, vol, p_center)

                if res:
                    # Construct Row
                    row = {
                        'Flow': flow, 'VolFlow': vol, 
                        'P_bar': res['P'], 'T_C': res['T'],
                        'MA': res['MA'], 'Power': res['Power'], 
                        'S6_Pres': res.get('S6_Pres', 0), 'Method': method
                    }
                    # Add Extended Data if present
                    for k in extended_keys:
                        row[k] = res.get(k, '')
                        
                    with open('hysys_automation/' + OUT_FILE, 'a', newline='') as f:
                        csv.DictWriter(f, fieldnames=keys).writerow(row)
                else:
                    print("  [FAIL]")

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
        opt = Optimizer2D(mgr, mdl)
        opt.run()
    finally:
        stop_event.set()

if __name__ == "__main__":
    main()
