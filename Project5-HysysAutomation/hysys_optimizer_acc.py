import os, time, csv, threading, math
import win32com.client
import win32gui, win32con

"""
HYSYS ACCELERATED OPTIMIZER (hysys_optimizer_acc.py)
----------------------------------------------------
Techniques Used:
1.  **Surrogate Modeling (ML/Interpolation)**: Uses historical verified data to predict 
    optimal Pressure (P) and Temperature (T) start points.
2.  **Gradient-Free Optimization (Secant Method)**: Replaces brute-force T-scan with 
    a root-finding algorithm to converge on Target MA = 2.0°C rapidly (O(log n)).
3.  **Adaptive Heuristics**: Dynamically adjusts search bounds based on model confidence.
4.  **Legacy Stability Integration**: Retains the "Hard Reset" safety net for robustness.

Target: Maximize Speed without compromising Stability.
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
MODEL_DATA_FILE = "hysys_automation/optimization_final_summary_verified.csv"
OUT_FILE = "optimization_accelerated_result.csv"

# Configuration
FLOWS = range(500, 1501, 100) # Full Range

class SurrogateModel:
    """
    Data-Driven Model to predict P and T for a given Flow.
    Mimics 'Machine Learning' via piecewise linear regression (Interpolation) on verified data.
    """
    def __init__(self, csv_path):
        self.data = []
        if os.path.exists(csv_path):
            with open(csv_path, 'r') as f:
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
            print(f"[MODEL] Loaded {len(self.data)} verified points for surrogate modeling.")
        else:
            print("[MODEL] Warning: No training data found. Using fallback heuristics.")

    def predict(self, flow):
        """Predicts Optimal P and T for a given Flow"""
        # 1. Exact Match
        for d in self.data:
            if abs(d['Flow'] - flow) < 1.0:
                return d['P'], d['T']
        
        # 2. Interpolation
        if len(self.data) < 2:
            return None, None # Not enough data
            
        # Find neighbors
        lower = None
        upper = None
        for d in self.data:
            if d['Flow'] <= flow: lower = d
            if d['Flow'] >= flow and upper is None: upper = d
            
        if lower and upper and lower != upper:
            ratio = (flow - lower['Flow']) / (upper['Flow'] - lower['Flow'])
            p_pred = lower['P'] + ratio * (upper['P'] - lower['P'])
            t_pred = lower['T'] + ratio * (upper['T'] - lower['T'])
            return round(p_pred, 2), round(t_pred, 1)
        
        # Extrapolation (Edges)
        if lower: return lower['P'], lower['T']
        if upper: return upper['P'], upper['T']
        
        return None, None # Failure

class HysysNodeManager:
    """Lightweight Node Manager"""
    def __init__(self, app):
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app.SimulationCases.Open(file_path)
        else:
            self.case = app.ActiveDocument
        self.case.Visible = True
        self.solver = self.case.Solver
        self.fs = self.case.Flowsheet
        
        # Nodes
        self.nodes = {
            'S1': self.fs.MaterialStreams.Item("1"),
            'S10': self.fs.MaterialStreams.Item("10"),
            'S6': self.fs.MaterialStreams.Item("6"),
            'S7': self.fs.MaterialStreams.Item("7"),
            'LNG': self.fs.Operations.Item("LNG-100"),
            'Spreadsheet': self.fs.Operations.Item("SPRDSHT-1"),
            'ADJ4': self.fs.Operations.Item("ADJ-4"),
        }
        self.adjs = [self.fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]

class AcceleratedOptimizer:
    def __init__(self, manager, model):
        self.mgr = manager
        self.model = model
        self.solver = manager.solver

    def wait_stable(self, timeout=20, stable_time=0.5): # Accelerated Timings
        time.sleep(0.2)
        start = time.time()
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.1) # Fast polling
            
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.mgr.nodes['S1'].Temperature.Value
                if curr < -200: return False
                if abs(curr - last_val) < 0.02: # Slightly loose tolerance for speed
                    if ref_start is None: ref_start = time.time()
                    elif time.time() - ref_start >= stable_time: return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.1)
            except: return False
        return True

    def recover_state(self, flow, p_bar):
        """Standard Hard Reset"""
        try:
            self.solver.CanSolve = False
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            for adj in self.mgr.adjs: adj.Reset()
            self.solver.CanSolve = True
            self.wait_stable(15, 0.5)
            for adj in self.mgr.adjs: adj.Reset()
            self.wait_stable(5, 0.5)
            return True
        except: return False

    def set_inputs(self, flow, p_bar, t_deg):
        """Fast Input Set"""
        try:
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = t_deg
            return self.wait_stable(15, 0.5)
        except: return False

    def get_metrics(self):
        try:
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            if ma < 0.5: # Anti-lag
                time.sleep(0.5)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
            return {
                'MA': ma,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'S7_Pres': self.mgr.nodes['S7'].Pressure.Value / 100.0
            }
        except: return None

    def optimize_temperature_secant(self, flow, p_bar, t_start):
        """
        Secant Method to find T where MA ~= 2.0 (Target)
        O(log n) complexity vs O(n) grid scan.
        """
        target_ma = 2.0
        
        # History for convergence check
        t0 = t_start + 2.0 # Warmer guess
        t1 = t_start - 1.0 # Cooler guess
        
        # 1. Evaluate T0
        if not self.set_inputs(flow, p_bar, t0): return None
        m0 = self.get_metrics()
        if not m0: return None
        y0 = m0['MA'] - target_ma
        
        best_point = {'T': t0, **m0} if m0['MA'] > 0.5 else None
        
        print(f"    [SECANT] T0={t0:.1f}C -> MA={m0['MA']:.2f}", end="", flush=True)

        for i in range(5): # Max 5 steps
            # 2. Evaluate T1
            if not self.set_inputs(flow, p_bar, t1): break
            m1 = self.get_metrics()
            if not m1: break
            y1 = m1['MA'] - target_ma
            
            print(f", T{i+1}={t1:.1f}C(MA={m1['MA']:.2f})", end="", flush=True)
            
            # Update Best Point
            if m1['MA'] > 0.5:
                if best_point is None or m1['Power'] < best_point['Power']:
                     # Strict constraints for 'Best'
                     if m1['MA'] >= 1.0 and m1['S7_Pres'] < 37.0:
                        best_point = {'T': t1, **m1}

            # 3. Check Convergence
            if abs(y1) < 0.2: # Close enough to MA=2.0
                print(" -> Converged!")
                break
            
            # 4. Secant Step
            if abs(y1 - y0) < 0.01: # Avoid division by zero
                # Bump T1 slightly to get gradient
                t_next = t1 - 0.5
            else:
                # T_new = T1 - y1 * (T1 - T0) / (y1 - y0)
                t_next = t1 - y1 * (t1 - t0) / (y1 - y0)
                
            # Safety Bounds (-90 to -125)
            t_next = max(-125.0, min(-90.0, t_next))
            
            # Shift
            t0, y0 = t1, y1
            t1 = t_next
            
        print("") # Newline
        return best_point

    def run(self):
        print(f">> ACCELERATED OPTIMIZATION START (Target Flow: {FLOWS[0]}-{FLOWS[-1]})")
        
        # CSV Init
        keys = ['Flow', 'P_bar', 'T_C', 'MA', 'Power', 'Status']
        if not os.path.exists('hysys_automation/' + OUT_FILE):
             with open('hysys_automation/' + OUT_FILE, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()
                
        for flow in FLOWS:
            print(f"\n>> Flow {flow} kg/h", flush=True)
            
            # 1. Predict Start Point
            p_pred, t_pred = self.model.predict(flow)
            if p_pred is None: p_pred, t_pred = 4.0, -100.0 # Fallback
            
            print(f"   [MODEL] Predicted Start: P={p_pred}, T={t_pred}")
            
            # 2. Local Pressure Search around Prediction
            # Instead of wide scan, check P_pred, P_pred-0.1, P_pred+0.1
            pressures = [p_pred, round(p_pred-0.1, 2), round(p_pred+0.1, 2)]
            # If Model is confident (high flow), maybe just P_pred? 
            # Let's stick to 3 points for robustness.
            
            best_res = None
            
            # Reset before P-Loop
            self.recover_state(flow, p_pred) 
            
            for p in pressures:
                print(f"   Scanning P={p}:", end="", flush=True)
                
                # Temperature Optimization (Secant)
                # Start secant near t_pred
                res = self.optimize_temperature_secant(flow, p, t_pred)
                
                if res:
                    if best_res is None or res['Power'] < best_res['Power']:
                        best_res = {'P': p, **res}
            
            if best_res:
                print(f"   >> OPTIMAL: P={best_res['P']}, T={best_res['T']:.1f}, Pwr={best_res['Power']:.1f} kW")
                row = {'Flow': flow, 'P_bar': best_res['P'], 'T_C': best_res['T'], 
                       'MA': best_res['MA'], 'Power': best_res['Power'], 'Status': 'Accelerated'}
                with open('hysys_automation/' + OUT_FILE, 'a', newline='') as f:
                     csv.DictWriter(f, fieldnames=keys).writerow(row)
            else:
                print("   >> FAILED.")

def main():
    stop_event = threading.Event()
    def kill_popups():
        while not stop_event.is_set():
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except: pass
            time.sleep(1.0)
    threading.Thread(target=kill_popups, daemon=True).start()

    try:
        try: app = win32com.client.GetActiveObject("HYSYS.Application")
        except: app = win32com.client.Dispatch("HYSYS.Application")
        
        # Init Model
        model = SurrogateModel(MODEL_DATA_FILE)
        
        mgr = HysysNodeManager(app)
        opt = AcceleratedOptimizer(mgr, model)
        opt.run()
        
    finally:
        stop_event.set()

if __name__ == "__main__":
    main()
