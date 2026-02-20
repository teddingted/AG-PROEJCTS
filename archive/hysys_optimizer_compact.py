import os, time, csv, threading, contextlib
import win32com.client
import win32gui, win32con, win32api

# --- Config ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_compact_result_v2.csv"
FLOWS = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]
PRESET_P = {
    500: 3.0, 600: 3.5, 700: 4.0, 800: 4.4, 900: 4.8, 1000: 5.2, 
    1100: 5.6, 1200: 6.0, 1300: 6.5, 1400: 7.2, 1500: 7.5 
}

# --- Utils ---
def dismiss_popup():
    print("[Popup Handler] Active")
    for _ in range(30):
        try:
            hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try: win32gui.SetForegroundWindow(hwnd)
                except: pass
                time.sleep(0.5)
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                break
        except: pass
        time.sleep(1)

# --- Core Engine ---
class HysysEngine:
    def __init__(self):
        self.app = win32com.client.Dispatch("HYSYS.Application")
        self.app.Visible = True
        fp = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        self.case = self.app.SimulationCases.Open(fp) if os.path.exists(fp) else self.app.ActiveDocument
        
        fs = self.case.Flowsheet
        self.solver = self.case.Solver
        self._s1 = fs.MaterialStreams.Item("1")
        self._s10 = fs.MaterialStreams.Item("10")
        self._s7 = fs.MaterialStreams.Item("7")
        self._adj4 = fs.Operations.Item("ADJ-4")
        self._lng = fs.Operations.Item("LNG-100")
        self._cmp = fs.Operations.Item("K-100")
        self._pwr = fs.Operations.Item("SPRDSHT-1").Cell("C8")
        self._adjs = [fs.Operations.Item(n) for n in ["ADJ-1", "ADJ-2", "ADJ-3", "ADJ-4"]]

    @contextlib.contextmanager
    def frozen(self):
        self.solver.CanSolve = False
        try: yield
        finally: self.solver.CanSolve = True

    # --- Properties ---
    @property
    def flow(self): return self._s10.MassFlow.Value * 3600.0
    @flow.setter
    def flow(self, val): self._s10.MassFlow.Value = val / 3600.0

    @property
    def pressure(self): return self._s1.Pressure.Value / 100.0
    @pressure.setter
    def pressure(self, val):
        curr = self.pressure
        if abs(curr - val) > 1.0: 
            self._s1.Pressure.Value = val * 100.0
            self._wait(5, 0.5)
        else:
            self._s1.Pressure.Value = val * 100.0

    @property
    def target_temp(self): return self._adj4.TargetValue.Value
    @target_temp.setter
    def target_temp(self, val): self._adj4.TargetValue.Value = val

    @property
    def is_healthy(self):
        try: return self._s1.Temperature.Value > -200 and self._cmp.EnergyValue > 0.1
        except: return False

    # --- Logic ---
    def _wait(self, timeout=30, stable=1.0):
        time.sleep(0.5)
        t0 = time.time()
        while self.solver.IsSolving and (time.time()-t0 < timeout): time.sleep(0.2)
        
        ref_t, last_v = None, -999
        while time.time()-t0 < timeout:
            try:
                v = self._s1.Temperature.Value
                if v < -200: return False
                if abs(v - last_v) < 0.01:
                    if not ref_t: ref_t = time.time()
                    elif time.time()-ref_t >= stable: return True
                else: ref_t = None
                last_v = v; time.sleep(0.2)
            except: return False
        return True

    def reset(self, flow, p):
        with self.frozen():
            self.flow = flow
            self.pressure = p
            self._s1.Temperature.Value = 40.0
            self.target_temp = -90.0
            for a in self._adjs: a.Reset()
        self._wait(20, 1.0)
        for a in self._adjs: a.Reset()
        self._wait(10, 1.0)
        return self.is_healthy

    def set_point(self, f, p, t):
        try:
            if not self.is_healthy:
                if not self.reset(f, p): return False
            
            self.flow = f
            self.pressure = p
            self.target_temp = t
            
            if not self._wait(20, 1.0): return False
            time.sleep(1.0) # Settle buffer
            return self.is_healthy
        except: return False

    def get_result(self):
        try:
            app = self._lng.MinApproach.Value
            # Anti-Lag
            if app < 0.5:
                time.sleep(1.2)
                app = self._lng.MinApproach.Value
            return {
                'app': app,
                'p7': self._s7.Pressure.Value / 100.0,
                'pwr': self._pwr.CellValue
            }
        except: return None

# --- Main ---
def main():
    threading.Thread(target=dismiss_popup, daemon=True).start()
    eng = HysysEngine()
    print("="*50 + "\nCOMPACT OPTIMIZER (DEEP SCAN)\n" + "="*50)
    
    with open(os.path.join(FOLDER, OUT_FILE), 'w', newline='') as f:
        csv.writer(f).writerow(["Flow", "P", "T", "App", "Power"])

    for flow in FLOWS:
        print(f"\nFlow {flow} kg/h")
        cp = PRESET_P[flow]
        eng.reset(flow, cp)
        
        best = None
        
        # Pressure Scan
        for p in [round(cp + i*0.1, 1) for i in range(-2, 3)]:
            print(f"  P={p}: ", end="", flush=True)
            deep_scan_range = None
            
            # 1. Coarse Scan (2C steps)
            for t in range(-90, -125, -2):
                if not eng.set_point(flow, p, t):
                    print("!", end=""); eng.reset(flow, p); continue
                
                res = eng.get_result()
                if not res: continue
                
                # TRIGGER: If positive approach (Potential exists), trigger Deep Scan
                if res['app'] > 0.0:
                    print(f" TRIG({t})", end="")
                    # Scan ample Space around this point (Higher AND Lower)
                    deep_scan_range = range(t + 4, t - 10, -1)
                    break
                print(".", end="") # Tight/Negative

            # 2. Deep Neighborhood Scan
            if deep_scan_range:
                print(f" -> DeepScan({deep_scan_range.start}..{deep_scan_range.stop})", end="")
                lbest = None
                for t in deep_scan_range:
                    if eng.set_point(flow, p, t):
                        m = eng.get_result()
                        # Strict Constraints
                        if m and 2.0 <= m['app'] <= 3.0 and m['p7'] <= 36.5:
                            if not lbest or m['pwr'] < lbest['pwr']:
                                lbest = {'p':p, 't':t, **m}
                                best = lbest if not best or lbest['pwr'] < best['pwr'] else best
                                print("*", end="")
                print(f" OK" if lbest else " (NoCstr)")
            else: print(" -")

        if best:
            print(f"  >> BEST: {best['p']}bar, {best['t']}C, {best['pwr']:.1f}kW")
            with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as f:
                csv.writer(f).writerow([flow, best['p'], best['t'], best['app'], best['pwr']])

if __name__ == "__main__":
    main()
