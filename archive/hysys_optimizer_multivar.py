import os, time, csv, threading, contextlib
import win32com.client
import win32gui, win32con, win32api

# --- Config ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_highload_result.csv"

# High Load Focus
FLOWS_MASS = [1300, 1400, 1500] 
FLOWS_VOL = range(3200, 3801, 50) 

PRESET_P = {
    1300: 6.5, 1400: 7.2, 1500: 7.5 
}

# --- Utils ---
def dismiss_popup():
    print("[Popup Handler] Active (Simple)")
    while True:
        try:
            hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try: win32gui.SetForegroundWindow(hwnd)
                except: pass
                time.sleep(0.5)
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(1.0) # Wait for it to close
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
        self._adj1 = fs.Operations.Item("ADJ-1") 
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

    @property
    def mass_flow(self): return self._s10.MassFlow.Value * 3600.0
    @mass_flow.setter
    def mass_flow(self, val): self._s10.MassFlow.Value = val / 3600.0

    @property # ADJ-1 Target (Volume Flow m3/h -> m3/s)
    def vol_flow(self): return self._adj1.TargetValue.Value * 3600.0
    @vol_flow.setter
    def vol_flow(self, val): self._adj1.TargetValue.Value = val / 3600.0

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

    def _wait(self, timeout=30, stable=1.0):
        time.sleep(0.5)
        try:
           t0 = time.time()
           while self.solver.IsSolving and (time.time()-t0 < timeout): time.sleep(0.2)
           
           ref_t, last_v = None, -999
           while time.time()-t0 < timeout:
               v = self._s1.Temperature.Value
               if v < -200: return False
               if abs(v - last_v) < 0.01:
                   if not ref_t: ref_t = time.time()
                   elif time.time()-ref_t >= stable: return True
               else: ref_t = None
               last_v = v; time.sleep(0.2)
        except: return False
        return True

    def reset(self, m_flow, v_flow, p):
        with self.frozen():
            self.mass_flow = m_flow
            self.vol_flow = v_flow
            self.pressure = p
            self._s1.Temperature.Value = 40.0
            self.target_temp = -90.0
            for a in self._adjs: a.Reset()
        self._wait(20, 1.0)
        for a in self._adjs: a.Reset()
        self._wait(10, 1.0)
        return self.is_healthy

    def set_point(self, mf, vf, p, t):
        try:
            if not self.is_healthy:
                if not self.reset(mf, vf, p): return False
            self.mass_flow = mf
            self.vol_flow = vf
            self.pressure = p
            self.target_temp = t
            if not self._wait(20, 1.0): return False
            time.sleep(1.0)
            return self.is_healthy
        except: return False

    def get_result(self):
        try:
            app = self._lng.MinApproach.Value
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
    print("="*60 + "\nHIGH LOAD OPTIMIZER (1300-1500)\n" + "="*60)
    
    # Resume Logic: Load existing progress
    done_points = set()
    if os.path.exists(os.path.join(FOLDER, OUT_FILE)):
        with open(os.path.join(FOLDER, OUT_FILE), 'r') as f:
            reader = csv.reader(f)
            next(reader, None) # Skip header
            for row in reader:
                if row:
                    try: done_points.add((int(row[0]), int(row[1])))
                    except: pass
    
    # Open in Append Mode
    write_header = not os.path.exists(os.path.join(FOLDER, OUT_FILE))
    with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as f:
        writer = csv.writer(f)
        if write_header: writer.writerow(["MassFlow", "VolFlow", "P", "T", "App", "Power"])

        for m_flow in FLOWS_MASS:
            print(f"\n>> MassFlow {m_flow} kg/h")
            cp = PRESET_P[m_flow]
            
            for v_flow in FLOWS_VOL:
                if (m_flow, v_flow) in done_points:
                    print(f"   [Vol {v_flow}]: SKIPPED (Already Done)")
                    continue

                print(f"   [Vol {v_flow}]: ", end="")
                eng.reset(m_flow, v_flow, cp)
                
                best = None
                
                # STABLE LINEAR SCAN: -90 to -115 (Step -1)
                # Removes "Deep Scan" complexity to avoid instability
                for p in [round(cp + i*0.1, 1) for i in range(-5, 6)]:
                    print(f"{p}:", end="", flush=True)
                    
                    for t in range(-90, -116, -1):
                        if not eng.set_point(m_flow, v_flow, p, t):
                            print("!", end="")
                            eng.reset(m_flow, v_flow, p) 
                            continue
                        
                        m = eng.get_result()
                        if not m: continue
                        
                        # Valid Point?
                        if 2.0 <= m['app'] <= 3.0 and m['p7'] <= 36.5:
                            if not best or m['pwr'] < best['pwr']:
                                best = {'p':p, 't':t, **m}
                                print("*", end="") # New Best
                            else:
                                print(".", end="") # Valid but not best
                        elif m['app'] < 0.5:
                            print("X", end="") # Cross
                        else:
                            print("-", end="") # Wide
                            
                    print("|", end="")
    
                if best:
                    print(f" BEST:{best['p']}b, {best['pwr']:.1f}kW")
                    with open(os.path.join(FOLDER, OUT_FILE), 'a', newline='') as f:
                        csv.writer(f).writerow([m_flow, v_flow, best['p'], best['t'], best['app'], best['pwr']])
                else:
                    print(" NO SOLUTION")

if __name__ == "__main__":
    main()
