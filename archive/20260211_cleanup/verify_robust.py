import os, time, csv, threading
import win32com.client

"""
ROBUST VERIFICATION SCRIPT
Based on lessons from hysys_optimizer_final.py
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"

class RobustVerifier:
    def __init__(self, app):
        # Attach to case
        file_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
        if os.path.exists(file_path):
            self.case = app.SimulationCases.Open(file_path)
        else:
            self.case = app.ActiveDocument
        
        # Cache objects
        fs = self.case.Flowsheet
        self.solver = self.case.Solver
        self.s1 = fs.MaterialStreams.Item("1")  # N2 Loop
        self.s10 = fs.MaterialStreams.Item("10")  # Reliq
        self.s6 = fs.MaterialStreams.Item("6")
        self.adj4 = fs.Operations.Item("ADJ-4")
        self.lng = fs.Operations.Item("LNG-100")
        self.comp = fs.Operations.Item("K-100")
        self.sprdsht = fs.Operations.Item("SPRDSHT-1")
        self.adjs = [fs.Operations.Item(f"ADJ-{i}") for i in range(1, 5)]  # ADJ-1 to ADJ-4
    
    def is_healthy(self):
        """Health check: detect divergence early"""
        try:
            t = self.s1.Temperature.Value
            if t < -200: return False  # Cryogenic failure
            pwr = self.comp.EnergyValue
            if pwr < 0.1: return False  # Compressor trip
            return True
        except:
            return False
    
    def recover_state(self, flow, p_bar):
        """Hard reset with double-reset pattern"""
        print(f"   [RECOVERY] Hard Reset for Flow={flow}, P={p_bar}...", flush=True)
        try:
            # Phase 1: Pause solver and set safe state
            self.solver.CanSolve = False
            time.sleep(0.5)
            
            self.s10.MassFlow.Value = flow / 3600.0
            self.s1.Pressure.Value = p_bar * 100.0
            self.s1.Temperature.Value = 40.0  # Safe warm start
            self.adj4.TargetValue.Value = -90.0
            
            # Reset all adjusts
            for adj in self.adjs:
                try: adj.Reset()
                except: pass
            
            # Phase 2: Resume solver
            self.solver.CanSolve = True
            self.wait_stable(20, 1.0)
            
            # Phase 3: Double reset to clear latches
            for adj in self.adjs:
                try: adj.Reset()
                except: pass
            self.wait_stable(10, 1.0)
            
            return self.is_healthy()
        except:
            return False
    
    def wait_stable(self, timeout=30, stable_time=1.0):
        """Wait for temperature stability"""
        time.sleep(0.5)
        start = time.time()
        
        # Wait for solver idle
        while time.time() - start < timeout:
            if not self.solver.IsSolving: break
            time.sleep(0.2)
        
        # Wait for temperature stability
        ref_start, last_val = None, -9999
        while time.time() - start < timeout:
            try:
                curr = self.s1.Temperature.Value
                if curr < -200: return False
                
                if abs(curr - last_val) < 0.01:
                    if ref_start is None:
                        ref_start = time.time()
                    elif time.time() - ref_start >= stable_time:
                        return True
                else:
                    ref_start = None
                last_val = curr
                time.sleep(0.2)
            except:
                return False
        return True
    
    def set_inputs_safe(self, flow, p_bar, t_deg):
        """Set inputs with safety checks and gradual changes"""
        try:
            # Pre-check health
            if not self.is_healthy():
                if not self.recover_state(flow, p_bar):
                    return False
            
            # Set flow
            self.s10.MassFlow.Value = flow / 3600.0
            
            # Gradual pressure change if > 1 bar difference
            current_p = self.s1.Pressure.Value
            if abs(current_p - p_bar * 100) > 100:
                self.s1.Pressure.Value = p_bar * 100.0
                self.wait_stable(5, 0.5)
            else:
                self.s1.Pressure.Value = p_bar * 100.0
            
            # Set temperature
            self.adj4.TargetValue.Value = t_deg
            
            # Wait for convergence
            if not self.wait_stable(20, 1.0):
                return False
            
            # CRITICAL: Settle buffer for laggy adjusts
            time.sleep(1.5)
            
            # Post-check health
            if not self.is_healthy():
                return False
            
            return True
        except:
            return False
    
    def get_metrics(self):
        """Collect metrics with anti-lag logic"""
        try:
            ma = self.lng.MinApproach.Value
            s6_p = self.s6.Pressure.Value / 100.0
            pwr = self.sprdsht.Cell("C8").CellValue
            
            # Anti-lag: double-check suspicious values
            if ma < 0.5 or ma > 100:
                time.sleep(1.5)
                ma = self.lng.MinApproach.Value
            
            return {
                'MA': ma,
                'S6_Pres': s6_p,
                'Power': pwr,
                'N2_Flow': self.s1.MassFlow.Value * 3600.0,
                'N2_Temp': self.s1.Temperature.Value,
                'N2_Pres': self.s1.Pressure.Value / 100.0
            }
        except:
            return None
    
    def verify_point(self, flow, p, t):
        """Verify a single operating point"""
        print(f"\n>> VERIFYING: Flow={flow} kg/h, P={p} bar, T={t}°C", flush=True)
        
        # Step 1: Hard reset to this point
        if not self.recover_state(flow, p):
            print("   [FAIL] Recovery failed", flush=True)
            return None
        
        # Step 2: Apply exact target
        if not self.set_inputs_safe(flow, p, t):
            print("   [FAIL] Could not set inputs", flush=True)
            return None
        
        # Step 3: Collect metrics
        m = self.get_metrics()
        if not m:
            print("   [FAIL] Could not read metrics", flush=True)
            return None
        
        # Validate
        valid = 'OK' if (m['MA'] > 0.1 and m['S6_Pres'] < 38.0 and m['Power'] > 10.0) else 'FAIL'
        
        print(f"   RES: Power={m['Power']:.2f} kW, MA={m['MA']:.2f}°C, S6_P={m['S6_Pres']:.2f} bar [{valid}]", flush=True)
        
        return {
            'Flow': flow,
            'P_bar': p,
            'T_C': t,
            **m,
            'Status': valid,
            'Timestamp': time.strftime("%H:%M:%S")
        }
    
    def run(self):
        # Load verification points
        points = []
        csv_path = 'hysys_automation/optimization_final_result.csv'
        with open(csv_path, 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row['MassFlow'] and row['Note'] == 'VERIFIED':  # Only verified points
                    points.append({
                        'flow': float(row['MassFlow']),
                        'p': float(row['P_bar']),
                        't': float(row['T_C'])
                    })
        
        # Sort ascending (500 -> 1300)
        points.sort(key=lambda x: x['flow'])
        
        print("="*60)
        print(" ROBUST VERIFICATION (500->1300 kg/h)")
        print("="*60)
        print(f"Points to verify: {len(points)}")
        
        results = []
        for pt in points:
            res = self.verify_point(pt['flow'], pt['p'], pt['t'])
            if res:
                results.append(res)
        
        # Save results
        if results:
            out_file = 'hysys_automation/robust_verification_report.csv'
            with open(out_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=results[0].keys())
                writer.writeheader()
                writer.writerows(results)
            print(f"\n[OK] Saved to {out_file}")
            
            # Print summary
            print("\n" + "="*80)
            print(f"{'Flow':<6} {'P':<6} {'T':<6} | {'Power':<10} {'MA':<6} {'S6_P':<8} | {'Status'}")
            print("-"*80)
            for r in results:
                print(f"{r['Flow']:<6.0f} {r['P_bar']:<6.1f} {r['T_C']:<6.0f} | {r['Power']:<10.2f} {r['MA']:<6.2f} {r['S6_Pres']:<8.2f} | {r['Status']}")
            print("="*80)

def main():
    # Popup killer
    def kill_popups():
        import win32gui, win32con
        while True:
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print("[POPUP] Closed", flush=True)
            except: pass
            time.sleep(0.5)
    threading.Thread(target=kill_popups, daemon=True).start()
    print("[POPUP KILLER] Started\n")
    
    # Connect
    try:
        app = win32com.client.GetActiveObject("HYSYS.Application")
        print("[CONNECT] Attached to HYSYS\n")
    except:
        app = win32com.client.Dispatch("HYSYS.Application")
        print("[CONNECT] Launched HYSYS\n")
    
    verifier = RobustVerifier(app)
    verifier.run()

if __name__ == "__main__":
    main()
