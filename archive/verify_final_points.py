import sys
import os
import time
import csv
# import pandas as pd # Removed

# Add current directory to path
sys.path.append(os.getcwd())

from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

class FinalVerifier:
    def __init__(self):
        self.mgr = HysysNodeManager(connect_hysys())
        # ... (rest of init same) ...
        self.mgr.status_logger = lambda msg: print(f"[STATUS] {msg}", flush=True)

        # Register Nodes (CORRECTED MAPPING & SCALING)
        # Scale: 3600 for kg/s <-> kg/h
        self.mgr.register_node('reliq.mass_flow', 'Stream', '10', 'MassFlow', scale=3600.0, unit='kg/h') 
        self.mgr.register_node('n2.mass_flow', 'Stream', '1', 'MassFlow', scale=3600.0, unit='kg/h')
        
        self.mgr.register_node('n2.temp', 'Stream', '1', 'Temperature')
        # Scale: 100 for kPa <-> bar
        self.mgr.register_node('n2.pressure', 'Stream', '1', 'Pressure', scale=100.0, unit='bar')
        
        # Scale: 1.0 for C
        self.mgr.register_node('control.target_temp', 'Block', 'ADJ-4', 'TargetValue')
        
        self.mgr.register_node('result.min_approach', 'Operation', 'LNG-100', 'MinApproach')
        self.mgr.register_node('result.ua', 'Operation', 'LNG-100', 'UA')
        self.mgr.register_node('result.lmtd', 'Operation', 'LNG-100', 'LMTD')
        self.mgr.register_node('result.compressor_power', 'Stream', 'Q-100', 'Power')
        self.mgr.register_node('result.s6_pressure', 'Stream', '6', 'Pressure', scale=100.0, unit='bar')
        self.mgr.register_node('result.s1_molar_flow', 'Stream', '1', 'MolarFlow', scale=3600.0) # kmol/s -> kmol/h approx? Actually scalar check needed.

    def run_kick(self):
        print("   [RECOVERY] KICK: Spiking Mass Flow...", flush=True)
        # 1. Force Solving OFF (via Manager if possible, or just write values)
        # Spike Flow
        self.mgr.write('n2.mass_flow', 20000.0, verify=False)
        self.mgr.write('n2.pressure', 6.0, verify=False)
        self.mgr.write('control.target_temp', -90.0, verify=False)
        self.mgr.wait_stable(5, 1.0)
        
        # Reset Adjusts
        self.reset_adjusts()
        self.mgr.wait_stable(10, 1.0)

    def reset_adjusts(self):
        print("   [RECOVERY] Resetting Adjust Blocks...", flush=True)
        try:
            fs = self.mgr.case.Flowsheet
            for i in range(1, 6):
                try:
                    fs.Operations.Item(f"ADJ-{i}").Reset()
                except: pass
            time.sleep(2.0)
        except: pass

    def verify_point(self, flow, p, t):
        print(f"\n>> VERIFYING: Reliq Flow={flow} kg/h, P={p} bar, T={t} C", flush=True)
        
        # 1. Set Inputs
        self.mgr.write('reliq.mass_flow', float(flow), verify=False)
        self.mgr.write('n2.pressure', float(p), verify=False)
        self.mgr.write('control.target_temp', float(t), verify=False)
        
        # 2. Wait for stable convergence
        self.mgr.wait_stable(20, 1.0)
        
        # 3. Check Adjusts (Strict) & Power Sanity
        fs = self.mgr.case.Flowsheet
        all_solved = True
        
        # Check Power > 10 kW (Sanity)
        pwr = self.mgr.read('result.compressor_power')
        if pwr is None or pwr < 10.0:
            print(f"   [FAIL] Low Power ({pwr}), assuming stuck.", flush=True)
            all_solved = False
            
        if all_solved:
            for i in range(1, 6):
                adj_name = f"ADJ-{i}"
                try:
                    adj = fs.Operations.Item(adj_name)
                    if not adj.IsIgnored:
                        try:
                            err = abs(adj.Variable("Error").Value)
                            tol = adj.Variable("Tolerance").Value
                            if err > tol:
                                print(f"   [FAIL] {adj_name} Not Solved (Err: {err:.2e})")
                                all_solved = False
                        except: pass
                except: pass
            
        if not all_solved:
            print("   -> Retrying with Reset & Kick...", flush=True)
            self.run_kick()
            
            # Re-apply inputs
            print(f"   -> Re-applying inputs...", flush=True)
            self.mgr.write('reliq.mass_flow', float(flow), verify=False)
            self.mgr.write('n2.pressure', float(p), verify=False)
            self.mgr.write('control.target_temp', float(t), verify=False)
            self.mgr.wait_stable(30, 2.0)

        # 4. Collect Data
        data = {
            'Target_Reliq_Flow': flow,
            'Input_N2_P': p,
            'Input_ADJ4_T': t,
            'Actual_N2_Flow': self.mgr.read('n2.mass_flow'),     # S1
            'Actual_Reliq_Flow': self.mgr.read('reliq.mass_flow'), # S10
            'N2_Pres': self.mgr.read('n2.pressure'),
            'N2_Temp': self.mgr.read('n2.temp'),
            'S6_Pres': self.mgr.read('result.s6_pressure'),
            'Power_kW': self.mgr.read('result.compressor_power'),
            'MA_C': self.mgr.read('result.min_approach'),
            'UA': self.mgr.read('result.ua'),
            'LMTD': self.mgr.read('result.lmtd'),
            'Timestamp': time.strftime("%H:%M:%S")
        }
        
        # Validation
        val_ma = 'OK' if (data['MA_C'] and data['MA_C'] > 0.1) else 'FAIL'
        val_pres = 'OK' if (data['S6_Pres'] and data['S6_Pres'] < 38.0) else 'FAIL'
        
        data['Valid_MA'] = val_ma
        data['Valid_Pres'] = val_pres
        
        print(f"   RES: Power={data['Power_kW']:.2f} kW, MA={data['MA_C']:.2f} C, S6_P={data['S6_Pres']:.2f} bar")
        return data

    def prepare_simulation(self):
        print("\n[INIT] Preparing Simulation State...", flush=True)
        # 1. Force Safe State
        # Set Reliq Flow High (1500)
        self.mgr.write('reliq.mass_flow', 1500.0, verify=False)
        # Set N2 Pressure High
        self.mgr.write('n2.pressure', 7.0, verify=False)
        # Set Target Temp Safe
        self.mgr.write('control.target_temp', -95.0, verify=False)
        
        # KICK: Spike N2 Flow (Stream 1) manually to wake up loop if stuck 
        # (Though ADJ-1 should control it, sometimes it needs a helper push)
        self.mgr.write('n2.mass_flow', 20000.0, verify=False)
        
        time.sleep(2.0)
        
        # 2. Reset Adjusts
        self.reset_adjusts()
        
        # 3. Wait for Stability
        if self.mgr.wait_stable(30, 2.0):
            print("[INIT] System Stabilized.", flush=True)
        else:
             print("[INIT] Warning: System not fully stable after init.", flush=True)

    def run(self):
        # Define Points (From CSV + 700 manual)
        points = []
        
        # Load CSV points
        csv_path = r'hysys_automation/optimization_final_result.csv'
        try:
            with open(csv_path, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row['MassFlow']:
                        points.append({
                            'flow': float(row['MassFlow']),
                            'p': float(row['P_bar']),
                            't': float(row['T_C'])
                        })
        except Exception as e:
             print(f"Could not read CSV: {e}")
        
        # Add 700 point if missing
        start_flows = [p['flow'] for p in points]
        if 700.0 not in start_flows:
            points.append({'flow': 700, 'p': 4.0, 't': -112})
            
        # Sort in REVERSE order (1500 -> 500)
        points.sort(key=lambda x: x['flow'], reverse=True)
        
        self.prepare_simulation()
        
        results = []
        print("="*60)
        print(" FINAL VERIFICATION RUN ")
        print("="*60)
        
        for pt in points:
            res = self.verify_point(pt['flow'], pt['p'], pt['t'])
            results.append(res)
            
        # Save Report (CSV Module)
        out_file = r'hysys_automation/final_verification_report.csv'
        fieldnames = list(results[0].keys()) if results else []
        with open(out_file, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
        print(f"\nSaved to {out_file}")
        
        # Print Table (Manual)
        print("\n" + "="*80)
        print(f"{'Flow':<6} {'P_In':<6} {'T_In':<6} | {'Power':<10} {'MA':<6} {'S6_P':<8} | {'Status'}")
        print("-" * 80)
        for r in results:
            status = "OK" if r['Valid_MA'] == 'OK' and r['Valid_Pres'] == 'OK' else "FAIL"
            pwr = r['Power_kW'] if r['Power_kW'] else 0.0
            ma = r['MA_C'] if r['MA_C'] else 0.0
            s6 = r['S6_Pres'] if r['S6_Pres'] else 0.0
            
            print(f"{r['Target_Reliq_Flow']:<6.0f} {r['Input_N2_P']:<6.1f} {r['Input_ADJ4_T']:<6.0f} | {pwr:<10.2f} {ma:<6.2f} {s6:<8.2f} | {status}")
        print("="*80)

if __name__ == "__main__":
    # Popup Killer
    import threading
    def kill_popups():
        import win32gui, win32con
        while True:
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print("[POPUP KILLER] Closed HYSYS popup", flush=True)
            except: pass
            time.sleep(0.5)
    threading.Thread(target=kill_popups, daemon=True).start()
    print("[POPUP KILLER] Started - monitoring for HYSYS popups...\n")
    
    FinalVerifier().run()
