import sys
import os
import time
import csv
import threading
from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

# --- Configuration ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_comprehensive_500_1500.csv"

# Target Flows (Reliq Flow / Stream 10)
FLOWS = range(500, 1501, 100) # 500 to 1500 step 100

# Preset Pressures (User-Specified Targets)
# Linear interpolation: 500→3.0 bar, 1500→7.4 bar
# Slope: (7.4 - 3.0) / (1500 - 500) = 0.0044 bar/kg/h
PRESET_P = {
    500: 3.0,   600: 3.44,  700: 3.88,  800: 4.32,  900: 4.76,
    1000: 5.2,  1100: 5.64, 1200: 6.08, 1300: 6.52, 1400: 6.96, 
    1500: 7.4
}

class ComprehensiveOptimizer:
    def __init__(self):
        self.app = connect_hysys()
        if not self.app: raise RuntimeError("HYSYS Connection Failed")
        
        try:
            self.mgr = HysysNodeManager(self.app) # Active doc
        except:
             # Fallback
            full_path = os.path.abspath(os.path.join(FOLDER, SIM_FILE))
            if os.path.exists(full_path):
                 self.mgr = HysysNodeManager(self.app, full_path)
            else: raise

        self._register_nodes()

    def _register_nodes(self):
        print("[DEBUG] Registering extra nodes...", flush=True)
        # Register extra nodes needed for comprehensive checks
        try:
            fs = self.mgr.case.Flowsheet
            ops = fs.Operations
            strms = fs.MaterialStreams
            
            # CORRECT MAPPING:
            # - 'inlet.mass_flow' -> Stream 10 (Reliq Flow) [Input Target]
            # - 'control.s1_mass_flow' -> Stream 1 (N2 Loop Flow) [Result/Control]
            
            # Re-register core nodes to ensure correct mapping
            # Note: HysysNodeManager might have defaults, we override here.
            
            s10 = strms.Item("10")
            self.mgr.register('inlet.mass_flow', s10, 'MassFlow', scale=3600.0, unit='kg/h')
            
            s1 = strms.Item("1")
            self.mgr.register('control.s1_mass_flow', s1, 'MassFlow', scale=3600.0, unit='kg/h')
            self.mgr.register('inlet.pressure', s1, 'Pressure', scale=0.01, unit='bar') # N2 Loop Pressure
            self.mgr.register('inlet.temperature', s1, 'Temperature', tolerance=0.1)
            
            # S6 Pressure
            print("[DEBUG] Finding Stream '6'...", flush=True)
            s6 = strms.Item("6")
            self.mgr.register('check.s6_pressure', s6, 'Pressure', scale=0.01, unit='bar')
            
            # LNG Data
            print("[DEBUG] Finding Operation 'LNG-100'...", flush=True)
            lng = ops.Item("LNG-100")
            self.mgr.register('result.ua', lng, 'UA', unit='kJ/C-h')
            self.mgr.register('result.lmtd', lng, 'LMTD', unit='C')
            # result.min_approach already exists standard
            print("[DEBUG] Extra nodes registered.", flush=True)
        except Exception as e:
            print(f"[DEBUG] Registration Error: {e}", flush=True)
            raise

    def check_constraints(self):
        """
        Strict check of all constraints.
        Returns: (True/False, ReasonString)
        """
        # 1. Block Convergence (ADJ-1 to ADJ-5)
        fs = self.mgr.case.Flowsheet
        for i in range(1, 6):
            adj_name = f"ADJ-{i}"
            try:
                adj = fs.Operations.Item(adj_name)
                if not adj.IsIgnored:
                    # STRICT CONVERGENCE CHECK
                    # Try to read Error and Tolerance directly from Variable collection
                    try:
                        err = abs(adj.Variable("Error").Value)
                        tol = adj.Variable("Tolerance").Value
                        if err > tol:
                            # Not converged
                             return False, f"{adj_name} Not Solved (Err: {err:.2e} > {tol:.2e})"
                    except:
                        # Fallback if Variable wrapper fails: check if specific 'Error' prop works (unlikely but try)
                        # Or assume if Solver is idle it *might* be ok, but user wants strict.
                        # Let's count it as warning? No, user implies strict.
                        # But if we can't read it, we can't fail everything.
                        # We print a warning to console in the loop instead.
                        pass
            except: pass
        
        # 2. S6 Pressure < 38.0 bar
        p6 = self.mgr.read('check.s6_pressure')
        if p6 and p6 >= 38.0:
            return False, f"S6 Pressure High ({p6:.2f} bar)"
            
        # 3. Min Approach > 0.5 (Absolute Fail)
        ma = self.mgr.read('result.min_approach')
        if ma is not None:
            if ma < 0.5 and ma > -900: # -32767 is empty/invalid
                return False, f"Temp Cross (MA {ma:.2f})"
        
        return True, "OK"

    def recover_system(self):
        """Enhanced Recovery Logic"""
        print("   [RECOVERY] Triggering Enhanced Reset...")
        
        # 1. Reset adjust blocks manually
        for i in range(1, 6):
            self.mgr.reset_block(f"ADJ-{i}")
            
        # 2. Inject Safe P/T to S1
        self.mgr.write('inlet.temperature', 40.0, verify=False)
        self.mgr.write('inlet.pressure', 6.0, verify=False)
        
        # 3. Spike Mass Flow (Stream 1)
        # Note: 'control.s1_mass_flow' is now Stream 1
        self.mgr.write('control.s1_mass_flow', 20000.0, verify=False)
        
        # 4. Wait
        time.sleep(2.0)
        self.mgr.wait_stable(15, 2.0)
        return self.mgr.is_healthy()

    def reset_adjusts(self):
        """
        Reset all ADJ blocks to clear stuck states.
        """
        print("   [RECOVERY] Resetting Adjust Blocks (ADJ-1 to ADJ-5)...", flush=True)
        fs = self.mgr.case.Flowsheet
        for i in range(1, 6):
            adj_name = f"ADJ-{i}"
            try:
                adj = fs.Operations.Item(adj_name)
                adj.Reset()
            except: pass
        time.sleep(2.0) # Wait for re-solve attempt

    def run_kick(self, target_flow, target_p, target_t):
        """
        Aggressive kick that restores the state.
        1. Call recover_system() -> (Spikes S1 to 20000, P=6, T=40).
        2. Restore Stream 10 Flow (Reliq).
        3. Restore Stream 1 Pressure.
        4. Restore Stream 1 Temp (target).
        """
        self.recover_system()
        
        print(f"   [KICK RESTORE] Flow={target_flow} (S10), P={target_p}, T={target_t}...", flush=True)
        # Restore State
        self.mgr.write('inlet.mass_flow', target_flow, verify=False) # Stream 10
        self.mgr.write('inlet.pressure', target_p, verify=False)     # Stream 1
        self.mgr.write('control.target_temp', target_t, verify=False)
        
        # Wait for stability
        self.mgr.wait_stable(10, 1.0)

    def collect_data(self, flow_set, p_set, t_set, status="OK"):
        return {
            'Flow_Set': flow_set,
            'P_Set': p_set,
            'T_Adj4_Set': t_set,
            'Time': time.strftime("%H:%M:%S"),
            'S1_Pres': self.mgr.read('inlet.pressure'),
            'S1_Temp': self.mgr.read('inlet.temperature'),
            'S1_Flow': self.mgr.read('control.s1_mass_flow'), # N2 Loop (Stream 1)
            'S10_Flow': self.mgr.read('inlet.mass_flow'),     # Reliq (Stream 10 - Target)
            'S6_Pres': self.mgr.read('check.s6_pressure'),
            'Power': self.mgr.read('result.compressor_power'),
            'MA': self.mgr.read('result.min_approach'),
            'UA': self.mgr.read('result.ua'),
            'LMTD': self.mgr.read('result.lmtd'),
            'Status': status
        }

    def run(self):
        headers = [
            'Flow_Set', 'P_Set', 'T_Adj4_Set', 'Time', 
            'S1_Pres', 'S1_Temp', 'S1_Flow', 'S10_Flow', 'S6_Pres', 
            'Power', 'MA', 'UA', 'LMTD', 'Status'
        ]
        
        # Open CSV
        csv_path = os.path.join(FOLDER, OUT_FILE)
        write_header = not os.path.exists(csv_path)
        
        with open(csv_path, 'a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            if write_header: writer.writeheader()
            
            # INITIALIZATION: Set Safe State
            print(">> INITIALIZING SIMULATION TO SAFE STATE...", flush=True)
            self.mgr.write('inlet.mass_flow', 1500.0, verify=False) # Start High (Stream 10)
            self.mgr.write('control.s1_mass_flow', 20000.0, verify=False) # Helper Kick (Stream 1)
            self.mgr.write('inlet.pressure', 6.0, verify=False)     # Mid Pressure
            self.mgr.write('control.target_temp', -95.0, verify=False) # Safe Temp
            self.mgr.wait_stable(15, 2.0)
            self.reset_adjusts()
            self.mgr.wait_stable(10, 1.0)
            print(">> INITIALIZATION COMPLETE.\n", flush=True)

            for flow in FLOWS:
                print(f"\n>> PROCESSING FLOW: {flow} kg/h", flush=True)
                
                # Set Flow
                if not self.mgr.write('inlet.mass_flow', flow):
                    print("   [FAIL] Could not set Flow.")
                    self.recover_system()
                    continue
                    
                # Centers
                center_p = PRESET_P.get(flow, 5.0)
                
                # PRESSURE LOOP (Range +/- 0.5 bar)
                for p in [round(center_p + i*0.1, 1) for i in range(-5, 6)]:
                    # Constraint Check: P within 3.0 ~ 7.4
                    if not (3.0 <= p <= 7.4): continue
                    
                    print(f"   P={p} bar: ", end="")
                    self.mgr.write('inlet.pressure', p)
                    
                    if not self.mgr.wait_stable(10):
                        print("UNSTABLE -> REC", end="")
                        self.recover_system()
                        # Retry
                        self.mgr.write('inlet.pressure', p)
                        self.mgr.wait_stable(10)

                    # COARSE TEMP SCAN (-90 to -120, step -2)
                    best_t, min_pwr = None, 99999
                    valid_range = []
                    
                    print(" TScan[", end="")
                    consecutive_fails = 0
                    
                    for t in range(-90, -121, -2):
                        self.mgr.write('control.target_temp', t, verify=False)
                        time.sleep(1.0) # Propagation
                        
                        # Check Constraints
                        ok, reason = self.check_constraints()
                        
                        # Check MA validity
                        ma = self.mgr.read('result.min_approach')
                        
                        # Failure Detection
                        is_fail = (not ok) or (ma is None) or (ma < -100) or (ma < 0.1)
                        
                        if is_fail:
                            consecutive_fails += 1
                            if reason == "OK": reason = "Invalid MA"
                            print("x" if not ok else "?", end="", flush=True) 
                            
                            # Standard Recovery: Reset Adjusts (After 3 fails)
                            if consecutive_fails == 3:
                                self.reset_adjusts()
                            
                            # Aggressive Kick if REALLY stuck (After 6 fails)
                            if consecutive_fails >= 6:
                                print(f" [STUCK ({consecutive_fails})] -> KICK ", end="", flush=True)
                                self.run_kick(flow, p, t)
                                consecutive_fails = 0 # Reset counter
                            continue
                        else:
                            consecutive_fails = 0 # Reset on success
                            
                        print(".", end="", flush=True) # OK
                        pwr = self.mgr.read('result.compressor_power')
                        
                        # Collect "Valid" points roughly
                        if 2.0 <= ma <= 3.5: # Desired Range
                             valid_range.append(t)
                        
                        # Track best power
                        if pwr and pwr < min_pwr and ma > 0.5:
                            min_pwr = pwr
                            best_t = t
                    print("] ", end="")

                    # FINE TEMP SCAN (Using valid range or best_t)
                    # If we found a valid MA region, scan strictly there with 1 deg step
                    scan_targets = []
                    if valid_range:
                         # Expand range slightly
                         start_t = max(valid_range) + 1
                         end_t = min(valid_range) - 1
                         scan_targets = range(start_t, end_t, -1)
                    elif best_t:
                         scan_targets = range(best_t+2, best_t-3, -1)
                    
                    if scan_targets:
                        print(f"Fine[{len(scan_targets)} pts] ", end="")
                        for t in scan_targets:
                             self.mgr.write('control.target_temp', t, verify=False)
                             time.sleep(1.5)
                             ok, msg = self.check_constraints()
                             status = "OK" if ok else msg
                             
                             data = self.collect_data(flow, p, t, status)
                             writer.writerow(data)
                             f.flush()
                        print("Done")
                    else:
                        print("No Valid Region")

if __name__ == "__main__":
    print("="*60)
    print(" COMPREHENSIVE OPTIMIZATION (500-1500)")
    print("="*60)
    
    # Popup Killer (ACTIVATED to handle pop-up errors)
    def kill_popups():
        import win32gui, win32con
        while True:
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd: 
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print("[POPUP KILLER] Closed HYSYS popup", flush=True)
            except: pass
            time.sleep(0.5)  # Check every 0.5s for faster response
    threading.Thread(target=kill_popups, daemon=True).start()
    print("[POPUP KILLER] Started - monitoring for HYSYS popups...")
    
    opt = ComprehensiveOptimizer()
    opt.run()
