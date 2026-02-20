"""
HYSYS UNIFIED AGENT
-------------------
Combines the robust "Node-Based" architecture with the sophisticated "Anti-Lag/Fine-Tuning" logic.
Features:
- NodeManager (Verified Writes, Emergency Reset)
- Smart Presets (Historical Data)
- Coarse -> Fine Optimization Search
- Anti-Lag Stabilization
"""
import os
import csv
import time
import glob
import win32com.client
import threading
import win32gui, win32con, win32api
from hysys_node_manager import HysysNodeManager

def dismiss_popup():
    """Background thread to kill HYSYS popups"""
    print("[Popup Handler] Active")
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
                time.sleep(1.0)
        except: pass
        time.sleep(1)

# --- Configuration ---
SEARCH_DIR = os.path.dirname(os.path.abspath(__file__))
AG_BEGINNING_DIR = os.path.dirname(SEARCH_DIR)
OUTPUT_FILE = os.path.join(SEARCH_DIR, "optimization_unified_result.csv")

# Optimal Starting Pressures (from History)
PRESET_P = {
    500: 3.0, 600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 7.2, 1500: 7.5
}

# Proven Stable Points (Flow: {Pressure, Temp})
ANCHOR_POINTS = {
    500: {'p': 3.0, 't': -111.0},
    600: {'p': 3.5, 't': -108.0},
    700: {'p': 4.0, 't': -106.0}, 
    800: {'p': 4.3, 't': -104.0},
    900: {'p': 4.8, 't': -102.0},
    1000: {'p': 5.3, 't': -100.0},
    1100: {'p': 5.8, 't': -100.0},
    1200: {'p': 6.2, 't': -98.0},
    1300: {'p': 6.5, 't': -100.0},
    1400: {'p': 7.0, 't': -98.0},
    1500: {'p': 7.4, 't': -98.0}
}
# Full Range
MASS_FLOWS = [500, 600, 700, 800, 900, 1000, 1100, 1200, 1300, 1400, 1500]

class UnifiedAgent:
    def __init__(self):
        self.app = None
        self.mgr = None
        self.results = []
        
    def start(self):
        try:
            self.app = win32com.client.GetActiveObject("HYSYS.Application")
            print("[AGENT] Attached to existing HYSYS.")
        except:
            print("[AGENT] Launching new HYSYS...")
            self.app = win32com.client.Dispatch("HYSYS.Application")
            self.app.Visible = True

    def scan_files(self):
        print(f"[AGENT] Scanning 'Efficiency Increase*.hsc'...")
        pattern = os.path.join(AG_BEGINNING_DIR, "**/Efficiency Increase*.hsc")
        files = glob.glob(pattern, recursive=True)
        return files

    def run(self):
        print("="*60 + "\nUNIFIED AGENT STARTING\n" + "="*60)
        
        # Start Popup Handler
        threading.Thread(target=dismiss_popup, daemon=True).start()
        
        self.start()
        files = self.scan_files()
        if not files:
            print("[AGENT] No files found.")
            return

        for f in files:
            self.optimize_file(f)
            
        if self.mgr:
            self.mgr.close()

    def optimize_file(self, file_path):
        print(f"\n[PROCESSING] {os.path.basename(file_path)}")
        try:
            self.mgr = HysysNodeManager(self.app, file_path)
            
            # Initial Reset to Standard
            if not self.mgr.emergency_reset():
                print("[CRITICAL] Initial Reset Failed.")
                return

            print(f"\n{'MassFlow':<10} {'Pressure':<10} {'Temp':<10} {'App':<10} {'Power':<10} {'Status':<15}")
            print("-" * 70)

            for mf in MASS_FLOWS:
                self.optimize_operating_point(mf)
                
        except Exception as e:
            print(f"[ERROR] File Processing Failed: {e}")
            if self.mgr: self.mgr.close()

    def optimize_operating_point(self, mf):
        """
        Executes the 'Coarse -> Fine' optimization strategy for a given Mass Flow.
        with SAFETY GUARDS (Temp Check, Adjust Auto-Reset).
        """
        # 0. STREAM 1 TEMP GUARD (User Request)
        # Verify valid inlet temp range (0~50). If crazy, force 40.
        try:
            t_in = self.mgr.read('inlet.temperature')
            if t_in < 0 or t_in > 50:
                print(f" [GUARD] Stream 1 Temp {t_in:.1f} is unstable. Resetting to 40C.")
                self.mgr.write('inlet.temperature', 40.0)
                time.sleep(1.0)
        except: pass

        # 1. Determine safe starting Pressure
        center_p = PRESET_P.get(mf, 6.0)
        
        # 2. Reset / Transition
        try:
            if not self.mgr.batch_write({'inlet.mass_flow': mf, 'inlet.pressure': center_p}):
                 self.mgr.emergency_reset()
                 if not self.mgr.batch_write({'inlet.mass_flow': mf, 'inlet.pressure': center_p}):
                     print(f"{mf:<10} {center_p:<10} {'N/A':<10} {'N/A':<10} {'N/A':<10} {'SKIP (Input)':<15}")
                     return
        except Exception as e:
            self.mgr.emergency_reset()
            return

        # 3. Pressure Scan (Center +/- small range)
        p_range = [round(center_p + i*0.1, 1) for i in range(-2, 3)]
        
        best_global = None
        
        for p in p_range:
            try:
                if abs(self.mgr.read('inlet.pressure') - p) > 0.05:
                    if not self.mgr.write('inlet.pressure', p): continue
            except: continue

            # 4. Temperature Search
            best_local = self.search_temperature(mf, p)
            if best_local:
                if not best_global or best_local['power'] < best_global['power']:
                    best_global = best_local

        # 5. Report Best
        if best_global:
            bg = best_global
            print(f"{mf:<10} {bg['p']:<10} {bg['t']:<10} {bg['app']:<10.2f} {bg['power']:<10.1f} {'OPTIMAL':<15}")
            self.save_result(mf, bg)
        else:
            print(f"{mf:<10} {center_p:<10} {'N/A':<10} {'N/A':<10} {'N/A':<10} {'NO SOLUTION':<15}")

    def search_temperature(self, mf, p):
        """
        Coarse Scan -> Anti-Lag -> Fine Scan
        """
        valid_range_ts = []
        coarse_temps = range(-90, -116, -2)
        
        for t in coarse_temps:
            try:
                # ACT (Soft Touch)
                if not self.mgr.write('control.target_temp', t, verify=False):
                    continue
                
                # STABILITY WAIT (Lesson from Final)
                # Give HYSYS a moment to solve before checking blocks
                time.sleep(0.5) 
                
                # RECOVERY CHECK (with Double Reset + Full Restore)
                if not self.mgr.check_blocks():
                     print("R", end="") 
                     # 1. Try simple Adjust Reset first
                     self.mgr.reset_block("Adjust-1")
                     self.mgr.reset_block("Adjust-4")
                     time.sleep(2.0)
                     
                     if not self.mgr.check_blocks():
                         # 2. FULL RESET if simple didn't work
                         # This clears the "stuck" state for next iteration
                         print("X", end="")
                         if not self.recover_simulation_state(mf, p):
                             # If we can't recover to anchor, this P/Flow combo is dead.
                             # Break loop to avoid infinite spin
                             print(" [ABORT SCANS] Recovery Failed")
                             break 
                         continue 

                # OBSERVE (with Anti-Lag)
                state = self.mgr.get_state()
                app = state['min_approach']
                
                if app < 0.5:
                    time.sleep(1.5)
                    state = self.mgr.get_state()
                    app = state['min_approach']
                
                if 0.5 <= app <= 3.5:
                    valid_range_ts.append(t)
                
                if len(valid_range_ts) > 0 and app < 0.5: break
                     
            except Exception as e:
                pass
        
        # B. Fine Scan
        best = None
        if valid_range_ts:
            t_max = max(valid_range_ts) + 1
            t_min = min(valid_range_ts) - 1
            fine_temps = range(t_max, t_min - 1, -1)
            
            for t in fine_temps:
                try:
                    if not self.mgr.write('control.target_temp', t, verify=False): continue
                    
                    time.sleep(0.5) # Stability Wait
                    
                    if not self.mgr.check_blocks():
                        # Recovery
                        self.mgr.reset_block("Adjust-1")
                        self.mgr.reset_block("Adjust-4")
                        time.sleep(1.5)
                        if not self.mgr.check_blocks(): continue

                    if not self.mgr.is_healthy(): continue
                    
                    state = self.mgr.get_state()
                    app = state['min_approach']
                    p7 = state['p7']
                    pwr = state['power']
                    
                    if 2.0 <= app <= 3.0 and p7 <= 36.5:
                        if not best or pwr < best['power']:
                            best = {'p': p, 't': t, 'app': app, 'power': pwr, 'p7': p7}
                except: pass
        
        return best

    def recover_simulation_state(self, mf, p):
        """
        Restores simulation to a PROVEN ANCHOR state.
        Uses historical data to jump to a valid point.
        """
        try:
            # 1. Select Anchor
            if mf in ANCHOR_POINTS:
                anchor = ANCHOR_POINTS[mf]
                t_reset = anchor['t'] + 10.0 # Warm offset for stability
                p_reset = anchor['p']
            else:
                # Fallback Safe
                t_reset = -90.0
                p_reset = PRESET_P.get(mf, 6.0)

            print(f" [RECOVER] Anchor Reset: {mf}kg/h, {p_reset}bar, {t_reset}C")
            
            self.mgr.solver.CanSolve = False
            
            # Reset Inputs
            
            # S1 MASS FLOW KICK
            self.mgr.write('control.s1_mass_flow', 20000.0, verify=False)
            time.sleep(0.5)
            
            updates = {
                'inlet.mass_flow': mf,
                'inlet.pressure': p_reset,
                'inlet.temperature': 40.0,      
                'control.target_temp': t_reset   
            }
            for k, v in updates.items():
                self.mgr.write(k, v, verify=False)
                
            # Reset Adjusts
            self.mgr.reset_block("Adjust-1")
            self.mgr.reset_block("Adjust-4")
            
            self.mgr.solver.CanSolve = True
            
            # stability wait
            self.mgr.wait_stable(timeout=15, stable_duration=2.0)
            
            # Second Reset
            self.mgr.reset_block("Adjust-1")
            self.mgr.reset_block("Adjust-4")
            time.sleep(1.0)
            
            # Final Verify
            if not self.mgr.check_blocks():
                print(" [RECOVER FAIL] Anchor Point Unstable.")
                return False
            
            print(" [RECOVER SUCCESS] Anchor Point Stabilized.")
            return True
        except: return False

    def save_result(self, mf, data):
        """Append result to CSV"""
        file_exists = os.path.isfile(OUTPUT_FILE)
        try:
            with open(OUTPUT_FILE, 'a', newline='') as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["MassFlow", "Pressure", "Temp", "MinApproach", "Power", "P7"])
                writer.writerow([mf, data['p'], data['t'], data['app'], data['power'], data['p7']])
        except Exception as e:
            print(f"[ERROR] Could not save result: {e}")

if __name__ == "__main__":
    agent = UnifiedAgent()
    agent.run()
