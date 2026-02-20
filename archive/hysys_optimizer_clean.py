"""
HYSYS Agentic Optimizer
Autonomously manages HYSYS instances, scans for files, and optimizes process parameters.
Architecture: Observe -> Decide -> Act
"""
import os
import csv
import time
import glob
import win32com.client
from hysys_node_manager import HysysNodeManager

# Configuration
SEARCH_DIR = os.path.dirname(os.path.abspath(__file__)) # Current dir by default
# Or use the parent directory as requested: "ag-beginning"
AG_BEGINNING_DIR = os.path.dirname(SEARCH_DIR) 

OUTPUT_FILE = os.path.join(SEARCH_DIR, "optimization_agent_results.csv")

class HysysAgent:
    def __init__(self):
        self.app = None
        self.mgr = None
        self.results = []
        
    def start_hysys(self):
        """Ensure HYSYS is running"""
        try:
            self.app = win32com.client.GetActiveObject("HYSYS.Application")
            print("[AGENT] Attached to existing HYSYS instance.")
        except:
            print("[AGENT] Launching new HYSYS instance...")
            self.app = win32com.client.Dispatch("HYSYS.Application")
            self.app.Visible = True

    def scan_files(self):
        """Find target .hsc files"""
        # Looking for files recursively in AG-BEGINNING
        print(f"[AGENT] Scanning for 'Efficiency Increase*.hsc' files in: {AG_BEGINNING_DIR}")
        pattern = os.path.join(AG_BEGINNING_DIR, "**/Efficiency Increase*.hsc")
        files = glob.glob(pattern, recursive=True)
        print(f"[AGENT] Found {len(files)} simulation files.")
        for f in files:
            print(f"  - {os.path.basename(f)}")
        return files

    def run_optimization_cycle(self, file_path):
        """Main agent loop for a single file"""
        print(f"\n[AGENT] PROCESSING: {os.path.basename(file_path)}")
        
        # 1. Load File & Initialize Manager
        try:
            self.mgr = HysysNodeManager(self.app, file_path)
        except Exception as e:
            print(f"[AGENT] Failed to load file: {e}")
            return
            

        # 2. Define Parameter Space (500 ~ 1500 kg/h)
        # 500 and 1500 are fixed points.
        mass_flows = [500, 1300, 1400, 1500] 
        # General search range for non-fixed points (covering intermediate values like 1300/1400)
        pressures = [6.0, 6.2, 6.4, 6.5, 6.6, 6.8, 7.0, 7.2, 7.4, 7.5]
        temperatures = range(-90, -116, -2)
        
        # Fixed Value Lookup (mass_flow -> {pressure, temp_est})
        # Data derived from 'optimization_final_result.csv' (Successful Runs)
        FIXED_POINTS = {
            500:  {'pressure': 3.0, 'temp_est': -111.0}, # Low flow requires much lower pressure
            1500: {'pressure': 7.4, 'temp_est': -98.0}   # High flow safe point
        }

        # Table Header
        print(f"\n{'MassFlow':<10} {'Pressure':<10} {'BestTemp':<10} {'MinApp':<10} {'Power':<10} {'P7':<10} {'Status':<15}")
        print("-" * 80)
        
        # 3. Execution Loop
        for mf in mass_flows:
            
            # Reset to safe base before starting a major branch
            if not self.mgr.emergency_reset():
                print(f"{mf:<10} {'RESET FAIL':<60}")
                break
                
            # Determine Pressure Range
            if mf in FIXED_POINTS:
                # If fixed point, we ONLY run that specific pressure/setup? 
                # User said: "500과 1500은 set-table 기준의 fixed value 사용"
                # Implies we use specific values. Let's assume we test ONLY that point.
                target_p = FIXED_POINTS[mf]['pressure']
                search_ps = [target_p]
            else:
                search_ps = pressures
            
            for p in search_ps:
                # A. Move to Operation Point
                # Try to transition safely
                try:
                    # 1. First establish Mass Flow & Pressure
                    # Note: We might need to adjust "Safe Start" based on target.
                    # For now, trust verified write.
                    
                    if not self.mgr.batch_write({'inlet.mass_flow': mf, 'inlet.pressure': p}):
                        print(f"{mf:<10} {p:<10} {'TRANSITION FAIL':<50}")
                        continue
                        
                except Exception as e:
                     print(f"{mf:<10} {p:<10} {'ERROR':<50}")
                     continue

                # B. Optimize Temperature
                best = None
                status_msg = "NO SOLUTION"
                
                # If Fixed Point, maybe we just verify one temperature?
                # Or optimize around the estimate? Let's optimize around estimate.
                if mf in FIXED_POINTS:
                     t_start = int(FIXED_POINTS[mf]['temp_est'])
                     t_range = range(t_start, t_start-5, -2) # Small search
                else:
                     t_range = temperatures

                for t in t_range:
                    try:
                        # ACT
                        if not self.mgr.write('control.target_temp', t):
                            continue # Write failed
                        
                        # VERIFY BLOCKS (Specific Unit Checks)
                        if not self.mgr.check_blocks():
                             # print(f" [Blocks Unsolved @ {t}C] ", end="")
                             continue

                        # OBSERVE (Global Health)
                        if not self.mgr.is_healthy():
                             # print(f" [Unstable @ {t}C] ", end="")
                             continue
                        
                        # JUDGE
                        state = self.mgr.get_state()
                        app_val = state['min_approach']
                        pwr = state['power']
                        p7 = state['p7']
                        
                        # Logic: 2.0 <= Approach <= 3.0 AND P7 <= 36.5
                        if 2.0 <= app_val <= 3.0 and p7 <= 36.5:
                            if not best or pwr < best['power']:
                                best = {'t': t, 'app': app_val, 'power': pwr, 'p7': p7}
                                status_msg = "OPTIMAL"
                        elif app_val < 0.5:
                            pass # Too close
                        else:
                            pass # Valid but not optimal
                            
                    except Exception as e:
                        # print("E", end="")
                        self.mgr.emergency_reset()
                        self.mgr.batch_write({'inlet.mass_flow': mf, 'inlet.pressure': p})
                
                # RECORD & REPORT
                if best:
                    print(f"{mf:<10} {p:<10} {best['t']:<10} {best['app']:<10.2f} {best['power']:<10.1f} {best['p7']:<10.1f} {status_msg:<15}")
                    self.save_result(os.path.basename(file_path), mf, p, best)
                else:
                     print(f"{mf:<10} {p:<10} {'N/A':<10} {'N/A':<10} {'N/A':<10} {'N/A':<10} {status_msg:<15}")

    def save_result(self, filename, mf, p, data):
        """Append result to CSV"""
        file_exists = os.path.isfile(OUTPUT_FILE)
        with open(OUTPUT_FILE, 'a', newline='') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Filename", "MassFlow", "Pressure", "BestTemp", "MinApproach", "Power", "P7"])
            
            writer.writerow([filename, mf, p, data['t'], data['app'], data['power'], data['p7']])

    def run(self):
        print("="*60)
        print("HYSYS AGENTIC OPTIMIZER STARTING")
        print("="*60)
        
        self.start_hysys()
        files = self.scan_files()
        
        if not files:
            print("[AGENT] No files found to process.")
            return

        for f in files:
            self.run_optimization_cycle(f)
            
        print("\n[AGENT] All tasks completed.")
        if self.mgr:
            self.mgr.close()

if __name__ == "__main__":
    agent = HysysAgent()
    agent.run()
