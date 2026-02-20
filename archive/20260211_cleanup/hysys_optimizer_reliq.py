import sys
import os
import time
import csv
import threading
from hysys_automation.hysys_node_manager import HysysNodeManager, connect_hysys

# --- Configuration ---
SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_reliq_1300_1500.csv"

# Target Flows (Reliq Flow / Stream 10)
FLOWS = [1300, 1350, 1400, 1450, 1500]

# Optimally calibrated starting pressures (Approximation)
PRESET_P = {
    1300: 6.5, 1350: 6.8, 1400: 7.2, 1450: 7.3, 1500: 7.5
}

class ReliqOptimizer:
    def __init__(self):
        self.app = connect_hysys()
        if not self.app:
            raise RuntimeError("Could not connect to HYSYS")
        
        # Ensure file is open (optional, usually active doc is enough per user workflow)
        # But let's try to be safe if no doc is active
        try:
            self.mgr = HysysNodeManager(self.app)
        except:
             # Try opening default
            full_path = os.path.abspath(os.path.join(FOLDER, SIM_FILE))
            if os.path.exists(full_path):
                 self.mgr = HysysNodeManager(self.app, full_path)
            else:
                 raise

        self._register_extra_nodes()

    def _register_extra_nodes(self):
        """Register additional nodes needed for this specific optimization"""
        # We need specific data points: UA, LMTD, MinApproach
        # Manager already has 'result.min_approach' (LNG-100 MinApproach)
        
        fs = self.mgr.case.Flowsheet
        ops = fs.Operations
        
        # UA and LMTD from LNG-100
        # Check standard property names for Exchanger
        lng = ops.Item("LNG-100")
        self.mgr.register('result.ua', lng, 'UA', unit='kJ/C-h')
        self.mgr.register('result.lmtd', lng, 'LMTD', unit='C')
        
    def hard_reset(self):
        """
        Hard Reset Strategy:
        Spike Stream 1 Mass Flow to 20,000 kg/h to force ADJ-1 to re-evaluate/reset,
        then wait for stability.
        """
        print("   [RESET] Triggering Mass Flow Spike (20,000 kg/h)...")
        # 1. Spike
        self.mgr.write('control.s1_mass_flow', 20000.0, verify=False)
        time.sleep(2.0)
        
        # 2. Wait for ADJ-1 to restore it (approx 29,143 / 39 C)
        # We wait for "Input Temperature" to stabilize around 39-40C 
        # or Flow to stabilize around 29000? 
        # Actually, we just need the system to be 'solvable' again.
        
        print("   [RESET] Waiting for recovery (15s)...")
        self.mgr.wait_stable(timeout=20, stable_duration=2.0)
        
        # Check health
        if self.mgr.is_healthy():
            print("   [RESET] System Recovered.")
            return True
        else:
            print("   [RESET] Recovery Failed (Unstable).")
            return False

    def collect_data(self, flow_set, p_set):
        """Collect all requested data points"""
        return {
            'Reliq_Flow_Set': flow_set,
            'Pressure_Inlet_Set': p_set,
            'Time': time.strftime("%H:%M:%S"),
            'Pressure_Inlet_Act': self.mgr.read('inlet.pressure'),
            'Temp_Inlet': self.mgr.read('inlet.temperature'),
            'Flow_Inlet_Act': self.mgr.read('control.s1_mass_flow'), # Stream 1 Mass Flow
            'Reliq_Flow_Act': self.mgr.read('inlet.mass_flow'), # Stream 10 Mass Flow
            'Power_Comp': self.mgr.read('result.compressor_power'),
            'Min_Approach': self.mgr.read('result.min_approach'),
            'LMTD': self.mgr.read('result.lmtd'),
            'UA': self.mgr.read('result.ua'),
            'Pressure_Outlet': self.mgr.read('outlet.p7'),
            'Status': 'OK' if self.mgr.is_healthy() else 'Unstable'
        }

    def run(self):
        # Init CSV
        headers = [
            'Reliq_Flow_Set', 'Pressure_Inlet_Set', 'Time', 
            'Pressure_Inlet_Act', 'Temp_Inlet', 'Flow_Inlet_Act', 'Reliq_Flow_Act',
            'Power_Comp', 'Min_Approach', 'LMTD', 'UA', 'Pressure_Outlet', 'Status'
        ]
        
        csv_path = os.path.join(FOLDER, OUT_FILE)
        file_exists = os.path.exists(csv_path)
        
        with open(csv_path, 'a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            if not file_exists: writer.writeheader()
            
            for flow in FLOWS:
                print(f"\n>> Optimizing Reliq Flow: {flow} kg/h")
                
                # 1. Set Flow (Stream 10)
                if not self.mgr.write('inlet.mass_flow', flow):
                    print("   [FAIL] Could not set Reliq Flow.")
                    continue
                
                # 2. Pressure Scan
                center_p = PRESET_P.get(flow, 7.0)
                # Range: Center +/- 0.3 bar
                pressures = [round(center_p + i*0.1, 1) for i in range(-3, 4)]
                
                for p in pressures:
                    print(f"   P={p} bar: ", end="", flush=True)
                    
                    # Apply Pressure
                    # If we are far off, maybe do it gradually? 
                    # check current
                    curr_p = self.mgr.read('inlet.pressure')
                    if curr_p and abs(curr_p - p) > 1.0:
                         # Intermediate step
                         self.mgr.write('inlet.pressure', (curr_p+p)/2)
                         time.sleep(1.0)

                    self.mgr.write('inlet.pressure', p)
                    
                    # Wait for stability
                    is_stable = self.mgr.wait_stable(timeout=15, stable_duration=1.0)
                    
                    if not is_stable or not self.mgr.is_healthy():
                        print("UNSTABLE -> RESET", end="")
                        self.hard_reset()
                        # Retry once?
                        print(" -> RETRY", end="")
                        self.mgr.write('inlet.pressure', p)
                        if not self.mgr.wait_stable(timeout=15):
                             print(" -> FAIL")
                             continue
                    
                    # Collect Data
                    data = self.collect_data(flow, p)
                    writer.writerow(data)
                    f.flush() # Ensure write
                    
                    # Feedback
                    print(f" OK (Pwr: {data['Power_Comp']:.1f} kW, MA: {data['Min_Approach']:.2f} C)")

        print("\nOptimization Complete.")
        self.mgr.dispose()

if __name__ == "__main__":
    print("="*60)
    print(" RELIQ FLOW OPTIMIZATION (1300-1500)")
    print("="*60)
    
    # Popup Killer
    def kill_popups():
        import win32gui, win32con, win32api
        while True:
            try:
                hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
                if hwnd:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except: pass
            time.sleep(1)
            
    t = threading.Thread(target=kill_popups, daemon=True)
    t.start()
    
    try:
        opt = ReliqOptimizer()
        opt.run()
    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")
