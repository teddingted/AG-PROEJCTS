import os, time, csv, threading, json, signal
import win32com.client
import win32gui, win32con
import pywintypes
from datetime import datetime
import logging

"""
HYSYS OPTIMIZER - DISPATCH EDITION (Robust Background Execution)
-----------------------------------------------------------------
Features:
- Independent HYSYS process (Dispatch instead of GetActiveObject)
- Automatic crash recovery with retry logic
- Checkpoint/resume system for interruption handling
- Timeout detection to prevent infinite hangs
- Background execution (Visible=False)
- Comprehensive error logging
"""

SIM_FILE = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
FOLDER = "hysys_automation"
OUT_FILE = "optimization_2d_extended.csv"
CHECKPOINT_FILE = "hysys_automation/optimization_checkpoint.json"
LOG_FILE = f"hysys_automation/logs/optimizer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
MODEL_DATA_FILE = "hysys_automation/optimization_final_summary_verified.csv"

# Configuration
FLOWS = [1200, 1300]
VOL_FLOWS = range(3500, 3801, 50)

DISPATCH_CONFIG = {
    'visible': False,              # Background mode
    'timeout_per_point': 300,      # 5 min max per optimization point
    'max_retries': 3,              # Retry failed points 3 times
    'checkpoint_interval': 1,      # Save progress after each point
    'cleanup_on_exit': True,       # Kill HYSYS when done
    'restart_every_n': 20,         # Restart HYSYS every N points (prevent memory leaks)
}

ANCHOR_POINTS = {
    500: {'P': 3.00, 'T': -111.0},
    1500: {'P': 7.40, 'T': -99.0}
}

PRESET_P = {
    600: 3.5, 700: 4.0, 800: 4.4,
    900: 4.8, 1000: 5.2, 1100: 5.6, 1200: 6.0,
    1300: 6.5, 1400: 6.81
}

# Setup logging
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class TimeoutError(Exception):
    """Raised when operation exceeds timeout"""
    pass

class CheckpointManager:
    """Manages optimization progress checkpoints"""
    
    def __init__(self, checkpoint_file):
        self.file = checkpoint_file
        self.data = self.load()
    
    def load(self):
        """Load checkpoint from file"""
        if os.path.exists(self.file):
            try:
                with open(self.file, 'r') as f:
                    return json.load(f)
            except:
                logger.warning(f"Failed to load checkpoint, starting fresh")
        return {'completed': [], 'last_update': None}
    
    def save(self):
        """Save checkpoint to file"""
        self.data['last_update'] = datetime.now().isoformat()
        with open(self.file, 'w') as f:
            json.dump(self.data, f, indent=2)
    
    def is_completed(self, flow, vol):
        """Check if a point has been processed"""
        return [flow, vol] in self.data['completed']
    
    def mark_completed(self, flow, vol):
        """Mark a point as completed"""
        if [flow, vol] not in self.data['completed']:
            self.data['completed'].append([flow, vol])
            self.save()
            logger.info(f"Checkpoint: Marked ({flow}, {vol}) as completed")

class HysysProcessManager:
    """Manages HYSYS process lifecycle with crash recovery"""
    
    def __init__(self, config):
        self.config = config
        self.app = None
        self.case = None
        self.process_count = 0
        
    def start_hysys(self):
        """Start a new HYSYS instance"""
        logger.info("Starting new HYSYS instance (Dispatch mode)...")
        try:
            self.app = win32com.client.Dispatch("HYSYS.Application")
            self.app.Visible = self.config['visible']
            
            sim_path = os.path.abspath(os.path.join(os.getcwd(), FOLDER, SIM_FILE))
            logger.info(f"Opening case: {sim_path}")
            
            self.case = self.app.SimulationCases.Open(sim_path)
            logger.info(f"Active Case: {self.case.Title.Value}")
            
            self.process_count += 1
            return True
        except Exception as e:
            logger.error(f"Failed to start HYSYS: {e}")
            return False
    
    def stop_hysys(self):
        """Gracefully close HYSYS"""
        logger.info("Stopping HYSYS instance...")
        try:
            if self.case:
                self.case.Close()
            if self.app:
                self.app.Quit()
        except:
            logger.warning("Graceful close failed, forcing termination")
            self._force_kill_hysys()
        finally:
            self.app = None
            self.case = None
    
    def restart_hysys(self):
        """Restart HYSYS (for crash recovery or periodic refresh)"""
        logger.info("Restarting HYSYS...")
        self.stop_hysys()
        time.sleep(3)
        return self.start_hysys()
    
    def _force_kill_hysys(self):
        """Force kill HYSYS processes"""
        try:
            os.system("taskkill /F /IM hysys.exe 2>nul")
            time.sleep(2)
        except:
            pass
    
    def should_restart(self):
        """Check if HYSYS should be restarted (prevent memory leaks)"""
        if self.config['restart_every_n'] > 0:
            return self.process_count % self.config['restart_every_n'] == 0
        return False

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
            logger.info(f"[MODEL] Loaded {len(self.data)} points.")

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
    def __init__(self, process_mgr):
        self.process_mgr = process_mgr
        self.case = process_mgr.case
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
    def __init__(self, manager, model, process_mgr, checkpoint_mgr):
        self.mgr = manager
        self.model = model
        self.process_mgr = process_mgr
        self.checkpoint = checkpoint_mgr
        self.solver = manager.solver
        self.timeout = DISPATCH_CONFIG['timeout_per_point']
        self.max_retries = DISPATCH_CONFIG['max_retries']

    def execute_with_timeout(self, func, *args, **kwargs):
        """Execute function with timeout"""
        result = [None]
        exception = [None]
        
        def target():
            try:
                result[0] = func(*args, **kwargs)
            except Exception as e:
                exception[0] = e
        
        thread = threading.Thread(target=target)
        thread.daemon = True
        thread.start()
        thread.join(timeout=self.timeout)
        
        if thread.is_alive():
            logger.error(f"Timeout exceeded ({self.timeout}s)")
            raise TimeoutError(f"Operation timed out after {self.timeout}s")
        
        if exception[0]:
            raise exception[0]
        
        return result[0]

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
            
            self.mgr.nodes['S10'].MassFlow.Value = flow / 3600.0
            vol_s = vol_flow / 3600.0
            logger.debug(f"[RESET] Vol={vol_flow} m3/h -> {vol_s:.4f} m3/s")
            self.mgr.nodes['ADJ1'].TargetValue.Value = vol_s
            
            self.mgr.nodes['S1'].Pressure.Value = p_bar * 100.0
            self.mgr.nodes['S1'].Temperature.Value = 40.0
            self.mgr.nodes['ADJ4'].TargetValue.Value = -90.0
            
            for adj in self.mgr.adjs: adj.Reset()
            self.solver.CanSolve = True
            self.wait_stable(20, 1.0)
            
            for adj in self.mgr.adjs: adj.Reset()
            self.wait_stable(10, 1.0)
            return True
        except Exception as e:
            logger.error(f"Recover state failed: {e}")
            return False

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
        except Exception as e:
            logger.error(f"Set inputs failed: {e}")
            return False

    def get_metrics(self):
        try:
            ma = self.mgr.nodes['LNG'].MinApproach.Value
            if ma < 0.5:
                time.sleep(1.0)
                ma = self.mgr.nodes['LNG'].MinApproach.Value
            
            base_metrics = {
                'MA': ma,
                'Power': self.mgr.nodes['Spreadsheet'].Cell("C8").CellValue,
                'S6_Pres': self.mgr.nodes['S6'].Pressure.Value / 100.0,
                'S7_Pres': self.mgr.nodes['S7'].Pressure.Value / 100.0
            }
            
            ss = self.mgr.nodes['Spreadsheet']
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
            except Exception as e:
                logger.warning(f"Failed to read extended metrics: {e}")
                
            return base_metrics
        except Exception as e:
            logger.error(f"Get metrics failed: {e}")
            return None

    def strategy_grid_scan(self, flow, vol_flow, p_center):
        logger.info(f"[GRID] P={p_center}...")
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
        
        if best:
            logger.info(f"[GRID] OK ({best['Power']:.1f} kW)")
        else:
            logger.warning("[GRID] FAIL")
        return best

    def optimize_point(self, flow, vol):
        """Optimize single point with retry logic"""
        for attempt in range(self.max_retries):
            try:
                logger.info(f"Processing ({flow}, {vol}) - Attempt {attempt + 1}/{self.max_retries}")
                
                # Use Grid Scan for high flows
                p_center = PRESET_P.get(flow, 6.0)
                p_pred, _ = self.model.predict(flow)
                if p_pred: p_center = p_pred
                
                result = self.execute_with_timeout(self.strategy_grid_scan, flow, vol, p_center)
                
                if result:
                    logger.info(f"✓ Success: ({flow}, {vol}) -> P={result['P']}, Power={result['Power']:.1f}")
                    return result, "Grid"
                else:
                    logger.warning(f"✗ Failed: ({flow}, {vol}) - No valid solution")
                    
            except TimeoutError:
                logger.error(f"Timeout on ({flow}, {vol}) - Restarting HYSYS...")
                self.process_mgr.restart_hysys()
                self.mgr = HysysNodeManager(self.process_mgr)
                self.solver = self.mgr.solver
                
            except pywintypes.com_error as e:
                logger.error(f"COM Error on ({flow}, {vol}): {e} - Restarting HYSYS...")
                self.process_mgr.restart_hysys()
                self.mgr = HysysNodeManager(self.process_mgr)
                self.solver = self.mgr.solver
                
            except Exception as e:
                logger.error(f"Unexpected error on ({flow}, {vol}): {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(5)
                    
        logger.error(f"✗✗ Failed after {self.max_retries} attempts: ({flow}, {vol})")
        return None, "Failed"

    def run(self):
        logger.info("="*60)
        logger.info(">> DISPATCH-BASED OPTIMIZER STARTED")
        logger.info("="*60)
        
        keys = ['Flow', 'VolFlow', 'P_bar', 'T_C', 'MA', 'Power', 'S6_Pres', 'Method']
        extended_keys = ['SS_MassFlow', 'SS_VolFlow', 'SS_SuctionP', 'SS_DischargeP', 
                         'SS_Ratio', 'SS_Duty_kW', 'SS_LMTD', 'SS_MinAppr', 'SS_UA_kWC', 'SS_ExpInletT']
        keys.extend(extended_keys)
        
        output_path = f'{FOLDER}/{OUT_FILE}'
        if not os.path.exists(output_path):
            with open(output_path, 'w', newline='') as f:
                csv.DictWriter(f, fieldnames=keys).writeheader()

        total_points = len(FLOWS) * len(VOL_FLOWS)
        completed = 0
        
        for flow in FLOWS:
            logger.info(f"\n{'='*60}\nMASS FLOW: {flow} kg/h\n{'='*60}")
            
            for vol in VOL_FLOWS:
                if self.checkpoint.is_completed(flow, vol):
                    logger.info(f"Skipping ({flow}, {vol}) - Already completed")
                    completed += 1
                    continue

                logger.info(f"Processing ({flow}, {vol}) [{completed+1}/{total_points}]")
                
                res, method = self.optimize_point(flow, vol)

                if res:
                    row = {
                        'Flow': flow, 'VolFlow': vol, 
                        'P_bar': res['P'], 'T_C': res['T'],
                        'MA': res['MA'], 'Power': res['Power'], 
                        'S6_Pres': res.get('S6_Pres', 0), 'Method': method
                    }
                    for k in extended_keys:
                        row[k] = res.get(k, '')
                        
                    with open(output_path, 'a', newline='') as f:
                        csv.DictWriter(f, fieldnames=keys).writerow(row)
                    
                    self.checkpoint.mark_completed(flow, vol)
                    completed += 1
                    logger.info(f"Progress: {completed}/{total_points} ({100*completed/total_points:.1f}%)")
                
                # Periodic HYSYS restart
                if self.process_mgr.should_restart():
                    logger.info("Performing periodic HYSYS restart...")
                    self.process_mgr.restart_hysys()
                    self.mgr = HysysNodeManager(self.process_mgr)
                    self.solver = self.mgr.solver
        
        logger.info("="*60)
        logger.info(f"OPTIMIZATION COMPLETE: {completed}/{total_points} points")
        logger.info("="*60)

def main():
    process_mgr = None
    try:
        logger.info("Initializing Dispatch-based optimizer...")
        
        process_mgr = HysysProcessManager(DISPATCH_CONFIG)
        if not process_mgr.start_hysys():
            logger.error("Failed to start HYSYS. Exiting.")
            return
        
        checkpoint_mgr = CheckpointManager(CHECKPOINT_FILE)
        model = SurrogateModel(MODEL_DATA_FILE)
        mgr = HysysNodeManager(process_mgr)
        opt = Optimizer2D(mgr, model, process_mgr, checkpoint_mgr)
        
        opt.run()
        
    except KeyboardInterrupt:
        logger.info("Interrupted by user. Saving checkpoint...")
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
    finally:
        if process_mgr and DISPATCH_CONFIG['cleanup_on_exit']:
            process_mgr.stop_hysys()
        logger.info("Optimizer terminated.")

if __name__ == "__main__":
    main()
