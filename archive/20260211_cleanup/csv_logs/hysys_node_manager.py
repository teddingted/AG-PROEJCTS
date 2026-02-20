"""
HYSYS Node Manager - Robust Agentic Version
Provides OPC-UA style path-based access to HYSYS COM objects with:
- Automatic File Loading
- Popup Suppression
- Verified Writes (Read-back check)
- Emergency Recovery
- Robust Error Handling for COM Failures
"""
import time
import os
import win32com.client
import pythoncom
from contextlib import contextmanager

class HysysNode:
    """Represents a single addressable node in HYSYS with verification capabilities"""
    def __init__(self, obj, property_name, scale=1.0, unit="", tolerance=0.01):
        self.obj = obj
        self.property_name = property_name
        self.scale = scale
        self.unit = unit
        self.tolerance = tolerance # Tolerance for verification

    def read(self):
        """Read current value with scaling"""
        try:
            prop = getattr(self.obj, self.property_name)
            # Check if property has .Value attribute (COM wrapper) or is direct value
            if hasattr(prop, 'Value'):
                val = prop.Value
            else:
                val = prop
            
            if val is None:
                return None
            return val * self.scale
        except Exception as e:
            # print(f" [READ ERROR] Failed to read {self.property_name}: {e}")
            return None

    def write(self, value):
        """Write value with inverse scaling"""
        try:
            target_val = value / self.scale
            prop = getattr(self.obj, self.property_name)
            
            if hasattr(prop, 'Value'):
                getattr(self.obj, self.property_name).Value = target_val
            else:
                setattr(self.obj, self.property_name, target_val)
                
            return True
        except Exception as e:
            print(f" [WRITE ERROR] Failed to write {value} to {self.property_name}: {e}")
            return False

    def verify(self, target_value):
        """Verify if current value matches target within tolerance"""
        current = self.read()
        if current is None: return False
        
        # simple abs diff
        diff = abs(current - target_value)
        is_ok = diff < self.tolerance
        if not is_ok:
            # print(f" [VERIFY FAIL] Target: {target_value}, Actual: {current}, Diff: {diff}")
            pass
        return is_ok

class HysysNodeManager:
    """
    Manages HYSYS COM objects via path-based addressing.
    Acts as the 'Hardware Abstraction Layer' for the Agent.
    """
    def __init__(self, app=None, case_path=None):
        self.app = app
        self.case = None
        self.solver = None
        self.nodes = {}
        
        if app:
            self._attach_to_app(app, case_path)
            
    def _attach_to_app(self, app, case_path):
        """Attach to HYSYS application and load/get case"""
        self.app = app
        
        # Suppress Popups
        print("[NODE MANAGER] Suppressing Popups (Interactive=False)...")
        try:
            self.app.Interactive = False
        except Exception as e:
            print(f"[NODE MANAGER] Warning: Could not set Interactive=False ({e}). Popups may appear.")
        
        try:
            self.app.ScreenUpdating = True 
        except:
            pass
        
        if case_path:
            abs_path = os.path.abspath(case_path)
            print(f"[NODE MANAGER] Opening Case: {abs_path}")
            try:
                self.case = self.app.SimulationCases.Open(abs_path)
            except Exception as e:
                print(f"[NODE MANAGER] Error opening case: {e}")
                # Fallback to active document if open fails
                try:
                    self.case = self.app.ActiveDocument
                    print("[NODE MANAGER] Fallback: Attached to Active Document")
                except:
                    raise RuntimeError("Could not open case or verify active document.") 
        else:
            print("[NODE MANAGER] Attaching to Active Document")
            try:
                self.case = self.app.ActiveDocument
            except Exception as e:
                 raise RuntimeError(f"Could not get ActiveDocument: {e}")

        if not self.case:
             raise RuntimeError("HYSYS Case object is None after attachment attempts.")

        self.case.Visible = True
        try:
            self.solver = self.case.Solver
        except:
             print("[NODE MANAGER] Warning: Could not get Solver object.")
             self.solver = None

        self._build_standard_nodes()

    def _safe_get_item(self, collection, name):
        """Safely retrieve an item from a COM collection"""
        try:
            return collection.Item(name)
        except:
            return None

    def _build_standard_nodes(self):
        """Register commonly used nodes from current simulation"""
        try:
            fs = self.case.Flowsheet
        except:
            print("[NODE MANAGER] CRITICAL: Could not access Flowsheet.")
            return

        # Helper to register safely
        def safe_reg(path, collection, item_name, prop, **kwargs):
            obj = self._safe_get_item(collection, item_name)
            if obj:
                self.register(path, obj, prop, **kwargs)
            else:
                print(f"[NODE MANAGER] Warning: Node '{item_name}' (for {path}) not found.")

        try:
            # Streams
            strms = fs.MaterialStreams
            safe_reg('inlet.temperature', strms, "1", 'Temperature', tolerance=0.1)
            safe_reg('inlet.pressure', strms, "1", 'Pressure', scale=0.01, unit='bar')  # kPa -> bar
            safe_reg('inlet.mass_flow', strms, "10", 'MassFlow', scale=3600.0, unit='kg/h')  # kg/h
            safe_reg('control.s1_mass_flow', strms, "1", 'MassFlow', scale=3600.0, unit='kg/h') 
            
            safe_reg('outlet.p7', strms, "7", 'Pressure', scale=0.01, unit='bar')
            
            # Operations
            ops = fs.Operations
            safe_reg('control.vol_flow', ops, "ADJ-1", 'TargetValue', scale=3600.0, unit='m3/h')  
            safe_reg('control.target_temp', ops, "ADJ-4", 'TargetValue', unit='C', tolerance=0.5)
            safe_reg('result.min_approach', ops, "LNG-100", 'MinApproach', unit='C')
            safe_reg('result.compressor_power', ops, "K-100", 'EnergyValue', unit='kW')
            
            # Spreadsheet Cell
            ss = self._safe_get_item(ops, "SPRDSHT-1")
            if ss:
                try:
                    pwr_cell = ss.Cell("C8")
                    self.register('result.total_power', pwr_cell, 'CellValue', unit='kW')
                except:
                    print("[NODE MANAGER] Warning: Could not access SPRDSHT-1 Cell C8")

            print(f"[NODE MANAGER] Registered {len(self.nodes)} nodes.", flush=True)
            if not self.nodes:
                print("[NODE MANAGER] WARNING: No nodes were registered! Check HYSYS connection/file.")
        except Exception as e:
            print(f"[NODE MANAGER] Error building nodes: {e}")
            # Do not raise, allow partial functionality

    def register_node(self, path, category, name, prop, scale=1.0, unit="", tolerance=0.01):
        """Dynamic registration by category and name"""
        try:
            fs = self.case.Flowsheet
            obj = None
            
            if category == 'Stream': 
                # Try Material Stream First
                try: obj = fs.MaterialStreams.Item(name)
                except: pass
                # Try Energy Stream Second
                if not obj:
                    try: obj = fs.EnergyStreams.Item(name)
                    except: pass
            elif category in ['Operation', 'Block']: 
                try: obj = fs.Operations.Item(name)
                except: pass
            
            if obj:
                self.register(path, obj, prop, scale, unit, tolerance)
                return True
        except: pass
        print(f"[REGISTER ERROR] Could not register {path} ({category}: {name})")
        return False

    def register(self, path, obj, property_name, scale=1.0, unit="", tolerance=0.01):
        """Register a new node for path-based access"""
        self.nodes[path] = HysysNode(obj, property_name, scale, unit, tolerance)
    
    def read(self, path):
        """Read value by node path"""
        if path not in self.nodes:
            # raise KeyError(f"Node '{path}' not registered")
            print(f" [READ] Warning: Node '{path}' not registered")
            return None
        return self.nodes[path].read()
    
    def write(self, path, value, verify=True, max_retries=3):
        """
        Robust Write: Write -> Wait -> Verify
        Returns True if successful, False if failed.
        """
        if path not in self.nodes:
            print(f" [WRITE] Error: Node '{path}' not registered")
            return False
        
        node = self.nodes[path]
        
        for attempt in range(max_retries):
            # Write
            print(f"[DEBUG MSG] Writing {path} = {value}...", flush=True)
            if not node.write(value):
                print(f" [WRITE FAIL] Attempt {attempt+1}/{max_retries} for {path}", flush=True)
                time.sleep(0.5)
                continue
            
            if not verify:
                return True
                
            # Wait & Verify
            time.sleep(0.5) # Minimum settling time
            if node.verify(value):
                return True
            else:
                # print(f" [VERIFY RETRY] Attempt {attempt+1}/{max_retries} for {path}")
                pass
        
        print(f" [WRITE FINAL FAIL] Could not verify write to {path}", flush=True)
        return False

    def batch_write(self, updates):
        """
        Atomic multi-node update with solver freeze.
        ALL writes must succeed for function to return True.
        """
        # print(f"[DEBUG] Batch write: {list(updates.keys())}", flush=True)
        all_ok = True
        with self._frozen():
            for path, value in updates.items():
                # Verify ONLY for inputs that don't depend on solver
                # Use verify=True generally but handle failures gracefully
                if not self.write(path, value, verify=True): 
                    all_ok = False
        
        # Wait for solver to process changes
        self.wait_stable()
        return all_ok
    
    
    def emergency_reset(self):
        """
        HARD RESET to safe known state.
        Used when simulation diverges or becomes unstable.
        """
        print("\n[!!! EMERGENCY RESET TRIGGERED !!!]")
        
        try:
            # 1. Force Solver OFF
            if self.solver: self.solver.CanSolve = False
            time.sleep(0.5)
            
            # 2. Reset to Safe Values
            safe_state = {
                'inlet.mass_flow': 15000, # User requested high value
                'inlet.pressure': 7.4,
                'inlet.temperature': 40.0,
                'control.target_temp': -90.0
            }
            
            print(" [RESET] Applying Safe State...")
            
            # Try setting stream 1 flow if registered (sometimes needed to unstick)
            if 'control.s1_mass_flow' in self.nodes:
                 self.write('control.s1_mass_flow', 20000.0, verify=False)
                 time.sleep(0.5)
    
            for path, val in safe_state.items():
                if path in self.nodes:
                    self.write(path, val, verify=False) 
            
            # Reset ALL Adjusts (ADJ-1 to ADJ-4)
            print(" [RESET] Resetting Adjusts...")
            for i in range(1, 5):
                 self.reset_block(f"ADJ-{i}")
    
            # 3. Force Solver ON
            print(" [RESET] Re-enabling Solver...")
            if self.solver: self.solver.CanSolve = True
            
            # 4. Wait for stability
            if self.wait_stable(timeout=60, stable_duration=5.0):
                print(" [RESET SUCCESS] System Stabilized.")
                return True
            else:
                print(" [RESET FAIL] System did not stabilize.")
                return False
        except Exception as e:
            print(f" [RESET CRITICAL ERROR] {e}")
            return False

    def reset_block(self, name):
        """Attempts to reset a specific operation"""
        try:
            fs = self.case.Flowsheet
            op = self._safe_get_item(fs.Operations, name)
            if op and hasattr(op, 'Reset'):
                op.Reset()
                # print(f" [RECOVERY] Reset {name}")
                return True
        except:
            return False

    def check_blocks(self):
        """
        Check if critical blocks are healthy.
        Returns: True if all healthy, False if any failed.
        """
        if not self.case: return False
        
        fs = self.case.Flowsheet
        ops = fs.Operations
        
        checks = {
            'Compressor (K-100)': "K-100",
            'Adjust-1': "ADJ-1",
            'Adjust-4': "ADJ-4",
            'LNG Exchanger': "LNG-100"
        }
        
        all_ok = True
        
        for name, item_name in checks.items():
            try:
                obj = self._safe_get_item(ops, item_name)
                if not obj: continue # Skip if missing

                # 1. Ignored Check
                try:
                    if getattr(obj, 'Ignored', False):
                        continue 
                except: pass

                # 2. Operational Checks
                if "Compressor" in name:
                     try:
                         if obj.EnergyValue <= 1e-6:
                             print(f" [BLOCK FAIL] {name} (Zero Power)")
                             all_ok = False
                     except: pass

                elif "LNG" in name:
                     try:
                         val = obj.MinApproach.Value
                         if val is None or val < -900: 
                             print(f" [BLOCK FAIL] {name} (Unsolved)")
                             all_ok = False
                     except: pass
                        
            except Exception as e:
                print(f" [BLOCK ERROR] {name}: {e}")
                all_ok = False
        
        return all_ok


    @contextmanager
    def _frozen(self):
        """Context manager to freeze/unfreeze solver"""
        if not self.solver:
            yield
            return

        was_solving = self.solver.CanSolve
        self.solver.CanSolve = False
        try:
            yield
        finally:
            self.solver.CanSolve = was_solving
            time.sleep(0.5) # Allow solver to kick in
    
    def wait_stable(self, timeout=30, stable_duration=1.0):
        """
        Smart wait for simulation to stabilize.
        Checks:
        1. Solver.IsSolving
        2. Key process variables (Temperature) stability
        """
        if not self.solver: return False
        
        time.sleep(0.5)
        t0 = time.time()
        
        # 1. Wait for Solver Engine
        # Add safeguard against infinite loop if IsSolving stays True
        while True:
            try:
                if not self.solver.IsSolving:
                    break
            except:
                break # If COM fails, assume stopped or crashed
            
            if time.time() - t0 > timeout:
                print(" [SOLVER TIMEOUT]", end="", flush=True)
                return False
            time.sleep(0.2)
        
        # 2. Wait for Process Value Stability (Temperature)
        ref_t = None
        last_temp = -9999
        
        while time.time() - t0 < timeout:
            try:
                temp = self.read('inlet.temperature')
                if temp is None:
                    # If read fails, maybe simulation is busy or closed.
                    # Wait a bit and retry logic, or fail?
                    time.sleep(0.5)
                    continue 
            except:
                return False
                
            if abs(temp - last_temp) < 0.05: # Stricter stability check
                if not ref_t:
                    ref_t = time.time()
                elif time.time() - ref_t >= stable_duration:
                    return True # Stable for duration
            else:
                ref_t = None # Reset timer if changed
            
            last_temp = temp
            time.sleep(0.2)
        
        print(" [STABILITY TIMEOUT]", end="", flush=True)
        return False
    
    def is_healthy(self):
        """Check if simulation is in valid/converged state"""
        try:
            temp = self.read('inlet.temperature')
            power = self.read('result.compressor_power')
            # Basic sanity checks
            if temp is None or power is None: return False
            if temp < -273 or temp > 500: return False
            if power < 0.1: return False
            return True
        except:
            return False
            
    def get_state(self):
        """Get current operating state as dict"""
        return {
            'mass_flow': self.read('inlet.mass_flow'),
            'pressure': self.read('inlet.pressure'),
            'temperature': self.read('inlet.temperature'),
            'min_approach': self.read('result.min_approach'),
            'power': self.read('result.total_power'),
            'p7': self.read('outlet.p7')
        }

    def dispose(self):
        """Explicit cleanup of COM objects"""
        print("[NODE MANAGER] Disposing Objects...", end=" ")
        try:
            if self.app:
                # self.app.Interactive = True # Restore interactive mode
                pass
        except: pass
        
        self.nodes = {}
        self.solver = None
        self.case = None
        self.app = None
        print("Done.")

def connect_hysys():
    """Robust connection to HYSYS Application"""
    print("[CONNECT] Attempting to connect to HYSYS...", flush=True)
    app = None
    try:
        # Try to attach to existing active object
        app = win32com.client.GetActiveObject("HYSYS.Application")
        print("[CONNECT] Attached to Active HYSYS Instance.", flush=True)
    except Exception:
        # Fallback: Create new instance (or attach via Dispatch if registered but not active in ROT)
        print("[CONNECT] Active object not found. Trying Dispatch...", flush=True)
        try:
            app = win32com.client.Dispatch("HYSYS.Application")
            print("[CONNECT] Dispatched HYSYS Application.")
        except Exception as e:
            print(f"[CONNECT] CRITICAL ERROR: Could not dispatch HYSYS. {e}")
            return None
    
    
    # CRITICAL: Ensure Visible so it doesn't vanish if script ends or fails
    try:
        app.Visible = True
    except:
        print("[CONNECT] Warning: Could not set Visible=True")

    return app

# Usage Example
if __name__ == "__main__":
    # Test Connect
    print("Testing ROBUST Node Manager...")
    try:
        app = connect_hysys()
        if app:
            mgr = HysysNodeManager(app)
            
            t = mgr.read('inlet.temperature')
            print(f"Current Temp: {t}")
            
            # Test Verified Write if possible
            # mgr.write('inlet.mass_flow', 1500)
            
            # mgr.emergency_reset()
            
            mgr.dispose()
        else:
            print("Could not connect.")

    except Exception as e:
        print(f"Test Failed: {e}")
