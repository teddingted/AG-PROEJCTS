import os
import win32com.client
import threading
import time
import win32gui
import win32con
import win32api

def dismiss_popup():
    """
    Background thread to detect and dismiss the 'Aspen HYSYS' popup 
    that appears when opening a v14 case in v15.
    """
    print("[Popup Handler] Started...")
    # Try for a limited time (e.g., 30 seconds)
    max_retries = 30
    found = False
    
    for _ in range(max_retries):
        try:
            # Find window by class and title
            # #32770 is the standard class for Windows dialogs
            hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
            
            if hwnd:
                print(f"[Popup Handler] Popup found! Handle: {hwnd}")
                
                # Make sure it's not minimized and bring to front
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except:
                    pass
                
                time.sleep(0.5)
                
                # Send 'Enter' key to click the default button (usually OK/Yes)
                print("[Popup Handler] Sending ENTER key...")
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                
                found = True
                break
        except Exception as e:
            print(f"[Popup Handler] Error checking window: {e}")
            
        time.sleep(1)
        
    if found:
        print("[Popup Handler] Popup dismissed successfully.")
    else:
        print("[Popup Handler] No popup detected within timeout.")

def get_spreadsheet_data():
    # 1. Define File Path - Using absolute path
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))

    print(f"Target File: {file_path}")
    
    # 2. Connect to HYSYS
    print("Connecting to HYSYS Application...")
    try:
        app = win32com.client.Dispatch("HYSYS.Application")
        # specific for v15 popup handling: ensure app is visible so dialogs can be interacted with
        app.Visible = True
    except Exception as e:
        print(f"Failed to connect to HYSYS Application: {e}")
        return

    # 3. Open Case
    case = None
    try:
        if os.path.exists(file_path):
            print("File found. Opening case...")
            
            # Start the popup handler in the background
            t = threading.Thread(target=dismiss_popup)
            t.daemon = True # Daemon thread handles exit if main program finishes
            t.start()
            
            case = app.SimulationCases.Open(file_path)
            # Ensure it is visible
            case.Visible = True
        else:
            print("File not found at constructed path.")
            print("Attempting to use currently active case...")
            case = app.ActiveDocument
    except Exception as e:
        print(f"Error opening case: {e}")
        try:
            case = app.ActiveDocument
        except:
            pass
            
    if case is None:
        print("Could not access any HYSYS case.")
        return

    print(f"Case Open: {case.Name}")

    # 4. Access Spreadsheet
    ss_name = "SPRDSHT-1"
    ss = None
    try:
        ss = case.Flowsheet.Operations.Item(ss_name)
    except Exception as e:
        print(f"Could not find operation '{ss_name}': {e}")
        return


    print(f"Accessed Spreadsheet: {ss_name}")

    print(f"Accessed Spreadsheet: {ss_name}")

    # 6. Main Read Logic
    headers = []
    values = []
    units = []
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'] # 0-9


    def convert_to_display_unit(raw_val, unit_str):
        """
        Convert HYSYS internal units to the target display units.
        Factors derived from observation of Internal vs Expected.
        """
        if raw_val is None:
            return None
        
        try:
            val = float(raw_val)
        except:
            return raw_val

        # Conversion Rules
        if unit_str == 'kg/h':
            # Internal (kg/s) -> kg/h
            return val * 3600.0
        elif unit_str == 'm3/h':
            # Internal (m3/s) -> m3/h
            return val * 3600.0
        elif unit_str == 'bar':
            # Internal (kPa) -> bar
            return val / 100.0
        elif unit_str == 'W/C':
             # Internal (kW/C or similar scaled) -> W/C
             # Observed: 157.14 -> 157141 => * 1000
             return val * 1000.0
        
        # Default: No conversion (Temp C, Duty kW seem to match according to user)
        return val

    def get_cell_data(cell_addr_str):
        """
        Get value and unit from a cell.
        Returns: (value_str, unit_str)
        """
        val_str = "Error"
        unit_str = ""
        
        try:
            cell = ss.Cell(cell_addr_str)
        except Exception as e:
            return f"Err(GetCell): {e}", ""

        # 1. Get Unit first (needed for conversion)
        try:
             if hasattr(cell, 'Units'):
                 unit_str = str(cell.Units)
        except:
             pass

        # 2. Get Value
        raw_val = None
        try:
            raw_val = cell.CellValue
        except:
            pass
            
        if raw_val is None:
            try:
                raw_val = cell.CellText
            except:
                pass

        # 3. Convert and Format
        converted_val = convert_to_display_unit(raw_val, unit_str)
        
        if isinstance(converted_val, (int, float)):
            val_str = f"{converted_val:.2f}"
        elif converted_val is None:
            val_str = ""
        else:
            val_str = str(converted_val)
            
        return val_str, unit_str

    for i, col_name in enumerate(cols):
        # Row 11 (Header/Parameter Name) - Usually just text, no unit
        h_addr = f"{col_name}11"
        h_val, _ = get_cell_data(h_addr)
        # Force string for header if it looks numeric
        
        # Row 12 (Value) - Has value and potential unit
        v_addr = f"{col_name}12"
        v_val, v_unit = get_cell_data(v_addr)
        
        headers.append(h_val)
        values.append(v_val)
        units.append(v_unit)

    # 7. Print Table
    print("\n" + "="*100)
    print(f"{'Col':<5} | {'Row 11 (Parameter)':<40} | {'Row 12 (Value)':<15} | {'Unit':<15}")
    print("-" * 100)
    for i, col in enumerate(cols):
        h = headers[i]
        v = values[i]
        u = units[i]
        
        # Clean NoneStr
        if h == "None": h = ""
        if v == "None": v = ""
        
        print(f"{col:<5} | {h:<40} | {v:<15} | {u:<15}")
    print("="*100 + "\n")



if __name__ == "__main__":
    get_spreadsheet_data()
