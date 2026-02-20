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

    # 6. Main Read Logic
    headers = []
    values = []
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'] # 0-9

    def get_cell_value_safe(cell_addr_str, cell_obj=None):
        """
        Try multiple ways to get data from a cell.
        HYSYS cells can behave differently depending on content (Variable vs Text).
        """
        # If we don't have the object yet, try to get it
        if cell_obj is None:
            try:
                cell_obj = ss.Cell(cell_addr_str)
            except Exception as e:
                return f"Err(GetCell): {e}"

        # Try getting values from properties
        # Order: CellValue -> CellText -> Formula -> Variable.Name (if connected)
        
        # 1. CellValue (Standard)
        try:
            return str(cell_obj.CellValue)
        except:
            pass
            
        # 2. CellText (Better for text labels)
        try:
            return str(cell_obj.CellText)
        except:
            pass
            
        # 3. Label (Sometimes used)
        try:
            return str(cell_obj.Label)
        except:
            pass

        # 4. Formula (If it's just a raw number/string entered as formula)
        try:
            return str(cell_obj.Formula)
        except:
            pass

        return "Err(NoVal)"

    for i, col_name in enumerate(cols):
        # Row 11 (Header/Parameter Name)
        # Using string address 'A11' etc. as (int, int) failed in tests
        h_addr = f"{col_name}11"
        h_val = get_cell_value_safe(h_addr)
        
        # Row 12 (Value)
        v_addr = f"{col_name}12"
        v_val = get_cell_value_safe(v_addr)
        
        headers.append(h_val)
        values.append(v_val)

    # 7. Print Table
    print("\n" + "="*80)
    print(f"{'Col':<5} | {'Row 11 (Parameter)':<40} | {'Row 12 (Value)':<20}")
    print("-" * 80)
    for i, col in enumerate(cols):
        h = headers[i]
        v = values[i]
        # Clean NoneStr
        if h == "None": h = ""
        if v == "None": v = ""
        
        print(f"{col:<5} | {h:<40} | {v:<20}")
    print("="*80 + "\n")

    # Keep the app open for a moment or close it? 
    # Usually better to keep it open if debugging, or close.
    # For now, we leave it open as per original script.


if __name__ == "__main__":
    get_spreadsheet_data()
