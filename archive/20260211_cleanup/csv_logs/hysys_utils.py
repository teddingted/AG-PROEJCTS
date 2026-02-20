import win32com.client
import threading
import time
import win32gui
import win32con
import win32api
import os

def dismiss_popup():
    """
    Background thread to detect and dismiss the 'Aspen HYSYS' popup 
    that appears when opening a v14 case in v15.
    """
    print("[Popup Handler] Started...")
    max_retries = 30
    found = False
    
    for _ in range(max_retries):
        try:
            hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
            if hwnd:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except:
                    pass
                time.sleep(0.5)
                # VK_RETURN
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                found = True
                break
        except:
            pass
        time.sleep(1)

def convert_to_display_unit(raw_val, unit_str):
    if raw_val is None:
        return None
    try:
        val = float(raw_val)
    except:
        return raw_val

    if unit_str == 'kg/h': return val * 3600.0
    elif unit_str == 'm3/h': return val * 3600.0
    elif unit_str == 'bar': return val / 100.0
    elif unit_str == 'W/C': return val * 1000.0
    
    return val

def read_data_from_case(case):
    ss_name = "SPRDSHT-1"
    try:
        ss = case.Flowsheet.Operations.Item(ss_name)
    except Exception as e:
        print(f"Could not find operation '{ss_name}': {e}")
        return None

    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'] # 0-9
    
    def get_cell_data(cell_addr_str):
        val_str = "Error"
        unit_str = ""
        try:
            cell = ss.Cell(cell_addr_str)
        except Exception as e:
            return f"Err: {e}", ""

        # Unit
        try:
             if hasattr(cell, 'Units'):
                 unit_str = str(cell.Units)
        except:
             pass

        # Value
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

        converted_val = convert_to_display_unit(raw_val, unit_str)
        
        if isinstance(converted_val, (int, float)):
            val_str = f"{converted_val:.2f}"
        elif converted_val is None:
            val_str = ""
        else:
            val_str = str(converted_val)
            
        return val_str, unit_str

    results = {'headers': [], 'values': [], 'units': []}
    for col_name in cols:
        h_val, _ = get_cell_data(f"{col_name}11")
        v_val, v_unit = get_cell_data(f"{col_name}12")
        results['headers'].append(h_val)
        results['values'].append(v_val)
        results['units'].append(v_unit)
        
    return results

def print_data_table(data):
    if not data: return
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    headers = data['headers']
    values = data['values']
    units = data['units']

    print("\n" + "="*100)
    print(f"{'Col':<5} | {'Row 11 (Parameter)':<40} | {'Row 12 (Value)':<15} | {'Unit':<15}")
    print("-" * 100)
    for i, col in enumerate(cols):
        h = headers[i] if headers[i] != "None" else ""
        v = values[i] if values[i] != "None" else ""
        u = units[i]
        print(f"{col:<5} | {h:<40} | {v:<15} | {u:<15}")
    print("="*100 + "\n")
