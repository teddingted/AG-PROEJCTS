import os
import win32com.client
import threading
import time
import win32gui
import win32con
import win32api

def dismiss_popup():
    max_retries = 10
    for _ in range(max_retries):
        try:
            hwnd = win32gui.FindWindow("#32770", "Aspen HYSYS")
            if hwnd:
                win32gui.SetForegroundWindow(hwnd)
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                return
        except:
            pass
        time.sleep(1)

def check_celltext():
    filename = "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc"
    folder = "hysys_automation"
    file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))

    try:
        app = win32com.client.Dispatch("HYSYS.Application")
        app.Visible = True
    except:
        return

    t = threading.Thread(target=dismiss_popup)
    t.daemon = True
    t.start()

    try:
        case = app.SimulationCases.Open(file_path)
        case.Visible = True
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        
        # Check Col A (Mass Flow)
        cell_a12 = ss.Cell("A12")
        print(f"A12 CellValue: {cell_a12.CellValue}")
        print(f"A12 CellText: '{cell_a12.CellText}'")
        
        # Check Col B (Vol Flow)
        cell_b12 = ss.Cell("B12")
        print(f"B12 CellValue: {cell_b12.CellValue}")
        print(f"B12 CellText: '{cell_b12.CellText}'")
        
        # Check Col C (Pressure)
        cell_c12 = ss.Cell("C12")
        print(f"C12 CellValue: {cell_c12.CellValue}")
        print(f"C12 CellText: '{cell_c12.CellText}'")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_celltext()
