import win32com.client
import sys
import os

def check_spreadsheet():
    try:
        try:
            app = win32com.client.GetActiveObject("HYSYS.Application")
        except:
            app = win32com.client.Dispatch("HYSYS.Application")
        
        case = app.ActiveDocument
        if not case:
            print("No active case found.")
            return

        fs = case.Flowsheet
        ss = fs.Operations.Item("SPRDSHT-1")
        
        if not ss:
            print("Spreadsheet 'SPRDSHT-1' not found.")
            return
            
        print(f"Spreadsheet Name: {ss.Name}")
        
        # Check Cells A12 to J12
        cells = ['A12', 'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12', 'J12']
        
        print("\nChecking Cells A12-J12:")
        for cell_id in cells:
            try:
                cell = ss.Cell(cell_id)
                val = cell.CellValue
                print(f"  [{cell_id}] = {val} (Type: {type(val)})")
            except Exception as e:
                print(f"  [{cell_id}] ERROR: {e}")

    except Exception as e:
        print(f"Review Error: {e}")

if __name__ == "__main__":
    check_spreadsheet()
