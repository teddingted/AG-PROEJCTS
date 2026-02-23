import os
import win32com.client
from hysys_utils import read_data_from_case, print_data_table

def main():
    app = win32com.client.Dispatch("HYSYS.Application")
    case = app.ActiveDocument
    
    print("Checking SPRDSHT-1...")
    try:
        ss = case.Flowsheet.Operations.Item("SPRDSHT-1")
        # Print first few rows/cols
        for r in range(1, 15):
            row_str = ""
            for c in ["A", "B", "C", "D"]:
                cell = f"{c}{r}"
                try:
                    val = ss.Cell(cell).CellValue
                    txt = ss.Cell(cell).CellText
                    row_str += f" | {cell}: {txt} ({val})"
                except:
                    pass
            print(row_str)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
