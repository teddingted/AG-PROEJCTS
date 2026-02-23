import os
import win32com.client
import time
from hysys_utils import dismiss_popup

def main():
    filename = "ETHANE_CHS_MIMIC_REV1_FULL_LOOP.hsc"
    folder = "hysys_automation"
    # If the script is run from the directory where the file is, just use the filename
    # checking if current directory wraps 'hysys_automation'
    if folder in os.getcwd():
        file_path = os.path.abspath(filename)
    else:
        file_path = os.path.abspath(os.path.join(os.getcwd(), folder, filename))

    
    print(f"Opening {file_path}...")
    
    app = win32com.client.Dispatch("HYSYS.Application")
    app.Visible = True
    
    if os.path.exists(file_path):
        case = app.SimulationCases.Open(file_path)
        case.Visible = True
    else:
        print("File not found.")
        return

    # Wait a bit for everything to load
    time.sleep(5)
    
    try:
        # 1. Inspect 1_propene
        print("\n--- Inspecting 1_propene ---")
        try:
            s_propene = case.Flowsheet.MaterialStreams.Item("1_propene")
            print(f"Found '1_propene'")
            print(f"Temperature: {s_propene.Temperature.Value:.2f} C")
            print(f"Pressure: {s_propene.Pressure.Value:.2f} kPa")
            print(f"Mass Flow: {s_propene.MassFlow.Value*3600:.2f} kg/h")
            
            comps = s_propene.ComponentMassFraction.Values
            comp_names = case.Flowsheet.FluidPackage.Components.Names
            
            print("Composition (Mass Fraction):")
            for i, name in enumerate(comp_names):
                print(f"  {name}: {comps[i]:.4f}")
                    
        except Exception as e:
            print(f"Error accessing '1_propene': {e}")
            
        # 2. Inspect F_ethylene
        print("\n--- Inspecting F_ethylene ---")
        try:
            s_ethylene = case.Flowsheet.MaterialStreams.Item("F_ethylene")
            print(f"Found 'F_ethylene'")
            print(f"Temperature: {s_ethylene.Temperature.Value:.2f} C")
            print(f"Pressure: {s_ethylene.Pressure.Value:.2f} kPa")
             
        except Exception as e:
            print(f"Error accessing 'F_ethylene': {e}")

    except Exception as e:
        print(f"General Error: {e}")
    finally:
        # Don't close for now, just keep it open for debugging
        pass

if __name__ == "__main__":
    main()
