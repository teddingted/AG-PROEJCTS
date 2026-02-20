
# -*- coding: utf-8 -*-
import sys
import os

# UTF-8 Fix for Windows cp949
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'


import pandas as pd
import re
import os

# Paths
FDS_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\fds_content.txt"
CSV_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\ERSN_1s_2024-08-28.csv"
OUTPUT_DIR = r"c:\Users\Admin\Desktop\AG-BEGINNING\analysis_output"

def extract_controllers_from_fds():
    """
    Parse FDS text to extract all control loops
    Format in FDS:
    Purpose  To maintain...
    PV <tag>
    SP <value/source>
    OP <tag>
    Acting Reverse/Direct
    """
    print("=== EXTRACTING CONTROLLERS FROM FDS ===\n")
    
    with open(FDS_PATH, 'r', encoding='utf-8', errors='ignore') as f:
        fds_text = f.read()
    
    controllers = []
    
    # Pattern to find controller sections
    # Look for "Purpose" keyword followed by PV, SP, OP
    pattern = r'Purpose\s+(.*?)\nPV\s+(.*?)\nSP\s+(.*?)\nOP\s+(.*?)\nActing\s+(.*?)\n'
    
    matches = re.findall(pattern, fds_text, re.DOTALL)
    
    for match in matches:
        purpose, pv, sp, op, acting = match
        
        # Clean up text
        purpose = purpose.strip()
        pv = pv.strip()
        sp = sp.strip()
        op = op.strip()
        acting = acting.strip()
        
        # Extract controller name from OP tag
        # Usually in format like "71FCV0001" or "81TCV0002 and 81TCV0003"
        controller_match = re.search(r'(\d{2}[A-Z]{3}\d{4})', op)
        if controller_match:
            controller_name = controller_match.group(1)
        else:
            controller_name = "Unknown"
        
        controllers.append({
            'Controller': controller_name,
            'Purpose': purpose,
            'PV_Tag': pv,
            'SP_Source': sp,
            'OP_Tag': op,
            'Acting': acting
        })
    
    print(f"Extracted {len(controllers)} controllers from FDS")
    return pd.DataFrame(controllers)

def manual_controller_list():
    """
    Manually define all known controllers from FDS
    (In case automated extraction fails due to PDF formatting issues)
    """
    controllers = [
        # Section 3.1.2.1 - Load Control
        {
            'Section': '3.1.2.1',
            'Controller': 'Load Control',
            'Type': 'Cascade',
            'Purpose': 'Control re-liquefaction load',
            'PV_Tag': 'GMS Tank Pressure',
            'SP_Source': 'GMS Output (Max 1500 kg/h)',
            'OP_Tag': '71FIC0001 + 81PIC0003',
            'Acting': 'Complex'
        },
        # Section 3.1.2.2 - BOG Flow Controller
        {
            'Section': '3.1.2.2',
            'Controller': '71FIC0001',
            'Type': 'FIC',
            'Purpose': 'Maintain flow rate for re-liquefaction',
            'PV_Tag': '72FI0001',
            'SP_Source': 'From GMS (Auto) or Manual input',
            'OP_Tag': '71FCV0001',
            'Acting': 'Reverse'
        },
        # Section 3.1.2.3 - N2 Compander Suction Pressure
        {
            'Section': '3.1.2.3',
            'Controller': '81PIC0003',
            'Type': 'PIC',
            'Purpose': 'Maintain inlet pressure for N2 compander',
            'PV_Tag': '51PIT011A',
            'SP_Source': 'RSP from Pre-set table',
            'OP_Tag': '81XV0002, 81PCV0003, 81PCV0004 (Split Range)',
            'Acting': 'Reverse'
        },
        # Section 3.1.2.4 - Expander Inlet Temperature
        {
            'Section': '3.1.2.4',
            'Controller': '71TIC0001',
            'Type': 'TIC',
            'Purpose': 'Maintain expander inlet temperature for N2 compander',
            'PV_Tag': '81TE0005',
            'SP_Source': '-130°C (To be adjusted during commissioning)',
            'OP_Tag': '71FCV0001',
            'Acting': 'Reverse'
        },
        # Section 3.1.1.2 - LD Compressor Inlet Temperature
        {
            'Section': '3.1.1.2',
            'Controller': '81TIC0001',
            'Type': 'TIC',
            'Purpose': 'Maintain inlet temperature for LD compressor',
            'PV_Tag': '72TE0001',
            'SP_Source': '-50°C',
            'OP_Tag': '81TCV0002, 81TCV0003 (Split Range)',
            'Acting': 'Reverse'
        },
        # Additional controllers from Section 3.3.3
        {
            'Section': '3.3.3.1',
            'Controller': 'N2 Discharge Pressure',
            'Type': 'PIC',
            'Purpose': 'N2 discharge pressure control of N2 boosting compressor',
            'PV_Tag': 'Unknown',
            'SP_Source': 'Unknown',
            'OP_Tag': 'Unknown',
            'Acting': 'Unknown'
        },
        {
            'Section': '3.3.3.5',
            'Controller': '81PIC0001',
            'Type': 'PIC',
            'Purpose': 'N2 boosting compressor inlet pressure control',
            'PV_Tag': 'Unknown',
            'SP_Source': 'Unknown',
            'OP_Tag': 'Unknown',
            'Acting': 'Unknown'
        },
        {
            'Section': '3.3.3.6',
            'Controller': '81PIC0002',
            'Type': 'PIC',
            'Purpose': 'N2 buffer tank pressure control',
            'PV_Tag': 'Unknown',
            'SP_Source': 'Unknown',
            'OP_Tag': 'Unknown',
            'Acting': 'Unknown'
        },
        {
            'Section': '3.3.3.7',
            'Controller': '82PIC0001',
            'Type': 'PIC',
            'Purpose': 'N2 expansion vessel pressure control',
            'PV_Tag': 'Unknown',
            'SP_Source': 'Unknown',
            'OP_Tag': 'Unknown',
            'Acting': 'Unknown'
        },
        {
            'Section': '3.3.3.8',
            'Controller': '81PIC0005',
            'Type': 'PIC',
            'Purpose': 'N2 compander suction pressure control',
            'PV_Tag': 'Unknown',
            'SP_Source': 'Unknown',
            'OP_Tag': 'Unknown',
            'Acting': 'Unknown'
        }
    ]
    
    return pd.DataFrame(controllers)

def find_tags_in_csv(controller_df):
    """
    Cross-reference controller tags with CSV columns
    """
    print("\n=== CROSS-REFERENCING WITH CSV DATA ===\n")
    
    # Load CSV column names only
    df_sample = pd.read_csv(CSV_PATH, nrows=1)
    csv_columns = df_sample.columns.tolist()
    
    print(f"CSV contains {len(csv_columns)} columns")
    
    # For each controller, find matching tags
    results = []
    
    for idx, row in controller_df.iterrows():
        controller_name = row['Controller']
        pv_tag = row['PV_Tag']
        op_tag = row['OP_Tag']
        
        # Search for PV tag in CSV
        pv_matches = [col for col in csv_columns if pv_tag in col or col in pv_tag]
        
        # Search for OP tag in CSV  
        # OP might be split range, extract all mentioned tags
        op_tags_mentioned = re.findall(r'\d{2}[A-Z]{2,4}\d{4}', op_tag)
        op_matches = []
        for op_candidate in op_tags_mentioned:
            op_matches.extend([col for col in csv_columns if op_candidate in col])
        
        results.append({
            'Controller': controller_name,
            'Type': row.get('Type', 'Unknown'),
            'Purpose': row['Purpose'][:50] + '...' if len(row['Purpose']) > 50 else row['Purpose'],
            'PV_Expected': pv_tag,
            'PV_Found_in_CSV': ', '.join(pv_matches) if pv_matches else 'NOT FOUND',
            'OP_Expected': op_tag,
            'OP_Found_in_CSV': ', '.join(op_matches) if op_matches else 'NOT FOUND',
            'Tag_Match_Rate': f"{(len(pv_matches) > 0) + (len(op_matches) > 0)}/2"
        })
    
    results_df = pd.DataFrame(results)
    
    # Save
    save_path = os.path.join(OUTPUT_DIR, "controller_sensor_mapping.csv")
    results_df.to_csv(save_path, index=False)
    print(f"\nSaved controller-sensor mapping to {save_path}")
    
    # Summary
    total_controllers = len(results_df)
    pv_found = results_df['PV_Found_in_CSV'].apply(lambda x: x != 'NOT FOUND').sum()
    op_found = results_df['OP_Found_in_CSV'].apply(lambda x: x != 'NOT FOUND').sum()
    
    print(f"\n=== MATCH SUMMARY ===")
    print(f"Total Controllers: {total_controllers}")
    print(f"PV Tags Found: {pv_found}/{total_controllers} ({100*pv_found/total_controllers:.1f}%)")
    print(f"OP Tags Found: {op_found}/{total_controllers} ({100*op_found/total_controllers:.1f}%)")
    
    return results_df

def main():
    # Try automated extraction first
    auto_controllers = extract_controllers_from_fds()
    
    if not auto_controllers.empty:
        print("\nUsing auto-extracted controllers")
        controller_df = auto_controllers
    else:
        print("\nAuto-extraction failed, using manual list")
        controller_df = manual_controller_list()
    
    # Cross-reference with CSV
    mapping = find_tags_in_csv(controller_df)
    
    print("\n=== CONTROLLER MAPPING COMPLETE ===")
    print(mapping.to_string(index=False))

if __name__ == "__main__":
    main()
