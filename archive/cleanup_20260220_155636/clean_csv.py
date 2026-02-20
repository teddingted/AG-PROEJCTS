import csv
import os

FILE = "hysys_automation/optimization_2d_extended.csv"
TEMP = "hysys_automation/optimization_2d_extended_temp.csv"

# Remove 1100 kg/h rows where VolFlow >= 3700
count = 0
with open(FILE, 'r') as fin, open(TEMP, 'w', newline='') as fout:
    reader = csv.DictReader(fin)
    writer = csv.DictWriter(fout, fieldnames=reader.fieldnames)
    writer.writeheader()
    
    for row in reader:
        try:
            flow = int(float(row['Flow']))
            vol = int(float(row['VolFlow']))
            
            if flow == 1100 and vol >= 3700:
                print(f"Removing {flow}, {vol}")
                count += 1
                continue
        except: pass
        writer.writerow(row)

os.replace(TEMP, FILE)
print(f"Cleaned {count} rows.")
