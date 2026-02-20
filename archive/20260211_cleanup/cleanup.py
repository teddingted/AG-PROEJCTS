import os
import shutil
import glob

FOLDER = "hysys_automation"
ARCHIVE = os.path.join(FOLDER, "archive")

if not os.path.exists(ARCHIVE):
    os.makedirs(ARCHIVE)

# Files to keep
KEEP_FILES = [
    "final_optimizer.py",
    "hysys_utils.py",
    "Efficiency Increase_1.5 t_2 min. temp without heat recovery.hsc",
    "optimize_robust.py", # Keep as backup for now
    "archive",
    "__pycache__"
]

# Move all other python files and logs
all_files = glob.glob(os.path.join(FOLDER, "*"))

for file_path in all_files:
    filename = os.path.basename(file_path)
    
    # Skip if item is to be kept
    if filename in KEEP_FILES:
        continue
        
    # Skip directories (except archive which is handled by KEEP_FILES logic implicitly but good to be explicit)
    if os.path.isdir(file_path):
        continue
        
    # Move files
    try:
        shutil.move(file_path, os.path.join(ARCHIVE, filename))
        print(f"Moved: {filename}")
    except Exception as e:
        print(f"Error moving {filename}: {e}")

print("Cleanup complete.")
