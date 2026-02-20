import os
import shutil
import datetime

root_dir = r"C:\Users\Admin\Desktop\AG-BEGINNING"
archive_root = os.path.join(root_dir, "archive")
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
target_archive_dir = os.path.join(archive_root, f"cleanup_{timestamp}")

# Files and Folders to KEEP
keep_items = [
    "AutoPlotDigitizerV2_Windows_Port",
    "hysys_automation",
    "HiWayCalculator",
    "hysys_optimizer_unified.py",
    "hysys_optimizer_dispatch.py",
    "hysys_optimizer_hybrid.py",
    "hysys_optimizer_multidim.py",
    "hysys_optimizer_acc.py",
    "hysys_optimizer_2d.py",
    "test_hysys_connection.py",
    "OPTIMIZATION_REPORT_AND_STRATEGY.md",
    "HYSYS_AUTOMATION_SMRY_KR.md",
    "archive",                # Do not move the archive folder itself
    "cleanup_ag_beginning.py", # Do not move self while running
    ".gitignore",             # Good practice to keep
    "task.md",                # Keep active task files if present locally (unlikely but safe)
    "implementation_plan.md"
]

print(f"Starting cleanup of {root_dir}")
print(f"Target Archive: {target_archive_dir}")

if not os.path.exists(target_archive_dir):
    os.makedirs(target_archive_dir)
    print("Created archive directory.")

moved_count = 0

# Get list of all items in root directory
all_items = os.listdir(root_dir)

for item in all_items:
    # Skip items in the keep list
    if item in keep_items:
        print(f"Skipping (KEEP): {item}")
        continue
    
    src_path = os.path.join(root_dir, item)
    dst_path = os.path.join(target_archive_dir, item)
    
    try:
        # Move file or directory
        print(f"Moving: {item} -> {dst_path}")
        shutil.move(src_path, dst_path)
        moved_count += 1
    except Exception as e:
        print(f"ERROR moving {item}: {e}")

print("--------------------------------------------------")
print(f"Cleanup Complete. Moved {moved_count} items to archive.")
print("--------------------------------------------------")
