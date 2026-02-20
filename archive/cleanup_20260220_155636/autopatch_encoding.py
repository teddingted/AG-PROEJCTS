# -*- coding: utf-8 -*-
"""
Auto-patch all Python scripts with UTF-8 encoding fix
"""

import os
import sys

# Force UTF-8 for this script itself
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')

ENCODING_FIX = '''# -*- coding: utf-8 -*-
import sys
import os

# UTF-8 Fix for Windows cp949
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

'''

def patch_python_file(filepath):
    """Add UTF-8 encoding fix to a Python file if not already present"""
    
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # Check if already patched
    if 'PYTHONIOENCODING' in content or 'reconfigure(encoding' in content:
        print(f"[SKIP] {os.path.basename(filepath)} - Already patched")
        return False
    
    # Check if it has shebang or encoding declaration
    lines = content.split('\n')
    insert_pos = 0
    
    # Skip shebang and existing encoding declarations
    for i, line in enumerate(lines):
        if line.startswith('#!') or line.startswith('# -*- coding'):
            insert_pos = i + 1
        elif line.strip() and not line.startswith('#'):
            break
    
    # Insert encoding fix
    new_content = '\n'.join(lines[:insert_pos]) + '\n' + ENCODING_FIX + '\n'.join(lines[insert_pos:])
    
    # Write back
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"[PATCHED] {os.path.basename(filepath)}")
    return True

def main():
    base_dir = r'c:\Users\Admin\Desktop\AG-BEGINNING'
    
    python_files = [
        'analyze_bog.py',
        'visualize_bog_events.py',
        'analyze_control_logic.py',
        'deep_analysis.py',
        'validate_fds_sequence.py',
        'extract_all_controllers.py',
        'extract_fds.py'
    ]
    
    print("=== Auto-patching Python files with UTF-8 fix ===\n")
    
    patched = 0
    for filename in python_files:
        filepath = os.path.join(base_dir, filename)
        if os.path.exists(filepath):
            if patch_python_file(filepath):
                patched += 1
        else:
            print(f"[NOT FOUND] {filename}")
    
    print(f"\n=== Complete: {patched} files patched ===")

if __name__ == "__main__":
    main()
