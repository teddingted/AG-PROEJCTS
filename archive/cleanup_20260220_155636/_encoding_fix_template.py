# -*- coding: utf-8 -*-
"""
BOG Analysis Script Template
Resolves Windows cp949 encoding issues
"""

import sys
import os

# ============================================
# ENCODING FIX FOR WINDOWS (cp949 → UTF-8)
# ============================================
if sys.platform == 'win32':
    # Force UTF-8 for stdout/stderr
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    
    # Set environment variable for subprocess
    os.environ['PYTHONIOENCODING'] = 'utf-8'

# Now all print() statements will work with Korean and Unicode characters
# ============================================

import pandas as pd
import numpy as np

def safe_print(message):
    """
    Fallback print function that handles encoding errors gracefully
    """
    try:
        print(message)
    except UnicodeEncodeError:
        # Replace problematic characters with ASCII equivalents
        safe_msg = message.encode('ascii', 'replace').decode('ascii')
        print(safe_msg)

# Example usage:
if __name__ == "__main__":
    print("✓ UTF-8 encoding test: 한글 테스트")
    print("한국어와 Unicode 문자 (✓✗→←) 모두 정상 출력")
    
    # Use safe_print if you want extra safety
    safe_print("✓ This will always work even if encoding fails")
