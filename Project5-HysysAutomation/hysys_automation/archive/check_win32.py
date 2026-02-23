try:
    import win32gui
    import win32api
    import win32con
    print("Imports successful")
except ImportError as e:
    print(f"Import failed: {e}")
