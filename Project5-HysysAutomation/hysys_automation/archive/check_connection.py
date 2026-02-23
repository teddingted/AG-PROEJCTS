import sys

def check_hysys_connection():
    print("Connecting to Aspen HYSYS...")
    try:
        import win32com.client
    except ImportError:
        print("Error: 'pywin32' library is not installed.")
        print("Please install it using: pip install pywin32")
        return

    try:
        # Try to connect to an active HYSYS instance or start a new one
        # "HYSYS.Application" is the standard ProgID.
        app = win32com.client.Dispatch("HYSYS.Application")
        print("Successfully connected to HYSYS!")
        
        # Print version if available to confirm interaction
        # Note: Depending on the version, the property might be different, but Version is common.
        try:
            print(f"HYSYS Version: {app.Version}")
        except:
            pass
        
        print("Connection test passed.")
        
    except Exception as e:
        print("Failed to connect to HYSYS.")
        print(f"Error details: {e}")
        print("\nTroubleshooting tips:")
        print("1. Ensure Aspen HYSYS is installed on this machine.")
        print("2. Ensure 'pywin32' is installed (pip install pywin32).")
        print("3. Try running this script as Administrator if permissions are an issue.")

if __name__ == "__main__":
    check_hysys_connection()
