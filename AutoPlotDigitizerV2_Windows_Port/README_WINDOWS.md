# AutoPlotDigitizer V2 - Windows Porting Kit

This folder contains the complete source code and build configuration to create a standalone Windows executable.

## Contents
*   `main.py`: Entry point of the application.
*   `gui/`: Graphical User Interface code.
*   `core/`: Image processing logic.
*   `requirements.txt`: List of Python libraries required.
*   `AutoPlotDigitizer.spec`: Configuration for PyInstaller.

## How to Build (On Windows)

1.  **Install Python**: Ensure Python 3.10 or later is installed.
2.  **Open Terminal**: Open PowerShell or Command Prompt in this folder.
3.  **Install Dependencies**:
    ```powershell
    pip install -r requirements.txt
    pip install pyinstaller
    ```
4.  **Build EXE**:
    ```powershell
    pyinstaller AutoPlotDigitizer.spec
    ```
5.  **Run**:
    *   Go to the newly created `dist` folder.
    *   Run `AutoPlotDigitizer.exe`.

## Troubleshooting
If you encounter errors regarding `cv2`:
```powershell
pip install opencv-python-headless
```
