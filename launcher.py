"""Launcher for GW Parser.
Compiled into an EXE via PyInstaller. Handles venv detection and setup.
"""
import sys
import os
import subprocess
from pathlib import Path

# Determine base directory (works both normally and when frozen by PyInstaller)
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

VENV_DIR = BASE_DIR / ".venv"
GUI_SCRIPT = BASE_DIR / "gw_parser_gui.py"
SETUP_SCRIPT = BASE_DIR / "setup_venv.py"

def get_system_python() -> str:
    """Find the system Python executable that created this frozen exe."""
    # sys.executable points to the temp extracted exe; we need the real system python.
    # We look for it in the standard PATH.
    python_cmd = "python"
    if sys.platform == "win32":
        python_cmd = "python"
        
    try:
        result = subprocess.run(
            [python_cmd, "--version"], 
            capture_output=True, 
            text=True
        )
        if result.returncode == 0:
            return python_cmd
    except Exception:
        pass
        
    return "python3"

def main():
    os.chdir(str(BASE_DIR))
    
    # If venv doesn't exist, run the bootstrapper
    if not VENV_DIR.exists():
        print("=" * 50)
        print("GW Parser - First Run Setup")
        print("=" * 50)
        print("Python environment not found.")
        print("Setting up virtual environment...\n")
        
        if not SETUP_SCRIPT.exists():
            print("ERROR: setup_venv.py is missing!")
            print("Please ensure all program files were extracted together.")
            input("\nPress Enter to exit...")
            sys.exit(1)
            
        system_python = get_system_python()
        result = subprocess.run(
            [system_python, str(SETUP_SCRIPT)],
            cwd=str(BASE_DIR)
        )
        
        if result.returncode != 0 or not VENV_DIR.exists():
            print("\nERROR: Environment setup failed.")
            print("Please make sure Python is installed and in your system PATH.")
            input("\nPress Enter to exit...")
            sys.exit(1)
            
        print("\nSetup complete! Launching GUI...\n")

    # Determine the venv python path
    if sys.platform == "win32":
        venv_python = VENV_DIR / "Scripts" / "python.exe"
    else:
        venv_python = VENV_DIR / "bin" / "python"
        
    if not venv_python.exists():
        print(f"ERROR: Virtual environment python not found at {venv_python}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Launch the actual GUI in the venv
    subprocess.run([str(venv_python), str(GUI_SCRIPT)], cwd=str(BASE_DIR))

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Fatal Error: {e}")
        input("\nPress Enter to exit...")