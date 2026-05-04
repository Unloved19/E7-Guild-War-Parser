"""Virtual environment bootstrapper for GW Parser.

Run this first to set up the Python environment.
Creates a venv in the program folder and installs all dependencies.
"""

import sys
import os
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
VENV_DIR = BASE_DIR / ".venv"
REQUIREMENTS_FILE = BASE_DIR / "requirements.txt"

# Packages that need special handling
PIP_INDEX_URL = "https://pypi.org/simple/"

# Core requirements (also written to requirements.txt for pip)
REQUIREMENTS = [
    "opencv-python>=4.8.0",
    "easyocr>=1.7.0",
    "numpy>=1.24.0",
    "mss>=9.0.0",
    "Pillow>=10.0.0",
    "PyMuPDF>=1.23.0",
    "gspread>=5.12.0",
    "google-auth-oauthlib>=1.1.0",
    "google-auth-httplib2>=0.1.0",
    "google-api-python-client>=2.100.0",
    "openpyxl>=3.1.0",
    "customtkinter>=5.2.0",
    "requests>=2.31.0",
]


def get_venv_python() -> Path:
    """Get the Python executable path inside the venv."""
    if sys.platform == "win32":
        return VENV_DIR / "Scripts" / "python.exe"
    else:
        return VENV_DIR / "bin" / "python"


def get_venv_pip() -> Path:
    """Get the pip executable path inside the venv."""
    if sys.platform == "win32":
        return VENV_DIR / "Scripts" / "pip.exe"
    else:
        return VENV_DIR / "bin" / "pip"


def create_venv() -> bool:
    """Create the virtual environment."""
    print("=" * 50)
    print("GW Parser - Environment Setup")
    print("=" * 50)
    print()
    
    if VENV_DIR.exists():
        print(f"Virtual environment already exists at: {VENV_DIR}")
        print("To recreate, delete the .venv folder and run this again.")
        return True

    print(f"Creating virtual environment at: {VENV_DIR}")
    print()
    
    result = subprocess.run(
        [sys.executable, "-m", "venv", str(VENV_DIR)],
        capture_output=False,
    )
    
    if result.returncode != 0:
        print("\nERROR: Failed to create virtual environment.")
        print("Make sure you have the 'venv' module installed.")
        print("On some Linux systems, you may need: sudo apt install python3-venv")
        return False
    
    print("Virtual environment created successfully.")
    return True


def write_requirements() -> None:
    """Write requirements.txt for reference."""
    content = "# GW Parser Requirements\n"
    content += "# Generated automatically by setup_venv.py\n\n"
    for req in REQUIREMENTS:
        content += f"{req}\n"
    REQUIREMENTS_FILE.write_text(content, encoding="utf-8")
    print(f"Written {REQUIREMENTS_FILE.name}")


def install_requirements() -> bool:
    """Install all requirements into the venv."""
    pip = get_venv_pip()
    
    if not pip.exists():
        print(f"\nERROR: pip not found at {pip}")
        return False

    print(f"\nInstalling dependencies...")
    print("(This may take a few minutes, especially for easyocr models)")
    print("-" * 50)
    
    # Install each package with output shown
    for i, req in enumerate(REQUIREMENTS, 1):
        print(f"\n[{i}/{len(REQUIREMENTS)}] Installing {req}...")
        result = subprocess.run(
            [str(pip), "install", req, "--index-url", PIP_INDEX_URL],
            capture_output=False,
        )
        
        if result.returncode != 0:
            print(f"\nWARNING: Failed to install {req}")
            print("You can try installing it manually later:")
            print(f"  {pip} install {req}")
            # Continue with other packages
    
    print("\n" + "-" * 50)
    print("Dependency installation complete.")
    return True


def verify_installation() -> bool:
    """Verify that key packages can be imported."""
    python = get_venv_python()
    
    print("\nVerifying installation...")
    
    test_script = """
import sys
errors = []

try:
    import cv2
except ImportError as e:
    errors.append(f"opencv-python: {e}")

try:
    import easyocr
except ImportError as e:
    errors.append(f"easyocr: {e}")

try:
    import numpy
except ImportError as e:
    errors.append(f"numpy: {e}")

try:
    import mss
except ImportError as e:
    errors.append(f"mss: {e}")

try:
    import PIL
except ImportError as e:
    errors.append(f"Pillow: {e}")

try:
    import fitz
except ImportError as e:
    errors.append(f"PyMuPDF: {e}")

try:
    import gspread
except ImportError as e:
    errors.append(f"gspread: {e}")

try:
    import openpyxl
except ImportError as e:
    errors.append(f"openpyxl: {e}")

try:
    import customtkinter
except ImportError as e:
    errors.append(f"customtkinter: {e}")

if errors:
    print("ERRORS:")
    for err in errors:
        print(f"  - {err}")
    sys.exit(1)
else:
    print("All packages verified successfully.")
    sys.exit(0)
"""
    
    result = subprocess.run(
        [str(python), "-c", test_script],
        capture_output=True,
        text=True,
    )
    
    print(result.stdout)
    if result.stderr:
        print(result.stderr)
    
    return result.returncode == 0


def create_run_script() -> None:
    """Create a convenience script to run the GUI."""
    if sys.platform != "win32":
        # On Linux/Mac, create a shell script
        run_script = BASE_DIR / "run_gui.sh"
        content = f"""#!/bin/bash
# Run GW Parser GUI
cd "{BASE_DIR}"
"{get_venv_python()}" "{BASE_DIR / 'gw_parser_gui.py'}" "$@"
"""
        run_script.write_text(content)
        run_script.chmod(0o755)
        print(f"Created {run_script.name}")
    else:
        # On Windows, create a batch file
        run_script = BASE_DIR / "run_gui.bat"
        content = f"""@echo off
REM Run GW Parser GUI
cd /d "{BASE_DIR}"
"{get_venv_python()}" "{BASE_DIR / 'gw_parser_gui.py'}" %*
"""
        run_script.write_text(content)
        print(f"Created {run_script.name}")


def main():
    print()
    
    # Step 1: Create venv
    if not create_venv():
        sys.exit(1)
    
    # Step 2: Write requirements.txt
    write_requirements()
    
    # Step 3: Install packages
    if not install_requirements():
        print("\nWARNING: Some packages failed to install.")
        print("The program may not work correctly.")
    
    # Step 4: Verify
    if not verify_installation():
        print("\nWARNING: Verification failed.")
        print("Try running this script again, or install missing packages manually.")
    
    # Step 5: Create run script
    create_run_script()
    
    # Done
    print()
    print("=" * 50)
    print("Setup complete!")
    print("=" * 50)
    print()
    print("To run the GW Parser:")
    if sys.platform == "win32":
        print(f"  Double-click: run_gui.bat")
        print(f"  Or run: {get_venv_python()} gw_parser_gui.py")
    else:
        print(f"  Run: ./run_gui.sh")
        print(f"  Or run: {get_venv_python()} gw_parser_gui.py")
    print()
    print("Note: easyocr will download OCR models (~100MB) on first use.")
    print("      Make sure you have an internet connection when first running the parser.")
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}")
        sys.exit(1)