#!/usr/bin/env python3
"""
Outlook Email Extractor - Launcher

This script ensures all dependencies are installed and launches the application.
It automatically handles virtual environment creation and activation.
"""
import os
import sys
import subprocess
import platform
import venv
import shutil
import stat
import importlib.util
from pathlib import Path

# Constants
VENV_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'venv')
REQUIREMENTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'requirements.txt')
PYTHON_EXECUTABLE = sys.executable
IS_WINDOWS = platform.system() == 'Windows'
IS_MACOS = platform.system() == 'Darwin'
IS_LINUX = platform.system() == 'Linux'

def create_virtualenv():
    """Create a new virtual environment if it doesn't exist."""
    if not os.path.exists(VENV_DIR):
        print(f"Creating virtual environment in {VENV_DIR}...")
        venv.create(VENV_DIR, with_pip=True)
        print("Virtual environment created successfully.")
    else:
        print(f"Using existing virtual environment at {VENV_DIR}")

def get_venv_python():
    """Get the path to the Python executable in the virtual environment."""
    if IS_WINDOWS:
        return os.path.join(VENV_DIR, 'Scripts', 'python.exe')
    else:
        return os.path.join(VENV_DIR, 'bin', 'python')

def install_requirements():
    """Install required packages in the virtual environment."""
    venv_python = get_venv_python()
    
    # Upgrade pip first
    print("\nUpgrading pip...")
    subprocess.check_call([venv_python, "-m", "pip", "install", "--upgrade", "pip"])
    
    # Install PySimpleGUI from private PyPI server first
    print("\nInstalling PySimpleGUI from private PyPI...")
    subprocess.check_call([
        venv_python, "-m", "pip", "install", "--upgrade",
        "--extra-index-url", "https://PySimpleGUI.net/install",
        "PySimpleGUI"
    ])
    
    # Install other requirements
    print("\nInstalling other requirements...")
    subprocess.check_call([venv_python, "-m", "pip", "install", "-r", REQUIREMENTS])

def run_in_venv():
    """Run the application in the virtual environment."""
    venv_python = get_venv_python()
    project_root = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(project_root, 'outlook_extractor', 'run.py')
    
    # Set PYTHONPATH to include the project root
    env = os.environ.copy()
    if 'PYTHONPATH' in env:
        env['PYTHONPATH'] = f"{project_root}:{env['PYTHONPATH']}"
    else:
        env['PYTHONPATH'] = project_root
    
    print("\nStarting Outlook Email Extractor...")
    os.execle(venv_python, venv_python, script_path, env)

def main():
    """Main entry point for the launcher."""
    print("Outlook Email Extractor - Launcher")
    print("=" * 40)
    
    # Create and set up virtual environment
    create_virtualenv()
    install_requirements()
    
    # Run the application in the virtual environment
    try:
        run_in_venv()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nError: {str(e)}")
        print("\nPlease try the following steps to resolve the issue:")
        print("1. Delete the 'venv' folder and run this launcher again")
        print("2. Make sure you have Python 3.8+ installed")
        print("3. Check your internet connection")
        if IS_MACOS:
            print("4. On macOS, you might need to install Xcode Command Line Tools")
            print("   Run: xcode-select --install")
        sys.exit(1)

if __name__ == "__main__":
    main()
