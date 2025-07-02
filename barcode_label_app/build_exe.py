#!/usr/bin/env python3
"""
Build script to create Windows executable for Barcode Label Generator
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def run_command(cmd, description=""):
    """Run a command and handle errors"""
    print(f"\n{'='*60}")
    print(f"Running: {description}")
    print(f"Command: {cmd}")
    print(f"{'='*60}")
    
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    
    if result.stdout:
        print("STDOUT:")
        print(result.stdout)
    
    if result.stderr:
        print("STDERR:")
        print(result.stderr)
    
    if result.returncode != 0:
        print(f"❌ Command failed with return code {result.returncode}")
        return False
    else:
        print("✅ Command completed successfully")
        return True

def main():
    print("🚀 Building Barcode Label Generator Windows Executable")
    print("=" * 60)
    
    # Get the current directory
    current_dir = Path(__file__).parent.absolute()
    print(f"Working directory: {current_dir}")
    
    # Check if we're in a virtual environment
    if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        print("✅ Virtual environment detected")
    else:
        print("⚠️  Warning: Not in a virtual environment. Consider using one.")
    
    # Install/upgrade pip and required packages
    print("\n📦 Installing required packages...")
    if not run_command("pip install --upgrade pip", "Upgrading pip"):
        return False
    
    if not run_command("pip install -r requirements.txt", "Installing requirements"):
        return False
    
    # Clean previous builds
    print("\n🧹 Cleaning previous builds...")
    for dir_name in ['build', 'dist', '__pycache__']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Removed {dir_name}")
    
    # Remove .spec file if exists
    spec_file = "simple_barcode_app.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"Removed {spec_file}")
    
    # Create the executable
    print("\n🔨 Creating Windows executable...")
    
    # PyInstaller command with all necessary options
    pyinstaller_cmd = [
        "pyinstaller",
        "--onefile",                    # Create single executable file
        "--windowed",                   # No console window (GUI app)
        "--name=BarcodeGenerator",      # Name of the executable
        "--icon=logo.png",              # Use logo as icon (will be converted)
        "--add-data=logo.png;.",        # Include logo file
        "--add-data=data;data",         # Include data folder
        "--hidden-import=PIL._tkinter_finder",  # Fix PIL+tkinter issues
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.messagebox",
        "--collect-all=treepoem",       # Include all treepoem files
        "--collect-all=PIL",            # Include all PIL files
        "--collect-all=pandas",         # Include all pandas files
        "simple_barcode_app.py"         # Main script
    ]
    
    cmd_str = " ".join(pyinstaller_cmd)
    
    if not run_command(cmd_str, "Building executable with PyInstaller"):
        print("\n❌ PyInstaller failed. Trying alternative approach...")
        
        # Try without windowed mode for debugging
        pyinstaller_cmd_debug = [
            "pyinstaller",
            "--onefile",
            "--name=BarcodeGenerator",
            "--add-data=logo.png;.",
            "--add-data=data;data",
            "--hidden-import=PIL._tkinter_finder",
            "--hidden-import=tkinter",
            "--hidden-import=tkinter.ttk",
            "--collect-all=treepoem",
            "simple_barcode_app.py"
        ]
        
        cmd_debug_str = " ".join(pyinstaller_cmd_debug)
        if not run_command(cmd_debug_str, "Building executable (debug mode)"):
            return False
    
    # Check if executable was created
    exe_path = os.path.join("dist", "BarcodeGenerator.exe")
    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path) / (1024 * 1024)  # Size in MB
        print(f"\n🎉 Success! Executable created:")
        print(f"   📍 Location: {os.path.abspath(exe_path)}")
        print(f"   📏 Size: {file_size:.1f} MB")
        
        # Create a distribution folder with all necessary files
        dist_folder = "BarcodeGenerator_Distribution"
        if os.path.exists(dist_folder):
            shutil.rmtree(dist_folder)
        
        os.makedirs(dist_folder)
        
        # Copy executable
        shutil.copy2(exe_path, dist_folder)
        
        # Copy sample files
        if os.path.exists("logo.png"):
            shutil.copy2("logo.png", dist_folder)
        
        if os.path.exists("data"):
            shutil.copytree("data", os.path.join(dist_folder, "data"))
        
        # Create README for distribution
        readme_content = """# Barcode Label Generator

## Installation
1. Simply run BarcodeGenerator.exe
2. No additional installation required!

## Usage
1. The app will look for 'data/serial_tracker.xlsx' by default
2. You can browse to select a different Excel file
3. You can browse to select a custom logo image
4. Adjust positions using the sliders
5. Enter a serial number to lookup data
6. Generate and save labels

## Files Included
- BarcodeGenerator.exe - Main application
- logo.png - Default logo (you can change this)
- data/ - Sample Excel data folder

## Support
If you encounter any issues, make sure you have:
- Windows 10 or later
- Administrator privileges (for first run)

Generated with PyInstaller
"""
        
        with open(os.path.join(dist_folder, "README.txt"), "w") as f:
            f.write(readme_content)
        
        print(f"\n📦 Distribution folder created: {os.path.abspath(dist_folder)}")
        print(f"   This folder contains everything needed to run the app!")
        
    else:
        print(f"\n❌ Executable not found at {exe_path}")
        return False
    
    print("\n" + "="*60)
    print("✅ Build completed successfully!")
    print("📋 Next steps:")
    print("   1. Test the executable on your Windows machine")
    print("   2. Share the entire 'BarcodeGenerator_Distribution' folder")
    print("   3. Users just need to run BarcodeGenerator.exe")
    print("="*60)
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1) 