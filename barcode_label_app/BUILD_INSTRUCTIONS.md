# Building Windows Executable for Barcode Label Generator

This guide will help you create a standalone Windows executable (.exe) file from the Python barcode label generator application.

## Prerequisites

### 1. Python Installation
- Python 3.8 or later
- Download from [python.org](https://python.org)
- **Important**: During installation, check "Add Python to PATH"

### 2. Required Tools
- Git (optional, for version control)
- Windows 10 or later (for building)

## Quick Start (Recommended)

### Option 1: Use the Batch File (Easiest)
1. Open Command Prompt or PowerShell as Administrator
2. Navigate to the `barcode_label_app` folder
3. Double-click `build.bat` or run:
   ```batch
   build.bat
   ```

### Option 2: Use Python Script
1. Open Command Prompt or PowerShell as Administrator
2. Navigate to the `barcode_label_app` folder
3. Run:
   ```bash
   python build_exe.py
   ```

## Manual Build Process

If you prefer to understand each step:

### Step 1: Set Up Virtual Environment (Recommended)
```bash
# Create virtual environment
python -m venv venv

# Activate it (Windows)
venv\Scripts\activate

# Activate it (Git Bash/WSL)
source venv/Scripts/activate
```

### Step 2: Install Dependencies
```bash
# Upgrade pip
pip install --upgrade pip

# Install all requirements
pip install -r requirements.txt
```

### Step 3: Build Executable
```bash
# Basic build
pyinstaller --onefile --windowed --name=BarcodeGenerator simple_barcode_app.py

# Advanced build with all assets
pyinstaller --onefile --windowed --name=BarcodeGenerator --add-data="logo.png;." --add-data="data;data" --collect-all=treepoem simple_barcode_app.py
```

## Build Output

After successful build, you'll find:

```
barcode_label_app/
├── dist/
│   └── BarcodeGenerator.exe          # Main executable
├── build/                            # Temporary build files
├── BarcodeGenerator_Distribution/    # Ready-to-share folder
│   ├── BarcodeGenerator.exe
│   ├── logo.png
│   ├── data/
│   │   └── serial_tracker.xlsx
│   └── README.txt
└── BarcodeGenerator.spec            # PyInstaller spec file
```

## Distribution

### What to Share
Share the entire `BarcodeGenerator_Distribution` folder, which contains:
- `BarcodeGenerator.exe` - The main application
- `logo.png` - Default logo file
- `data/` - Sample Excel data
- `README.txt` - Instructions for end users

### End User Requirements
- Windows 10 or later
- No Python installation needed
- No additional software required

## Troubleshooting

### Common Issues

#### 1. "Python not found"
**Solution**: Reinstall Python and ensure "Add to PATH" is checked

#### 2. "Permission denied"
**Solution**: Run Command Prompt as Administrator

#### 3. "Module not found" errors
**Solution**: 
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

#### 4. Large executable size (>100MB)
**Normal**: The executable includes Python runtime and all libraries

#### 5. Antivirus flags the executable
**Normal**: Many antivirus programs flag PyInstaller executables as suspicious. Add an exception or use `--debug` flag during development.

### Build Options Explained

| Option | Description |
|--------|-------------|
| `--onefile` | Create single executable file |
| `--windowed` | No console window (GUI only) |
| `--name=BarcodeGenerator` | Name of the executable |
| `--add-data="logo.png;."` | Include logo file |
| `--add-data="data;data"` | Include data folder |
| `--collect-all=treepoem` | Include all treepoem dependencies |
| `--hidden-import=tkinter` | Explicitly include tkinter |

### Advanced Customization

#### Custom Icon
```bash
pyinstaller --onefile --windowed --icon=custom_icon.ico --name=BarcodeGenerator simple_barcode_app.py
```

#### Debug Mode (shows console)
```bash
pyinstaller --onefile --name=BarcodeGenerator simple_barcode_app.py
```

#### Include Additional Files
```bash
pyinstaller --onefile --windowed --add-data="additional_file.txt;." --name=BarcodeGenerator simple_barcode_app.py
```

## Build Script Features

The `build_exe.py` script automatically:
- ✅ Checks for virtual environment
- ✅ Installs/updates dependencies
- ✅ Cleans previous builds
- ✅ Builds with optimal settings
- ✅ Creates distribution folder
- ✅ Includes all necessary files
- ✅ Generates user documentation
- ✅ Handles errors gracefully

## Testing the Executable

### Basic Test
1. Navigate to `BarcodeGenerator_Distribution`
2. Double-click `BarcodeGenerator.exe`
3. Verify the app opens without errors

### Full Test
1. Try loading the default Excel file
2. Browse for a different Excel file
3. Change the logo
4. Adjust positions with sliders
5. Generate a label
6. Save the label

## File Size Optimization

The executable will be large (50-150MB) because it includes:
- Python runtime
- All required libraries (PIL, pandas, tkinter, etc.)
- Barcode generation libraries
- Data processing libraries

This is normal for PyInstaller executables and ensures the app runs on any Windows machine without dependencies.

## Support

If you encounter issues:
1. Check the troubleshooting section above
2. Run with `--debug` flag to see detailed output
3. Check the build log for specific error messages
4. Ensure all dependencies are properly installed

## Security Notes

- Some antivirus software may flag PyInstaller executables
- This is a false positive due to the executable packing method
- You can add an exception for the executable
- For enterprise distribution, consider code signing

---

**Generated executable will be completely standalone and ready for distribution on any Windows system!** 