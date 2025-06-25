# Barcode Scanner & Label Generator

A Python Tkinter application that receives barcode scanner input, looks up data in an Excel tracker file, and generates formatted labels with THALES branding for printing.

## Features

- **Barcode Scanner Integration**: Receives input from physical barcode scanners
- **Data Parsing**: Automatically extracts part numbers and serial numbers from scanned text
- **Excel Lookup**: Finds matching records in your serial number tracker Excel file
- **Label Generation**: Creates formatted labels with THALES branding and QR codes
- **Print Ready**: High-resolution labels suitable for direct printing

## Prerequisites

- Python 3.7+
- Physical barcode scanner (acts as keyboard input)
- Excel file with serial number tracker data

## Installation

1. **Activate your virtual environment:**
   ```bash
   # From the project root directory
   ./env/bin/Activate.ps1  # For PowerShell on macOS
   # or
   source env/bin/activate  # For bash/zsh
   ```

2. **Install system dependencies (macOS with Homebrew):**
   ```bash
   brew install python-tk
   ```

3. **Navigate to the application directory:**
   ```bash
   cd barcode_label_app
   ```

4. **Check system requirements:**
   ```bash
   python check_requirements.py
   ```

5. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

6. **Verify your Excel file is in place:**
   - The file `data/serial_tracker.xlsx` should exist
   - It should contain columns like CPN, PRODUCT DESC, SL.From, SL.End, etc.

## Usage

1. **Start the application:**
   ```bash
   python main.py
   ```

2. **Using the application:**
   - Place cursor in the "Scanned Data" input field
   - Scan your barcode with the physical scanner
   - The app will automatically parse part numbers and serial numbers
   - Click "Lookup in Excel" to find matching records
   - Click "Generate Label" to create the formatted label
   - Use "Save Label" or "Print Label" to output the result

## Application Workflow

1. **Scan Barcode** → Data appears in input field
2. **Parse Data** → Part number and serial number extracted
3. **Excel Lookup** → Matching record found in tracker
4. **Generate Label** → Formatted label created with QR code
5. **Save/Print** → Label saved or sent to printer

## File Structure

```
barcode_label_app/
├── main.py              # Main application entry point
├── config.py            # Configuration settings
├── data_parser.py       # Barcode data parsing
├── excel_handler.py     # Excel file operations
├── label_generator.py   # Label creation
├── ui_components.py     # UI interface
├── requirements.txt     # Python dependencies
├── data/
│   └── serial_tracker.xlsx  # Your Excel tracker file
├── assets/              # Logo and images (optional)
└── output_labels/       # Generated labels saved here
```

## Configuration

Edit `config.py` to customize:
- Barcode parsing patterns
- Label dimensions and styling
- Excel column mappings
- File paths

## Troubleshooting

- **Barcode not parsing**: Adjust parsing patterns in `config.py`
- **Excel lookup fails**: Verify part numbers match in Excel file
- **Label generation issues**: Check Excel data completeness
- **Print problems**: Verify default printer settings

## Support

For issues or customization needs, refer to the detailed implementation guide in the original instructions document. 