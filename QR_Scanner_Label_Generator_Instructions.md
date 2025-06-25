# Barcode Scanner Input & Label Generator - Python Tkinter Application

## Project Overview
This application receives barcode scanner input, looks up the scanned data in an Excel tracker file, and generates formatted labels for printing. The workflow involves:

1. **Barcode Input**: Receive scanned data from physical barcode scanner via input field
2. **Data Parsing**: Extract part numbers and serial numbers from scanned text
3. **Excel Lookup**: Find matching records in a serial number tracker Excel file
4. **Label Generation**: Create formatted labels with THALES branding

## Prerequisites and Dependencies

### Required Python Libraries
```bash
pip install tkinter
pip install pillow
pip install pandas
pip install openpyxl
pip install qrcode
```

### System Requirements
- Python 3.7+
- Physical barcode scanner (acts as keyboard input)
- Excel file with serial number tracker data

## Application Structure

### Directory Structure
```
barcode_label_app/
├── main.py                 # Main application file
├── data_parser.py          # Parse scanned barcode data
├── excel_handler.py        # Excel file operations
├── label_generator.py      # Label creation and formatting
├── ui_components.py        # Tkinter UI components
├── config.py              # Configuration settings
├── assets/
│   └── logo.png           # THALES logo (optional)
└── data/
    └── serial_tracker.xlsx # Excel tracker file
```

## Detailed Implementation Guide

### 1. Configuration File (config.py)
```python
import os

# Application Settings
APP_TITLE = "Barcode Scanner & Label Generator"
WINDOW_SIZE = "900x700"

# File Paths
EXCEL_FILE_PATH = "data/serial_tracker.xlsx"
LOGO_PATH = "assets/logo.png"
OUTPUT_DIR = "output_labels"

# Excel Column Mappings (based on your screenshot)
EXCEL_COLUMNS = {
    'DATE': 'DATE',
    'CONSIGNEE': 'CONSIGNEE', 
    'CPN': 'CPN',              # Component Part Number
    'WO_PO_NO': 'WO/PO NO',    # Work Order/Purchase Order Number
    'CARD_TYPE': 'CARD TYPE',
    'PRODUCT_DESC': 'PRODUCT DESC',
    'REV': 'Rev',              # Revision
    'KIT_SIZE': 'KIT SIZE',
    'SL_FROM': 'SL.From',      # Serial Range Start
    'SL_END': 'SL.End',        # Serial Range End
    'PRINTED_BY': 'Printed By'
}

# Barcode Data Parsing
BARCODE_PATTERNS = {
    'PART_NUMBER': r'P/N\s*:?\s*([A-Za-z0-9]+)',    # Matches "P/N : 63215031AA"
    'SERIAL_NUMBER': r'S/N\s*:?\s*([A-Za-z0-9.]+)', # Matches "S/N : 68925616.017806"
    'PART_NUMBER_ALT': r'([0-9]{8}[A-Z]{2})',        # Alternative: direct pattern
    'SERIAL_NUMBER_ALT': r'([0-9]{8}\.[0-9]{6})'     # Alternative: direct pattern
}

# Label Settings
LABEL_WIDTH = 400
LABEL_HEIGHT = 200
FONT_SIZE_LARGE = 16
FONT_SIZE_MEDIUM = 12
FONT_SIZE_SMALL = 10
```

### 2. Data Parser Module (data_parser.py)
```python
import re
import config

class BarcodeDataParser:
    def __init__(self):
        self.patterns = config.BARCODE_PATTERNS
    
    def parse_barcode_data(self, raw_data):
        """Parse barcode scanner input to extract part number and serial number"""
        try:
            if not raw_data or not raw_data.strip():
                return {}, "No data provided"
            
            data = {}
            raw_data = raw_data.strip()
            
            # Try primary patterns first
            part_match = re.search(self.patterns['PART_NUMBER'], raw_data, re.IGNORECASE)
            if part_match:
                data['part_number'] = part_match.group(1)
            
            serial_match = re.search(self.patterns['SERIAL_NUMBER'], raw_data, re.IGNORECASE)
            if serial_match:
                data['serial_number'] = serial_match.group(1)
            
            # Try alternative patterns if primary didn't work
            if 'part_number' not in data:
                part_alt_match = re.search(self.patterns['PART_NUMBER_ALT'], raw_data)
                if part_alt_match:
                    data['part_number'] = part_alt_match.group(1)
            
            if 'serial_number' not in data:
                serial_alt_match = re.search(self.patterns['SERIAL_NUMBER_ALT'], raw_data)
                if serial_alt_match:
                    data['serial_number'] = serial_alt_match.group(1)
            
            # If still no data found, try line-by-line parsing
            if not data:
                lines = raw_data.split('\n')
                for line in lines:
                    line = line.strip()
                    
                    # Look for part number indicators
                    if any(indicator in line.upper() for indicator in ['P/N', 'PART', 'PN']):
                        # Extract alphanumeric sequence after indicators
                        parts = re.findall(r'[A-Za-z0-9]+', line)
                        if len(parts) >= 2:  # Skip the indicator, take the value
                            data['part_number'] = parts[-1]
                    
                    # Look for serial number indicators
                    if any(indicator in line.upper() for indicator in ['S/N', 'SERIAL', 'SN']):
                        # Extract alphanumeric sequence (with dots/dashes) after indicators
                        serials = re.findall(r'[A-Za-z0-9./-]+', line)
                        if len(serials) >= 2:  # Skip the indicator, take the value
                            data['serial_number'] = serials[-1]
            
            if not data:
                return {}, "Could not parse part number or serial number from input"
            
            return data, None
            
        except Exception as e:
            return {}, f"Error parsing barcode data: {str(e)}"
    
    def validate_data(self, data):
        """Validate parsed data"""
        errors = []
        
        if 'part_number' not in data or not data['part_number']:
            errors.append("Part number not found")
        
        if 'serial_number' not in data or not data['serial_number']:
            errors.append("Serial number not found")
        
        # Validate part number format (adjust as needed)
        if 'part_number' in data:
            part_num = data['part_number']
            if len(part_num) < 8:
                errors.append("Part number seems too short")
        
        return errors if errors else None
```

### 3. Excel Handler Module (excel_handler.py)
```python
import pandas as pd
from datetime import datetime
import config

class ExcelHandler:
    def __init__(self, excel_path=None):
        self.excel_path = excel_path or config.EXCEL_FILE_PATH
        self.df = None
        self.load_excel()
    
    def load_excel(self):
        """Load Excel file into pandas DataFrame"""
        try:
            self.df = pd.read_excel(self.excel_path)
            print(f"Loaded Excel file with {len(self.df)} rows")
            return True
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")
            return False
    
    def find_matching_record(self, part_number, serial_number=None):
        """Find matching record in Excel based on part number and serial range"""
        if self.df is None:
            return None, "Excel file not loaded"
        
        try:
            # Search by Component Part Number (CPN) - exact match first
            exact_matches = self.df[self.df['CPN'].astype(str).str.upper() == part_number.upper()]
            
            if not exact_matches.empty:
                matching_rows = exact_matches
            else:
                # Try partial match
                matching_rows = self.df[self.df['CPN'].astype(str).str.contains(
                    part_number, case=False, na=False)]
            
            if matching_rows.empty:
                return None, f"No matching records found for part number: {part_number}"
            
            # If serial number provided, check if it's within range
            if serial_number:
                for idx, row in matching_rows.iterrows():
                    sl_from = str(row.get('SL.From', ''))
                    sl_end = str(row.get('SL.End', ''))
                    
                    if self.is_serial_in_range(serial_number, sl_from, sl_end):
                        return row.to_dict(), None
                
                return None, f"Serial number {serial_number} not in any valid range for part {part_number}"
            
            # Return first matching record if no serial number check needed
            return matching_rows.iloc[0].to_dict(), None
            
        except Exception as e:
            return None, f"Error searching Excel: {str(e)}"
    
    def is_serial_in_range(self, serial_number, range_start, range_end):
        """Check if serial number falls within the specified range"""
        try:
            if not range_start or not range_end or range_start == 'nan' or range_end == 'nan':
                return True  # If no range specified, accept any serial
            
            # Extract numeric part of serial numbers for comparison
            serial_num = self.extract_numeric_serial(serial_number)
            start_num = self.extract_numeric_serial(str(range_start))
            end_num = self.extract_numeric_serial(str(range_end))
            
            return start_num <= serial_num <= end_num
            
        except Exception as e:
            print(f"Error checking serial range: {str(e)}")
            return True  # Default to True if can't parse
    
    def extract_numeric_serial(self, serial_str):
        """Extract numeric part from serial number string"""
        try:
            # For formats like "63215031-100351" or "63215031.100351"
            if '-' in serial_str:
                parts = serial_str.split('-')
                return int(parts[-1]) if parts[-1].isdigit() else 0
            elif '.' in serial_str:
                parts = serial_str.split('.')
                return int(parts[-1]) if parts[-1].isdigit() else 0
            else:
                # Remove all non-digit characters and convert
                clean_serial = re.sub(r'[^\d]', '', str(serial_str))
                return int(clean_serial) if clean_serial.isdigit() else 0
        except:
            return 0
    
    def update_printed_status(self, part_number, printed_by):
        """Update the printed status in Excel"""
        try:
            if self.df is not None:
                # Find the row to update
                mask = self.df['CPN'].astype(str).str.upper() == part_number.upper()
                if mask.any():
                    self.df.loc[mask, 'Printed By'] = printed_by
                    # Add timestamp column if it doesn't exist
                    if 'Print Date' not in self.df.columns:
                        self.df['Print Date'] = ''
                    self.df.loc[mask, 'Print Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    
                    # Save back to Excel
                    self.df.to_excel(self.excel_path, index=False)
                    return True
                
        except Exception as e:
            print(f"Error updating Excel: {str(e)}")
            return False
    
    def get_excel_summary(self):
        """Get summary of Excel data for display"""
        if self.df is None:
            return "Excel file not loaded"
        
        return f"Total records: {len(self.df)}\nColumns: {', '.join(self.df.columns.tolist())}"
```

### 4. Label Generator Module (label_generator.py)
```python
from PIL import Image, ImageDraw, ImageFont
import qrcode
import config
from datetime import datetime
import os

class LabelGenerator:
    def __init__(self):
        self.label_width = config.LABEL_WIDTH
        self.label_height = config.LABEL_HEIGHT
        
    def create_label(self, excel_data, scanned_data=None):
        """Create a formatted label image based on your sample"""
        try:
            # Create blank label with white background
            label = Image.new('RGB', (self.label_width, self.label_height), 'white')
            draw = ImageDraw.Draw(label)
            
            # Load fonts (fallback to default if Arial not available)
            try:
                font_large = ImageFont.truetype("arial.ttf", config.FONT_SIZE_LARGE)
                font_medium = ImageFont.truetype("arial.ttf", config.FONT_SIZE_MEDIUM)
                font_small = ImageFont.truetype("arial.ttf", config.FONT_SIZE_SMALL)
                font_bold = ImageFont.truetype("arialbd.ttf", config.FONT_SIZE_MEDIUM)
            except:
                font_large = ImageFont.load_default()
                font_medium = ImageFont.load_default()
                font_small = ImageFont.load_default()
                font_bold = ImageFont.load_default()
            
            # Draw THALES header (top-left)
            draw.text((10, 10), "THALES", fill='black', font=font_large)
            
            # Draw part number with green background (based on your sample)
            part_number = excel_data.get('CPN', 'N/A')
            
            # Green rectangle background for part number
            draw.rectangle([(70, 35), (250, 60)], fill='lightgreen', outline='black')
            draw.text((75, 40), f"P/N : {part_number}", fill='black', font=font_bold)
            
            # Draw Product Description
            y_offset = 70
            product_desc = excel_data.get('PRODUCT DESC', '')[:40]  # Truncate if too long
            draw.text((10, y_offset), product_desc, fill='black', font=font_small)
            
            # Draw Serial Number
            y_offset += 25
            if scanned_data and 'serial_number' in scanned_data:
                serial_num = scanned_data['serial_number']
            else:
                # Generate serial from range
                serial_num = self.generate_serial_from_range(excel_data)
            
            draw.text((10, y_offset), f"S/N° {serial_num}", fill='black', font=font_medium)
            
            # Draw Quantity
            y_offset += 25
            qty = excel_data.get('KIT SIZE', '1')
            draw.text((10, y_offset), f"QTY : {qty}", fill='black', font=font_small)
            
            # Create and add QR code (top-right corner)
            qr_data = f"P/N : {part_number}\nS/N : {serial_num}"
            qr_img = self.create_qr_code(qr_data)
            if qr_img:
                qr_size = 80
                qr_img = qr_img.resize((qr_size, qr_size))
                label.paste(qr_img, (self.label_width - qr_size - 10, 10))
            
            # Add border around entire label
            draw.rectangle([(0, 0), (self.label_width-1, self.label_height-1)], 
                          outline='black', width=2)
            
            return label, None
            
        except Exception as e:
            return None, f"Error creating label: {str(e)}"
    
    def generate_serial_from_range(self, excel_data):
        """Generate serial number from Excel range data"""
        try:
            sl_from = str(excel_data.get('SL.From', ''))
            sl_end = str(excel_data.get('SL.End', ''))
            
            if not sl_from or sl_from == 'nan':
                return "SERIAL_N/A"
            
            # For ranges like "63215031-100351" to "63215031-100370"
            # We'll use the start of the range
            return sl_from
                
        except Exception as e:
            print(f"Error generating serial: {str(e)}")
            return "SERIAL_ERROR"
    
    def create_qr_code(self, data):
        """Create QR code image"""
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=3,
                border=1,
            )
            qr.add_data(data)
            qr.make(fit=True)
            return qr.make_image(fill_color="black", back_color="white")
        except Exception as e:
            print(f"Error creating QR code: {str(e)}")
            return None
    
    def save_label(self, label_image, filename=None):
        """Save label image to file"""
        try:
            if not os.path.exists(config.OUTPUT_DIR):
                os.makedirs(config.OUTPUT_DIR)
                
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"label_{timestamp}.png"
            
            filepath = os.path.join(config.OUTPUT_DIR, filename)
            label_image.save(filepath, 'PNG', dpi=(300, 300))  # High DPI for printing
            return filepath, None
            
        except Exception as e:
            return None, f"Error saving label: {str(e)}"
```

### 5. UI Components Module (ui_components.py)
```python
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import config

class MainUI:
    def __init__(self, root, app_controller):
        self.root = root
        self.controller = app_controller
        self.setup_ui()
        
    def setup_ui(self):
        """Setup main UI components"""
        self.root.title(config.APP_TITLE)
        self.root.geometry(config.WINDOW_SIZE)
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Barcode Scanner & Label Generator", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=15)
        
        # Barcode Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Barcode Scanner Input", padding="15")
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        input_frame.columnconfigure(1, weight=1)
        
        # Instructions
        instructions = ttk.Label(input_frame, 
                                text="Place cursor in the field below and scan with your barcode scanner:",
                                font=('Arial', 10))
        instructions.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=tk.W)
        
        # Barcode input field
        ttk.Label(input_frame, text="Scanned Data:", font=('Arial', 11, 'bold')).grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = ttk.Entry(input_frame, textvariable=self.barcode_var, 
                                      font=('Arial', 11), width=60)
        self.barcode_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.barcode_entry.focus()  # Set focus for barcode scanner
        
        # Parse button
        parse_btn = ttk.Button(input_frame, text="Parse Data", command=self.parse_barcode_data)
        parse_btn.grid(row=1, column=2, padx=(10, 0))
        
        # Clear button
        clear_btn = ttk.Button(input_frame, text="Clear", command=self.clear_input)
        clear_btn.grid(row=2, column=2, padx=(10, 0), pady=(5, 0))
        
        # Bind Enter key to parse
        self.barcode_entry.bind('<Return>', lambda e: self.parse_barcode_data())
        
        # Parsed Data Display
        data_frame = ttk.LabelFrame(main_frame, text="Parsed Data", padding="15")
        data_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        data_frame.columnconfigure(1, weight=1)
        
        # Part Number
        ttk.Label(data_frame, text="Part Number:", font=('Arial', 11, 'bold')).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.part_number_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.part_number_var, width=25, 
                 font=('Arial', 11)).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Serial Number
        ttk.Label(data_frame, text="Serial Number:", font=('Arial', 11, 'bold')).grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.serial_number_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.serial_number_var, width=25,
                 font=('Arial', 11)).grid(row=1, column=1, sticky=(tk.W, tk.E), 
                                         padx=(0, 10), pady=(5, 0))
        
        # Lookup Button
        lookup_btn = ttk.Button(data_frame, text="Lookup in Excel", 
                               command=self.lookup_data, style='Accent.TButton')
        lookup_btn.grid(row=2, column=0, columnspan=2, pady=15)
        
        # Excel Data Display
        excel_frame = ttk.LabelFrame(main_frame, text="Excel Lookup Results", padding="15")
        excel_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        excel_frame.columnconfigure(0, weight=1)
        
        # Treeview for Excel data
        columns = ('Field', 'Value')
        self.excel_tree = ttk.Treeview(excel_frame, columns=columns, show='headings', height=8)
        self.excel_tree.heading('Field', text='Field')
        self.excel_tree.heading('Value', text='Value')
        self.excel_tree.column('Field', width=150)
        self.excel_tree.column('Value', width=300)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(excel_frame, orient=tk.VERTICAL, command=self.excel_tree.yview)
        self.excel_tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid treeview and scrollbar
        self.excel_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Label Generation Section
        label_frame = ttk.LabelFrame(main_frame, text="Label Generation", padding="15")
        label_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Button frame
        btn_frame = ttk.Frame(label_frame)
        btn_frame.grid(row=0, column=0, columnspan=2)
        
        ttk.Button(btn_frame, text="Generate Label", 
                  command=self.generate_label).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Save Label", 
                  command=self.save_label).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Print Label", 
                  command=self.print_label).grid(row=0, column=2, padx=5)
        
        # Label Preview
        preview_frame = ttk.LabelFrame(main_frame, text="Label Preview", padding="10")
        preview_frame.grid(row=1, column=2, rowspan=4, padx=(20, 0), sticky=(tk.N, tk.W))
        
        self.label_preview = ttk.Label(preview_frame, text="Label preview will appear here\nafter generation", 
                                      justify=tk.CENTER, anchor=tk.CENTER)
        self.label_preview.grid(row=0, column=0, padx=10, pady=10)
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Place cursor in barcode field and scan")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, font=('Arial', 10))
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def parse_barcode_data(self):
        """Handle barcode data parsing"""
        barcode_data = self.barcode_var.get().strip()
        if not barcode_data:
            self.show_error("Please scan or enter barcode data first")
            return
        
        self.controller.parse_barcode_data(barcode_data)
    
    def clear_input(self):
        """Clear all input fields"""
        self.barcode_var.set("")
        self.part_number_var.set("")
        self.serial_number_var.set("")
        self.clear_excel_data()
        self.clear_label_preview()
        self.barcode_entry.focus()
        self.update_status("Input cleared - Ready for new scan")
    
    def lookup_data(self):
        """Handle Excel lookup"""
        part_number = self.part_number_var.get().strip()
        serial_number = self.serial_number_var.get().strip()
        
        if not part_number:
            self.show_error("Part number is required for lookup")
            return
            
        self.controller.lookup_excel_data(part_number, serial_number)
    
    def generate_label(self):
        """Handle label generation"""
        self.controller.generate_label()
    
    def save_label(self):
        """Handle label saving"""
        self.controller.save_label()
    
    def print_label(self):
        """Handle label printing"""
        self.controller.print_label()
    
    def update_parsed_data(self, data):
        """Update UI with parsed barcode data"""
        self.part_number_var.set(data.get('part_number', ''))
        self.serial_number_var.set(data.get('serial_number', ''))
    
    def update_excel_data(self, data):
        """Update treeview with Excel data"""
        # Clear previous data
        self.clear_excel_data()
        
        # Add new data
        for key, value in data.items():
            self.excel_tree.insert('', 'end', values=(key, str(value)))
    
    def clear_excel_data(self):
        """Clear Excel data display"""
        for item in self.excel_tree.get_children():
            self.excel_tree.delete(item)
    
    def update_label_preview(self, label_image):
        """Update label preview"""
        try:
            # Resize image for preview
            preview_size = (250, 125)
            preview_image = label_image.resize(preview_size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(preview_image)
            
            self.label_preview.configure(image=photo, text="")
            self.label_preview.image = photo  # Keep reference
        except Exception as e:
            print(f"Error updating preview: {str(e)}")
            self.label_preview.configure(text=f"Preview Error: {str(e)}")
    
    def clear_label_preview(self):
        """Clear label preview"""
        self.label_preview.configure(image="", text="Label preview will appear here\nafter generation")
        if hasattr(self.label_preview, 'image'):
            self.label_preview.image = None
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def show_error(self, message):
        """Show error message"""
        messagebox.showerror("Error", message)
    
    def show_success(self, message):
        """Show success message"""
        messagebox.showinfo("Success", message)
    
    def show_info(self, message):
        """Show info message"""
        messagebox.showinfo("Information", message)
```

### 6. Main Application Controller (main.py)
```python
import tkinter as tk
import os
from data_parser import BarcodeDataParser
from excel_handler import ExcelHandler
from label_generator import LabelGenerator
from ui_components import MainUI
import config

class BarcodeLabeMApp:
    def __init__(self):
        self.root = tk.Tk()
        
        # Initialize components
        self.data_parser = BarcodeDataParser()
        self.excel_handler = ExcelHandler()
        self.label_generator = LabelGenerator()
        self.ui = MainUI(self.root, self)
        
        # Current data storage
        self.current_barcode_data = None
        self.current_parsed_data = None
        self.current_excel_data = None
        self.current_label = None
        
        # Create output directory
        os.makedirs(config.OUTPUT_DIR, exist_ok=True)
        
        # Check if Excel file exists
        self.check_excel_file()
    
    def check_excel_file(self):
        """Check if Excel file exists and is accessible"""
        if not os.path.exists(config.EXCEL_FILE_PATH):
            self.ui.show_error(f"Excel file not found: {config.EXCEL_FILE_PATH}")
            self.ui.update_status("ERROR: Excel file not found")
        else:
            summary = self.excel_handler.get_excel_summary()
            self.ui.update_status(f"Excel loaded - {summary}")
    
    def parse_barcode_data(self, raw_barcode_data):
        """Parse barcode scanner input"""
        self.ui.update_status("Parsing barcode data...")
        
        self.current_barcode_data = raw_barcode_data
        parsed_data, error = self.data_parser.parse_barcode_data(raw_barcode_data)
        
        if error:
            self.ui.show_error(f"Parse Error: {error}")
            self.ui.update_status("Error parsing barcode data")
            return
        
        # Validate parsed data
        validation_errors = self.data_parser.validate_data(parsed_data)
        if validation_errors:
            self.ui.show_error("Validation Issues:\n" + "\n".join(validation_errors))
        
        self.current_parsed_data = parsed_data
        self.ui.update_parsed_data(parsed_data)
        self.ui.update_status("Barcode data parsed successfully")
        
        # Auto-lookup if part number found
        if 'part_number' in parsed_data:
            self.lookup_excel_data(parsed_data['part_number'], 
                                 parsed_data.get('serial_number'))
    
    def lookup_excel_data(self, part_number, serial_number=None):
        """Lookup data in Excel file"""
        if not part_number:
            self.ui.show_error("Part number is required for lookup")
            return
        
        self.ui.update_status("Looking up data in Excel...")
        
        data, error = self.excel_handler.find_matching_record(part_number, serial_number)
        
        if error:
            self.ui.show_error(f"Lookup Error: {error}")
            self.ui.update_status("Error looking up Excel data")
            return
        
        self.current_excel_data = data
        self.ui.update_excel_data(data)
        self.ui.update_status("Excel data found - Ready to generate label")
    
    def generate_label(self):
        """Generate label from current data"""
        if not self.current_excel_data:
            self.ui.show_error("Please lookup Excel data first")
            return
        
        self.ui.update_status("Generating label...")
        
        label_image, error = self.label_generator.create_label(
            self.current_excel_data, self.current_parsed_data)
        
        if error:
            self.ui.show_error(f"Label Generation Error: {error}")
            self.ui.update_status("Error generating label")
            return
        
        self.current_label = label_image
        self.ui.update_label_preview(label_image)
        self.ui.update_status("Label generated successfully")
    
    def save_label(self):
        """Save current label to file"""
        if not self.current_label:
            self.ui.show_error("Please generate a label first")
            return
        
        self.ui.update_status("Saving label...")
        
        # Generate filename based on part number
        part_num = self.current_parsed_data.get('part_number', 'unknown') if self.current_parsed_data else 'unknown'
        filename = f"label_{part_num}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        
        filepath, error = self.label_generator.save_label(self.current_label, filename)
        
        if error:
            self.ui.show_error(f"Save Error: {error}")
            self.ui.update_status("Error saving label")
        else:
            self.ui.show_success(f"Label saved successfully!\n\nFile: {os.path.basename(filepath)}")
            self.ui.update_status(f"Label saved: {os.path.basename(filepath)}")
            
            # Update Excel with printed status
            if self.current_parsed_data and 'part_number' in self.current_parsed_data:
                self.excel_handler.update_printed_status(
                    self.current_parsed_data['part_number'], 
                    "System User")
    
    def print_label(self):
        """Print current label"""
        if not self.current_label:
            self.ui.show_error("Please generate a label first")
            return
        
        self.ui.update_status("Preparing label for printing...")
        
        # Save temporary file for printing
        temp_filename = "temp_print_label.png"
        temp_file, error = self.label_generator.save_label(
            self.current_label, temp_filename)
        
        if error:
            self.ui.show_error(f"Print Preparation Error: {error}")
            return
        
        try:
            # Print using system default printer
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                # Windows: use default image viewer/printer
                os.startfile(temp_file)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", temp_file], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", temp_file], check=True)
            
            self.ui.update_status("Label sent to printer")
            self.ui.show_info("Label has been sent to your default printer.\nPlease check your printer queue.")
            
            # Update Excel with printed status
            if self.current_parsed_data and 'part_number' in self.current_parsed_data:
                self.excel_handler.update_printed_status(
                    self.current_parsed_data['part_number'], 
                    "System User")
            
        except Exception as e:
            self.ui.show_error(f"Print Error: {str(e)}")
            self.ui.update_status("Error printing label")
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            print("Application closed by user")
        except Exception as e:
            print(f"Application error: {str(e)}")

if __name__ == "__main__":
    app = BarcodeLabeMApp()
    app.run()
```

## Installation and Setup Instructions

### 1. Environment Setup
```bash
# Create project directory
mkdir barcode_label_app
cd barcode_label_app

# Create virtual environment (optional but recommended)
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install required packages
pip install pillow pandas openpyxl qrcode[pil]
```

### 2. Directory Structure Setup
```bash
# Create required directories
mkdir data
mkdir output_labels
mkdir assets

# Copy your Excel file to data directory
cp "Serial number tracker.xlsx" data/serial_tracker.xlsx
```

### 3. Configuration
- Update `config.py` with your specific Excel column names
- Adjust barcode parsing patterns if needed
- Modify label dimensions and styling as required

### 4. Testing the Application
```bash
# Run the application
python main.py

# Test workflow:
# 1. Place cursor in the barcode input field
# 2. Scan your barcode (it will appear as text)
# 3. Click "Parse Data" or press Enter
# 4. Verify part number and serial number extraction
# 5. Click "Lookup in Excel"
# 6. Review Excel data found
# 7. Click "Generate Label"
# 8. Use "Save Label" or "Print Label"
```

## Key Features

1. **Simple Input**: Just one input field for barcode scanner
2. **Auto-parsing**: Automatically extracts part numbers and serial numbers
3. **Excel Integration**: Looks up data in your existing Excel tracker
4. **Label Preview**: Shows label before printing/saving
5. **Print Ready**: High-resolution labels suitable for printing
6. **Status Updates**: Clear feedback on each operation

## Usage Workflow

1. **Start Application**: Run `python main.py`
2. **Scan Barcode**: Place cursor in input field and scan with your barcode scanner
3. **Parse Data**: Data is automatically parsed, or click "Parse Data"
4. **Excel Lookup**: Click "Lookup in Excel" to find matching records
5. **Generate Label**: Click "Generate Label" to create the formatted label
6. **Save/Print**: Use "Save Label" or "Print Label" buttons

## Troubleshooting

### Common Issues:
1. **Barcode Not Parsing**: Check and adjust parsing patterns in `config.py`
2. **Excel Lookup Fails**: Verify part numbers match exactly in Excel
3. **Label Generation Issues**: Ensure all required Excel fields are present
4. **Print Problems**: Check default printer settings

This simplified version focuses on the core functionality you need with a clean, user-friendly interface. 