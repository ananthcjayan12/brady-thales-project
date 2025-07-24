#!/usr/bin/env python3
"""
Enhanced Barcode Scanner & Label Generator
- Excel file selection
- Live label preview
- Adjustable label elements
- Custom label design matching your format
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
import os
import json
from datetime import datetime
import win32print, win32ui, win32con
from PIL import Image, ImageDraw, ImageWin

# PDF generation imports
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm, inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.graphics.barcode import code128
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
from reportlab.lib.colors import black, blue
from reportlab.lib.utils import ImageReader

class EnhancedBarcodeLabelApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Enhanced Barcode Scanner & Label Generator")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)  # Set minimum size
        
        # Excel file path - default (relative to script location)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_file = os.path.join(script_dir, "data", "serial_tracker.xlsx")
        self.df = None
        
        # Settings file path
        self.settings_file = os.path.join(script_dir, "label_settings.json")
        
        # Current data
        self.current_excel_data = None
        self.current_label = None
        
        # UI variables (will be initialized in setup_ui)
        self.settings_status_var = None
        
        # Default label settings - Using exact measurements from debug_label_generator_pdf.py
        self.default_settings = {
            'width': 490,            # 173mm converted to pixels (173 * 2.834)
            'height': 170,           # 60mm converted to pixels (60 * 2.834)
            'logo_path': self.get_default_logo_path(),
            'logo_x': 14,            # 5mm * 2.834
            'logo_y': 6,             # 2mm * 2.834 
            'logo_width': 99,        # 35mm * 2.834
            'logo_height': 48,       # 17mm * 2.834
            'pd_x': 127,             # 45mm * 2.834
            'pd_y': 17,              # 6mm * 2.834
            'pn_x': 127,             # 45mm * 2.834
            'pn_y': 40,              # 14mm * 2.834
            'pr_x': 127,             # 45mm * 2.834
            'pr_y': 82,              # 29mm * 2.834
            'sn_x': 127,             # 45mm * 2.834
            'sn_y': 130,             # 46mm * 2.834
            'barcode_width': 255,    # 90mm * 2.834
            'barcode_height': 23,    # 8mm * 2.834
            # Font sizes for different text elements
            'font_company_size': 14,     # For company name/logo text
            'font_label_size': 10,       # For P/D, P/N, P/R, S/N labels
            'font_data_size': 9,         # For data text below barcodes
            'font_dlm_size': 8           # For DLM text
        }
        
        # Load settings from file or use defaults
        self.label_settings = self.load_settings()
        
        # Load Excel file
        self.load_excel()
        
        # Setup UI
        self.setup_ui()
        
        # Create output directory
        os.makedirs("output_labels", exist_ok=True)
        
        # Update UI from loaded settings
        self.update_ui_from_settings()
        
        # Generate initial preview
        self.update_preview()
    
    def get_default_logo_path(self):
        """Get default logo path relative to script location"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_dir, "logo.png")
        # Return logo path if exists, otherwise None
        return logo_path if os.path.exists(logo_path) else None
    
    def save_settings(self):
        """Save current label settings to JSON file"""
        try:
            # Update settings from UI
            self.update_label_settings()
            
            # Save to file
            with open(self.settings_file, 'w') as f:
                json.dump(self.label_settings, f, indent=2)
            
            messagebox.showinfo("Settings Saved", f"Label settings saved to:\n{os.path.basename(self.settings_file)}")
            self.status_var.set("Settings saved successfully")
            if hasattr(self, 'settings_status_var') and self.settings_status_var:
                self.settings_status_var.set("✓ Settings saved to file")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")
    
    def load_settings(self):
        """Load label settings from JSON file or return defaults"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    saved_settings = json.load(f)
                
                # Merge with defaults to ensure all keys exist
                settings = self.default_settings.copy()
                settings.update(saved_settings)
                
                print(f"Loaded settings from {self.settings_file}")
                return settings
            else:
                print("No saved settings found, using defaults")
                return self.default_settings.copy()
                
        except Exception as e:
            print(f"Error loading settings: {e}, using defaults")
            return self.default_settings.copy()
    
    def load_and_apply_settings(self):
        """Load settings from file and apply to UI"""
        try:
            if os.path.exists(self.settings_file):
                self.label_settings = self.load_settings()
                self.update_ui_from_settings()
                self.update_preview()
                messagebox.showinfo("Settings Loaded", "Settings loaded successfully!")
                self.status_var.set("Settings loaded from file")
                if hasattr(self, 'settings_status_var') and self.settings_status_var:
                    self.settings_status_var.set("✓ Settings loaded from saved file")
            else:
                messagebox.showinfo("No Settings File", "No saved settings file found. Using current settings.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings: {e}")
    
    def reset_settings(self):
        """Reset settings to defaults"""
        if messagebox.askyesno("Reset Settings", "Reset all label settings to default values?"):
            self.label_settings = self.default_settings.copy()
            self.update_ui_from_settings()
            self.update_preview()
            self.status_var.set("Settings reset to defaults")
            if hasattr(self, 'settings_status_var') and self.settings_status_var:
                self.settings_status_var.set("Using default settings")
    
    def update_ui_from_settings(self):
        """Update UI controls to match current settings"""
        try:
            # Update all the variable controls
            self.width_var.set(self.label_settings['width'])
            self.height_var.set(self.label_settings['height'])
            self.logo_x_var.set(self.label_settings['logo_x'])
            self.logo_y_var.set(self.label_settings['logo_y'])
            self.logo_width_var.set(self.label_settings['logo_width'])
            self.logo_height_var.set(self.label_settings['logo_height'])
            self.pd_x_var.set(self.label_settings['pd_x'])
            self.pd_y_var.set(self.label_settings['pd_y'])
            self.pn_x_var.set(self.label_settings['pn_x'])
            self.pn_y_var.set(self.label_settings['pn_y'])
            self.pr_x_var.set(self.label_settings['pr_x'])
            self.pr_y_var.set(self.label_settings['pr_y'])
            self.sn_x_var.set(self.label_settings['sn_x'])
            self.sn_y_var.set(self.label_settings['sn_y'])
            self.barcode_width_var.set(self.label_settings['barcode_width'])
            self.barcode_height_var.set(self.label_settings['barcode_height'])
            
            # Update font size controls
            self.font_company_size_var.set(self.label_settings.get('font_company_size', 14))
            self.font_label_size_var.set(self.label_settings.get('font_label_size', 10))
            self.font_data_size_var.set(self.label_settings.get('font_data_size', 9))
            self.font_dlm_size_var.set(self.label_settings.get('font_dlm_size', 8))
            
            # Update logo path
            logo_path = self.label_settings.get('logo_path')
            if logo_path and os.path.exists(logo_path):
                self.logo_path_var.set(logo_path)
            else:
                self.logo_path_var.set("No logo selected")
                
        except Exception as e:
            print(f"Error updating UI from settings: {e}")
    
    def load_excel(self):
        """Load Excel file"""
        try:
            self.df = pd.read_excel(self.excel_file)
            print(f"Loaded Excel file with {len(self.df)} rows")
            print(f"Columns: {list(self.df.columns)}")
        except Exception as e:
            print(f"Error loading Excel: {e}")
            self.df = None
    
    def setup_ui(self):
        """Setup enhanced UI with preview and controls"""
        # Create main paned window - better for smaller screens
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel - Controls (slightly larger weight for controls)
        left_frame = ttk.Frame(main_paned, padding="5")
        main_paned.add(left_frame, weight=3)
        
        # Right panel - Preview
        right_frame = ttk.Frame(main_paned, padding="5")
        main_paned.add(right_frame, weight=2)
        
        self.setup_left_panel(left_frame)
        self.setup_right_panel(right_frame)
    
    def setup_left_panel(self, parent):
        """Setup left control panel with Main and Settings tabs"""
        # Create notebook with Main and Settings tabs
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        main_tab = ttk.Frame(notebook)
        settings_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text='Main')
        notebook.add(settings_tab, text='Settings')

        # Title in Main tab
        title = ttk.Label(main_tab, text='Barcode Scanner & Label Generator', font=('Arial', 14, 'bold'))
        title.pack(pady=(0, 10))
        
        # Excel file selection in Main tab
        excel_frame = ttk.LabelFrame(main_tab, text="Excel File", padding="10")
        excel_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.excel_path_var = tk.StringVar(value=self.excel_file)
        ttk.Label(excel_frame, text="Excel file:").pack(anchor=tk.W)
        
        path_frame = ttk.Frame(excel_frame)
        path_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Entry(path_frame, textvariable=self.excel_path_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="Browse", command=self.browse_excel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(path_frame, text="Load", command=self.load_selected_excel).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Input section in Main tab
        input_frame = ttk.LabelFrame(main_tab, text="Barcode Input", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(input_frame, text="Scan or type serial number:").pack(anchor=tk.W)
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = ttk.Entry(input_frame, textvariable=self.barcode_var, font=('Arial', 12))
        self.barcode_entry.pack(fill=tk.X, pady=(5, 10))
        self.barcode_entry.focus()
        
        # Buttons
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Lookup", command=self.lookup_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Clear", command=self.clear_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="View Excel", command=self.view_excel).pack(side=tk.LEFT)
        
        # Bind Enter key
        self.barcode_entry.bind('<Return>', lambda e: self.lookup_data())
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-s>', lambda e: self.save_settings())
        self.root.bind('<Control-S>', lambda e: self.save_settings())
        
        # Action buttons in Main tab
        action_frame = ttk.LabelFrame(main_tab, text="Actions", padding="10")
        action_frame.pack(fill=tk.X)
        
        ttk.Button(action_frame, text="Update Preview", command=self.update_preview).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_frame, text="Save Label", command=self.save_label).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_frame, text="Print", command=self.print_label).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select Excel file and enter serial number for range lookup")
        status_bar = ttk.Label(main_tab, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))

        # Label controls in Settings tab
        controls_frame = ttk.LabelFrame(settings_tab, text="Label Settings", padding="10")
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.setup_label_controls(controls_frame)
        
        # Settings management buttons in Settings tab
        settings_mgmt_frame = ttk.LabelFrame(settings_tab, text="Settings Management", padding="5")
        settings_mgmt_frame.pack(fill=tk.X, pady=(5, 0))
        
        # First row of buttons
        settings_btn_frame1 = ttk.Frame(settings_mgmt_frame)
        settings_btn_frame1.pack(fill=tk.X, pady=(0, 2))
        
        ttk.Button(settings_btn_frame1, text="Save (Ctrl+S)", command=self.save_settings, width=12).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(settings_btn_frame1, text="Load", command=self.load_and_apply_settings, width=10).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(settings_btn_frame1, text="Reset", command=self.reset_settings, width=10).pack(side=tk.LEFT)
        
        # Settings status on second row
        self.settings_status_var = tk.StringVar()
        settings_status_label = ttk.Label(settings_mgmt_frame, textvariable=self.settings_status_var, 
                                        font=('Arial', 8), foreground='blue')
        settings_status_label.pack(pady=(2, 0))
        
        # Set initial settings status
        if os.path.exists(self.settings_file):
            self.settings_status_var.set("✓ Settings loaded from saved file")
        else:
            self.settings_status_var.set("Using default settings")
    
    def setup_label_controls(self, parent):
        """Setup label adjustment controls"""
        # Create a scrollable frame for controls - reduced height for smaller screens
        canvas = tk.Canvas(parent, height=150)
        scrollbar_ctrl = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_ctrl.set)

        # Add mouse wheel scrolling - cross-platform
        def _on_mousewheel(event):
            # Different platforms use different delta values
            if event.delta:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            else:
                # For Linux/Unix systems
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")

        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)  # Linux
            canvas.bind_all("<Button-5>", _on_mousewheel)  # Linux
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)

        # Label dimensions (83mm x 32mm) - more compact
        dims_frame = ttk.LabelFrame(scrollable_frame, text="Dimensions (83mm x 32mm)", padding="3")
        dims_frame.pack(fill=tk.X, pady=(0, 3))

        ttk.Label(dims_frame, text="Width:").grid(row=0, column=0, sticky=tk.W)
        self.width_var = tk.IntVar(value=self.label_settings['width'])
        ttk.Scale(dims_frame, from_=300, to=700, variable=self.width_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(dims_frame, textvariable=self.width_var).grid(row=0, column=2)

        ttk.Label(dims_frame, text="Height:").grid(row=1, column=0, sticky=tk.W)
        self.height_var = tk.IntVar(value=self.label_settings['height'])
        ttk.Scale(dims_frame, from_=150, to=350, variable=self.height_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW)
        ttk.Label(dims_frame, textvariable=self.height_var).grid(row=1, column=2)

        dims_frame.columnconfigure(1, weight=1)

        # Position controls for new label format - more compact
        pos_frame = ttk.LabelFrame(scrollable_frame, text="Positions", padding="3")
        pos_frame.pack(fill=tk.X, pady=(0, 3))

        # Logo position
        ttk.Label(pos_frame, text="Logo X:").grid(row=0, column=0, sticky=tk.W)
        self.logo_x_var = tk.IntVar(value=self.label_settings['logo_x'])
        ttk.Scale(pos_frame, from_=0, to=200, variable=self.logo_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW)
        
        ttk.Label(pos_frame, text="Logo Y:").grid(row=1, column=0, sticky=tk.W)
        self.logo_y_var = tk.IntVar(value=self.label_settings['logo_y'])
        ttk.Scale(pos_frame, from_=0, to=100, variable=self.logo_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW)

        # P/D position
        ttk.Label(pos_frame, text="P/D X:").grid(row=2, column=0, sticky=tk.W)
        self.pd_x_var = tk.IntVar(value=self.label_settings['pd_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.pd_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=2, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="P/D Y:").grid(row=3, column=0, sticky=tk.W)
        self.pd_y_var = tk.IntVar(value=self.label_settings['pd_y'])
        ttk.Scale(pos_frame, from_=0, to=250, variable=self.pd_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=3, column=1, sticky=tk.EW)

        # P/N position
        ttk.Label(pos_frame, text="P/N X:").grid(row=4, column=0, sticky=tk.W)
        self.pn_x_var = tk.IntVar(value=self.label_settings['pn_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.pn_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=4, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="P/N Y:").grid(row=5, column=0, sticky=tk.W)
        self.pn_y_var = tk.IntVar(value=self.label_settings['pn_y'])
        ttk.Scale(pos_frame, from_=0, to=250, variable=self.pn_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=5, column=1, sticky=tk.EW)

        # P/R position
        ttk.Label(pos_frame, text="P/R X:").grid(row=6, column=0, sticky=tk.W)
        self.pr_x_var = tk.IntVar(value=self.label_settings['pr_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.pr_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=6, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="P/R Y:").grid(row=7, column=0, sticky=tk.W)
        self.pr_y_var = tk.IntVar(value=self.label_settings['pr_y'])
        ttk.Scale(pos_frame, from_=0, to=250, variable=self.pr_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=7, column=1, sticky=tk.EW)

        # S/N position
        ttk.Label(pos_frame, text="S/N X:").grid(row=8, column=0, sticky=tk.W)
        self.sn_x_var = tk.IntVar(value=self.label_settings['sn_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.sn_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=8, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="S/N Y:").grid(row=9, column=0, sticky=tk.W)
        self.sn_y_var = tk.IntVar(value=self.label_settings['sn_y'])
        ttk.Scale(pos_frame, from_=0, to=250, variable=self.sn_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=9, column=1, sticky=tk.EW)

        pos_frame.columnconfigure(1, weight=1)

        # Logo settings - more compact
        logo_frame = ttk.LabelFrame(scrollable_frame, text="Logo", padding="3")
        logo_frame.pack(fill=tk.X, pady=(0, 3))

        ttk.Label(logo_frame, text="Logo File:").grid(row=0, column=0, sticky=tk.W)
        self.logo_path_var = tk.StringVar(value=self.label_settings['logo_path'] or "No logo selected")
        logo_entry = ttk.Entry(logo_frame, textvariable=self.logo_path_var, state='readonly')
        logo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(5, 0))

        ttk.Button(logo_frame, text="Browse", command=self.browse_logo).grid(row=0, column=2, padx=(5, 0))
        ttk.Button(logo_frame, text="Clear", command=self.clear_logo).grid(row=0, column=3, padx=(5, 0))

        # Logo size controls
        ttk.Label(logo_frame, text="Logo Width:").grid(row=1, column=0, sticky=tk.W)
        self.logo_width_var = tk.IntVar(value=self.label_settings['logo_width'])
        ttk.Scale(logo_frame, from_=50, to=300, variable=self.logo_width_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(logo_frame, textvariable=self.logo_width_var).grid(row=1, column=3)

        ttk.Label(logo_frame, text="Logo Height:").grid(row=2, column=0, sticky=tk.W)
        self.logo_height_var = tk.IntVar(value=self.label_settings['logo_height'])
        ttk.Scale(logo_frame, from_=20, to=100, variable=self.logo_height_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=2, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(logo_frame, textvariable=self.logo_height_var).grid(row=2, column=3)

        logo_frame.columnconfigure(1, weight=1)

        # Barcode settings - more compact
        barcode_frame = ttk.LabelFrame(scrollable_frame, text="Barcode", padding="3")
        barcode_frame.pack(fill=tk.X, pady=(0, 3))

        ttk.Label(barcode_frame, text="Barcode Width:").grid(row=0, column=0, sticky=tk.W)
        self.barcode_width_var = tk.IntVar(value=self.label_settings['barcode_width'])
        ttk.Scale(barcode_frame, from_=200, to=450, variable=self.barcode_width_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(barcode_frame, textvariable=self.barcode_width_var).grid(row=0, column=3)

        ttk.Label(barcode_frame, text="Barcode Height:").grid(row=1, column=0, sticky=tk.W)
        self.barcode_height_var = tk.IntVar(value=self.label_settings['barcode_height'])
        ttk.Scale(barcode_frame, from_=15, to=60, variable=self.barcode_height_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(barcode_frame, textvariable=self.barcode_height_var).grid(row=1, column=3)

        barcode_frame.columnconfigure(1, weight=1)

        # Font size settings - more compact
        font_frame = ttk.LabelFrame(scrollable_frame, text="Font Sizes", padding="3")
        font_frame.pack(fill=tk.X, pady=(0, 3))

        ttk.Label(font_frame, text="Company Font:").grid(row=0, column=0, sticky=tk.W)
        self.font_company_size_var = tk.IntVar(value=self.label_settings.get('font_company_size', 14))
        ttk.Scale(font_frame, from_=8, to=24, variable=self.font_company_size_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(font_frame, textvariable=self.font_company_size_var).grid(row=0, column=3)

        ttk.Label(font_frame, text="Label Font (P/D, P/N):").grid(row=1, column=0, sticky=tk.W)
        self.font_label_size_var = tk.IntVar(value=self.label_settings.get('font_label_size', 10))
        ttk.Scale(font_frame, from_=6, to=18, variable=self.font_label_size_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(font_frame, textvariable=self.font_label_size_var).grid(row=1, column=3)

        ttk.Label(font_frame, text="Data Font:").grid(row=2, column=0, sticky=tk.W)
        self.font_data_size_var = tk.IntVar(value=self.label_settings.get('font_data_size', 9))
        ttk.Scale(font_frame, from_=6, to=16, variable=self.font_data_size_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=2, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(font_frame, textvariable=self.font_data_size_var).grid(row=2, column=3)

        ttk.Label(font_frame, text="DLM Font:").grid(row=3, column=0, sticky=tk.W)
        self.font_dlm_size_var = tk.IntVar(value=self.label_settings.get('font_dlm_size', 8))
        ttk.Scale(font_frame, from_=5, to=14, variable=self.font_dlm_size_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=3, column=1, sticky=tk.EW, columnspan=2)
        ttk.Label(font_frame, textvariable=self.font_dlm_size_var).grid(row=3, column=3)

        font_frame.columnconfigure(1, weight=1)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar_ctrl.pack(side='right', fill='y')
    
    def setup_right_panel(self, parent):
        """Setup right preview panel"""
        # Preview section
        preview_frame = ttk.LabelFrame(parent, text="Label Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas for preview - more responsive
        self.preview_canvas = tk.Canvas(preview_frame, bg='white', width=400, height=220)
        self.preview_canvas.pack(expand=True, fill=tk.BOTH)
        
        # Preview info
        info_frame = ttk.Frame(preview_frame)
        info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.preview_info = ttk.Label(info_frame, text="Preview will update automatically")
        self.preview_info.pack()
    
    def on_setting_change(self, *args):
        """Called when any setting changes"""
        self.update_label_settings()
        self.update_preview()
    

    
    def update_label_settings(self):
        """Update internal label settings from UI"""
        self.label_settings.update({
            'width': self.width_var.get(),
            'height': self.height_var.get(),
            'logo_path': self.logo_path_var.get() if self.logo_path_var.get() != "No logo selected" else None,
            'logo_x': self.logo_x_var.get(),
            'logo_y': self.logo_y_var.get(),
            'logo_width': self.logo_width_var.get(),
            'logo_height': self.logo_height_var.get(),
            'pd_x': self.pd_x_var.get(),
            'pd_y': self.pd_y_var.get(),
            'pn_x': self.pn_x_var.get(),
            'pn_y': self.pn_y_var.get(),
            'pr_x': self.pr_x_var.get(),
            'pr_y': self.pr_y_var.get(),
            'sn_x': self.sn_x_var.get(),
            'sn_y': self.sn_y_var.get(),
            'barcode_width': self.barcode_width_var.get(),
            'barcode_height': self.barcode_height_var.get(),
            'font_company_size': self.font_company_size_var.get(),
            'font_label_size': self.font_label_size_var.get(),
            'font_data_size': self.font_data_size_var.get(),
            'font_dlm_size': self.font_dlm_size_var.get()
        })
    
    def browse_logo(self):
        """Browse for logo image file"""
        filename = filedialog.askopenfilename(
            title="Select Logo Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.logo_path_var.set(filename)
            self.label_settings['logo_path'] = filename
            self.update_preview()
    
    def clear_logo(self):
        """Clear the selected logo"""
        self.logo_path_var.set("No logo selected")
        self.label_settings['logo_path'] = None
        self.update_preview()
    
    def browse_excel(self):
        """Browse for Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
    
    def load_selected_excel(self):
        """Load the selected Excel file"""
        self.excel_file = self.excel_path_var.get()
        self.load_excel()
        if self.df is not None:
            self.status_var.set(f"Loaded Excel: {os.path.basename(self.excel_file)}")
        else:
            self.status_var.set("Failed to load Excel file")
    
    def lookup_data(self):
        """Range-based lookup - check if serial number is between SL.From and SL.End"""
        if self.df is None:
            messagebox.showerror("Error", "Excel file not loaded!")
            return
        
        serial_number = self.barcode_var.get().strip()
        if not serial_number:
            messagebox.showwarning("Warning", "Please enter a serial number!")
            return
        
        self.status_var.set(f"Searching for serial number: {serial_number}")
        
        # Find the SL.From and SL.End columns
        sl_from_col = self.find_column(['SL.From', 'SL From', 'SL_From', 'Serial From', 'From'])
        sl_end_col = self.find_column(['SL.End', 'SL End', 'SL_End', 'Serial End', 'End', 'To'])
        
        if not sl_from_col or not sl_end_col:
            messagebox.showerror("Error", 
                f"Could not find serial range columns!\n"
                f"Looking for columns like: SL.From, SL From, SL.End, SL End\n"
                f"Available columns: {', '.join(self.df.columns)}")
            return
        
        # Extract numeric part from serial number
        input_serial_num = self.extract_serial_number(serial_number)
        if input_serial_num is None:
            messagebox.showerror("Error", f"Could not extract numeric part from serial number: {serial_number}")
            return
        
        # Search for matching range
        found_rows = []
        
        for idx, row in self.df.iterrows():
            try:
                # Get range values
                from_val = row[sl_from_col]
                end_val = row[sl_end_col]
                
                # Skip rows with empty range values
                if pd.isna(from_val) or pd.isna(end_val):
                    continue
                
                # Extract numeric parts from range
                from_num = self.extract_serial_number(str(from_val))
                end_num = self.extract_serial_number(str(end_val))
                
                if from_num is None or end_num is None:
                    continue
                
                # Check if input serial number is within range
                if from_num <= input_serial_num <= end_num:
                    found_rows.append((idx, row))
                    print(f"Found match: {serial_number} ({input_serial_num}) is between {from_val} ({from_num}) and {end_val} ({end_num})")
                
            except Exception as e:
                print(f"Error processing row {idx}: {e}")
                continue
        
        if not found_rows:
            messagebox.showerror("Error", f"No range found for serial number: {serial_number}")
            self.current_excel_data = None
            self.status_var.set(f"No range found for serial: {serial_number}")
            self.update_preview()
            return
        
        # Use first match for label generation
        self.current_excel_data = found_rows[0][1].to_dict()
        self.update_preview()
        self.print_label()
        self.status_var.set(f"Scanned {serial_number} - sent to printer")
        self.barcode_var.set("")
        self.barcode_entry.focus()
    
    def generate_barcode(self, data, width=350, height=35):
        """Generate a clean Code128 barcode using ReportLab's built-in barcode - simple and reliable"""
        try:
            from reportlab.graphics.barcode import code128
            from reportlab.graphics.renderPM import drawToPIL
            from reportlab.graphics.shapes import Drawing
            from reportlab.lib.units import mm
            
            # Create a drawing with the barcode
            drawing = Drawing(width, height)
            
            # Calculate bar width to fit the desired width
            estimated_bars = len(data) * 11 + 35  # Rough estimate including start/stop/check
            target_bar_width = width / estimated_bars
            bar_width = max(0.5, target_bar_width)  # Minimum bar width for readability
            
            # Create the barcode with calculated bar width
            barcode = code128.Code128(data, 
                                     barWidth=bar_width,
                                     barHeight=height,
                                     humanReadable=False,  # We'll add text separately
                                     quiet=0)  # No quiet zones - we control positioning
            
            # Get the actual width of the generated barcode and scale if needed
            actual_width = barcode.width
            if actual_width > 0 and actual_width != width:
                scale_factor = width / actual_width
                barcode = code128.Code128(data, 
                                         barWidth=bar_width * scale_factor,
                                         barHeight=height,
                                         humanReadable=False,
                                         quiet=0)
            
            # Add barcode to drawing
            barcode.x = 0
            barcode.y = 0
            drawing.add(barcode)
            
            # Convert to PIL Image
            pil_img = drawToPIL(drawing, fmt='RGB', dpi=150)
            
            # Ensure it's the right size
            if pil_img.size != (width, height):
                pil_img = pil_img.resize((width, height), Image.Resampling.LANCZOS)
            
            return pil_img
            
        except Exception as e:
            print(f"ReportLab barcode error: {e}")
            # Fall back to simple pattern if ReportLab fails
            return self.generate_simple_barcode(data, width, height)


    def generate_label_image(self):
        """Generate 83mm x 32mm label with P/D, P/N, P/R, S/N fields"""
        settings = self.label_settings
        width = settings['width']
        height = settings['height']
        
        # Create image
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        
        # Load fonts with sizes from settings
        try:
            font_company = ImageFont.truetype("arial.ttf", settings.get('font_company_size', 14))
            font_label = ImageFont.truetype("arial.ttf", settings.get('font_label_size', 10))
            font_data = ImageFont.truetype("arial.ttf", settings.get('font_data_size', 9))
            font_dlm = ImageFont.truetype("arial.ttf", settings.get('font_dlm_size', 8))
        except:
            font_company = ImageFont.load_default()
            font_label = font_company
            font_data = font_company
            font_dlm = font_company
        
        # Draw border
        draw.rectangle([0, 0, width-1, height-1], outline='black', width=1)
        
        # 1. Add logo if available
        if settings['logo_path'] and os.path.exists(settings['logo_path']):
            try:
                logo_img = Image.open(settings['logo_path'])
                
                # Convert to RGB if needed
                if logo_img.mode != 'RGB':
                    logo_img = logo_img.convert('RGB')
                
                # Resize logo to specified dimensions
                logo_resized = logo_img.resize((settings['logo_width'], settings['logo_height']), Image.Resampling.LANCZOS)
                
                # Paste logo on the label
                img.paste(logo_resized, (settings['logo_x'], settings['logo_y']))
                
            except Exception as e:
                print(f"Error loading logo: {e}")
                # Fallback to text if logo fails
                draw.text((settings['logo_x'], settings['logo_y']), "CYIENT DLM", fill='black', font=font_company)
        else:
            # No logo - draw fallback text
            draw.text((settings['logo_x'], settings['logo_y']), "CYIENT DLM", fill='black', font=font_company)
        
        if self.current_excel_data:
            # Get field data from Excel
            pd_data = self.get_field_data(['P/D', 'PD', 'DESCRIPTION', 'DESC', 'PRODUCT'])
            pn_data = self.get_field_data(['P/N', 'PN', 'PART', 'CPN', 'PART_NUMBER'])
            pr_data = self.get_field_data(['P/R', 'PR', 'REVISION', 'REV', 'VERSION'])
            
            # S/N should be the lookup input value (the barcode that was scanned/entered)
            sn_data = self.barcode_var.get().strip() if hasattr(self, 'barcode_var') and self.barcode_var.get().strip() else "CDL2349-1195"
            
            # Default values if not found
            if not pd_data: pd_data = "SCB CCA"
            if not pn_data: pn_data = "CZ5S1000B"
            if not pr_data: pr_data = "02"
            
            # 3. P/D field (NO BARCODE - text only)
            draw.text((settings['pd_x'], settings['pd_y']), "P/D", fill='black', font=font_label)
            draw.text((settings['pd_x'] + 30, settings['pd_y']), pd_data, fill='black', font=font_data)
            
            # 4. P/N field
            draw.text((settings['pn_x'], settings['pn_y']), "P/N", fill='black', font=font_label)
            pn_barcode = self.generate_barcode(pn_data, settings['barcode_width'], settings['barcode_height'])
            if pn_barcode:
                img.paste(pn_barcode, (settings['pn_x'] + 30, settings['pn_y'] + 2))
            else:
                draw.text((settings['pn_x'] + 30, settings['pn_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['pn_x'] + 30, settings['pn_y'] + settings['barcode_height'] + 5), pn_data, fill='black', font=font_data)
            
            # 5. P/R field
            draw.text((settings['pr_x'], settings['pr_y']), "P/R", fill='black', font=font_label)
            pr_barcode = self.generate_barcode(pr_data, settings['barcode_width'], settings['barcode_height'])
            if pr_barcode:
                img.paste(pr_barcode, (settings['pr_x'] + 30, settings['pr_y'] + 2))
            else:
                draw.text((settings['pr_x'] + 30, settings['pr_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['pr_x'] + 30, settings['pr_y'] + settings['barcode_height'] + 5), pr_data, fill='black', font=font_data)
            
            # 6. S/N field
            draw.text((settings['sn_x'], settings['sn_y']), "S/N", fill='black', font=font_label)
            sn_barcode = self.generate_barcode(sn_data, settings['barcode_width'], settings['barcode_height'])
            if sn_barcode:
                img.paste(sn_barcode, (settings['sn_x'] + 30, settings['sn_y'] + 2))
            else:
                draw.text((settings['sn_x'] + 30, settings['sn_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['sn_x'] + 30, settings['sn_y'] + settings['barcode_height'] + 5), sn_data, fill='black', font=font_data)
        
        else:
            # Sample data when no lookup performed
            # Generate sample barcodes for preview (excluding P/D)
            sample_pn_barcode = self.generate_simple_barcode("CZ5S1000B", settings['barcode_width'], settings['barcode_height'])
            sample_pr_barcode = self.generate_simple_barcode("02", settings['barcode_width'], settings['barcode_height'])
            sample_sn_barcode = self.generate_simple_barcode("CDL2349-1195", settings['barcode_width'], settings['barcode_height'])
            
            # P/D (NO BARCODE - text only)
            draw.text((settings['pd_x'], settings['pd_y']), "P/D", fill='black', font=font_label)
            draw.text((settings['pd_x'] + 30, settings['pd_y']), "SCB CCA", fill='black', font=font_data)
            
            # P/N
            draw.text((settings['pn_x'], settings['pn_y']), "P/N", fill='black', font=font_label)
            img.paste(sample_pn_barcode, (settings['pn_x'] + 30, settings['pn_y']))
            draw.text((settings['pn_x'] + 30, settings['pn_y'] + settings['barcode_height'] + 2), "CZ5S1000B", fill='black', font=font_data)
            
            # P/R
            draw.text((settings['pr_x'], settings['pr_y']), "P/R", fill='black', font=font_label)
            img.paste(sample_pr_barcode, (settings['pr_x'] + 30, settings['pr_y']))
            draw.text((settings['pr_x'] + 30, settings['pr_y'] + settings['barcode_height'] + 2), "02", fill='black', font=font_data)
            
            # S/N
            draw.text((settings['sn_x'], settings['sn_y']), "S/N", fill='black', font=font_label)
            img.paste(sample_sn_barcode, (settings['sn_x'] + 30, settings['sn_y']))
            draw.text((settings['sn_x'] + 30, settings['sn_y'] + settings['barcode_height'] + 2), "CDL2349-1195", fill='black', font=font_data)
        
        return img
    
    def find_column(self, possible_names):
        """Find a column that matches one of the possible names"""
        if self.df is None:
            return None
            
        for possible_name in possible_names:
            for col in self.df.columns:
                if possible_name.upper() in str(col).upper():
                    return col
        return None
    
    def extract_serial_number(self, serial_str):
        """Extract numeric part from serial number string"""
        import re
        
        # Remove whitespace
        serial_str = str(serial_str).strip()
        
        # Try different patterns to extract numbers
        patterns = [
            r'(\d+)$',           # Numbers at the end
            r'(\d+)',            # Any numbers
            r'(\d+)-(\d+)',      # Pattern like CDL2349-1195, take the last number
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, serial_str)
            if matches:
                if isinstance(matches[0], tuple):
                    # For patterns like CDL2349-1195, take the last number
                    return int(matches[0][-1])
                else:
                    # Take the last match (most specific)
                    return int(matches[-1])
        
        # If no pattern matches, try to extract any digits and combine them
        digits = re.findall(r'\d', serial_str)
        if digits:
            try:
                return int(''.join(digits))
            except ValueError:
                pass
        
        return None
    
    def get_field_data(self, field_names):
        """Get data for a field from Excel using multiple possible column names"""
        if not self.current_excel_data:
            return None
            
        for field_name in field_names:
            for key, value in self.current_excel_data.items():
                if field_name.upper() in key.upper():
                    return str(value)
        return None
    
    def update_preview(self):
        """Update the label preview"""
        try:
            # Generate label
            self.current_label = self.generate_label_image()
            
            # Clear canvas
            self.preview_canvas.delete("all")
            
            # Convert PIL image to PhotoImage for display
            preview_img = self.current_label.copy()
            
            # Scale to fit canvas while maintaining aspect ratio
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:  # Canvas is ready
                img_width, img_height = preview_img.size
                
                # Calculate scaling
                scale_x = (canvas_width - 20) / img_width
                scale_y = (canvas_height - 20) / img_height
                scale = min(scale_x, scale_y, 1.0)  # Don't scale up
                
                new_width = int(img_width * scale)
                new_height = int(img_height * scale)
                
                preview_img = preview_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                # Convert to PhotoImage
                from PIL import ImageTk
                self.preview_photo = ImageTk.PhotoImage(preview_img)
                
                # Center on canvas
                x = (canvas_width - new_width) // 2
                y = (canvas_height - new_height) // 2
                
                self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.preview_photo)
                
                # Update info
                self.preview_info.config(text=f"Preview: {self.label_settings['width']}x{self.label_settings['height']}px")
            
        except Exception as e:
            print(f"Error updating preview: {e}")
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(225, 125, text=f"Preview Error: {e}", anchor=tk.CENTER)
    
    def save_label(self):
        """Save the current label as PDF"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"output_labels/label_{timestamp}.pdf"
            
            # Generate PDF label
            self.generate_pdf_label(filename)
            
            # Also save PNG preview for reference
            if self.current_label:
                png_filename = f"output_labels/label_{timestamp}.png"
                self.current_label.save(png_filename, 'PNG', dpi=(300, 300))
                
                messagebox.showinfo("Success", f"Label saved as:\nPDF: {filename}\nPNG Preview: {png_filename}")
                self.status_var.set(f"Label saved: {os.path.basename(filename)} + PNG preview")
            else:
                messagebox.showinfo("Success", f"Label saved as PDF:\n{filename}")
                self.status_var.set(f"Label saved: {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving label: {e}")
            print(f"Save error: {e}")  # For debugging
    
    def print_pdf_as_image(self, pdf_path):
        """Convert PDF to image and print using the exact same method as samplepdfprint.py"""
        try:
            # Convert PDF to image first
            try:
                import fitz  # PyMuPDF
                # Open PDF and convert first page to image
                pdf_document = fitz.open(pdf_path)
                page = pdf_document[0]
                # Render page as image with high DPI for printing
                mat = fitz.Matrix(3.0, 3.0)  # 3x zoom for better quality
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                
                # Create PIL Image from the data
                from io import BytesIO
                img = Image.open(BytesIO(img_data))
                pdf_document.close()
                
            except ImportError:
                # PyMuPDF not available, try using pdf2image
                try:
                    from pdf2image import convert_from_path
                    images = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
                    img = images[0]
                except ImportError:
                    return False  # No PDF conversion libraries available
            
            # Now print the image using the exact same method as samplepdfprint.py
            printer_name = win32print.GetDefaultPrinter()
            print(f"Printing to: {printer_name}")
            
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)
            
            # Get printable area
            printable_area = (hDC.GetDeviceCaps(win32con.HORZRES),
                            hDC.GetDeviceCaps(win32con.VERTRES))
            
            # Calculate scaling to fit printable area
            ratio = min(printable_area[0] / img.size[0], printable_area[1] / img.size[1])
            scaled_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
            
            # Resize image
            bmp = img.resize(scaled_size)
            dib = ImageWin.Dib(bmp)
            
            # Print the image
            hDC.StartDoc("Label Print")
            hDC.StartPage()
            
            # Center the image on the page
            x = (printable_area[0] - scaled_size[0]) // 2
            y = (printable_area[1] - scaled_size[1]) // 2
            
            dib.draw(hDC.GetHandleOutput(), (x, y, x + scaled_size[0], y + scaled_size[1]))
            
            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
            
            return True
            
        except Exception as e:
            print(f"PDF to image print error: {e}")
            return False

    def print_pdf_directly(self, pdf_path):
        """Convert PDF to image and print directly using win32print - exactly like samplepdfprint.py"""
        try:
            import os
            import tempfile
            
            # Convert relative path to absolute path
            abs_pdf_path = os.path.abspath(pdf_path)
            
            # Check if file exists
            if not os.path.exists(abs_pdf_path):
                print(f"PDF file not found: {abs_pdf_path}")
                return False
            
            print(f"Converting PDF to image for direct printing: {abs_pdf_path}")
            
            # Convert PDF to image first - try PyMuPDF first (doesn't need Poppler)
            try:
                # Try PyMuPDF first (more reliable on Windows, no external dependencies)
                import fitz
                pdf_document = fitz.open(abs_pdf_path)
                page = pdf_document[0]
                # Render page as image with high DPI for printing
                mat = fitz.Matrix(3.0, 3.0)  # 3x zoom for better quality
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                
                # Create PIL Image from the data
                from io import BytesIO
                img = Image.open(BytesIO(img_data))
                pdf_document.close()
                print("PDF converted to image using PyMuPDF")
                
                # Save the original converted image for verification
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                original_img_path = f"output_labels/printer_image_original_{timestamp}.png"
                img.save(original_img_path, 'PNG', dpi=(300, 300))
                print(f"Original printer image saved: {original_img_path}")
                
            except ImportError:
                try:
                    # Fallback: try pdf2image (requires Poppler)
                    from pdf2image import convert_from_path
                    images = convert_from_path(abs_pdf_path, dpi=300, first_page=1, last_page=1)
                    img = images[0]
                    print("PDF converted to image using pdf2image")
                    
                    # Save the original converted image for verification
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    original_img_path = f"output_labels/printer_image_original_{timestamp}.png"
                    img.save(original_img_path, 'PNG', dpi=(300, 300))
                    print(f"Original printer image saved: {original_img_path}")
                    
                except Exception as e:
                    print(f"pdf2image error (likely missing Poppler): {e}")
                    print("No PDF conversion libraries available or working")
                    return False
            except Exception as e:
                print(f"PyMuPDF error: {e}")
                try:
                    # Fallback: try pdf2image (requires Poppler)
                    from pdf2image import convert_from_path
                    images = convert_from_path(abs_pdf_path, dpi=300, first_page=1, last_page=1)
                    img = images[0]
                    print("PDF converted to image using pdf2image (fallback)")
                    
                    # Save the original converted image for verification
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    original_img_path = f"output_labels/printer_image_original_{timestamp}.png"
                    img.save(original_img_path, 'PNG', dpi=(300, 300))
                    print(f"Original printer image saved: {original_img_path}")
                    
                except Exception as e2:
                    print(f"pdf2image fallback error: {e2}")
                    print("Both PyMuPDF and pdf2image failed")
                    return False
            
            # Now print the image using the EXACT same method as samplepdfprint.py
            printer_name = win32print.GetDefaultPrinter()
            print(f"Printing to: {printer_name}")
            
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)
            
            # Get printable area (same as samplepdfprint.py)
            printable_area = (hDC.GetDeviceCaps(win32con.HORZRES),
                            hDC.GetDeviceCaps(win32con.VERTRES))
            
            # Calculate scaling to fit printable area (same as samplepdfprint.py)
            ratio = min(printable_area[0] / img.size[0], printable_area[1] / img.size[1])
            scaled_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
            
            # Resize image (same as samplepdfprint.py)
            bmp = img.resize(scaled_size)
            dib = ImageWin.Dib(bmp)
            
            # Save the final scaled image that goes to printer for verification
            scaled_img_path = f"output_labels/printer_image_scaled_{timestamp}.png"
            bmp.save(scaled_img_path, 'PNG', dpi=(300, 300))
            print(f"Scaled printer image saved: {scaled_img_path}")
            
            # Print the image (EXACT same code as samplepdfprint.py)
            hDC.StartDoc("PDF Label Print")
            hDC.StartPage()
            
            # Center the image on the page (same as samplepdfprint.py)
            x = (printable_area[0] - scaled_size[0]) // 2
            y = (printable_area[1] - scaled_size[1]) // 2
            
            dib.draw(hDC.GetHandleOutput(), (x, y, x + scaled_size[0], y + scaled_size[1]))
            
            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
            
            print("PDF printed successfully using direct win32print method")
            return True
            
        except Exception as e:
            print(f"Direct PDF print error: {e}")
            return False
    
    def print_label(self):
        """Generate and print label as PDF directly to printer - simple like samplepdfprint.py"""
        try:
            # Generate PDF label with current data
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_filename = f"output_labels/label_{timestamp}.pdf"
            
            # Generate the PDF first
            self.generate_pdf_label(pdf_filename)
            print(f"PDF generated: {pdf_filename}")
            
            # Print PDF directly to default printer (like samplepdfprint.py approach)
            if self.print_pdf_directly(pdf_filename):
                self.status_var.set(f"Label sent to printer: {os.path.basename(pdf_filename)}")
                # No popup message - just continue to next serial number
            else:
                # If printing fails, just notify user that PDF is saved
                self.status_var.set(f"PDF label saved: {pdf_filename} - Please print manually")
                messagebox.showinfo("PDF Generated", 
                                  f"Label saved as PDF: {pdf_filename}\n\nPrint failed - please open file and print manually.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error generating/printing PDF label: {e}")
            self.status_var.set(f"PDF error: {e}")
            print(f"PDF error: {e}")  # For debugging
    
    def view_excel(self):
        """Show Excel file contents"""
        if self.df is None:
            messagebox.showerror("Error", "Excel file not loaded!")
            return
        
        # Create Excel viewer window
        excel_window = tk.Toplevel(self.root)
        excel_window.title("Excel File Contents")
        excel_window.geometry("900x500")
        
        frame = ttk.Frame(excel_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"File: {self.excel_file}", font=('Arial', 12, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame, text=f"Rows: {len(self.df)} | Columns: {len(self.df.columns)}", 
                 font=('Arial', 10)).pack(anchor=tk.W, pady=(0, 10))
        
        # Create treeview to show data in table format
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview
        columns = list(self.df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=50)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Insert data (first 100 rows to avoid performance issues)
        for idx, row in self.df.head(100).iterrows():
            values = [str(row[col]) for col in columns]
            tree.insert('', tk.END, values=values)
        
        # Pack treeview and scrollbars
        tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        if len(self.df) > 100:
            ttk.Label(frame, text=f"Showing first 100 rows of {len(self.df)} total rows", 
                     font=('Arial', 9, 'italic')).pack(pady=(5, 0))
    
    def clear_all(self):
        """Clear all data"""
        self.barcode_var.set("")
        self.results_text.delete(1.0, tk.END)
        self.current_excel_data = None
        self.barcode_entry.focus()
        self.status_var.set("Cleared - Ready for new serial number lookup")
        self.update_preview()
    
    def run(self):
        """Start the application"""
        # Bind canvas resize event to update preview
        self.preview_canvas.bind('<Configure>', lambda e: self.root.after(100, self.update_preview))
        self.root.mainloop()

    def add_logo_to_canvas(self, canvas_obj, logo_path, x_mm, y_mm, width_mm, height_mm):
        """Add a logo image to the canvas at the specified position and size"""
        try:
            # Check if logo file exists
            if not os.path.exists(logo_path):
                print(f"Warning: Logo file not found at {logo_path}")
                return False
            
            # Convert mm to points
            x_pts = x_mm * mm
            y_pts = y_mm * mm
            width_pts = width_mm * mm
            height_pts = height_mm * mm
            
            # Load and draw the image
            canvas_obj.drawImage(logo_path, x_pts, y_pts, width_pts, height_pts)
            
            return True
            
        except Exception as e:
            print(f"Error loading logo from {logo_path}: {e}")
            # Draw a placeholder rectangle if logo fails to load
            canvas_obj.setStrokeColor(black)
            canvas_obj.setFillColor("lightgray")
            canvas_obj.rect(x_mm * mm, y_mm * mm, width_mm * mm, height_mm * mm, fill=1, stroke=1)
            
            # Add text placeholder
            canvas_obj.setFillColor(black)
            canvas_obj.setFont("Helvetica", 8)
            canvas_obj.drawString((x_mm + 2) * mm, (y_mm + height_mm/2) * mm, "LOGO")
            
            return False

    def create_barcode_directly(self, canvas_obj, data, x, y, width_mm, height_mm):
        """Create a barcode directly on the canvas using reportlab's built-in Code128 barcode"""
        try:
            # Convert mm to points
            width_pts = width_mm * mm
            height_pts = height_mm * mm
            
            # Calculate the bar width needed to achieve the desired total width
            # Code128 typically has about 11 bars per character + start/stop patterns
            # This is an approximation - we'll create the barcode and scale if needed
            estimated_bars = len(data) * 11 + 35  # Rough estimate including start/stop/check
            target_bar_width = width_pts / estimated_bars
            
            # Ensure minimum bar width for readability (0.3 points minimum)
            bar_width = max(0.3, target_bar_width)
            
            # Create the barcode with calculated bar width
            # Code128 automatically starts with the correct start pattern (first bar should be black)
            barcode = code128.Code128(data, 
                                     barWidth=bar_width,  # Calculated to achieve target width
                                     barHeight=height_pts,
                                     humanReadable=False,  # We'll add text separately
                                     quiet=0)  # No quiet zones - we control positioning
            
            # Get the actual width of the generated barcode
            actual_width = barcode.width
            
            # If the barcode is too wide or too narrow, scale it
            if actual_width > 0:
                scale_factor = width_pts / actual_width
                
                # Save the current graphics state
                canvas_obj.saveState()
                
                # Apply scaling and draw at the correct position
                canvas_obj.translate(x, y)
                canvas_obj.scale(scale_factor, 1.0)  # Scale width only, keep height
                barcode.drawOn(canvas_obj, 0, 0)
                
                # Restore the graphics state
                canvas_obj.restoreState()
            else:
                # Fallback: draw without scaling
                barcode.drawOn(canvas_obj, x, y)
            
            return True
        except Exception as e:
            print(f"Error creating barcode for '{data}': {e}")
            # Draw a placeholder rectangle if barcode fails
            canvas_obj.setStrokeColor(black)
            canvas_obj.setFillColor(black)
            canvas_obj.rect(x, y, width_mm * mm, height_mm * mm, fill=0, stroke=1)
            return False

    def flip_y(self, y_mm, label_height):
        """Convert top-left Y coordinate to bottom-left for reportlab"""
        return label_height - (y_mm * mm)

    def generate_pdf_label(self, filename=None):
        """Generate label as PDF using exact measurements from debug_label_generator_pdf.py"""
        
        # Label dimensions in mm (converted from pixels)
        label_width_mm = 173  # About 490 pixels
        label_height_mm = 60  # About 170 pixels
        
        # Convert to points for reportlab (1 mm = 2.834645669 points)
        label_width = label_width_mm * mm
        label_height = label_height_mm * mm
        
        # Configuration in mm
        config_mm = {
            'logo_x': 5,     # 14px ≈ 5mm
            'logo_y': 2,     # 6px ≈ 2mm  
            'logo_width': 35,  # Logo width
            'logo_height': 17, # Logo height
            'field_start_x': 45,  # 127px ≈ 45mm
            'text_offset': 15,     # Offset for barcode/text
            'barcode_width': 90,  # 255px ≈ 90mm
            'barcode_height': 8, # 23px ≈ 8mm
            'field_gap': 5.3,     # Gap between fields
            'text_bc_offset': 0,  # Text barcode offset
        }
        
        # Field positions in mm (converted from pixel positions)
        field_positions_mm = {
            'P/D': 6,   # 17px ≈ 6mm
            'P/N': 14,  # 40px ≈ 14mm  
            'P/R': 29,  # 82px ≈ 29mm
            'S/N': 46   # 130px ≈ 46mm
        }
        
        # Create PDF filename if not provided
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"output_labels/label_{timestamp}.pdf"
        
        # Create PDF canvas with exact label size
        c = canvas.Canvas(filename, pagesize=(label_width, label_height))
        
        # Draw border
        c.setStrokeColor(black)
        c.setLineWidth(0.5)
        c.rect(0, 0, label_width, label_height)
        
        # 1. Company logo area
        # Try to load logo from different possible locations
        logo_paths = [
            "logo.png",  # Current directory
            "assets/logo.png",  # Assets folder
            "../logo.png",  # Parent directory
            "assets/logo copy.png"  # Alternative logo
        ]
        
        logo_loaded = False
        if self.label_settings.get('logo_path') and os.path.exists(self.label_settings['logo_path']):
            logo_loaded = self.add_logo_to_canvas(c, self.label_settings['logo_path'], 
                                               config_mm['logo_x'], 
                                               self.flip_y(config_mm['logo_y'] + config_mm['logo_height'], label_height) / mm,
                                               config_mm['logo_width'], 
                                               config_mm['logo_height'])
        else:
            for logo_path in logo_paths:
                if os.path.exists(logo_path):
                    logo_loaded = self.add_logo_to_canvas(c, logo_path, 
                                                       config_mm['logo_x'], 
                                                       self.flip_y(config_mm['logo_y'] + config_mm['logo_height'], label_height) / mm,
                                                       config_mm['logo_width'], 
                                                       config_mm['logo_height'])
                    if logo_loaded:
                        print(f"Logo loaded from: {logo_path}")
                        break
        
        # Fallback to text if logo not found
        if not logo_loaded:
            c.setFont("Helvetica-Bold", 14)
            c.setFillColor(black)
            c.drawString(config_mm['logo_x'] * mm, self.flip_y(config_mm['logo_y'] + 4, label_height), "CYIENT")
            
            c.setFillColor(blue)
            c.drawString((config_mm['logo_x'] + 21) * mm, self.flip_y(config_mm['logo_y'] + 4, label_height), "DLM")
        
        # Get data from current excel data or use defaults
        if self.current_excel_data:
            pd_data = self.get_field_data(['P/D', 'PD', 'DESCRIPTION', 'DESC', 'PRODUCT']) or "SCB CCA"
            pn_data = self.get_field_data(['P/N', 'PN', 'PART', 'CPN', 'PART_NUMBER']) or "CZ5S1000B"
            pr_data = self.get_field_data(['P/R', 'PR', 'REVISION', 'REV', 'VERSION']) or "02"
            sn_data = self.barcode_var.get().strip() if hasattr(self, 'barcode_var') and self.barcode_var.get().strip() else "CDL2349-1195"
        else:
            pd_data = "SCB CCA"
            pn_data = "CZ5S1000B"
            pr_data = "02"
            sn_data = "CDL2349-1195"
        
        # 2. P/D field (text only, no barcode)
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 10)
        pd_y = field_positions_mm['P/D']
        c.drawString(config_mm['field_start_x'] * mm , self.flip_y(pd_y + 3, label_height), "P/D")
        
        c.setFont("Helvetica", 8)
        c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                     self.flip_y(pd_y + 3, label_height), pd_data)
        
        # 3. P/N field (with barcode)
        c.setFont("Helvetica-Bold", 10)
        pn_y = field_positions_mm['P/N']
        c.drawString(config_mm['field_start_x'] * mm, self.flip_y(pn_y + 3, label_height), "P/N")
        
        # Draw P/N barcode
        self.create_barcode_directly(c, pn_data, 
                               (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                               self.flip_y(pn_y + config_mm['barcode_height'] + 1, label_height),
                               config_mm['barcode_width'], config_mm['barcode_height'])
        
        # P/N text below barcode
        c.setFont("Helvetica", 8)
        c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                     self.flip_y(pn_y + config_mm['barcode_height'] + 4, label_height), pn_data)
        
        # 4. P/R field (with barcode)
        c.setFont("Helvetica-Bold", 10)
        pr_y = field_positions_mm['P/R']
        c.drawString(config_mm['field_start_x'] * mm, self.flip_y(pr_y + 3, label_height), "P/R")
        
        # Draw P/R barcode
        self.create_barcode_directly(c, pr_data, 
                               (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                               self.flip_y(pr_y + config_mm['barcode_height'] + 1, label_height),
                               config_mm['barcode_width'], config_mm['barcode_height'])
        
        # P/R text below barcode
        c.setFont("Helvetica", 8)
        c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                     self.flip_y(pr_y + config_mm['barcode_height'] + 4, label_height), pr_data)
        
        # 5. S/N field (with barcode)
        c.setFont("Helvetica-Bold", 10)
        sn_y = field_positions_mm['S/N']
        c.drawString(config_mm['field_start_x'] * mm, self.flip_y(sn_y + 3, label_height), "S/N")
        
        # Draw S/N barcode
        self.create_barcode_directly(c, sn_data, 
                               (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                               self.flip_y(sn_y + config_mm['barcode_height'] + 1, label_height),
                               config_mm['barcode_width'], config_mm['barcode_height'])
        
        # S/N text below barcode
        c.setFont("Helvetica", 8)
        c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                     self.flip_y(sn_y + config_mm['barcode_height'] + 4, label_height), sn_data)
        
        # Save the PDF
        c.save()
        
        print(f"PDF label saved: {filename}")
        return filename

if __name__ == "__main__":
    app = EnhancedBarcodeLabelApp()
    app.run()