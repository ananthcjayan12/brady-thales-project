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
        
        # Default label settings - 83mm x 32mm label (489px x 189px at 150 DPI)
        self.default_settings = {
            'width': 489,  # 83mm at 150 DPI
            'height': 189, # 32mm at 150 DPI
            'logo_path': self.get_default_logo_path(),
            'logo_x': 15,
            'logo_y': 5,
            'logo_width': 150,
            'logo_height': 40,
            'pd_x': 190,    # P/D field position
            'pd_y': 5,
            'pn_x': 190,    # P/N field position  
            'pn_y': 45,
            'pr_x': 190,    # P/R field position
            'pr_y': 85,
            'sn_x': 190,    # S/N field position
            'sn_y': 125,
            'barcode_width': 280,
            'barcode_height': 25
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
        ttk.Scale(dims_frame, from_=100, to=300, variable=self.height_var,
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
        ttk.Scale(pos_frame, from_=0, to=180, variable=self.pd_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=3, column=1, sticky=tk.EW)

        # P/N position
        ttk.Label(pos_frame, text="P/N X:").grid(row=4, column=0, sticky=tk.W)
        self.pn_x_var = tk.IntVar(value=self.label_settings['pn_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.pn_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=4, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="P/N Y:").grid(row=5, column=0, sticky=tk.W)
        self.pn_y_var = tk.IntVar(value=self.label_settings['pn_y'])
        ttk.Scale(pos_frame, from_=0, to=180, variable=self.pn_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=5, column=1, sticky=tk.EW)

        # P/R position
        ttk.Label(pos_frame, text="P/R X:").grid(row=6, column=0, sticky=tk.W)
        self.pr_x_var = tk.IntVar(value=self.label_settings['pr_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.pr_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=6, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="P/R Y:").grid(row=7, column=0, sticky=tk.W)
        self.pr_y_var = tk.IntVar(value=self.label_settings['pr_y'])
        ttk.Scale(pos_frame, from_=0, to=180, variable=self.pr_y_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=7, column=1, sticky=tk.EW)

        # S/N position
        ttk.Label(pos_frame, text="S/N X:").grid(row=8, column=0, sticky=tk.W)
        self.sn_x_var = tk.IntVar(value=self.label_settings['sn_x'])
        ttk.Scale(pos_frame, from_=0, to=480, variable=self.sn_x_var,
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=8, column=1, sticky=tk.EW)

        ttk.Label(pos_frame, text="S/N Y:").grid(row=9, column=0, sticky=tk.W)
        self.sn_y_var = tk.IntVar(value=self.label_settings['sn_y'])
        ttk.Scale(pos_frame, from_=0, to=180, variable=self.sn_y_var,
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
            'sn_y': self.sn_y_var.get()
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
    
    def generate_barcode(self, data, width=280, height=25):
        """Generate a clean Code128 barcode using treepoem (no text)"""
        # Attempt using treepoem only if Ghostscript is available
        try:
            import shutil
            # Check for Ghostscript executable
            if not (shutil.which('gs') or shutil.which('gswin64c')):
                raise ImportError('Ghostscript not found')
            import treepoem
            # Generate Code128 barcode
            barcode_img = treepoem.generate_barcode(
                barcode_type='code128',
                data=data,
                options={'includetext': False, 'height': 0.5, 'width': 0.02}
            )
            if barcode_img.mode != 'RGB':
                barcode_img = barcode_img.convert('RGB')
            return barcode_img.resize((width, height), Image.Resampling.LANCZOS)
        except Exception as e:
            # Fallback to simple barcode and report error
            print(f"Barcode generator error: {e}")
            self.status_var.set(f"Barcode error: {e}")
            return self.generate_simple_barcode(data, width, height)
    
    def generate_simple_barcode(self, data, width=280, height=25):
        """Fallback: Generate a simple barcode pattern"""
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        
        # Create Code128-like pattern manually
        # This is a simplified pattern generator
        import hashlib
        
        # Use hash to create consistent pattern
        hash_val = hashlib.md5(data.encode()).hexdigest()
        
        # Convert to pattern
        bar_width = 2
        x = 5  # Start with small margin
        
        # Generate bars based on data
        for i in range(0, len(hash_val), 2):
            try:
                hex_val = int(hash_val[i:i+2], 16)
                
                # Create pattern: each hex pair creates different bar patterns
                for bit in range(4):
                    if hex_val & (1 << bit):
                        # Draw black bar
                        draw.rectangle([x, 3, x + bar_width - 1, height - 3], fill='black')
                    x += bar_width
                    
                    if x >= width - 10:  # Leave margin at end
                        break
                
                if x >= width - 10:
                    break
                    
            except (ValueError, IndexError):
                continue
        
        return img
    
    def generate_label_image(self):
        """Generate 83mm x 32mm label with P/D, P/N, P/R, S/N fields"""
        settings = self.label_settings
        width = settings['width']
        height = settings['height']
        
        # Create image
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        
        # Load fonts
        try:
            font_company = ImageFont.truetype("arial.ttf", 14)  # For CYIENT
            font_label = ImageFont.truetype("arial.ttf", 10)    # For P/D, P/N, etc labels
            font_data = ImageFont.truetype("arial.ttf", 9)      # For data text
            font_dlm = ImageFont.truetype("arial.ttf", 8)       # For DLM
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
            
            # 3. P/D field
            draw.text((settings['pd_x'], settings['pd_y']), "P/D", fill='black', font=font_label)
            pd_barcode = self.generate_barcode(pd_data, settings['barcode_width'], settings['barcode_height'])
            if pd_barcode:
                img.paste(pd_barcode, (settings['pd_x'] + 30, settings['pd_y'] + 2))
            else:
                draw.text((settings['pd_x'] + 30, settings['pd_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['pd_x'] + 30, settings['pd_y'] + 28), pd_data, fill='black', font=font_data)
            
            # 4. P/N field
            draw.text((settings['pn_x'], settings['pn_y']), "P/N", fill='black', font=font_label)
            pn_barcode = self.generate_barcode(pn_data, settings['barcode_width'], settings['barcode_height'])
            if pn_barcode:
                img.paste(pn_barcode, (settings['pn_x'] + 30, settings['pn_y'] + 2))
            else:
                draw.text((settings['pn_x'] + 30, settings['pn_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['pn_x'] + 30, settings['pn_y'] + 28), pn_data, fill='black', font=font_data)
            
            # 5. P/R field
            draw.text((settings['pr_x'], settings['pr_y']), "P/R", fill='black', font=font_label)
            pr_barcode = self.generate_barcode(pr_data, settings['barcode_width'], settings['barcode_height'])
            if pr_barcode:
                img.paste(pr_barcode, (settings['pr_x'] + 30, settings['pr_y'] + 2))
            else:
                draw.text((settings['pr_x'] + 30, settings['pr_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['pr_x'] + 30, settings['pr_y'] + 28), pr_data, fill='black', font=font_data)
            
            # 6. S/N field
            draw.text((settings['sn_x'], settings['sn_y']), "S/N", fill='black', font=font_label)
            sn_barcode = self.generate_barcode(sn_data, settings['barcode_width'], settings['barcode_height'])
            if sn_barcode:
                img.paste(sn_barcode, (settings['sn_x'] + 30, settings['sn_y'] + 2))
            else:
                draw.text((settings['sn_x'] + 30, settings['sn_y'] + 2), "|||||||||||||||||||", fill='black', font=font_data)
            draw.text((settings['sn_x'] + 30, settings['sn_y'] + 28), sn_data, fill='black', font=font_data)
        
        else:
            # Sample data when no lookup performed
            # Also generate sample barcodes for preview
            sample_pd_barcode = self.generate_simple_barcode("SCB CCA", settings['barcode_width'], settings['barcode_height'])
            sample_pn_barcode = self.generate_simple_barcode("CZ5S1000B", settings['barcode_width'], settings['barcode_height'])
            sample_pr_barcode = self.generate_simple_barcode("02", settings['barcode_width'], settings['barcode_height'])
            sample_sn_barcode = self.generate_simple_barcode("CDL2349-1195", settings['barcode_width'], settings['barcode_height'])
            
            # P/D
            draw.text((settings['pd_x'], settings['pd_y']), "P/D", fill='black', font=font_label)
            img.paste(sample_pd_barcode, (settings['pd_x'] + 30, settings['pd_y'] + 2))
            draw.text((settings['pd_x'] + 30, settings['pd_y'] + 28), "SCB CCA", fill='black', font=font_data)
            
            # P/N
            draw.text((settings['pn_x'], settings['pn_y']), "P/N", fill='black', font=font_label)
            img.paste(sample_pn_barcode, (settings['pn_x'] + 30, settings['pn_y'] + 2))
            draw.text((settings['pn_x'] + 30, settings['pn_y'] + 28), "CZ5S1000B", fill='black', font=font_data)
            
            # P/R
            draw.text((settings['pr_x'], settings['pr_y']), "P/R", fill='black', font=font_label)
            img.paste(sample_pr_barcode, (settings['pr_x'] + 30, settings['pr_y'] + 2))
            draw.text((settings['pr_x'] + 30, settings['pr_y'] + 28), "02", fill='black', font=font_data)
            
            # S/N
            draw.text((settings['sn_x'], settings['sn_y']), "S/N", fill='black', font=font_label)
            img.paste(sample_sn_barcode, (settings['sn_x'] + 30, settings['sn_y'] + 2))
            draw.text((settings['sn_x'] + 30, settings['sn_y'] + 28), "CDL2349-1195", fill='black', font=font_data)
        
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
        """Save the current label"""
        if not self.current_label:
            messagebox.showwarning("Warning", "No label to save!")
            return
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"output_labels/label_{timestamp}.png"
            
            self.current_label.save(filename, 'PNG', dpi=(300, 300))
            
            messagebox.showinfo("Success", f"Label saved as:\n{filename}")
            self.status_var.set(f"Label saved: {os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving label: {e}")
    
    def print_label(self):
        """Print the current label"""
        if not self.current_label:
            messagebox.showwarning("Warning", "No label to print!")
            return
        
        try:
            import tempfile
            import subprocess
            import platform
            import shutil
            
            # Save to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                self.current_label.save(tmp.name, 'PNG', dpi=(300, 300))
                
                # Send directly to default printer on all platforms
                if platform.system() == "Windows":
                    # Try Win32 direct printing, fallback to PDF print on ImportError
                    try:
                        import win32print, win32ui, win32con
                        from PIL import ImageWin

                        printer_name = win32print.GetDefaultPrinter()
                        hDC = win32ui.CreateDC()
                        hDC.CreatePrinterDC(printer_name)
                        hDC.StartDoc("Label")
                        hDC.StartPage()

                        dib = ImageWin.Dib(self.current_label)
                        dib.draw(hDC.GetHandleOutput(), (0, 0, self.current_label.width, self.current_label.height))

                        hDC.EndPage()
                        hDC.EndDoc()
                        hDC.DeleteDC()
                    except ImportError as e:
                        messagebox.showerror("Error", f"Windows PDF print failure import: {e}")
                        # Fallback: save as PDF and use default print verb
                        try:
                            pdf_tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
                            self.current_label.save(pdf_tmp.name, 'PDF', dpi=(300,300))
                            os.startfile(pdf_tmp.name, 'print')
                        except Exception as e:
                            messagebox.showerror("Error", f"Windows PDF print failure: {e}")
                    except Exception as e:
                        messagebox.showerror("Error", f"Windows print failure: {e}")
                else:
                    subprocess.run(["lp", tmp.name], check=True)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error printing label: {e}")
    
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

if __name__ == "__main__":
    app = EnhancedBarcodeLabelApp()
    app.run()