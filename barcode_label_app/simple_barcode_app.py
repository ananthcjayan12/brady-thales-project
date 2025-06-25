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
from datetime import datetime

class EnhancedBarcodeLabelApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Enhanced Barcode Scanner & Label Generator")
        self.root.geometry("1200x800")
        
        # Excel file path - default
        self.excel_file = "data/serial_tracker.xlsx"
        self.df = None
        
        # Current data
        self.current_excel_data = None
        self.current_label = None
        
        # Label settings - adjustable
        self.label_settings = {
            'width': 400,
            'height': 200,
            'company_text': 'CYIENT',
            'dlm_text': 'DLM',
            'part_bg_color': 'lightgreen',
            'font_size_large': 16,
            'font_size_medium': 12,
            'font_size_small': 10,
            'qr_size': 80,
            'part_x': 70,
            'part_y': 35,
            'serial_x': 10,
            'serial_y': 100,
            'qty_x': 10,
            'qty_y': 160
        }
        
        # Load Excel file
        self.load_excel()
        
        # Setup UI
        self.setup_ui()
        
        # Create output directory
        os.makedirs("output_labels", exist_ok=True)
        
        # Generate initial preview
        self.update_preview()
    
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
        # Create main paned window
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left panel - Controls
        left_frame = ttk.Frame(main_paned, padding="10")
        main_paned.add(left_frame, weight=1)
        
        # Right panel - Preview
        right_frame = ttk.Frame(main_paned, padding="10")
        main_paned.add(right_frame, weight=1)
        
        self.setup_left_panel(left_frame)
        self.setup_right_panel(right_frame)
    
    def setup_left_panel(self, parent):
        """Setup left control panel"""
        # Title
        title = ttk.Label(parent, text="Barcode Scanner & Label Generator", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=(0, 20))
        
        # Excel file selection
        excel_frame = ttk.LabelFrame(parent, text="Excel File", padding="10")
        excel_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.excel_path_var = tk.StringVar(value=self.excel_file)
        ttk.Label(excel_frame, text="Excel file:").pack(anchor=tk.W)
        
        path_frame = ttk.Frame(excel_frame)
        path_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Entry(path_frame, textvariable=self.excel_path_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="Browse", command=self.browse_excel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(path_frame, text="Load", command=self.load_selected_excel).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Input section
        input_frame = ttk.LabelFrame(parent, text="Barcode Input", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(input_frame, text="Scan or type part number:").pack(anchor=tk.W)
        
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
        
        # Results section
        results_frame = ttk.LabelFrame(parent, text="Found Data", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.results_text = tk.Text(results_frame, height=8, font=('Courier', 9))
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Label controls
        controls_frame = ttk.LabelFrame(parent, text="Label Settings", padding="10")
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.setup_label_controls(controls_frame)
        
        # Action buttons
        action_frame = ttk.LabelFrame(parent, text="Actions", padding="10")
        action_frame.pack(fill=tk.X)
        
        ttk.Button(action_frame, text="Update Preview", command=self.update_preview).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_frame, text="Save Label", command=self.save_label).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_frame, text="Print", command=self.print_label).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select Excel file and enter part number")
        status_bar = ttk.Label(parent, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def setup_label_controls(self, parent):
        """Setup label adjustment controls"""
        # Create a scrollable frame for controls
        canvas = tk.Canvas(parent, height=200)
        scrollbar_ctrl = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_ctrl.set)
        
        # Label dimensions
        dims_frame = ttk.LabelFrame(scrollable_frame, text="Dimensions", padding="5")
        dims_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(dims_frame, text="Width:").grid(row=0, column=0, sticky=tk.W)
        self.width_var = tk.IntVar(value=self.label_settings['width'])
        ttk.Scale(dims_frame, from_=200, to=600, variable=self.width_var, 
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(dims_frame, textvariable=self.width_var).grid(row=0, column=2)
        
        ttk.Label(dims_frame, text="Height:").grid(row=1, column=0, sticky=tk.W)
        self.height_var = tk.IntVar(value=self.label_settings['height'])
        ttk.Scale(dims_frame, from_=100, to=400, variable=self.height_var, 
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW)
        ttk.Label(dims_frame, textvariable=self.height_var).grid(row=1, column=2)
        
        dims_frame.columnconfigure(1, weight=1)
        
        # Position controls
        pos_frame = ttk.LabelFrame(scrollable_frame, text="Positions", padding="5")
        pos_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Part number position
        ttk.Label(pos_frame, text="Part X:").grid(row=0, column=0, sticky=tk.W)
        self.part_x_var = tk.IntVar(value=self.label_settings['part_x'])
        ttk.Scale(pos_frame, from_=0, to=300, variable=self.part_x_var, 
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=0, column=1, sticky=tk.EW)
        
        ttk.Label(pos_frame, text="Part Y:").grid(row=1, column=0, sticky=tk.W)
        self.part_y_var = tk.IntVar(value=self.label_settings['part_y'])
        ttk.Scale(pos_frame, from_=0, to=150, variable=self.part_y_var, 
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=1, column=1, sticky=tk.EW)
        
        # QR code size
        ttk.Label(pos_frame, text="QR Size:").grid(row=2, column=0, sticky=tk.W)
        self.qr_size_var = tk.IntVar(value=self.label_settings['qr_size'])
        ttk.Scale(pos_frame, from_=40, to=120, variable=self.qr_size_var, 
                 orient=tk.HORIZONTAL, command=self.on_setting_change).grid(row=2, column=1, sticky=tk.EW)
        
        pos_frame.columnconfigure(1, weight=1)
        
        # Text settings
        text_frame = ttk.LabelFrame(scrollable_frame, text="Text", padding="5")
        text_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(text_frame, text="Company:").grid(row=0, column=0, sticky=tk.W)
        self.company_var = tk.StringVar(value=self.label_settings['company_text'])
        ttk.Entry(text_frame, textvariable=self.company_var).grid(row=0, column=1, sticky=tk.EW)
        self.company_var.trace_add('write', self.on_text_change)
        
        ttk.Label(text_frame, text="Sub text:").grid(row=1, column=0, sticky=tk.W)
        self.dlm_var = tk.StringVar(value=self.label_settings['dlm_text'])
        ttk.Entry(text_frame, textvariable=self.dlm_var).grid(row=1, column=1, sticky=tk.EW)
        self.dlm_var.trace_add('write', self.on_text_change)
        
        text_frame.columnconfigure(1, weight=1)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_ctrl.pack(side="right", fill="y")
    
    def setup_right_panel(self, parent):
        """Setup right preview panel"""
        # Preview section
        preview_frame = ttk.LabelFrame(parent, text="Label Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas for preview
        self.preview_canvas = tk.Canvas(preview_frame, bg='white', width=450, height=250)
        self.preview_canvas.pack(expand=True)
        
        # Preview info
        info_frame = ttk.Frame(preview_frame)
        info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.preview_info = ttk.Label(info_frame, text="Preview will update automatically")
        self.preview_info.pack()
    
    def on_setting_change(self, *args):
        """Called when any setting changes"""
        self.update_label_settings()
        self.update_preview()
    
    def on_text_change(self, *args):
        """Called when text settings change"""
        self.update_label_settings()
        self.update_preview()
    
    def update_label_settings(self):
        """Update internal label settings from UI"""
        self.label_settings.update({
            'width': self.width_var.get(),
            'height': self.height_var.get(),
            'part_x': self.part_x_var.get(),
            'part_y': self.part_y_var.get(),
            'qr_size': self.qr_size_var.get(),
            'company_text': self.company_var.get(),
            'dlm_text': self.dlm_var.get()
        })
    
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
        """Simple lookup - search in any column for the input"""
        if self.df is None:
            messagebox.showerror("Error", "Excel file not loaded!")
            return
        
        search_term = self.barcode_var.get().strip()
        if not search_term:
            messagebox.showwarning("Warning", "Please enter something to search for!")
            return
        
        self.status_var.set(f"Searching for: {search_term}")
        
        # Search in all string columns for the term
        found_rows = []
        
        for idx, row in self.df.iterrows():
            for col in self.df.columns:
                cell_value = str(row[col]).strip()
                if search_term.upper() in cell_value.upper():
                    found_rows.append((idx, row))
                    break  # Found in this row, move to next row
        
        if not found_rows:
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, f"No data found for: {search_term}\n\n")
            self.current_excel_data = None
            self.status_var.set(f"No data found for: {search_term}")
            self.update_preview()
            return
        
        # Show found data
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"Found {len(found_rows)} matching record(s):\n\n")
        
        for i, (idx, row) in enumerate(found_rows):
            self.results_text.insert(tk.END, f"=== Match {i+1} ===\n")
            for col, val in row.items():
                self.results_text.insert(tk.END, f"{col}: {val}\n")
            self.results_text.insert(tk.END, "\n")
        
        # Use first match for label generation
        self.current_excel_data = found_rows[0][1].to_dict()
        self.status_var.set(f"Found {len(found_rows)} match(es) - Preview updated")
        self.update_preview()
    
    def generate_label_image(self):
        """Generate label image matching your design"""
        settings = self.label_settings
        width = settings['width']
        height = settings['height']
        
        # Create image
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        
        # Load fonts
        try:
            font_large = ImageFont.truetype("arial.ttf", settings['font_size_large'])
            font_medium = ImageFont.truetype("arial.ttf", settings['font_size_medium'])
            font_small = ImageFont.truetype("arial.ttf", settings['font_size_small'])
            font_bold = ImageFont.truetype("arialbd.ttf", settings['font_size_medium'])
        except:
            font_large = ImageFont.load_default()
            font_medium = font_large
            font_small = font_large
            font_bold = font_large
        
        # Draw border
        draw.rectangle([0, 0, width-1, height-1], outline='black', width=2)
        
        # Draw company name (CYIENT)
        draw.text((10, 10), settings['company_text'], fill='black', font=font_large)
        
        # Draw sub text (DLM)
        company_width = draw.textlength(settings['company_text'], font=font_large)
        draw.text((company_width + 15, 15), settings['dlm_text'], fill='black', font=font_small)
        
        if self.current_excel_data:
            # Get part number
            part_number = None
            for key, value in self.current_excel_data.items():
                if 'CPN' in key.upper() or 'PART' in key.upper():
                    part_number = str(value)
                    break
            
            if part_number:
                # Draw part number with green background (like your image)
                part_text = f"P/N : {part_number}"
                text_width = draw.textlength(part_text, font=font_bold)
                
                # Green background box
                box_padding = 5
                draw.rectangle([
                    settings['part_x'], settings['part_y'],
                    settings['part_x'] + text_width + box_padding * 2,
                    settings['part_y'] + 25
                ], fill=settings['part_bg_color'], outline='black')
                
                # Part number text
                draw.text((settings['part_x'] + box_padding, settings['part_y'] + 5), 
                         part_text, fill='black', font=font_bold)
            
            # Get product description
            product_desc = None
            for key, value in self.current_excel_data.items():
                if 'DESC' in key.upper() or 'PRODUCT' in key.upper():
                    product_desc = str(value)
                    break
            
            if product_desc:
                # Draw PID line (like VSYS-BOX CVCS in your image)
                draw.text((10, settings['part_y'] + 35), f"PID : {product_desc[:25]}", 
                         fill='black', font=font_small)
            
            # Get serial number
            serial_number = None
            for key, value in self.current_excel_data.items():
                if 'SERIAL' in key.upper() or 'S/N' in key.upper() or 'SL.' in key.upper():
                    serial_number = str(value)
                    break
            
            if serial_number:
                # Draw serial number section
                draw.text((10, settings['serial_y']), f"S/N :", fill='black', font=font_small)
                
                # Create barcode for serial number
                try:
                    import code128
                    barcode_img = code128.image(serial_number, height=20)
                    barcode_img = barcode_img.resize((150, 20))
                    img.paste(barcode_img, (60, settings['serial_y']))
                except:
                    # Fallback - just draw the serial number
                    draw.text((60, settings['serial_y']), serial_number, fill='black', font=font_small)
                
                # Serial number text below barcode
                draw.text((60, settings['serial_y'] + 25), serial_number, fill='black', font=font_small)
            
            # Get quantity
            qty = None
            for key, value in self.current_excel_data.items():
                if 'QTY' in key.upper() or 'QUANTITY' in key.upper() or 'SIZE' in key.upper():
                    qty = str(value)
                    break
            
            if not qty:
                qty = "1"  # Default
            
            draw.text((10, settings['qty_y']), f"QTY : {qty}", fill='black', font=font_small)
            
            # Add QR code (top right)
            if part_number:
                qr_data = f"P/N:{part_number}"
                if serial_number:
                    qr_data += f"\nS/N:{serial_number}"
                
                qr = qrcode.QRCode(version=1, box_size=3, border=1)
                qr.add_data(qr_data)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white")
                qr_img = qr_img.resize((settings['qr_size'], settings['qr_size']))
                img.paste(qr_img, (width - settings['qr_size'] - 10, 10))
        
        else:
            # Sample data when no lookup performed
            draw.text((settings['part_x'] + 5, settings['part_y'] + 5), 
                     "P/N : Sample Part", fill='black', font=font_bold)
            draw.rectangle([settings['part_x'], settings['part_y'], 
                          settings['part_x'] + 150, settings['part_y'] + 25], 
                         fill=settings['part_bg_color'], outline='black')
            
            draw.text((10, settings['part_y'] + 35), "PID : Sample Product", fill='black', font=font_small)
            draw.text((10, settings['serial_y']), "S/N : Sample Serial", fill='black', font=font_small)
            draw.text((10, settings['qty_y']), "QTY : 1", fill='black', font=font_small)
        
        return img
    
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
            
            # Save to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                self.current_label.save(tmp.name, 'PNG', dpi=(300, 300))
                
                # Open with default program (which should allow printing)
                if platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", tmp.name])
                elif platform.system() == "Windows":
                    os.startfile(tmp.name)
                else:  # Linux
                    subprocess.run(["xdg-open", tmp.name])
                
                messagebox.showinfo("Print", "Label opened in default program for printing")
                self.status_var.set("Label sent to default program for printing")
                
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
        self.status_var.set("Cleared - Ready for new search")
        self.update_preview()
    
    def run(self):
        """Start the application"""
        # Bind canvas resize event to update preview
        self.preview_canvas.bind('<Configure>', lambda e: self.root.after(100, self.update_preview))
        self.root.mainloop()

if __name__ == "__main__":
    app = EnhancedBarcodeLabelApp()
    app.run()