#!/usr/bin/env python3
"""
Simple Barcode Scanner & Label Generator
One file - does everything simply!
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
import os
from datetime import datetime

class SimpleBarcodeLabelApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Simple Barcode Scanner & Label Generator")
        self.root.geometry("800x600")
        
        # Excel file path
        self.excel_file = "data/serial_tracker.xlsx"
        self.df = None
        
        # Current data
        self.current_excel_data = None
        self.current_label = None
        
        # Load Excel file
        self.load_excel()
        
        # Setup UI
        self.setup_ui()
        
        # Create output directory
        os.makedirs("output_labels", exist_ok=True)
    
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
        """Setup simple UI"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title = ttk.Label(main_frame, text="Barcode Scanner & Label Generator", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=(0, 20))
        
        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="Barcode Input", padding="15")
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(input_frame, text="Scan or type part number:").pack(anchor=tk.W)
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = ttk.Entry(input_frame, textvariable=self.barcode_var, 
                                      font=('Arial', 12), width=50)
        self.barcode_entry.pack(fill=tk.X, pady=(5, 10))
        self.barcode_entry.focus()
        
        # Buttons
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Lookup", command=self.lookup_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear", command=self.clear_all).pack(side=tk.LEFT)
        
        # Bind Enter key
        self.barcode_entry.bind('<Return>', lambda e: self.lookup_data())
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Found Data", padding="15")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Text widget to show found data
        self.results_text = tk.Text(results_frame, height=10, font=('Courier', 10))
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Label generation section
        label_frame = ttk.LabelFrame(main_frame, text="Label Generation", padding="15")
        label_frame.pack(fill=tk.X)
        
        button_frame2 = ttk.Frame(label_frame)
        button_frame2.pack()
        
        ttk.Button(button_frame2, text="Generate Label", command=self.generate_label).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame2, text="Save Label", command=self.save_label).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame2, text="View Excel", command=self.view_excel).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Enter part number and press Lookup")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
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
            self.results_text.insert(tk.END, "Available data in Excel:\n")
            
            # Show first few rows as examples
            for i, row in self.df.head().iterrows():
                self.results_text.insert(tk.END, f"Row {i}: {dict(row)}\n")
            
            self.current_excel_data = None
            self.status_var.set(f"No data found for: {search_term}")
            return
        
        # Show found data
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"Found {len(found_rows)} matching record(s):\n\n")
        
        for i, (idx, row) in enumerate(found_rows):
            self.results_text.insert(tk.END, f"=== Match {i+1} (Row {idx}) ===\n")
            for col, val in row.items():
                self.results_text.insert(tk.END, f"{col}: {val}\n")
            self.results_text.insert(tk.END, "\n")
        
        # Use first match for label generation
        self.current_excel_data = found_rows[0][1].to_dict()
        self.status_var.set(f"Found {len(found_rows)} match(es) - Ready to generate label")
    
    def generate_label(self):
        """Generate a simple label"""
        if not self.current_excel_data:
            messagebox.showwarning("Warning", "Please lookup data first!")
            return
        
        try:
            # Create label image
            width, height = 400, 200
            img = Image.new('RGB', (width, height), 'white')
            draw = ImageDraw.Draw(img)
            
            # Use default font
            try:
                font_large = ImageFont.truetype("arial.ttf", 16)
                font_medium = ImageFont.truetype("arial.ttf", 12)
                font_small = ImageFont.truetype("arial.ttf", 10)
            except:
                font_large = ImageFont.load_default()
                font_medium = ImageFont.load_default()
                font_small = ImageFont.load_default()
            
            # Draw border
            draw.rectangle([0, 0, width-1, height-1], outline='black', width=2)
            
            # Draw THALES header
            draw.text((10, 10), "THALES", fill='black', font=font_large)
            
            y_pos = 40
            
            # Draw available data
            for key, value in self.current_excel_data.items():
                if pd.notna(value) and str(value).strip():
                    text = f"{key}: {str(value)[:30]}"  # Limit length
                    draw.text((10, y_pos), text, fill='black', font=font_small)
                    y_pos += 15
                    if y_pos > height - 30:  # Don't go off the label
                        break
            
            # Add QR code if we have a part number
            part_number = None
            for key, value in self.current_excel_data.items():
                if 'CPN' in key.upper() or 'PART' in key.upper():
                    part_number = str(value)
                    break
            
            if part_number:
                qr_data = f"Part: {part_number}"
                qr = qrcode.QRCode(version=1, box_size=3, border=1)
                qr.add_data(qr_data)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white")
                qr_img = qr_img.resize((60, 60))
                img.paste(qr_img, (width-70, 10))
            
            self.current_label = img
            
            # Show success message
            messagebox.showinfo("Success", "Label generated successfully!\nClick 'Save Label' to save it.")
            self.status_var.set("Label generated - Ready to save")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating label: {e}")
            self.status_var.set(f"Error generating label: {e}")
    
    def save_label(self):
        """Save the generated label"""
        if not self.current_label:
            messagebox.showwarning("Warning", "Please generate a label first!")
            return
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"output_labels/label_{timestamp}.png"
            
            self.current_label.save(filename, 'PNG', dpi=(300, 300))
            
            messagebox.showinfo("Success", f"Label saved as:\n{filename}")
            self.status_var.set(f"Label saved: {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving label: {e}")
    
    def view_excel(self):
        """Show what's in the Excel file"""
        if self.df is None:
            messagebox.showerror("Error", "Excel file not loaded!")
            return
        
        # Open a new window to show Excel data
        excel_window = tk.Toplevel(self.root)
        excel_window.title("Excel File Contents")
        excel_window.geometry("800x400")
        
        frame = ttk.Frame(excel_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Excel file: {self.excel_file}", font=('Arial', 12, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame, text=f"Total rows: {len(self.df)}", font=('Arial', 10)).pack(anchor=tk.W, pady=(0, 10))
        
        # Text widget to show data
        text_widget = tk.Text(frame, font=('Courier', 9))
        scrollbar_excel = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar_excel.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_excel.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Show column headers
        text_widget.insert(tk.END, "COLUMNS:\n")
        text_widget.insert(tk.END, f"{list(self.df.columns)}\n\n")
        
        # Show first 20 rows
        text_widget.insert(tk.END, "FIRST 20 ROWS:\n")
        for idx, row in self.df.head(20).iterrows():
            text_widget.insert(tk.END, f"Row {idx}: {dict(row)}\n")
    
    def clear_all(self):
        """Clear everything"""
        self.barcode_var.set("")
        self.results_text.delete(1.0, tk.END)
        self.current_excel_data = None
        self.current_label = None
        self.barcode_entry.focus()
        self.status_var.set("Cleared - Ready for new search")
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = SimpleBarcodeLabelApp()
    app.run()