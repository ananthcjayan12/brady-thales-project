#!/usr/bin/env python3
"""
Debug Label Generator - Create exact label format as expected using direct PDF generation
This script generates labels directly as PDF using reportlab for maximum print clarity
"""

from reportlab.pdfgen import canvas
from reportlab.lib.units import mm, inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.graphics.barcode import code128
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
from reportlab.lib.colors import black, blue
import os

def create_barcode_directly(canvas_obj, data, x, y, width_mm, height_mm):
    """Create a barcode directly on the canvas using reportlab's built-in Code128 barcode"""
    try:
        # Convert mm to points
        width_pts = width_mm * mm
        height_pts = height_mm * mm
        
        # Create the barcode with proper parameters
        barcode = code128.Code128(data, 
                                 barWidth=0.8,  # Thin bar width in points
                                 barHeight=height_pts,
                                 humanReadable=False)  # We'll add text separately
        
        # Draw the barcode directly on the canvas
        barcode.drawOn(canvas_obj, x, y)
        
        return True
    except Exception as e:
        print(f"Error creating barcode for '{data}': {e}")
        # Draw a placeholder rectangle if barcode fails
        canvas_obj.setStrokeColor(black)
        canvas_obj.setFillColor(black)
        canvas_obj.rect(x, y, width_mm * mm, height_mm * mm, fill=0, stroke=1)
        return False

def create_perfect_pdf_label():
    """Create a perfectly aligned label directly as PDF using reportlab"""
    
    print("Creating PDF label with perfect alignment...")
    
    # Label dimensions in mm (converted from pixels: 489px ≈ 172.8mm, 170px ≈ 60mm)
    label_width_mm = 173  # About 489 pixels
    label_height_mm = 60  # About 170 pixels
    
    # Convert to points for reportlab (1 mm = 2.834645669 points)
    label_width = label_width_mm * mm
    label_height = label_height_mm * mm
    
    # Configuration in mm
    config_mm = {
        'logo_x': 5,     # 15px ≈ 5.3mm
        'logo_y': 2,     # 5px ≈ 1.8mm  
        'field_start_x': 65,  # 110px ≈ 38.9mm
        'text_offset': 9,     # 25px ≈ 8.8mm
        'barcode_width': 170,  # 280px ≈ 98.8mm
        'barcode_height': 8, # 35px ≈ 12.3mm
        'field_gap': 5.3,     # 15px ≈ 5.3mm
        'text_bc_offset': 6,  # 25px ≈ 8.8mm
    }
    
    # Field positions in mm (converted from pixel positions)
    field_positions_mm = {
        'P/D': config_mm['logo_y'] + 4,   # 20px ≈ 7.1mm
        'P/N': 14,  # 40px ≈ 14.1mm  
        'P/R': 29,  # 90px ≈ 31.8mm
        'S/N': 46   # 140px ≈ 49.4mm
    }
    
    # Create PDF filename
    filename = "debug_label_PERFECT_direct.pdf"
    
    # Create PDF canvas with exact label size
    c = canvas.Canvas(filename, pagesize=(label_width, label_height))
    
    # Set up coordinate system (reportlab uses bottom-left as origin, we want top-left)
    # We'll flip Y coordinates by subtracting from label height
    
    def flip_y(y_mm):
        """Convert top-left Y coordinate to bottom-left for reportlab"""
        return label_height - (y_mm * mm)
    
    # Draw border
    c.setStrokeColor(black)
    c.setLineWidth(0.5)
    c.rect(0, 0, label_width, label_height)
    
    # 1. Company logo/text area
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(black)
    c.drawString(config_mm['logo_x'] * mm, flip_y(config_mm['logo_y'] + 4), "CYIENT")
    
    c.setFillColor(blue)
    c.drawString((config_mm['logo_x'] + 21) * mm, flip_y(config_mm['logo_y'] + 4), "DLM")
    
    # 2. P/D field (text only, no barcode)
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 10)
    pd_y = field_positions_mm['P/D']
    c.drawString(config_mm['field_start_x'] * mm , flip_y(pd_y + 3), "P/D")
    
    c.setFont("Helvetica", 8)
    c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, flip_y(pd_y + 3), "SCB CCA")
    
    # 3. P/N field (with barcode)
    c.setFont("Helvetica-Bold", 10)
    pn_y = field_positions_mm['P/N']
    c.drawString(config_mm['field_start_x'] * mm, flip_y(pn_y + 3), "P/N")
    
    # Draw P/N barcode
    create_barcode_directly(c, "CZ5S1000B", 
                           (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                           flip_y(pn_y + config_mm['barcode_height'] + 1),
                           config_mm['barcode_width'], config_mm['barcode_height'])
    
    # P/N text below barcode
    c.setFont("Helvetica", 8)
    c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                 flip_y(pn_y + config_mm['barcode_height'] + 4), "CZ5S1000B")
    
    # 4. P/R field (with barcode)
    c.setFont("Helvetica-Bold", 10)
    pr_y = field_positions_mm['P/R']
    c.drawString(config_mm['field_start_x'] * mm, flip_y(pr_y + 3), "P/R")
    
    # Draw P/R barcode
    create_barcode_directly(c, "02", 
                           (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                           flip_y(pr_y + config_mm['barcode_height'] + 1),
                           config_mm['barcode_width'], config_mm['barcode_height'])
    
    # P/R text below barcode
    c.setFont("Helvetica", 8)
    c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                 flip_y(pr_y + config_mm['barcode_height'] + 4), "02")
    
    # 5. S/N field (with barcode)
    c.setFont("Helvetica-Bold", 10)
    sn_y = field_positions_mm['S/N']
    c.drawString(config_mm['field_start_x'] * mm, flip_y(sn_y + 3), "S/N")
    
    # Draw S/N barcode
    create_barcode_directly(c, "CDL2349-1195", 
                           (config_mm['field_start_x'] + config_mm['text_offset']) * mm, 
                           flip_y(sn_y + config_mm['barcode_height'] + 1),
                           config_mm['barcode_width'], config_mm['barcode_height'])
    
    # S/N text below barcode
    c.setFont("Helvetica", 8)
    c.drawString((config_mm['field_start_x'] + config_mm['text_offset']+config_mm['text_bc_offset']) * mm, 
                 flip_y(sn_y + config_mm['barcode_height'] + 4), "CDL2349-1195")
    
    # Save the PDF
    c.save()
    
    print(f"Saved: {filename}")
    print("PDF generated with maximum clarity using reportlab!")
    
    # Return configuration for main app (convert back to pixels for consistency)
    final_config = {
        'width': int(label_width_mm * 2.834),  # Convert mm back to approx pixels
        'height': int(label_height_mm * 2.834),
        'pd_x': int(config_mm['field_start_x'] * 2.834), 
        'pd_y': int(field_positions_mm['P/D'] * 2.834),
        'pn_x': int(config_mm['field_start_x'] * 2.834), 
        'pn_y': int(field_positions_mm['P/N'] * 2.834),
        'pr_x': int(config_mm['field_start_x'] * 2.834), 
        'pr_y': int(field_positions_mm['P/R'] * 2.834),
        'sn_x': int(config_mm['field_start_x'] * 2.834), 
        'sn_y': int(field_positions_mm['S/N'] * 2.834),
        'barcode_width': int(config_mm['barcode_width'] * 2.834), 
        'barcode_height': int(config_mm['barcode_height'] * 2.834),
        'text_offset': int(config_mm['text_offset'] * 2.834)
    }
    
    print("\nConfig for main app (pixel values):")
    print(f'"width": {final_config["width"]},')
    print(f'"height": {final_config["height"]},')
    print(f'"pd_x": {final_config["pd_x"]}, "pd_y": {final_config["pd_y"]},')
    print(f'"pn_x": {final_config["pn_x"]}, "pn_y": {final_config["pn_y"]},')
    print(f'"pr_x": {final_config["pr_x"]}, "pr_y": {final_config["pr_y"]},')
    print(f'"sn_x": {final_config["sn_x"]}, "sn_y": {final_config["sn_y"]},')
    print(f'"barcode_width": {final_config["barcode_width"]}, "barcode_height": {final_config["barcode_height"]}')
    
    return final_config

def create_multiple_test_labels():
    """Create multiple test labels to fine-tune spacing and barcode size"""
    
    print("\nCreating test labels with different barcode sizes...")
    
    # Test different barcode heights
    test_heights = [8, 10, 12, 15]  # in mm
    
    for height_mm in test_heights:
        filename = f"debug_label_test_height_{height_mm}mm.pdf"
        
        # Label dimensions
        label_width = 173 * mm
        label_height = 60 * mm
        
        c = canvas.Canvas(filename, pagesize=(label_width, label_height))
        
        def flip_y(y_mm):
            return label_height - (y_mm * mm)
        
        # Draw border
        c.setStrokeColor(black)
        c.setLineWidth(0.5)
        c.rect(0, 0, label_width, label_height)
        
        # Title
        c.setFont("Helvetica-Bold", 8)
        c.drawString(5 * mm, flip_y(2), f"Test: Barcode Height {height_mm}mm")
        
        # Test barcode with different height
        c.setFont("Helvetica-Bold", 10)
        c.drawString(39 * mm, flip_y(15), "P/N")
        
        # Create barcode with test height
        create_barcode_directly(c, "CZ5S1000B", 48 * mm, flip_y(15 + height_mm), 99, height_mm)
        
        c.setFont("Helvetica", 8)
        c.drawString(48 * mm, flip_y(15 + height_mm + 3), "CZ5S1000B")
        
        c.save()
        print(f"Created: {filename}")

if __name__ == "__main__":
    print("Debug Label Generator - Direct PDF Output")
    print("========================================")
    
    # Create the main perfect PDF label
    perfect_config = create_perfect_pdf_label()
    
    # Create test labels with different barcode heights
    create_multiple_test_labels()
    
    print("\nDone! Generated files:")
    print("- debug_label_PERFECT_direct.pdf (main output)")
    print("- debug_label_test_height_*.pdf (barcode height tests)")
    print("\nThe PDF files will have maximum print clarity!")
    print("Open them to verify the layout and barcode quality.")
