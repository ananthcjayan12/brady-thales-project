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
        'text_offset': 15,     # 25px ≈ 8.8mm
        'barcode_width': 70,  # 280px ≈ 98.8mm
        'barcode_height': 8, # 35px ≈ 12.3mm
        'field_gap': 5.3,     # 15px ≈ 5.3mm
        'text_bc_offset': 0,  # 25px ≈ 8.8mm
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

def create_barcode_width_tests():
    """Create test labels with different barcode widths to demonstrate the effect"""
    
    print("\nCreating barcode width tests...")
    
    # Test different barcode widths in mm
    test_widths = [60, 80, 100, 120, 140]  # Different widths in mm
    
    for width_mm in test_widths:
        filename = f"debug_label_width_test_{width_mm}mm.pdf"
        
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
        c.drawString(5 * mm, flip_y(3), f"Test: Barcode Width {width_mm}mm")
        
        # Test barcode with different width
        c.setFont("Helvetica-Bold", 10)
        c.drawString(5 * mm, flip_y(15), "P/N")
        
        # Create barcode with test width
        create_barcode_directly(c, "CZ5S1000B", 25 * mm, flip_y(25), width_mm, 10)
        
        c.setFont("Helvetica", 8)
        c.drawString(25 * mm, flip_y(28), "CZ5S1000B")
        
        # Also test with longer data
        c.setFont("Helvetica-Bold", 10)
        c.drawString(5 * mm, flip_y(40), "S/N")
        
        create_barcode_directly(c, "CDL2349-1195-EXTENDED", 25 * mm, flip_y(50), width_mm, 10)
        
        c.setFont("Helvetica", 8)
        c.drawString(25 * mm, flip_y(53), "CDL2349-1195-EXTENDED")
        
        c.save()
        print(f"Created: {filename}")

def create_barcode_start_test():
    """Create a test PDF to verify that barcodes always start with a black bar"""
    filename = "debug_barcode_start_test.pdf"
    
    # Create PDF with letter size
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Test various data strings to ensure they all start with black bars
    test_data = [
        "A",
        "1", 
        "CZ5S1000B",
        "02",
        "CDL2349-1195",
        "HELLO",
        "12345",
        "ABC123"
    ]
    
    y_pos = 250
    
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y_pos + 20, "Code128 Barcode Start Test - All should start with BLACK bar")
    
    for i, data in enumerate(test_data):
        c.setFont("Helvetica", 10)
        c.drawString(50, y_pos, f"Data: '{data}'")
        
        # Create barcode - should always start with black bar per Code128 spec
        create_barcode_directly(c, data, 150, y_pos - 5, 60, 8)
        
        # Add note about first bar
        c.setFont("Helvetica", 8)
        c.drawString(220, y_pos, "← First bar should be BLACK")
        
        y_pos -= 25
    
    c.save()
    print(f"Created barcode start test: {filename}")

if __name__ == "__main__":
    print("Debug Label Generator - Direct PDF Output")
    print("========================================")
    
    # Create the main perfect PDF label
    perfect_config = create_perfect_pdf_label()
    
    # Create test labels with different barcode heights
    create_multiple_test_labels()
    
    # Create test labels with different barcode widths
    create_barcode_width_tests()
    
    # Test barcode start with black bar
    create_barcode_start_test()
    
    print("\nDone! Generated files:")
    print("- debug_label_PERFECT_direct.pdf (main output)")
    print("- debug_label_test_height_*.pdf (barcode height tests)")
    print("- debug_label_width_test_*.pdf (barcode width tests)")
    print("- debug_barcode_start_test.pdf (verify first bar is black)")
    print("\nThe PDF files will have maximum print clarity!")
    print("Open the width test files to see how barcode width changes.")
    print("The start test file verifies that all barcodes begin with a black bar per Code128 spec.")
    print("You can adjust the 'barcode_width' value in the config to get the desired width.")
