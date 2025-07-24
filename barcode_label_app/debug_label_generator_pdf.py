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
from reportlab.lib.utils import ImageReader
from PIL import Image
import os

def add_logo_to_canvas(canvas_obj, logo_path, x_mm, y_mm, width_mm, height_mm):
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
        'logo_width': 35,  # Logo width
        'logo_height': 17, # Logo height
        'field_start_x': 45,  # 110px ≈ 38.9mm
        'text_offset': 15,     # 25px ≈ 8.8mm
        'barcode_width': 90,  # 280px ≈ 98.8mm
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
    
    # 1. Company logo area
    # Try to load logo from different possible locations
    logo_paths = [
        "logo.png",  # Current directory
        "assets/logo.png",  # Assets folder
        "../logo.png",  # Parent directory
        "assets/logo copy.png"  # Alternative logo
    ]
    
    logo_loaded = False
    for logo_path in logo_paths:
        if os.path.exists(logo_path):
            logo_loaded = add_logo_to_canvas(c, logo_path, 
                                           config_mm['logo_x'], 
                                           flip_y(config_mm['logo_y'] + config_mm['logo_height']) / mm,  # Convert back to mm
                                           config_mm['logo_width'], 
                                           config_mm['logo_height'])
            if logo_loaded:
                print(f"Logo loaded from: {logo_path}")
                break
    
    # Fallback to text if logo not found
    if not logo_loaded:
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(black)
        c.drawString(config_mm['logo_x'] * mm, flip_y(config_mm['logo_y'] + 4), "CYIENT")
        
        c.setFillColor(blue)
        c.drawString((config_mm['logo_x'] + 21) * mm, flip_y(config_mm['logo_y'] + 4), "DLM")
    
    # Add logo image (uncomment to use)
    # logo_path = "path/to/your/logo.png"
    # add_logo_to_canvas(c, logo_path, config_mm['logo_x'], config_mm['logo_y'], 50, 20)
    
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


if __name__ == "__main__":
    print("Debug Label Generator - Direct PDF Output")
    print("========================================")
    
    # Create the main perfect PDF label
    perfect_config = create_perfect_pdf_label()
    
 
    print("\nDone! Generated files:")
    print("- debug_label_PERFECT_direct.pdf (main output with logo)")
    print("- debug_logo_test_*.pdf (logo size tests)")
    print("\nThe PDF files will have maximum print clarity!")
    print("Check the logo test files to find the best logo size.")
    print("You can adjust the 'logo_width' and 'logo_height' values in the config.")
    print("Logo will be loaded from: logo.png, assets/logo.png, or fallback to text.")
