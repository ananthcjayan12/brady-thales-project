#!/usr/bin/env python3
"""
Debug Label Generator - Create exact label format as expected using direct PDF generation
This script generates labels directly as PDF using reportlab for maximum print clarity
"""

from reportlab.pdfgen import canvas
from reportlab.lib.units import mm, inch
from reportlab.lib.pagesizes import letter
from reportlab.graphics.barcode import code128
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
import hashlib
import os
import platform

def load_font(size):
    """Load the best available font for the current system"""
    system = platform.system()
    
    # Common font paths by OS
    font_paths = []
    
    if system == "Windows":
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "C:/Windows/Fonts/segoeui.ttf"
        ]
    elif system == "Darwin":  # macOS
        font_paths = [
            "/System/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Times.ttc"
        ]
    else:  # Linux and others
        font_paths = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/TTF/arial.ttf",
            "/usr/share/fonts/arial.ttf"
        ]
    
    # Try each font path
    for font_path in font_paths:
        try:
            if os.path.exists(font_path):
                return ImageFont.truetype(font_path, size)
        except:
            continue
    
    # Try generic font names
    generic_fonts = ["arial.ttf", "Arial.ttf", "helvetica.ttf", "Helvetica.ttf"]
    for font_name in generic_fonts:
        try:
            return ImageFont.truetype(font_name, size)
        except:
            continue
    
    # Fall back to default font
    print(f"Warning: Could not load any system fonts, using default font for size {size}")
    return ImageFont.load_default()

def generate_simple_barcode(data, width=280, height=30):
    """Generate a simple barcode pattern with proper bars"""
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # Create Code128-like pattern manually
    import hashlib
    
    # Use hash to create consistent pattern
    hash_val = hashlib.md5(data.encode()).hexdigest()
    
    # Calculate bar dimensions for proper barcode appearance
    margin = 20
    usable_width = width - (2 * margin)
    bar_count = usable_width - 20  # Ensure we have enough bars
    
    if bar_count == 0:
        bar_count = 20  # Minimum bars
        
    narrow_bar = max(1, usable_width // (bar_count * 3))  # Narrow bar width
    wide_bar = narrow_bar * 2  # Wide bar width
    
    x = margin
    
    # Create start pattern (typical for Code128)
    start_pattern = [1, 1, 0, 1, 0, 1, 1, 0]  # Start B pattern
    for bar in start_pattern:
        bar_width = wide_bar if bar else narrow_bar
        if bar:
            draw.rectangle([x, 3, x + bar_width - 1, height - 3], fill='black')
        x += bar_width
        if x >= width - margin:
            break
    
    # Generate data bars based on hash
    for i in range(0, min(len(hash_val), 20), 2):
        if x >= width - margin - 20:  # Leave space for stop pattern
            break
            
        try:
            hex_val = int(hash_val[i:i+2], 16)
            
            # Create alternating bar pattern based on hex value
            for bit in range(4):
                bar_width = narrow_bar if (hex_val & (1 << bit)) else wide_bar
                if hex_val & (1 << bit):
                    draw.rectangle([x, 3, x + bar_width - 1, height - 3], fill='black')
                x += bar_width
                if x >= width - margin - 20:
                    break
                    
        except (ValueError, IndexError):
            continue
    
    # Add stop pattern
    if x < width - margin:
        stop_pattern = [1, 1, 0, 0, 1, 1, 1]  # Stop pattern
        for bar in stop_pattern:
            if x >= width - 5:
                break
            bar_width = narrow_bar
            if bar:
                draw.rectangle([x, 3, x + bar_width - 1, height - 3], fill='black')
            x += bar_width
    
    return img



def create_perfect_alignment_label():
    """Create a label with perfect alignment and field_vertical_gap = 15"""
    
    print("Creating perfectly aligned label with field_vertical_gap = 15...")
    
    # Perfect alignment configuration
    config = {
        'width': 489,
        'height': 170,
        'logo_x': 15, 'logo_y': 5,
        'field_start_x': 110,
        'barcode_width': 280, 
        'barcode_height': 35,  # Increased from 20 to 35
        'text_offset': 25,
        'field_vertical_gap': 15
    }
    
    # Manual positioning for perfect alignment (adjusted for larger barcode height)
    field_positions = {
        'P/D': 20,   # Start position
        'P/N': 40,   # P/D + 15 gap
        'P/R': 90,   # P/N + barcode block (35 height + 3 text gap + 10 text + 2 spacing) + 15 gap
        'S/N': 140   # P/R + barcode block + 15 gap
    }
    
    # Create image
    img = Image.new('RGB', (config['width'], config['height']), 'white')
    draw = ImageDraw.Draw(img)
    
    # Load fonts
    font_company = load_font(14)
    font_label = load_font(10)
    font_data = load_font(8)
    
    # Draw border
    draw.rectangle([0, 0, config['width']-1, config['height']-1], outline='black', width=1)
    
    # 1. Company logo/text area
    draw.text((config['logo_x'], config['logo_y']), "CYIENT", fill='black', font=font_company)
    draw.text((config['logo_x'] + 60, config['logo_y']), "DLM", fill='blue', font=font_company)
    
    # 2. P/D field (NO BARCODE - text only, aligned)
    pd_y = field_positions['P/D']
    draw.text((config['field_start_x'], pd_y), "P/D", fill='black', font=font_label)
    draw.text((config['field_start_x'] + config['text_offset'], pd_y), "SCB CCA", fill='black', font=font_data)
    
    # 3. P/N field (with barcode, properly aligned)
    pn_y = field_positions['P/N']
    draw.text((config['field_start_x'], pn_y), "P/N", fill='black', font=font_label)
    pn_barcode = generate_simple_barcode("CZ5S1000B", config['barcode_width'], config['barcode_height'])
    barcode_y = pn_y   # Small gap after label
    img.paste(pn_barcode, (config['field_start_x'] + config['text_offset'], barcode_y))
    data_text_y = barcode_y + config['barcode_height'] + 3
    draw.text((config['field_start_x'] + config['text_offset'], data_text_y), "CZ5S1000B", fill='black', font=font_data)
    
    # 4. P/R field (with barcode, properly aligned)
    pr_y = field_positions['P/R']
    draw.text((config['field_start_x'], pr_y), "P/R", fill='black', font=font_label)
    pr_barcode = generate_simple_barcode("02", config['barcode_width'], config['barcode_height'])
    barcode_y = pr_y  # Small gap after label
    img.paste(pr_barcode, (config['field_start_x'] + config['text_offset'], barcode_y))
    data_text_y = barcode_y + config['barcode_height'] + 3
    draw.text((config['field_start_x'] + config['text_offset'], data_text_y), "02", fill='black', font=font_data)
    
    # 5. S/N field (with barcode, properly aligned)
    sn_y = field_positions['S/N']
    draw.text((config['field_start_x'], sn_y), "S/N", fill='black', font=font_label)
    sn_barcode = generate_simple_barcode("CDL2349-1195", config['barcode_width'], config['barcode_height'])
    barcode_y = sn_y   # Small gap after label
    img.paste(sn_barcode, (config['field_start_x'] + config['text_offset'], barcode_y))
    data_text_y = barcode_y + config['barcode_height'] + 3
    draw.text((config['field_start_x'] + config['text_offset'], data_text_y), "CDL2349-1195", fill='black', font=font_data)
    
    # Save the perfect alignment label as PDF
    filename = "debug_label_PERFECT.pdf"
    
    # Convert to RGB if needed for PDF output
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Save as PDF with high resolution
    img.save(filename, 'PDF', dpi=(300, 300), quality=95)
    print(f"Saved: {filename}")
    
    # Also save as PNG for quick preview
    png_filename = "debug_label_PERFECT.png"
    img.save(png_filename, 'PNG', dpi=(300, 300))
    print(f"Also saved PNG preview: {png_filename}")
    
    # Return config for main app
    final_config = {
        'width': config['width'],
        'height': config['height'],
        'pd_x': config['field_start_x'], 'pd_y': field_positions['P/D'],
        'pn_x': config['field_start_x'], 'pn_y': field_positions['P/N'],
        'pr_x': config['field_start_x'], 'pr_y': field_positions['P/R'],
        'sn_x': config['field_start_x'], 'sn_y': field_positions['S/N'],
        'barcode_width': config['barcode_width'], 
        'barcode_height': config['barcode_height'],
        'text_offset': config['text_offset']
    }
    
    print("\nPerfect alignment config for main app:")
    print(f'"width": {final_config["width"]},')
    print(f'"height": {final_config["height"]},')
    print(f'"pd_x": {final_config["pd_x"]}, "pd_y": {final_config["pd_y"]},')
    print(f'"pn_x": {final_config["pn_x"]}, "pn_y": {final_config["pn_y"]},')
    print(f'"pr_x": {final_config["pr_x"]}, "pr_y": {final_config["pr_y"]},')
    print(f'"sn_x": {final_config["sn_x"]}, "sn_y": {final_config["sn_y"]},')
    print(f'"barcode_width": {final_config["barcode_width"]}, "barcode_height": {final_config["barcode_height"]}')
    
    return final_config

if __name__ == "__main__":
    print("Debug Label Generator")
    print("====================")


    print("Creating perfect alignment label...")
    
    # Create the perfect alignment version
    perfect_config = create_perfect_alignment_label()
    
    print("\nDone! Check the generated files:")
    print("- debug_label_PERFECT.pdf (main output)")
    print("- debug_label_PERFECT.png (preview)")
    print("\nUse the config values from debug_label_PERFECT.pdf as it should have the best alignment.")
