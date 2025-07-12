"""
This script creates a tkinter GUI that generates a sample PNG image
and prints it to the default printer when the button is clicked.
Requires: pillow, pywin32 (install via: pip install pillow pywin32)
"""
import sys
import tempfile
import tkinter as tk
from tkinter import messagebox

try:
    from PIL import Image, ImageDraw, ImageWin
    import win32print, win32ui, win32con
except ImportError:
    print("Required modules missing. Please install with: pip install pillow pywin32")
    sys.exit(1)


def generate_sample_image(path):
    width, height = 600, 400
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)
    text = "Sample Image"
    # Use default font and get text size using font object
    try:
        from PIL import ImageFont
        font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        textwidth = bbox[2] - bbox[0]
        textheight = bbox[3] - bbox[1]
    except:
        # Fallback for older Pillow versions
        textwidth, textheight = 200, 20  # Approximate size
    x = (width - textwidth) // 2
    y = (height - textheight) // 2
    draw.text((x, y), text, fill='black')
    image.save(path)


def print_image(path):
    printer_name = win32print.GetDefaultPrinter()
    print(printer_name)
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    bmp = Image.open(path)
    printable_area = (hDC.GetDeviceCaps(win32con.HORZRES),
                      hDC.GetDeviceCaps(win32con.VERTRES))
    ratio = min(printable_area[0] / bmp.size[0], printable_area[1] / bmp.size[1])
    scaled_size = (int(bmp.size[0] * ratio), int(bmp.size[1] * ratio))
    bmp = bmp.resize(scaled_size)
    dib = ImageWin.Dib(bmp)
    hDC.StartDoc("Sample Print")
    hDC.StartPage()
    x = (printable_area[0] - scaled_size[0]) // 2
    y = (printable_area[1] - scaled_size[1]) // 2
    dib.draw(hDC.GetHandleOutput(), (x, y, x + scaled_size[0], y + scaled_size[1]))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()


def on_print():
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        tmp_path = tmp.name
        tmp.close()
        generate_sample_image(tmp_path)
        print_image(tmp_path)
        messagebox.showinfo("Success", "Printed successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    root.title("Sample Print")
    root.geometry("300x100")
    btn = tk.Button(root, text="Print Sample Image", command=on_print)
    btn.pack(expand=True, fill='both', padx=10, pady=10)
    root.mainloop()


if __name__ == '__main__':
    main()
