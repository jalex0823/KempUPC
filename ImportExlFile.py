import tkinter as tk
from tkinter import filedialog
import pandas as pd
from barcode import UPCA
from barcode.writer import SVGWriter
from fpdf import FPDF
import os
import webbrowser
from wand.image import Image
from svgutils.compose import Figure, SVG
from svgwrite import Drawing
from svgwrite.text import Text
from svglib.svglib import svg2rlg  # New import
from svgutils.transform import fromfile, Group  # Updated import

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsb")])  # Include .xlsb files
    if not filepath:
        return
    if filepath.endswith('.xlsb'):
        data = pd.read_excel(filepath, engine='pyxlsb')  # Use pyxlsb engine for .xlsb files
    else:
        data = pd.read_excel(filepath)
    
    if data is None:
        message_label.config(text="Failed to read Excel file.")
        return

    pdf = FPDF()
    writer = SVGWriter()

    # Generate all SVG files
    svg_files = []
    for i in range(0, len(data), 22):  # Iterate over data 22 rows at a time
        # Create a new SVG file for each group of rows
        svg_filename = os.path.join("C:\\Temp", f"barcode_{i//22}.svg")
    
        # Ensure the directory exists
        os.makedirs(os.path.dirname(svg_filename), exist_ok=True)
    
        # Create a new SVG drawing with the size of a letter page in pixels (assuming 96 DPI)
        dwg = Drawing(svg_filename, size=('816', '1056'))  # 8.5 x 11 inches in pixels
        
        for j in range(22):  # Iterate over the rows in the current group
            if i + j < len(data):
                upc = str(data.loc[i+j, 'UPC-A']).zfill(12)
                barcode = UPCA(upc, writer=SVGWriter())
                barcode_filename = svg_filename.replace('.svg', '')  # Remove the extension
                barcode.save(barcode_filename)
        
                # Load the barcode SVG and add it to the drawing
                barcode_svg = fromfile(barcode_filename + '.svg').getroot()
                barcode_svg.moveto(x, y)  # Move the barcode to the correct position
        
                # Add the barcode SVG to the drawing
                dwg.add(barcode_svg)
        
                # Add texts to the drawing
                dwg.add(Text(data.loc[i+j, 'Item Description'], insert=(x, y)))  # Add title above the barcode
                dwg.add(Text(upc, insert=(x, y + 60)))  # Add UPC number beneath the barcode
                dwg.add(Text(data.loc[i+j, 'Item#'], insert=(x, y + 80)))  # Add note beneath the barcode
        
        dwg.saveas(svg_filename)
        svg_files.append(svg_filename)

    # Combine SVG files into a single SVG file
    fig = Figure()
    for svg_file in svg_files:
        subfig = SVG(svg_file)
        fig.append(subfig)
    fig.save(os.path.join("C:\\Temp", "combined.svg"))
    message_label.config(text="SVG files combined successfully.")

def open_output():
    combined_svg_path = os.path.join("C:\\Temp", "combined.svg")
    if os.path.exists(combined_svg_path):
        webbrowser.open_new(os.path.abspath(combined_svg_path))
    else:
        message_label.config(text="No SVG file to open. Please generate the SVG first.")

root = tk.Tk()
root.geometry("280x200")  # Set the size of the window to be smaller
root.configure(bg='white')  # Set the window background to white
root.title("Create UPC Labels Form")  # Set the title of the form

open_button = tk.Button(root, text="Import Excel Spreadsheet", command=open_file)
open_button.grid(row=0, column=0, sticky='ew', padx=(0, 10))  # Add padding to the right of the button

output_button = tk.Button(root, text="Output UPC SVG File", command=open_output)
output_button.grid(row=0, column=1, sticky='ew')  # No padding needed here

message_label = tk.Label(root, text="", bg='white')  # Set the label background to white
message_label.grid(row=1, column=0, columnspan=2)

root.mainloop()