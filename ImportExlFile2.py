import tkinter as tk
from tkinter import filedialog
import pandas as pd
from barcode import UPCA
from barcode.writer import SVGWriter
import os
import webbrowser
import svgwrite
import glob
from PyPDF2 import PdfMerger
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF

TEMP_DIR = "C:\\Temp"

# Add a global variable to store the UPC numbers
upc_numbers = []

def import_excel_and_create_svg():
    global upc_numbers
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsb")])
    if not filepath:
        return

    if filepath.endswith('.xlsb'):
        data = pd.read_excel(filepath, engine='pyxlsb')
    else:
        data = pd.read_excel(filepath)

    required_columns = ['Item#', 'Item Description', 'UPC-A']
    if data is None or not set(required_columns).issubset(data.columns):
        message_label.config(text="Failed to read Excel file or missing columns.")
        return

    writer = SVGWriter()
    
    for i in range(0, len(data), 22):
        svg_filename = os.path.join(TEMP_DIR, f"labels_{i//22}.svg")
        dwg = svgwrite.Drawing(svg_filename, size=('816px', '1056px'))
        
        for j in range(22):
            if i + j < len(data):
                upc = str(data.loc[i + j, 'UPC-A']).zfill(12)  # Fill with 12 zeros
                upc_numbers.append(upc)  # Store the UPC number
                item_description = data.loc[i + j, 'Item Description']
                item_number = str(data.loc[i + j, 'Item#'])  # Convert item_number to string
                                
                barcode_svg_filename = os.path.join(TEMP_DIR, f"barcode_{i + j}")
                
                barcode = UPCA(upc, writer=SVGWriter())
                barcode.save(barcode_svg_filename)
                
                x = 70 + (j % 2) * 420  # Increase x coordinate for padding
                y = 35 + (j // 2) * 130  # Increase y coordinate for padding, adjusted to add space below item number
                
                # Assume max_width is the maximum width of your item descriptions
                max_width = 200

                # Add a border above the item description
                dwg.add(svgwrite.shapes.Line(start=(x + 55, y), end=(x + 55 + max_width, y), stroke=svgwrite.rgb(0, 0, 0, '%')))
                
                # Add the item description above the barcode and center it
                dwg.add(svgwrite.text.Text(item_description, insert=(x + 135, y + 20), text_anchor="middle"))  # Adjusted y coordinate to center item description

                # Add the barcode image and center it
                dwg.add(svgwrite.image.Image(barcode_svg_filename + ".svg", insert=(x + 90, y + 25), size=('180px', '50px')))  # Adjusted x and y coordinates

                # Add the item number below the barcode and UPC number and left justify it
                dwg.add(svgwrite.text.Text("Item Number: " + item_number, insert=(x + 55, y + 110), text_anchor="start"))  # Adjusted x coordinate and set text_anchor="start"

                # Add a border underneath the item number
                dwg.add(svgwrite.shapes.Line(start=(x + 55, y + 115), end=(x + 55 + max_width, y + 115), stroke=svgwrite.rgb(0, 0, 0, '%')))
        
        dwg.saveas(svg_filename)
        message_label.config(text="SVG files created successfully.")
            
def combine_svg_to_pdf():
    global upc_numbers
    pdf_files = []
    svg_files = glob.glob(os.path.join(TEMP_DIR, "labels_*.svg"))  # Find all SVG files that start with "labels_"
    for svg_file_path in svg_files:
        if not os.path.exists(svg_file_path):
            print(f"SVG file {svg_file_path} does not exist.")
            continue

        pdf_file = os.path.splitext(svg_file_path)[0] + ".pdf"
        try:
            drawing = svg2rlg(svg_file_path)
            renderPDF.drawToFile(drawing, pdf_file)
            pdf_files.append(pdf_file)
        except Exception as e:
            print(f"Failed to convert {svg_file_path} to PDF: {e}")

    # Combine all PDFs into one using PyPDF2
    merger = PdfMerger()

    for pdf in pdf_files:
        merger.append(pdf)

    combined_pdf_path = os.path.join(TEMP_DIR, "combined.pdf")
    merger.write(combined_pdf_path)
    merger.close()
    message_label.config(text="PDF file created successfully.")

    # Delete the individual PDF and SVG files
    for file_path in pdf_files + svg_files:
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")

def open_output():
    webbrowser.open('combined.pdf')

root = tk.Tk()
root.geometry("280x200")
root.configure(bg='white')
root.title("Create UPC Labels Form")

open_button = tk.Button(root, text="Import Excel Spreadsheet", command=import_excel_and_create_svg)
open_button.grid(row=0, column=0, sticky='ew', padx=(10, 10), pady=(10, 5))

combine_button = tk.Button(root, text="Combine SVG to PDF", command=combine_svg_to_pdf)
combine_button.grid(row=1, column=0, sticky='ew', padx=(10, 10), pady=(5, 5))

output_button = tk.Button(root, text="Output UPC PDF File", command=open_output)
output_button.grid(row=2, column=0, sticky='ew', padx=(10, 10), pady=(5, 10))

message_label = tk.Label(root, text="", bg='white')
message_label.grid(row=3, column=0, columnspan=2)

root.mainloop()