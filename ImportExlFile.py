import tkinter as tk
from tkinter import filedialog
import pandas as pd
from barcode import UPCA
from barcode.writer import SVGWriter
from PyPDF2 import PdfMerger
import os
import webbrowser
from wand.image import Image as WandImage
from svgwrite import Drawing
from svgwrite.text import Text
from svgwrite.image import Image as SvgImage
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from fpdf import FPDF
import time
import svgwrite
import base64
import subprocess
from PIL import ImageFont

TEMP_DIR = "C:\\Temp"

# Ensure that the TEMP_DIR directory exists
if not os.path.isdir(TEMP_DIR):
    print(f"Directory does not exist: {TEMP_DIR}")
    os.makedirs(TEMP_DIR)

def file_path_to_data_url(file_path):
    if not os.path.isfile(file_path):
        print(f"File does not exist: {file_path}")
        return None

    with open(file_path, 'rb') as f:
        return 'data:image/svg+xml;base64,' + base64.b64encode(f.read()).decode()

def text_length(text):
    font = ImageFont.load_default()
    return font.getsize(text)[0]

# Rest of your code...

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsb")])
    if not filepath:
        return
    if filepath.endswith('.xlsb'):
        data = pd.read_excel(filepath, engine='pyxlsb')
    else:
        data = pd.read_excel(filepath)

    if data is None or not {'UPC-A', 'Item Description', 'Item#'}.issubset(data.columns):
        message_label.config(text="Failed to read Excel file.")
        return

    pdf = FPDF()
    writer = SVGWriter()

    svg_files = []
    pdf_files = []
    for i in range(0, len(data), 22):
        svg_filename = os.path.join(TEMP_DIR, f"labels_{i//22}.svg")
        dwg = Drawing(svg_filename, size=('816px', '1056px'))

        for j in range(22):  # Iterate over the rows in the current group
            if i + j < len(data):
                upc = str(data.loc[i+j, 'UPC-A']).zfill(11)  # Adjusted to 11 digits
                barcode_svg_filename = os.path.join(TEMP_DIR, f"barcode_{i+j}.svg")

                # Print the UPC value and the barcode SVG filename
                print(f"UPC: {upc}, Barcode SVG Filename: {barcode_svg_filename}")

                barcode = UPCA(upc, writer=SVGWriter())
                barcode.save(barcode_svg_filename)

                # Check if the barcode SVG file exists
                if not os.path.isfile(barcode_svg_filename):
                    print(f"File was not created: {barcode_svg_filename}")

                # Convert the file path to a data URL
                barcode_svg_data_url = file_path_to_data_url(barcode_svg_filename)

                # Calculate the position of the barcode and its related texts
                x = 50 + (j % 2) * 350  # 50 for the first column, 400 for the second column
                y = 20 + (j // 2) * 90  # 20, 110, 200, ..., 990 for the rows

                dwg.add(Text(data.loc[i+j, 'Item Description'], insert=(x, y)))  # Add title above the barcode
                if barcode_svg_data_url is not None:
                    dwg.add(svgwrite.image.Image(barcode_svg_data_url, insert=(x, y + 30), size=('100px', '50px')))  # Add barcode beneath the title
                dwg.add(Text(upc, insert=(x, y + 80)))  # Add UPC number beneath the barcode

                item_number = str(data.loc[i+j, 'Item#'])
                item_number_width = text_length(item_number)  # Only pass the item_number
                dwg.add(Text(item_number, insert=(x + 50 - item_number_width / 2, y + 100)))  # Add note beneath the barcode centered

        dwg.saveas(svg_filename)
        svg_files.append(svg_filename)

    # Add this code after the barcode SVG files are generated
    for i in range(len(data)):
        barcode_svg_filename = os.path.join(TEMP_DIR, f"barcode_{i}.svg")
        if os.path.isfile(barcode_svg_filename):
            webbrowser.open_new(os.path.abspath(barcode_svg_filename))
        else:
            print(f"Barcode SVG file does not exist: {barcode_svg_filename}")

    for svg_file in svg_files:
        pdf_filename = os.path.splitext(svg_file)[0] + '.pdf'
        subprocess.run(['inkscape', '-z', '-f', svg_file, '-A', pdf_filename])
        pdf_files.append(pdf_filename)

    with PdfMerger() as merger:
        for pdf_file in pdf_files:
            while not os.path.isfile(pdf_file):
                print("Waiting for all files to be created...")
                time.sleep(1)  # wait for 1 second before checking again
            merger.append(pdf_file)

        merger.write(os.path.join(TEMP_DIR, "combined.pdf"))

    message_label.config(text="PDF files combined successfully.")

def open_output():
    combined_pdf_path = os.path.join(TEMP_DIR, "combined.pdf")
    if os.path.isfile(combined_pdf_path):
        webbrowser.open_new(os.path.abspath(combined_pdf_path))
    else:
        message_label.config(text="No PDF file to open. Please generate the PDF first.")

root = tk.Tk()
root.geometry("280x200")
root.configure(bg='white')
root.title("Create UPC Labels Form")

open_button = tk.Button(root, text="Import Excel Spreadsheet", command=open_file)
open_button.grid(row=0, column=0, sticky='ew', padx=(0, 10))

output_button = tk.Button(root, text="Output UPC PDF File", command=open_output)
output_button.grid(row=0, column=1, sticky='ew')

message_label = tk.Label(root, text="", bg='white')
message_label.grid(row=1, column=0, columnspan=2)

root.mainloop()