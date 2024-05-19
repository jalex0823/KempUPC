import os
import pandas as pd
from barcode import UPCA
from barcode.writer import ImageWriter
from io import BytesIO
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import svgwrite
from PyPDF2 import PdfMerger
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from tkinter import ttk
from ttkthemes import ThemedTk

TEMP_DIR = "C:\\Temp"

# Add a global variable to store the UPC numbers
upc_numbers = []

def calculate_check_digit(upca):
    odd_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 0)
    even_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 1)
    total_sum = odd_sum * 3 + even_sum
    check_digit = (10 - (total_sum % 10)) % 10
    return str(check_digit)

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

    writer = ImageWriter()
    
    for i in range(0, len(data), 22):
        svg_filename = os.path.join(TEMP_DIR, f"labels_{i//22}.svg")
        dwg = svgwrite.Drawing(svg_filename, size=('816px', '1056px'))
        
        for j in range(22):
            if i + j < len(data):
                upc = str(data.loc[i + j, 'UPC-A']).zfill(11)  # Fill with 11 zeros
                check_digit = calculate_check_digit(upc)
                upc_code = upc + check_digit
                upc_numbers.append(upc_code)  # Store the UPC number
                item_description = data.loc[i + j, 'Item Description']
                item_number = str(data.loc[i + j, 'Item#'])  # Convert item_number to string
                
                # Generate barcode
                barcode = UPCA(upc_code, writer=writer)
                barcode_io = BytesIO()
                barcode.write(barcode_io)
                
                # Load the barcode image
                barcode_image = Image.open(barcode_io)
                barcode_image_path = os.path.join(TEMP_DIR, f"barcode_{i + j}.png")
                barcode_image.save(barcode_image_path)
                
                x = 70 + (j % 2) * 420  # Increase x coordinate for padding
                y = 35 + (j // 2) * 130  # Increase y coordinate for padding, adjusted to add space below item number
                
                # Assume max_width is the maximum width of your item descriptions
                max_width = 200

                # Add a border above the item description
                dwg.add(svgwrite.shapes.Line(start=(x + 55, y), end=(x + 55 + max_width, y), stroke=svgwrite.rgb(0, 0, 0, '%')))
                
                # Add the item description above the barcode and center it
                dwg.add(svgwrite.text.Text(item_description, insert=(x + 135, y + 20), text_anchor="middle"))  # Adjusted y coordinate to center item description

                # Add the barcode image and center it
                dwg.add(svgwrite.image.Image(barcode_image_path, insert=(x + 90, y + 25), size=('180px', '50px')))  # Adjusted x and y coordinates

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
        if os.path.exists(pdf):  # Check if the PDF file exists
            try:
                with open(pdf, 'rb') as f:
                    merger.append(f)
            except FileNotFoundError:
                pass  # Do nothing if the file does not exist
            except Exception as e:
                print(f"Failed to append {pdf} to PDF Merger: {e}")
        else:
            print(f"PDF file {pdf} does not exist.")

    combined_pdf_path = os.path.join(TEMP_DIR, "combined.pdf")
    merger.write(combined_pdf_path)
    merger.close()
    message_label.config(text="PDF file created successfully.")

    # Open the combined PDF file
    os.startfile(combined_pdf_path)

    # Delete the individual PDF and SVG files
    for file_path in pdf_files + svg_files:
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")

    # Remove all SVG files after the combined PDF is created
    for svg_file in glob.glob(os.path.join(TEMP_DIR, "*.svg")):
        try:
            os.remove(svg_file)
        except Exception as e:
            print(f"Failed to delete {svg_file}: {e}")

def open_output():
    webbrowser.open('combined.pdf')

root = ThemedTk(theme="arc")  # Use the "arc" theme
root.geometry("280x200")
root.configure(bg='white')
root.title("Create UPC Labels Form")

# Create a style
style = ttk.Style()
style.configure("TButton", font=("Arial", 10), padding=10)
style.map("TButton",
          foreground=[('pressed', 'red'), ('active', 'blue')],
          background=[('pressed', '!disabled', 'black'), ('active', 'white')])

# Use ttk.Button instead of tk.Button and specify the width
open_button = ttk.Button(root, text="Import Excel Spreadsheet", command=import_excel_and_create_svg, width=30)
open_button.place(relx=0.5, rely=0.25, anchor='center')  # Place the button at the center of the window

combine_button = ttk.Button(root, text="Create UPC-A Codes", command=combine_svg_to_pdf, width=30)
combine_button.place(relx=0.5, rely=0.55, anchor='center')  # Place the button at the center of the window

message_label = tk.Label(root, text="", bg='white')
message_label.place(relx=0.5, rely=0.85, anchor='center')  # Place the label at the center of the window

root.mainloop()
