import tkinter as tk
from tkinter import filedialog
import pandas as pd
from barcode import EAN13
from barcode.writer import ImageWriter
from fpdf import FPDF
import os
import webbrowser

def open_file():
    global pdf_filename
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not filepath:
        return
    data = pd.read_excel(filepath)

    # Check if all required data is present
    required_columns = ['item', 'description', 'item_number']  # Add other required columns to this list
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        message_label.config(text=f"Missing required data: {', '.join(missing_columns)}")
        return
    else:
        message_label.config(text="All required data is present. Ready for output.")

    # Generate UPCs and barcodes
    pdf = FPDF()
    for i, row in data.iterrows():
        upc = str(row['item']).zfill(12)
        barcode = EAN13(upc, writer=ImageWriter())
        filename = f"barcode_{i}.png"
        barcode.save(filename)

        # Add barcode and item details to PDF
        pdf.add_page()
        pdf.set_font("Arial", size = 15)
        pdf.cell(200, 10, txt = row['description'], ln = True, align = 'C')
        pdf.image(filename, x=10, y=20, w=63.5)  # Set width to 2.5 inches (approximately 63.5 mm)
        pdf.cell(200, 110, txt = upc, ln = True, align = 'C')
        pdf.cell(200, 120, txt = str(row['item_number']), ln = True, align = 'C')

    # Save PDF
    pdf_filename = 'output.pdf'
    pdf.output(pdf_filename)
    print('UPC labels saved to output.pdf')

def open_output():
    if pdf_filename:
        webbrowser.open(pdf_filename)

root = tk.Tk()
root.geometry("280x200")  # Set the size of the window to be smaller
root.title("Create UPC Labels Form")  # Set the title of the form
open_button = tk.Button(root, text="Import Excel Spreadsheet", command=open_file)
open_button.grid(row=0, column=0, sticky='ew', padx=(0, 10))  # Add padding to the right of the button

output_button = tk.Button(root, text="Output UPC PDF File", command=open_output)
output_button.grid(row=0, column=1, sticky='ew')  # No padding needed here

message_label = tk.Label(root, text="")
message_label.grid(row=1, column=0, columnspan=2)

pdf_filename = None

root.mainloop()