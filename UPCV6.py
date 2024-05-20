import os
import pandas as pd
import requests
from io import BytesIO
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from tkinter import ttk
from ttkthemes import ThemedTk
import glob
import webbrowser

TEMP_DIR = "C:\\Temp"

# Add a global variable to store the UPC numbers
upc_numbers = []

# This list will store the Excel data
excel_data = []

def calculate_check_digit(upca):
    odd_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 0)
    even_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 1)
    total_sum = odd_sum * 3 + even_sum
    check_digit = (10 - (total_sum % 10)) % 10
    return str(check_digit)

def import_excel_and_create_png():
    global upc_numbers
    global excel_data  # Add this line to access the global variable
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

    for i in range(len(data)):
        upc = str(data.loc[i, 'UPC-A']).zfill(11)  # Fill with 11 zeros
        check_digit = calculate_check_digit(upc)
        upc_code = upc + check_digit
        upc_numbers.append(upc_code)  # Store the UPC number
        item_description = data.loc[i, 'Item Description']
        item_number = str(data.loc[i, 'Item#'])  # Convert item_number to string

        # Append the row data to the excel_data list
        excel_data.append({
            'Item_description': item_description,
            'Item_number': item_number,
            'UPC_code': upc_code
        })

        # Fetch the barcode image from bwip-js API
        url = f"https://bwipjs-api.metafloor.com/?bcid=upca&text={upc_code}&scale=3&rotate=N&includetext"
        response = requests.get(url)
        barcode_image = Image.open(BytesIO(response.content)).convert("RGBA")

        # Create a new white image of the same size
        white_background = Image.new("RGBA", barcode_image.size, "WHITE")
        white_background.paste(barcode_image, (0, 0), barcode_image)

        barcode_image_path = os.path.join(TEMP_DIR, f"barcode_{i}.png")
        white_background.save(barcode_image_path)

        # Verify if the file has been created
        if os.path.exists(barcode_image_path):
            print(f"Barcode saved as {barcode_image_path}")
        else:
            print("Error in saving the barcode.")

    message_label.config(text="PNG files created successfully.")

# Rest of your code...
def combine_png_to_pdf():
    # Create a new PDF
    c = canvas.Canvas("combined.pdf")

    for row_data in excel_data:
        # Retrieve the item description and item number
        item_description = row_data['Item_description']
        item_number = row_data['Item_number']

        # Load the barcode image
        barcode_image = Image.open(f"{TEMP_DIR}/barcode_{item_number}.png")

        # Create a new PDF page
        c.showPage()

        # Add the barcode image to the left side of the page
        c.drawInlineImage(barcode_image, 0, 0)

        # Add the item description and item number to the right side of the page, right justified
        text = c.beginText(200, 100)  # Adjust the coordinates as needed
        text.textLine(f"Item #: {item_number}")
        text.textLine(item_description)
        c.drawText(text)

    # Save the PDF
    c.save()

root = ThemedTk(theme="equilux")  # Use the "equilux" theme
root.geometry('200x200')  # Set the form size to 200x200 pixels

style = ttk.Style()
style.configure("TButton", background="white")
style.map("TButton",
      foreground=[('active', 'blue')],
      background=[('active', 'white')])

button1_text = "Import Excel and Create PNG"
button2_text = "Combine PNG to PDF"

# Determine the maximum length of the button texts
max_length = max(len(button1_text), len(button2_text))

button1 = ttk.Button(root, text=button1_text, command=import_excel_and_create_png, width=max_length, style="TButton")
button1.pack(padx=10, pady=10)

button2 = ttk.Button(root, text=button2_text, command=combine_png_to_pdf, width=max_length, style="TButton")
button2.pack(padx=10, pady=10)

message_label = tk.Label(root, text="")
message_label.pack()

# Run the application
root.mainloop()