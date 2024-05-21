import os
import pandas as pd
import requests
from io import BytesIO
from PIL import Image as PILImage
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from tkinter import ttk
from ttkthemes import ThemedTk
from reportlab.lib.units import inch
from reportlab.platypus import Image as ReportLabImage
from tkinter import messagebox

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
        barcode_image = PILImage.open(BytesIO(response.content)).convert("RGBA")

        # Resize the image to fit within the max dimensions
        max_width = 170
        max_height = 100
        barcode_image.thumbnail((max_width, max_height))

        # Create a new white image of the same size as the thumbnail
        white_background = PILImage.new("RGBA", (max_width, max_height), "WHITE")
        white_background.paste(barcode_image, (0, 0), barcode_image)

        barcode_image_path = os.path.join(TEMP_DIR, f"barcode_{i}.png")
        white_background.save(barcode_image_path)

        # Verify if the file has been created
        if os.path.exists(barcode_image_path):
            print(f"Barcode saved as {barcode_image_path}")
        else:
            print("Error in saving the barcode.")

    message_label.config(text="PNG files created successfully.")

def combine_png_to_pdf():
    # Define the dimensions of the PDF
    width, height = letter  # 8.5 x 11 inches

    # Define the margins, spacing, and image size
    top_margin = 0.5 * inch
    bottom_margin = 0.5 * inch
    left_margin = 0.5 * inch
    right_margin = 0.166 * inch
    row_spacing = 1 * inch
    column_spacing = 1 * inch
    column_width = 4 * inch
    column_height = 2 * inch
    max_image_width = 170 / 72 * inch  # Convert pixels to inches
    max_image_height = 100 / 72 * inch  # Convert pixels to inches

    # Create a canvas with the defined dimensions
    combined_pdf_path = os.path.join(TEMP_DIR, "combined.pdf")
    c = canvas.Canvas(combined_pdf_path, pagesize=letter)

    # Get all PNG files in the directory
    png_files = [f for f in os.listdir(TEMP_DIR) if f.endswith('.png')]

    x = left_margin
    y = height - top_margin - column_height

    for i, png_file in enumerate(png_files):
        # Get the file path
        png_file_path = os.path.join(TEMP_DIR, png_file)

        # Create a reportlab Image from the file path
        img = ReportLabImage(png_file_path)

        # Resize the image to fit within the defined image size
        img_width = min(img.drawWidth, max_image_width)
        img_height = min(img.drawHeight, max_image_height)
        img.drawWidth = img_width
        img.drawHeight = img_height

        # Calculate the x-coordinate for the image to be right-indented
        img_x = x + column_width - img_width

        # Draw the image on the canvas at the calculated position
        img.drawOn(c, img_x, y)

        # Move to the next column or row
        if (i + 1) % 2 == 0:
            x = left_margin
            y -= column_height + row_spacing
        else:
            x += column_width + column_spacing

        # Add a page break if we've reached the bottom margin
        if y < bottom_margin + column_height:
            c.showPage()
            x = left_margin
            y = height - top_margin - column_height

    # Save the PDF
    c.save()
    
    # Display a success message
    messagebox.showinfo("Success", "PDF created successfully.")

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
