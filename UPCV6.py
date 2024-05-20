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

def calculate_check_digit(upca):
    odd_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 0)
    even_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 1)
    total_sum = odd_sum * 3 + even_sum
    check_digit = (10 - (total_sum % 10)) % 10
    return str(check_digit)

def import_excel_and_create_images(image_format='png'):
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

    for i in range(len(data)):
        upc = str(data.loc[i, 'UPC-A']).zfill(11)  # Fill with 11 zeros
        check_digit = calculate_check_digit(upc)
        upc_code = upc + check_digit
        upc_numbers.append(upc_code)  # Store the UPC number
        item_description = data.loc[i, 'Item Description']
        item_number = str(data.loc[i, 'Item#'])  # Convert item_number to string

        # Fetch the barcode image from bwip-js API
        url = f"https://bwipjs-api.metafloor.com/?bcid=upca&text={upc_code}&scale=5&rotate=N&includetext&bgcolor=ffffff"
        response = requests.get(url)
        barcode_image = Image.open(BytesIO(response.content))

        # Create a white background image
        background = Image.new('RGB', barcode_image.size, (255, 255, 255))

        # Paste the barcode image onto the background
        background.paste(barcode_image, mask=barcode_image.split()[3])  # 3 is the alpha channel

        barcode_image_path = os.path.join(TEMP_DIR, f"barcode_{i}.png")
        background.save(barcode_image_path, 'PNG')

        # Verify if the file has been created
        if os.path.exists(barcode_image_path):
            print(f"Barcode saved as {barcode_image_path}")
        else:
            print("Error in saving the barcode.")

    message_label.config(text="PNG files created successfully.")
def combine_images_to_pdf(image_format='png'):
    image_files = glob.glob(os.path.join(TEMP_DIR, f"barcode_*.{image_format}"))  # Find all image files that start with "barcode_"
    combined_pdf_path = os.path.join(TEMP_DIR, "combined.pdf")
    
    c = canvas.Canvas(combined_pdf_path, pagesize=letter)
    width, height = letter
    
    for i, image_file in enumerate(image_files):
        if not os.path.exists(image_file):
            print(f"{image_format.upper()} file {image_file} does not exist.")
            continue
    
        # Add each image file to a new page in the PDF
        if i % 44 == 0 and i != 0:
            c.showPage()
    
        # Adjust x and y coordinates for each barcode
        x = 70 if i % 2 == 0 else width / 2
        y = height - ((i // 2) % 22 * 100) - 100  # Reduce the buffer between rows
    
        c.drawImage(image_file, x, y, 120, 50)  # Reduce the width of the barcode
    
    c.save()
    message_label.config(text="PDF file created successfully.")
    
    # Open the combined PDF file
    os.startfile(combined_pdf_path)
    
    # Optionally delete the individual image files after the combined PDF is created
    for image_file in image_files:
        try:
            os.remove(image_file)
        except Exception as e:
            print(f"Failed to delete {image_file}: {e}")
    
root = ThemedTk(theme="arc")  # Use the "arc" theme
root.geometry("300x250")
root.configure(bg='white')
root.title("Create UPC Labels Form")

# Create a style
style = ttk.Style()
style.configure("TButton", font=("Arial", 10), padding=10)
style.map("TButton",
          foreground=[('pressed', 'red'), ('active', 'blue')],
          background=[('pressed', '!disabled', 'black'), ('active', 'white')])

open_button_images = ttk.Button(root, text="Import Excel & Create Images", command=lambda: import_excel_and_create_images('png'), width=30)
open_button_images.place(relx=0.5, rely=0.4, anchor='center')  # Place the button at the center of the window

combine_button_images = ttk.Button(root, text="Combine Images to PDF", command=lambda: combine_images_to_pdf('png'), width=30)
combine_button_images.place(relx=0.5, rely=0.8, anchor='center')  # Place the button at the center of the window

message_label = tk.Label(root, text="", bg='white')
message_label.place(relx=0.5, rely=0.95, anchor='center')  # Place the label at the center of the window

root.mainloop()    