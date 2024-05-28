import os
import pandas as pd
import requests
from io import BytesIO
from PIL import Image as PILImage, UnidentifiedImageError
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import letter
from tkinter import ttk
from ttkthemes import ThemedTk
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Image as ReportLabImage, Spacer, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

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
        upc_code = str(data.loc[i, 'UPC-A'])  # Convert UPC_code to string
        item_description = data.loc[i, 'Item Description']
        item_number = str(data.loc[i, 'Item#'])  # Convert item_number to string

        # Check if UPC_code is 11 digits long, if so, prepend a '0'
        if len(upc_code) == 11:
            upc_code = '0' + upc_code
        elif len(upc_code) != 12:
            message_label.config(text="Invalid UPC length at row {}".format(i + 1))
            return

        check_digit = calculate_check_digit(upc_code[:11])
        upc_code = upc_code[:11] + check_digit
        upc_numbers.append(upc_code)  # Store the UPC number

        # Append the row data to the excel_data list
        excel_data.append({
            'Item_description': item_description,
            'Item_number': item_number,
            'UPC_code': upc_code
        })

        # Fetch the barcode image from bwip-js API
        url = f"https://bwipjs-api.metafloor.com/?bcid=upca&text={upc_code}&scale=3&rotate=N&includetext"
        response = requests.get(url)

        if response.content:
            try:
                barcode_image = PILImage.open(BytesIO(response.content)).convert("RGBA")
            except UnidentifiedImageError:
                messagebox.showerror("Error", "The response content is not a valid image.")
                continue
        else:
            messagebox.showerror("Error", "The response content is empty.")
            continue

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
    # Define the output file path
    output_file_path = os.path.join(TEMP_DIR, "output.pdf")

    doc = SimpleDocTemplate(output_file_path, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    styleN = styles['Normal']

    # Get a list of all PNG files in the TEMP_DIR
    png_files = [f for f in os.listdir(TEMP_DIR) if f.endswith('.png')]

    # Sort the list of files to ensure they're in the correct order
    png_files.sort()

    for i in range(0, len(png_files), 2):
        # Create a list to hold the description and image for this row
        row = []

        # Add the first item's description and image to the row
        desc1 = Paragraph('<font size=8>{}</font>'.format(excel_data[i]['Item_description']), styleN)
        img1 = ReportLabImage(os.path.join(TEMP_DIR, png_files[i]), width=0.5*inch, height=0.5*inch)
        row.append([desc1, img1])

        # If there is a second item, add its description and image to the row
        if i+1 < len(png_files):
            desc2 = Paragraph('<font size=8>{}</font>'.format(excel_data[i+1]['Item_description']), styleN)
            img2 = ReportLabImage(os.path.join(TEMP_DIR, png_files[i+1]), width=0.5*inch, height=0.5*inch)
            row.append([desc2, img2])

        # Create a table with the descriptions and images in the row
        table = Table(row, colWidths=[3*inch, 3*inch])

        # Set the alignment of the second column to left
        table.setStyle(TableStyle([('LEFTPADDING', (1,0), (-1,-1), 0)]))

        # Add the table to the story
        story.append(table)

    doc.build(story)
    
    # Open the PDF file
    os.startfile(output_file_path)

    # Display a success message
    messagebox.showinfo("Success", "PDF created successfully.")
root = ThemedTk(theme="equilux")  # Use the "equilux" theme
root.geometry('300x200')  # Set the form size to 300x200 pixels

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
