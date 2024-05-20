import requests
from PIL import Image
from io import BytesIO
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def calculate_check_digit(upca):
    odd_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 0)
    even_sum = sum(int(x) for i, x in enumerate(upca) if i % 2 == 1)
    total_sum = odd_sum * 3 + even_sum
    check_digit = (10 - (total_sum % 10)) % 10
    return str(check_digit)

def generate_upca_barcode_from_excel(output_path):
    # Open a file dialog for the user to select the Excel file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Iterate over the UPC-A codes in column D
    for code in df['UPC-A']:
        # Calculate the check digit and complete the 12-digit UPC-A code
        check_digit = calculate_check_digit(str(code))
        upca_code = str(code) + check_digit

        # Use the bwip-js online API to generate the barcode
        url = f"https://bwipjs-api.metafloor.com/?bcid=upca&text={upca_code}&scale=3&rotate=N&includetext&bgcolor=ffffff"
        response = requests.get(url)

        # Save the barcode image to the specified output path
        png_file = os.path.join(output_path, f"{upca_code}.png")
        with open(png_file, "wb") as f:
            f.write(response.content)

        # Verify if the file has been created
        if os.path.exists(png_file):
            print(f"Barcode saved as {png_file}")
        else:
            print("Error in saving the barcode.")

# Generate UPC-A barcodes for the codes in the Excel file and save them to C:\Temp
generate_upca_barcode_from_excel('C:\\Temp')
