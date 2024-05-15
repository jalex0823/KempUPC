from barcode import UPCA
from barcode.writer import SVGWriter
from cairosvg import svg2png
from PIL import Image, ImageFont, ImageDraw
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
import io

def browse_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    return file_path

def generate_check_digit_upca(data):
    sum = 0
    for i, digit in enumerate(data):
        digit = int(digit)
        if i % 2 == 0:
            sum += 3 * digit
        else:
            sum += digit
    result = sum % 10
    if result != 0:
        result = 10 - result
    return str(result)

def filter_input_upca(data):
    return ''.join(filter(str.isdigit, data))

def Encode_UPCA(data):
    # Filter the input
    filtered_data = filter_input_upca(data)

    # Pad or truncate the data to 11 digits
    filtered_data = filtered_data.ljust(11, '0')[:11]

    # Generate the check digit
    check_digit = generate_check_digit_upca(filtered_data)

    # Add the check digit to the data
    filtered_data += check_digit

    barcode = UPCA(filtered_data, writer=SVGWriter(), write_text=False)
    barcode_svg = barcode.render() 
    return barcode_svg, filtered_data

def generate_barcode_png(upc_a):
    # Generate the barcode SVG
    barcode_svg, filtered_data = Encode_UPCA(upc_a)

    # Convert the SVG to PNG
    barcode_png = svg2png(bytestring=barcode_svg)

    # Create an image from the PNG
    image = Image.open(io.BytesIO(barcode_png))

    # Load the font
    font = ImageFont.truetype('C:\\ClairProjetKemp\\KempUPC\\BarcodeUPCAFunctions\\ConnectCodeUPCEANA_HRBS3_Trial.ttf', 18)

    # Create a draw object
    draw = ImageDraw.Draw(image)

    # Estimate the width of one character
    char_width = 25  # This is just an example, adjust as needed

    # Calculate the width of the text
    text_width = len(filtered_data) * char_width
    
    # Calculate the x position of the text
    text_x = (image.width - text_width) / 2

    # Draw the text on the image
    draw.text((text_x, image.height - 20), filtered_data, font=font, fill='black')

    # Save the image to a bytes buffer
    output = io.BytesIO()
    image.save(output, format='PNG')
    output.seek(0)
    return output.read()

def main():
    # Browse for the Excel file
    file_path = browse_file()
    if not file_path:
        print("No file selected. Exiting...")
        return
    
    # Load the Excel spreadsheet
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    
    # Create the directory if it doesn't exist
    output_dir = "C:\\Temp"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {output_dir}")
    
    # Extract data from the spreadsheet
    for row in sheet.iter_rows(min_row=2, values_only=True):
        item_number = row[0]
        item_description = row[1]
        group = row[2]
        upc_a = str(row[3])  # Convert to string in case the value is not already a string
        
        # Generate barcode PNG
        barcode_png = generate_barcode_png(upc_a)
        
        # Save the barcode PNG to a file in C:\Temp
        png_file_path = os.path.join(output_dir, f"{item_number}.png")
        with open(png_file_path, 'wb') as f:
            f.write(barcode_png)
        
        print(f"Barcode PNG generated for Item {item_number}")

    # Close the workbook
    wb.close()
    print("Barcode PNGs generated successfully.")

if __name__ == "__main__":
    main()