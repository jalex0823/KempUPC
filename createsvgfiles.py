import os
import pandas as pd
import svgwrite
from barcode import UPCA
from barcode.writer import SVGWriter
from tkinter import Tk
from tkinter import filedialog

TEMP_DIR = "C:\\Temp"

def generate_svg_barcodes(data):
    writer = SVGWriter()
    svg_files = []

    for i, row in data.iterrows():
        upc = str(row['UPC-A']).zfill(11)
        item_description = row['Item Description']
        item_number = row['Item#']

        svg_filename = os.path.join(TEMP_DIR, f"barcode_{i}")
        dwg = svgwrite.Drawing(svg_filename, size=('816px', '1056px'))

        barcode = UPCA(upc, writer=writer)
        barcode.save(svg_filename)

        x = 50 + (i % 2) * 350
        y = 20 + (i // 2) * 90

        dwg.add(svgwrite.text.Text(item_description, insert=(x, y)))
        dwg.add(svgwrite.image.Image(svg_filename + '.svg', insert=(x, y + 20), size=('100px', '50px')))
        dwg.add(svgwrite.text.Text(upc, insert=(x + 50, y + 70), text_anchor="middle"))
        dwg.add(svgwrite.text.Text(item_number, insert=(x + 50, y + 100), text_anchor="middle"))

        dwg.saveas(svg_filename + '.svg')
        svg_files.append(svg_filename + '.svg')

    return svg_files

def main():
    # Create a root window and hide it
    root = Tk()
    root.withdraw()

    # Open a file dialog and get the selected file
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx *.xlsb'), ('All Files', '*.*')])

    # If a file was selected, process it
    if file_path:
        data = pd.read_excel(file_path)
        svg_files = generate_svg_barcodes(data)
        print("SVG files generated:", svg_files)

    # Destroy the root window
    root.destroy()

if __name__ == "__main__":
    main()