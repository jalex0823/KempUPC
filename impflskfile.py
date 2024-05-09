from flask import Flask, request, send_file
import pandas as pd
from barcode import EAN13
from barcode.writer import ImageWriter
from fpdf import FPDF

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    data = pd.read_excel(file)

    # Check if all required data is present
    if 'item' not in data.columns:
        return 'Missing required data', 400
@app.route('/download', methods=['GET'])
def download_file():
    pdf_filename = 'output.pdf'
    try:
        return send_file(pdf_filename, as_attachment=True)
    except FileNotFoundError:
        return 'File not found', 404
    # Generate UPCs and barcodes
    pdf = FPDF()
    for i, row in data.iterrows():
        upc = str(row['item']).zfill(12)
        barcode = EAN13(upc, writer=ImageWriter())
        filename = f"barcode_{i}.png"
        barcode.save(filename)

        # Add barcode to PDF
        pdf.add_page()
        pdf.image(filename, x=10, y=10, w=100)

    # Save PDF
    pdf_filename = 'output.pdf'
    pdf.output(pdf_filename)

    return send_file(pdf_filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)