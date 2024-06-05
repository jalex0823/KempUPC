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
from reportlab.platypus import SimpleDocTemplate, Image as ReportLabImage, Spacer, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER

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
    global excel_data
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsb")])
    if not filepath:
        return

    if filepath.endswith('.xlsb'):
        data = pd.read_excel(filepath, engine='pyxlsb')
    else:
        data = pd.read_excel(filepath)

    required_columns = ['Item#', 'Item Description', 'Group', 'UPC-A']
    if data is None or not set(required_columns).issubset(data.columns):
        message_label.config(text="Failed to read Excel file or missing columns.")
        return

    num_rows = len(data)
    progress_bar['maximum'] = num_rows

    for i in range(num_rows):
        upc_code = str(data.loc[i, 'UPC-A'])
        item_description = data.loc[i, 'Item Description']
        item_number = str(data.loc[i, 'Item#'])
        group_name = data.loc[i, 'Group']

        if pd.isna(group_name):
            group_name = "UNKNOWN"
        else:
            group_name = str(group_name)

        if len(upc_code) == 11:
            upc_code = '0' + upc_code
        elif len(upc_code) != 12:
            message_label.config(text="Invalid UPC length at row {}".format(i + 1))
            return

        check_digit = calculate_check_digit(upc_code[:11])
        upc_code = upc_code[:11] + check_digit
        upc_numbers.append(upc_code)

        excel_data.append({
            'Item_description': item_description,
            'Item_number': item_number,
            'UPC_code': upc_code,
            'Group': group_name
        })

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

        max_width = 170
        max_height = 100
        barcode_image.thumbnail((max_width, max_height))

        white_background = PILImage.new("RGBA", (max_width, max_height), "WHITE")
        white_background.paste(barcode_image, (0, 0), barcode_image)

        barcode_image_path = os.path.join(TEMP_DIR, f"barcode_{i}.png")
        white_background.save(barcode_image_path)

        if os.path.exists(barcode_image_path):
            log_message(f"Barcode saved as {barcode_image_path}")
        else:
            log_message("Error in saving the barcode.")

        progress_bar['value'] = i + 1
        root.update_idletasks()

    combine_png_to_pdf()

def combine_png_to_pdf():
    output_file_path = os.path.join(TEMP_DIR, "output.pdf")

    doc = SimpleDocTemplate(output_file_path, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    styleH = styles['Heading1']
    styleH.alignment = TA_CENTER

    png_files = [f for f in os.listdir(TEMP_DIR) if f.endswith('.png')]
    png_files.sort()

    grouped_data = {}
    for i in range(len(excel_data)):
        group_name = excel_data[i]['Group']
        if group_name not in grouped_data:
            grouped_data[group_name] = []
        grouped_data[group_name].append({
            'description': excel_data[i]['Item_description'],
            'number': excel_data[i]['Item_number'],
            'png_file': os.path.join(TEMP_DIR, f"barcode_{i}.png")
        })

    first_group = True
    for group_name, items in grouped_data.items():
        if not first_group:
            story.append(PageBreak())
        else:
            first_group = False

        group_header = Paragraph(f'<b>{group_name.upper()}</b>', styleH)
        story.append(group_header)
        story.append(Spacer(1, 12))

        for i in range(0, len(items), 2):
            row = []

            # First item in the row
            desc1 = Paragraph('<font size=10>{}</font>'.format(items[i]['description']), styleN)
            item_num1 = Paragraph('<font size=10>{}</font>'.format(items[i]['number']), styleN)
            img_path1 = items[i]['png_file']
            with PILImage.open(img_path1) as img:
                original_width, original_height = img.size
                new_width = 2 * inch
                new_height = (new_width / original_width) * original_height
            img1 = ReportLabImage(img_path1, width=new_width, height=new_height)

            item1 = Table([[desc1], [img1], [item_num1]], colWidths=[3 * inch])
            item1.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                ('ALIGN', (0, 2), (0, 2), 'LEFT'),
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('PAD', (0, 0), (-1, -1), 5)
            ]))
            row.append(item1)

            # Second item in the row (if it exists)
            if i + 1 < len(items):
                desc2 = Paragraph('<font size=10>{}</font>'.format(items[i + 1]['description']), styleN)
                item_num2 = Paragraph('<font size=10>{}</font>'.format(items[i + 1]['number']), styleN)
                img_path2 = items[i + 1]['png_file']
                with PILImage.open(img_path2) as img:
                    original_width, original_height = img.size
                    new_width = 2 * inch
                    new_height = (new_width / original_width) * original_height
                img2 = ReportLabImage(img_path2, width=new_width, height=new_height)

                item2 = Table([[desc2], [img2], [item_num2]], colWidths=[3 * inch])
                item2.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                    ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                    ('ALIGN', (0, 2), (0, 2), 'LEFT'),
                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                    ('PAD', (0, 0), (-1, -1), 5)
                ]))
                row.append(item2)

            # Combine the row into a table
            # Combine the row into a table
            table = Table([row], colWidths=[3.5 * inch, 3.5 * inch])  # Increase the column widths
            table.setStyle(TableStyle([
                ('RIGHTPADDING', (0, 0), (-1, -1), 20),
                ('LEFTPADDING', (0, 0), (-1, -1), 20),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ]))
            story.append(table)
            story.append(Spacer(1 * inch, 0))

    doc.build(story)
    os.startfile(output_file_path)
    messagebox.showinfo("Success", "All process completed successfully.")

def log_message(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)

root = ThemedTk(theme="equilux")
root.geometry('600x500')
root.title("UPC-A Label Generator")

style = ttk.Style()
style.configure("TButton", background="white")
style.map("TButton", background=[("active", "lightgray")])

import_button = ttk.Button(root, text="Import Excel and Create PNG", command=import_excel_and_create_png)
import_button.pack(padx=20, pady=10)

progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='determinate')
progress_bar.pack(padx=20, pady=10)

log_text = tk.Text(root, wrap=tk.WORD, font=("Courier", 10))
log_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

scrollbar = tk.Scrollbar(log_text, command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=scrollbar.set)

root.mainloop()