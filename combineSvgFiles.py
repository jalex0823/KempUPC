import fitz
import os
from reportlab.platypus import Flowable, SimpleDocTemplate, Table
from reportlab.lib.units import inch
from reportlab.graphics import renderPDF
from reportlab.lib.pagesizes import letter
from svglib.svglib import svg2rlg
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import TableStyle

class SVGFlowable(Flowable):
    def __init__(self, svg_file, width=200):
        super().__init__()
        self.svg_file = svg_file
        self.width = width
        self.height = self.calculate_svg_height()

    def calculate_svg_height(self):
        drawing = svg2rlg(self.svg_file)
        aspect_ratio = drawing.width / drawing.height
        return self.width / aspect_ratio

    def wrap(self, available_width, available_height):
        return self.width, self.height

    def draw(self):
        drawing = svg2rlg(self.svg_file)
        renderPDF.draw(drawing, self.canv, 0, 0)

def merge_svgs_to_pdf(folder_path, output_pdf_name):
    svg_files = sorted([file.zfill(11) for file in os.listdir(folder_path) if file.endswith(".svg")])

    output_path = os.path.join(folder_path, output_pdf_name)
    doc = SimpleDocTemplate(output_path, pagesize=letter, topMargin=0, bottomMargin=0)  # Removed margins

    styles = getSampleStyleSheet()
    story = []
    data = []
    row_heights = []
    for i in range(0, len(svg_files), 2):
        row = []
        max_height = 0
        for j in range(i, min(i + 2, len(svg_files))):
            svg_file = svg_files[j]
            flowable = SVGFlowable(os.path.join(folder_path, svg_file), width=200)
            row.append(flowable)
            max_height = max(max_height, flowable.height)
        data.append(row)
        row_heights.append(max_height)

    table = Table(data, colWidths=[doc.width/2.0]*2, rowHeights=row_heights)
    table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'),
                               ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(table)

    doc.build(story)
    print(f"PDF '{output_pdf_name}' has been successfully created at {output_path}.")

    return output_path

def overlay_pdfs(folder_path, output_pdf_name, merged_svg_pdf, combined_pdf):
    combined_pdf_path = os.path.join(folder_path, combined_pdf)
    merged_svg_pdf_path = os.path.join(folder_path, merged_svg_pdf)
    combined_doc = fitz.open(combined_pdf_path)
    merged_doc = fitz.open(merged_svg_pdf_path)

    for page_number in range(len(combined_doc)):
        page = combined_doc[page_number]

        if page_number < len(merged_doc):
            overlay_page = merged_doc.load_page(page_number)
            overlay_rect = overlay_page.rect

            overlay_rect.y0 -= 0
            overlay_rect.y1 -= 0
            
            page.show_pdf_page(overlay_rect, merged_doc, overlay_page.number)

    output_path = os.path.join(folder_path, output_pdf_name)

    combined_doc.save(output_path)

    print(f"PDFs '{combined_pdf}' and '{merged_svg_pdf}' have been overlaid into '{output_pdf_name}' at {output_path}.")

folder_path = r"C:\Temp"
merged_svg_pdf = "Merged.pdf"
pdf_path = merge_svgs_to_pdf(folder_path, merged_svg_pdf)
combined_pdf = "Combined.pdf"
overlay_pdfs(folder_path, "Overlayed.pdf", merged_svg_pdf, combined_pdf)