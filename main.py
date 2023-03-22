import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

files = glob.glob('Invoice\\*.xlsx')

for file in files:
    data = pd.read_excel(file, sheet_name='Sheet 1')
    pdf = FPDF(orientation="P", unit="mm",format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    extract_filename = Path(file).stem
    invoice_no = extract_filename.split("-")[0]
    date = extract_filename.split("-")[1]
    pdf.set_font(family="Times", style="B", size=17)
    pdf.cell(w=50, h=18, txt=f"Invoice No: {invoice_no}", ln=1)
    pdf.cell(w=70, h=18, txt=f"Date: {date}", ln=1)
    pdf.output(f"Converted_PDFs\\{extract_filename}.pdf")
