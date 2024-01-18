import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    filename = Path(filepath).stem
    invoice_num = filename.split("-")[0]

    date = filename.split("-")[1]

    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"invoice nr. {invoice_num}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date. {date}")
    pdf.output(f"PDFs/{filename}.pdf")

