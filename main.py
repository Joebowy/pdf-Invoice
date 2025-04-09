# importing the data
import pandas as pd
import glob
from fpdf import FPDF
# extracting file name
from pathlib import Path

# having excess to all the data
filepaths = glob.glob("invoices/*.xlsx")

# Reading the excel file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # set the page for the pdf
    pdf = FPDF(orientation="p", unit="mm", format="A4")
    pdf.add_page()
    # extracting filename
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    data = filename.split("-")[1]
    # set the font for pdf
    pdf.set_font(family="Times", size=16, style="B")
    # write in pdf
    pdf.cell(w=50, h=8, txt=f"Invoice No.{invoice_nr}", ln=1)

    # adding data
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Data: {data}")

    # printing the pdf output for each
    pdf.output(f"PDFs/{filename}.pdf")
