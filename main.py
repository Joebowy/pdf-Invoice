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
    #extracting filename
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    # set the font for pdf
    pdf.set_font(family="Times",size=14, style="B")
    # write in pdf
    pdf.cell(w=40, h=8, txt=f"Invoice No.{invoice_nr}")
    # printing the pdf output for each
    pdf.output(f"PDFs/{filename}.pdf")



