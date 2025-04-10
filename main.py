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
    pdf.cell(w=50, h=8, txt=f"Data: {data}",ln=2)

    # adding header
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)
    column = [item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times",size=14,style="B")
    pdf.set_text_color(80,80,80)
    pdf.cell(w=25,h=8,txt=column[0],border=1)
    pdf.cell(w=55, h=8, txt=column[1], border=1)
    pdf.cell(w=48, h=8, txt=column[2], border=1)
    pdf.cell(w=35, h=8, txt=column[3], border=1)
    pdf.cell(w=25, h=8, txt=column[4], border=1,ln=1)

    #adding rows to the table
    for index,row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=25, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=55, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=48, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["total_price"]), border=1, ln=1)

        # calculating the total price
        total_sum=df["total_price"].sum()
    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=25, h=8, txt="", border=1)
    pdf.cell(w=55, h=8, txt="", border=1)
    pdf.cell(w=48, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt=str(total_sum), border=1, ln=1)

    # adding total sum sentence
    pdf.set_font(family="Times", size=16)
    pdf.cell(w=25, h=8, txt=f"The total price of the products is {str(total_sum)} Euro", ln=1)

    # adding company logo
    pdf.set_font(family="Times", size=16)
    pdf.cell(w=25, h=8, txt="Joe Company Limited")


    # printing the pdf output for each
    pdf.output(f"PDFs/{filename}.pdf")
