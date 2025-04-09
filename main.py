# importing the data
import pandas as pd
import glob

# having excess to all the data
filepaths = glob.glob("invoices/*.xlsx")

# Reading the excel file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)