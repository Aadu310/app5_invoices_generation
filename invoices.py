import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # pdf will be in portrait mode and in A4 format
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # add the new pdf page
    pdf.add_page()
    filename = Path(filepath).stem
    # extracts the name of the filename, this case 10001 is extracted.
    invoice_nr=filename.split("-")[0]
    # sets the font family,size and style of the letters
    pdf.set_font(family="Times",size=16, style="B")
    # sets the cell width and height and text is printed. Here txt name is dynamically taken
    pdf.cell(w=50,h=8, txt=f"Invoice num.{invoice_nr}")
    # creates a pdf folder and related files are created.
    pdf.output(f"PDFs/{filename}.pdf")
