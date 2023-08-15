import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*xlsx")


for filepath in filepaths:

    # pdf will be in portrait mode and in A4 format
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # add the new pdf page
    pdf.add_page()
    filename = Path(filepath).stem
    # extracts the name of the filename, this case 10001 is extracted.
    invoice_nr, date = filename.split("-")
    #or it can  be written as
    # date = filename.split("-")[1]
    # sets the font family,size and style of the letters
    pdf.set_font(family="Times",size=16, style="B")
    # sets the cell width and height and text is printed. Here txt name is dynamically taken
    pdf.cell(w=50,h=8, txt=f"Invoice num.{invoice_nr}",ln=1)
    # creates a pdf folder and related files are created.

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:{date}",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_"," ").title() for item in columns]
    pdf.set_font(family= "Times",size=10,style="B")
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1],  border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=20, h=8, txt=columns[4], border=1, ln=1)

    # read the excel file, add the rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30,h=8,txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=20, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # add the total sum
    total_sum = df["total_price"].sum()
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=20, h=8, txt=str(total_sum), border=1, ln=1)

    # add total sum sentence
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")