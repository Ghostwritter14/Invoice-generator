import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    # adding the pdf
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # reading the filepath
    filename= Path(filepath).stem
    invoice_no, Date = filename.split("-")

    # Setting the font for the main header
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_no}", ln=1)

    # Date of the invoice
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {Date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Columns header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Adding content in the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Final column to add the total price of the invoice
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add a final note of price
    pdf.set_font(family="Times", size=13, style="B")
    pdf.cell(w=30, h=8, txt=f"Your Total Price is {total_sum} rupees", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=30, h=8, txt=f"Big Bazar")
    pdf.image("bb.png", w=40)

    #Output
    pdf.output(f"PDFs/{filename}.pdf")
