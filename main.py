import pandas as pd
import glob
from fpdf import FPDF, XPos, YPos
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font('Times', 'B', 16)
    pdf.cell(w=50, h=8, text=f'Invoice nr.{invoice_nr}', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Times', 'B', 16)
    pdf.cell(w=50, h=8, text=f'Date: {invoice_date}', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # Define Table Header
    columns = [item.replace("_", " ").title()for item in df.columns]
    pdf.set_font('Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=33, h=8, text=columns[0], border=1)
    pdf.cell(w=70, h=8, text=columns[1], border=1)
    pdf.cell(w=33, h=8, text=columns[2], border=1)
    pdf.cell(w=33, h=8, text=columns[3], border=1)
    pdf.cell(w=25, h=8, text=columns[4], border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Define Table Rows
    for index, row in df.iterrows():
        pdf.set_font('Times',size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=33, h=8, text=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, text=str(row["product_name"]), border=1)
        pdf.cell(w=33, h=8, text=str(row["amount_purchased"]), border=1)
        pdf.cell(w=33, h=8, text=str(row["price_per_unit"]), border=1)
        pdf.cell(w=25, h=8, text=str(row["total_price"]), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)


    total_sum = df["total_price"].sum()
    # Define Table Footer
    pdf.set_font('Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=33, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=33, h=8, border=1)
    pdf.cell(w=33, h=8, border=1)
    pdf.cell(w=25, h=8, text=str(total_sum), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Add total sum sentence
    pdf.set_font('Times', size=10, style='B')
    pdf.cell(w=30, h=8, text=f"The total price is {total_sum}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Add company name and logo
    pdf.set_font('Times', size=10, style='B')
    pdf.cell(w=30, h=8, text=f"pythonhow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f'PDFs/{filename}.pdf')


