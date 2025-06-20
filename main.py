import pandas as pd
import glob
from fpdf import FPDF, XPos, YPos
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    print(invoice_nr, invoice_date)
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font('Times', 'B', 16)
    pdf.cell(w=50, h=8, text=f'Invoice nr.{invoice_nr}', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Times', 'B', 16)
    pdf.cell(w=50, h=8, text=f'Date: {invoice_date}', align='C')

    pdf.output(f'PDFs/{filename}.pdf')


