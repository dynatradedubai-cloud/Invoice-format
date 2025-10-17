
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from io import BytesIO
from datetime import datetime

st.title("Invoice Formatter")

uploaded_file = st.file_uploader("Upload Excel Dump", type=["xlsx"])

if uploaded_file:
    dump_df = pd.read_excel(uploaded_file, engine="openpyxl")
    dump_df = dump_df[dump_df["Sno"].apply(lambda x: str(x).isdigit())]

    # Load template
    template_path = "Format.xlsx"
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # Fill header fields
    customer_cd = dump_df["Customer"].iloc[0]
    invoice_nos = ", ".join(sorted(dump_df["Document No"].dropna().unique()))
    today_date = datetime.today().strftime("%d/%m/%Y")

    ws["C4"] = customer_cd
    ws["H2"] = invoice_nos
    ws["H4"] = today_date

    # Fill invoice table
    start_row = 6
    columns = ["Part No", "Brand", "Manf.Part", "Qty on Order", "Price", "Description", "COO", "HSCODE"]
    total_amount = 0

    for idx, row in enumerate(dump_df[columns].itertuples(index=False), start=1):
        excel_row = start_row + idx - 1
        qty = row[3]
        price = row[4]
        amount = qty * price
        total_amount += amount

        ws.cell(row=excel_row, column=1, value=idx)  # Sno
        ws.cell(row=excel_row, column=2, value=row[0])  # Part No
        ws.cell(row=excel_row, column=3, value=row[1])  # Brand
        ws.cell(row=excel_row, column=4, value=row[2])  # Manf.Part
        ws.cell(row=excel_row, column=5, value=qty)  # Qty
        ws.cell(row=excel_row, column=6, value=round(price, 2))  # Unit Price
        ws.cell(row=excel_row, column=7, value=round(amount, 2))  # Amount
        ws.cell(row=excel_row, column=8, value=row[5])  # Description
        ws.cell(row=excel_row, column=9, value=row[6])  # COO
        ws.cell(row=excel_row, column=10, value=row[7])  # HSCODE

    # Fill summary section
    ws["G34"] = round(total_amount, 2)
    ws["G35"] = 0
    ws["G36"] = round(total_amount, 2)
    ws["G37"] = round(total_amount * 0.05, 2)
    ws["G38"] = round(total_amount * 1.05, 2)

    # Adjust column widths
    for col in range(1, 11):
        max_length = 0
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col)].width = max_length + 2

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    st.download_button("Download Formatted Invoice", data=output.getvalue(), file_name="Formatted_Invoice.xlsx")
