
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import io

st.title("Invoice Formatter")

uploaded_dump = st.file_uploader("Upload Excel Dump", type=["xlsx"])
template_file = "Format.xlsx"

if uploaded_dump:
    df = pd.read_excel(uploaded_dump, sheet_name=0, engine="openpyxl")
    df = df[df["Sno"] != "Sno"]
    df = df.reset_index(drop=True)
    df["Sno"] = range(1, len(df) + 1)
    df["Amount"] = df["Qty on Order"] * df["Price"]
    invoice_data = df[["Sno", "Part No", "Brand", "Manf.Part", "Qty on Order", "Price", "Amount", "Description", "COO", "HSCODE"]]
    invoice_data.columns = ["Sno", "Part No", "Brand", "Manf.Part", "Qty", "Unit Price (in AED)", "Amount (in AED)", "Description", "COO", "HSCODE"]

    wb = load_workbook(template_file)
    ws = wb.active

    start_row = 6
    for i, row in invoice_data.iterrows():
        for j, value in enumerate(row):
            ws.cell(row=start_row + i, column=j + 1, value=value)

    ws["C4"] = df["Customer"].iloc[0]
    ws["H2"] = ", ".join(sorted(df["Document No"].unique()))
    ws["H4"] = datetime.today().strftime("%d/%m/%Y")

    total_amount = invoice_data["Amount (in AED)"].sum()
    ws["G34"] = total_amount
    ws["G35"] = 0
    ws["G36"] = total_amount
    ws["G37"] = round(total_amount * 0.05, 2)
    ws["G38"] = round(total_amount * 1.05, 2)

    output = io.BytesIO()
    wb.save(output)
    st.download_button("Download Invoice", data=output.getvalue(), file_name="invoice.xlsx")
