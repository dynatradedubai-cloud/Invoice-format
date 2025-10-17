
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime

st.title("Invoice Formatter")

uploaded_file = st.file_uploader("Upload Excel Dump", type=["xlsx"])
if uploaded_file:
    dump_df = pd.read_excel(uploaded_file, engine="openpyxl")
    template_path = "Format.xlsx"
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    dump_df = dump_df[dump_df["Sno"] != "Sno"]
    dump_df = dump_df.reset_index(drop=True)

    invoice_data = pd.DataFrame({
        "Sno": range(1, len(dump_df) + 1),
        "Part No": dump_df["Part No"],
        "Brand": dump_df["Brand"],
        "Manf.Part": dump_df["Manf.Part"],
        "Qty": dump_df["Qty on Order"],
        "Unit Price (in AED)": dump_df["Price"],
        "Amount (in AED)": dump_df["Qty on Order"] * dump_df["Price"],
        "Description": dump_df["Description"],
        "COO": dump_df["COO"],
        "HSCODE": dump_df["HSCODE"]
    })

    start_row = 6
    for i, row in invoice_data.iterrows():
        for j, value in enumerate(row):
            ws.cell(row=start_row + i, column=j + 1, value=value)

    customer_cd = dump_df["Customer"].iloc[0]
    invoice_numbers = ", ".join(dump_df["Document No"].dropna().unique())
    invoice_date = datetime.today().strftime("%d/%m/%Y")

    ws["C4"] = customer_cd
    ws["H2"] = invoice_numbers
    ws["H4"] = invoice_date

    total_amount = invoice_data["Amount (in AED)"].sum()
    discount = 0
    net_amount = total_amount - discount
    vat = round(net_amount * 0.05, 2)
    total_aed = net_amount + vat

    ws["G34"] = total_amount
    ws["G35"] = discount
    ws["G36"] = net_amount
    ws["G37"] = vat
    ws["G38"] = total_aed

    output_path = "invoice.xlsx"
    wb.save(output_path)
    with open(output_path, "rb") as f:
        st.download_button("Download Invoice", f, file_name="invoice.xlsx")
