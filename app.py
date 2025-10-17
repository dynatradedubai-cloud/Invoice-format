
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

st.title("Invoice Formatter")

uploaded_file = st.file_uploader("Upload Dump Excel File", type=["xlsx"])
if uploaded_file:
    dump_df = pd.read_excel(uploaded_file, engine="openpyxl")
    dump_df = dump_df[dump_df["Sno"] != "Sno"]
    dump_df.reset_index(drop=True, inplace=True)

    wb = openpyxl.load_workbook("Format.xlsx")
    ws = wb.active

    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    invoice_data = []
    serial_number = 1
    for _, row in dump_df.iterrows():
        qty = row["Qty on Order"]
        unit_price = row["Price"]
        amount = qty * unit_price
        invoice_data.append([
            serial_number,
            row["Part No"],
            row["Brand"],
            row["Manf.Part"],
            qty,
            unit_price,
            amount,
            row["Description"],
            row["COO"],
            row["HSCODE"]
        ])
        serial_number += 1

    start_row = 6
    for i, data_row in enumerate(invoice_data):
        for j, value in enumerate(data_row):
            ws.cell(row=start_row + i, column=j + 1, value=value)

    ws["C4"] = dump_df.iloc[0]["Customer"]
    ws["H2"] = ", ".join(sorted(set(dump_df["Document No"].astype(str))))
    ws["H4"] = datetime.today().strftime("%d/%m/%Y")

    summary_start = start_row + len(invoice_data)
    total_amount = sum(row[6] for row in invoice_data)
    discount = 0
    net_amount = total_amount - discount
    vat = round(net_amount * 0.05, 2)
    total_aed = net_amount + vat

    ws[f"D{summary_start}"] = "TOTAL AMOUNT"
    ws[f"G{summary_start}"] = total_amount
    ws[f"D{summary_start+1}"] = "DISCOUNT"
    ws[f"G{summary_start+1}"] = discount
    ws[f"D{summary_start+2}"] = "NET AMOUNT"
    ws[f"G{summary_start+2}"] = net_amount
    ws[f"D{summary_start+3}"] = "VAT 5%"
    ws[f"G{summary_start+3}"] = vat
    ws[f"D{summary_start+4}"] = "TOTAL AED"
    ws[f"G{summary_start+4}"] = total_aed

    output_file = "invoice.xlsx"
    wb.save(output_file)
    with open(output_file, "rb") as f:
        st.download_button("Download Invoice", f, file_name=output_file)
