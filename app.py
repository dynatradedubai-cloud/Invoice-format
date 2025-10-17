
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime
from invoice_formatter import format_invoice

st.title("Invoice Formatter App")

uploaded_file = st.file_uploader("Upload Excel Dump File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    wb = format_invoice(df)
    output = BytesIO()
    wb.save(output)
    st.download_button("Download Formatted Invoice", output.getvalue(), file_name="Formatted_Invoice.xlsx")
