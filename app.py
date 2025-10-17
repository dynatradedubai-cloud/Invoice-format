import streamlit as st
import pandas as pd
from invoice_formatter import format_invoice

st.set_page_config(page_title="Invoice Generator", layout="wide")

st.title("📄 Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel Dump File", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully.")
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    output_path = "formatted_invoice.xlsx"
    format_invoice(df, output_path)
    with open(output_path, "rb") as f:
        st.download_button("Download Formatted Invoice", f, file_name="formatted_invoice.xlsx")
