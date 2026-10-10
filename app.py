import streamlit as st
import pandas as pd
import fitz  
from docxtpl import DocxTemplate

st.title("Bioanalytical Auto-QC & Reporting")
st.header("1. Upload Files")

col1, col2, col3 = st.columns(3)
with col1:
    csv_file = st.file_uploader("Upload TSQ CSV", type=["csv"])
with col2:
    pdf_file = st.file_uploader("Upload Protocol PDF", type=["pdf"])
with col3:
    docx_template = st.file_uploader("Upload Word Template", type=["docx"])

if csv_file is not None:
    st.header("2. Data Preview")
    df = pd.read_csv(csv_file)
    st.dataframe(df.head())

if st.button("Run QC and Generate Report"):
    if csv_file and pdf_file and docx_template:
        st.success("Files received! The QC rules will go here.")
    else:
        st.error("Please upload all three files to proceed.")
