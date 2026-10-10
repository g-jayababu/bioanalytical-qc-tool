import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docxtpl import DocxTemplate

st.title("Bioanalytical Auto-QC & Reporting")
st.header("1. Upload Files")

col1, col2, col3 = st.columns(3)
with col1:
    csv_file = st.file_uploader("Upload TSQ Export", type=["csv", "xls", "xlsx"])
with col2:
    pdf_file = st.file_uploader("Upload Protocol PDF", type=["pdf"])
with col3:
    docx_template = st.file_uploader("Upload Word Template", type=["docx"])

if csv_file is not None:
    st.header("2. Data Preview")
    
    if csv_file.name.endswith('.csv'):
        df = pd.read_csv(csv_file)
    else:
        df = pd.read_excel(csv_file)
        
    st.dataframe(df.head())

if st.button("Run QC and Generate Report"):
    if csv_file and pdf_file and docx_template:
        st.success("Files received! Processing data...")
        
        # 1. Read the file into memory again for processing
        if csv_file.name.endswith('.csv'):
            df = pd.read_csv(csv_file)
        else:
            df = pd.read_excel(csv_file)
        
        # 2. Clean the data: Drop any completely empty rows
        df = df.dropna(how='all')
        
        # 3. Tell the user how many rows were found
        st.write(f"Data loaded successfully. Total rows: {len(df)}")
        
        # 4. Separate the data into groups for QC checking
        try:
            calibrators = df[df['Sample Type'] == 'Standard']
            qcs = df[df['Sample Type'] == 'QC']
            blank_samples = df[df['Sample Type'] == 'Blank']
            unknowns = df[df['Sample Type'] == 'Unknown']
            
            st.info(f"Found {len(calibrators)} Calibrators, {len(qcs)} QCs, and {len(blank_samples)} Blanks.")
            
            st.subheader("QC Samples Found:")
            st.dataframe(qcs.head())
            
        except KeyError:
            st.error("Error: Could not find a column named 'Sample Type'. Please open your export file and check the exact column header name.")
            
    else:
        st.error("Please upload all three files to proceed.")
