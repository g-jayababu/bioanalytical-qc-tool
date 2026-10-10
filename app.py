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
        st.success("Files received! Processing data...")
        
        # 1. Read the CSV into memory
        df = pd.read_csv(csv_file)
        
        # 2. Clean the data: Drop any completely empty rows
        df = df.dropna(how='all')
        
        # 3. Tell the user how many rows were found
        st.write(f"Data loaded successfully. Total rows: {len(df)}")
        
        # 4. Separate the data into groups for QC checking
        # We use a 'try/except' block so the app tells us what went wrong instead of crashing
        try:
            # The TSQ software usually labels these in a 'Sample Type' column
            calibrators = df[df['Sample Type'] == 'Standard']
            qcs = df[df['Sample Type'] == 'QC']
            blank_samples = df[df['Sample Type'] == 'Blank']
            unknowns = df[df['Sample Type'] == 'Unknown']
            
            st.info(f"Found {len(calibrators)} Calibrators, {len(qcs)} QCs, and {len(blank_samples)} Blanks.")
            
            # Show a preview of just the QCs on the screen
            st.subheader("QC Samples Found:")
            st.dataframe(qcs.head())
            
        except KeyError:
            st.error("Error: Could not find a column named 'Sample Type'. Please open your CSV in Excel and check the exact column header name.")
            
    else:
        st.error("Please upload all three files to proceed.")
