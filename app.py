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

def load_data(file):
    """Smart reader that handles CSV, XLSX, true XLS, and disguised TSV instrument exports."""
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        elif file.name.endswith('.xlsx'):
            return pd.read_excel(file, engine='openpyxl')
        elif file.name.endswith('.xls'):
            try:
                # Try reading as a proper older Excel file first
                return pd.read_excel(file, engine='xlrd')
            except Exception:
                # If it fails, it is likely a disguised TSQ text file
                file.seek(0)
                return pd.read_csv(file, sep='\t')
    except Exception as e:
        st.error(f"Error reading the data file: {e}")
        return None

if csv_file is not None:
    st.header("2. Data Preview")
    csv_file.seek(0) # Reset file pointer before reading
    df = load_data(csv_file)
    
    if df is not None:
        st.dataframe(df.head())

if st.button("Run QC and Generate Report"):
    if csv_file and pdf_file and docx_template:
        st.success("Files received! Processing data...")
        
        # Reset file pointer before reading a second time on button click
        csv_file.seek(0)
        df = load_data(csv_file)
        
        if df is not None:
            # Clean the data: Drop any completely empty rows
            df = df.dropna(how='all')
            
            st.write(f"Data loaded successfully. Total rows: {len(df)}")
            
            # Separate the data into groups for QC checking
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
