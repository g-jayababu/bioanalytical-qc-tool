import streamlit as st
import pandas as pd
import io
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

st.header("1. Upload Files")

col1, col2 = st.columns(2)
with col1:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx", "xlsm"])
with col2:
    uploaded_files = st.file_uploader("Upload TSQ Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

def load_data(file):
    """Smart reader that handles true Excel files and UTF-16 TSQ text files disguised as .xls"""
    file_name_lower = file.name.lower()
    dummy_cols = list(range(100))
    
    if file_name_lower.endswith('.csv'):
        try:
            return pd.read_csv(file, names=dummy_cols, on_bad_lines='skip'), None
        except Exception as e:
            return None, f"CSV Error: {str(e)}"
            
    elif file_name_lower.endswith('.xlsx') or file_name_lower.endswith('.xlsm'):
        try:
            return pd.read_excel(file, engine='openpyxl'), None
        except Exception as e:
            return None, f"XLSX Error: {str(e)}"
            
    elif file_name_lower.endswith('.xls'):
        try:
            return pd.read_excel(file, engine='xlrd'), None
        except Exception as e1:
            try:
                file.seek(0)
                return pd.read_csv(file, sep='\t', names=dummy_cols, on_bad_lines='skip'), None
            except Exception as e2:
                try:
                    file.seek(0)
                    return pd.read_csv(file, sep='\t', encoding='utf-16', names=dummy_cols, on_bad_lines='skip'), None
                except Exception as e3:
                    try:
                        file.seek(0)
                        return pd.read_csv(file, sep='\t', encoding='latin1', names=dummy_cols, on_bad_lines='skip'), None
                    except Exception as e4:
                        return None, f"Failed all formats. Last error: {str(e4)}"
                        
    return None, "Unsupported file extension."

if st.button("Generate Dashboard"):
    if setup_file and uploaded_files:
        st.success("Files received! Processing data...")
        
        # --- 1. EXTRACT DATA FROM SETUP FILE ---
        try:
            setup_df = pd.read_excel(setup_file, header=None, engine='openpyxl')
            mol_name = str(setup_df.iloc[0, 1]).strip()
            study_num = str(setup_df.iloc[1, 1]).strip()
            lloq = pd.to_numeric(setup_df.iloc[5, 1], errors='coerce')
            uloq = pd.to_numeric(setup_df.iloc[6, 1], errors='coerce')
            
            nominals = {}
            for index in range(8, len(setup_df)):
                cc_id = str(setup_df.iloc[index, 0]).strip().upper()
                nom_val = pd.to_numeric(setup_df.iloc[index, 1], errors='coerce')
                if cc_id != 'NAN' and cc_id != 'NONE' and pd.notna(nom_val):
                    nominals[cc_id] = nom_val
            st.info(f"Setup Loaded | Molecule: {mol_name} | Study: {study_num} | LLOQ: {lloq} | ULOQ: {uloq}")
        except Exception as e:
            st.error(f"Failed to read Setup file. Error: {str(e)}")
            st.stop()

        # --- 2. PROCESS INSTRUMENT FILES ---
        global_ccqc_ids = {}
        global_samples = {}
        summary_log = []
        excel_tabs = {}
        
        for file in uploaded_files:
            file.seek(0)
            df, error_msg = load_data(file)
            
            if df is None:
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": error_msg})
                continue
                
            df = df.dropna(how='all')
            df = df.dropna(how='all', axis=1)
            df = df.reset_index(drop=True)
            
            # --- FIND THE TRUE HEADER ROW ---
            header_idx = None
            for idx, row in df.iterrows():
                row_string = str(row.values).lower()
                if 'sample name' in row_string or ('sample' in row_string and 'type' in row_string):
                    header_idx = idx
                    break
            
            if header_idx is not None:
                new_headers = df.iloc[header_idx]
                df = df.iloc[header_idx + 1:].copy()
                df.columns = [str(c).lower().strip() for c in new_headers]
            else:
                df.columns = [str(c).lower().strip() for c in df.columns]

            df = df.loc[:, ~df.columns.str.contains('^nan$', case=False, na=False)]

            name_col = next((c for c in df.columns if 'sample name' in c or 'sample' in c), None)
            type_col = next((c for c in df.columns if 'sample type' in c or 'type' in c), None)
            is_area_col = next((c for c in df.columns if ('is area' in c or 'is peak area' in c or 'istd' in c) and 'ratio' not in c and 'analyte' not in c), None)
            conc_col = next((c for c in df.columns if 'calc' in c and 'conc' in c), None)
            acc_col = next((c for c in df.columns if 'accur' in c), None)
            
            if not all([name_col, type_col, is_area_col, conc_col, acc_col]):
                missing = []
                if not name_col: missing.append("Sample Name")
                if not type_col: missing.append("Sample Type")
                if not is_area_col: missing.append("IS Area")
                if not conc_col: missing.append("Calc Conc")
                if not acc_col: missing.append("Accuracy")
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": f"Missing columns: {', '.join(missing)}"})
                continue
                
            file_dupes = []
            file_sample_dupes = []
            
            cc_qc_df = df[df[type_col].str.contains('Standard
