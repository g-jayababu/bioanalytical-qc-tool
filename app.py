import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

st.header("1. Upload Files")

col1, col2 = st.columns(2)
with col1:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx"])
with col2:
    uploaded_files = st.file_uploader("Upload TSQ Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

def load_data(file):
    """Smart reader that handles CSV, XLSX, true XLS, and disguised TSV instrument exports."""
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        elif file.name.endswith('.xlsx'):
            return pd.read_excel(file, engine='openpyxl')
        elif file.name.endswith('.xls'):
            try:
                return pd.read_excel(file, engine='xlrd')
            except Exception:
                file.seek(0)
                return pd.read_csv(file, sep='\t')
    except Exception as e:
        return None

if st.button("Generate Dashboard"):
    if setup_file and uploaded_files:
        st.success("Files received! Processing data...")
        
        # --- 1. EXTRACT DATA FROM SETUP FILE ---
        # We read without headers so we can target exact row/column indices (0-indexed)
        setup_df = pd.read_excel(setup_file, header=None)
        
        # Cell B1 (Row 0, Col 1) = Molecule Name
        mol_name = str(setup_df.iloc[0, 1]).strip()
        
        # Cell B2 (Row 1, Col 1) = Study Number
        study_num = str(setup_df.iloc[1, 1]).strip()
        
        # Cell B6 (Row 5, Col 1) = LLOQ
        lloq = pd.to_numeric(setup_df.iloc[5, 1], errors='coerce')
        
        # Cell B7 (Row 6, Col 1) = ULOQ
        uloq = pd.to_numeric(setup_df.iloc[6, 1], errors='coerce')
        
        # Rows 9 and below (Index 8 onwards): Nominal Concentrations
        # Column A (Col 0) = ID, Column B (Col 1) = Nominal Value
        nominals = {}
        for index in range(8, len(setup_df)):
            cc_id = str(setup_df.iloc[index, 0]).strip().upper()
            nom_val = pd.to_numeric(setup_df.iloc[index, 1], errors='coerce')
            
            if cc_id != 'NAN' and cc_id != 'NONE' and pd.notna(nom_val):
                nominals[cc_id] = nom_val

        # Display extracted setup data to confirm it worked
        st.info(f"Setup Loaded | Molecule: {mol_name} | Study: {study_num} | LLOQ: {lloq} | ULOQ: {uloq}")
        st.write(f"Loaded {len(nominals)} nominal concentrations (e.g., {list(nominals.items())[:3]}...)")

        # --- 2. PROCESS INSTRUMENT FILES ---
        global_ccqc_ids = {}
        global_samples = {}
        summary_log = []
        excel_tabs = {}
        
        for file in uploaded_files:
            file.seek(0)
            df = load_data(file)
            
            if df is None:
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": "Could not read file."})
                continue
                
            df = df.dropna(how='all')
            df.columns = [str(c).lower().strip() for c in df.columns]
            
            name_col = next((c for c in df.columns if 'sample name' in c or 'sample' in c), None)
            type_col = next((c for c in df.columns if 'sample type' in c or 'type' in c), None)
            is_area_col = next((c for c in df.columns if 'is area' in c or 'istd' in c), None)
            conc_col = next((c for c in df.columns if 'calc' in c and 'conc' in c), None)
            acc_col = next((c for c in df.columns if 'accur' in c), None)
            
            if not all([name_col, type_col, is_area_col, conc_col, acc_col]):
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": "Missing critical columns."})
                continue
                
            file_dupes = []
            file_sample_dupes = []
            
            cc_qc_df = df[df[type_col].str.contains('Standard|QC', case=False, na=False)].copy()
            cc_qc_df[is_area_col] = pd.to_numeric(cc_qc_df[is_area_col], errors='coerce')
            avg_is_area = cc_qc_df[is_area_col].mean()
            
            lower_is_limit = avg_is_area * 0.4
            upper_is_limit = avg_is_area * 1.6
            
            out_of_range_is = []
            above_uloq = []
            
            for index, row in df.iterrows():
                s_name = str(row[name_col]).strip().upper()
                s_type = str(row[type_col]).strip()
                s_conc = pd.to_numeric(row[conc_col], errors='coerce')
                s_is_area = pd.to_numeric(row[is_area_col], errors='coerce')
                
                if pd.notna(s_is_area) and ('Standard' in s_type or 'QC' in s_type or 'Unknown' in s_type):
                    if s_is_area < lower_is_limit or s_is_area > upper_is_limit:
                        if 'BLK' not in s_name and 'AQSTD' not in s_name:
                             out_of_range_is.append({"Sample Name": s_name, "Concentration": s_conc, "IS Area": s_is_area})
                
                if 'Unknown' in s_type and pd.notna(s_conc):
                    if s_conc > uloq:
                         above_uloq.append({"Sample Name": s_name, "Concentration": s_conc})
                
                if 'Standard' in s_type or 'QC' in s_type:
                    if s_name in global_ccqc_ids:
                        file_dupes.append(f"Duplicate CC/QC: {s_name}")
                    else:
                        global_ccqc_ids[s_name] = {'file': file.name, 'conc': s_conc}
                
                elif 'Unknown' in s_type and 'AQSTD' not in s_name:
                    if s_name in global_samples:
                        file_sample_dupes.append(f"Duplicate Sample: {s_name}")
                    else:
                        global_samples[s_name] = {'file': file.name, 'conc': s_conc}

            clean_df = df[[name_col, type_col, is_area_col, conc_col, acc_col]].copy()
            clean_df.columns = ["Sample Name", "Sample Type", "IS Area", "Calc. Conc.", "% Accuracy"]
            
            excel_tabs[file.name[:31]] = clean_df
            
            summary_log.append({
                "File Name": file.name, 
                "Status": "SUCCESS", 
                "Duplicate CC & QC": "Yes" if file_dupes else "No",
                "Duplicate Samples": "Yes" if file_sample_dupes else "No",
                "Out of Range IS Count": len(out_of_range_is),
