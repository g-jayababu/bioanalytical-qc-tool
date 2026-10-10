import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

st.header("1. Upload Files")

col1, col2, col3 = st.columns(3)
with col1:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx", "xlsm"])
with col2:
    uploaded_files = st.file_uploader("Upload TSQ Quantis Plus Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True)
with col3:
    repeat_files = st.file_uploader("Upload Repeat Files (Optional)", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

def load_data(file):
    file_name_lower = file.name.lower()
    dummy_cols = list(range(100))
    if file_name_lower.endswith('.csv'):
        try: return pd.read_csv(file, names=dummy_cols, on_bad_lines='skip'), None
        except Exception as e: return None, f"CSV Error: {str(e)}"
    elif file_name_lower.endswith('.xlsx') or file_name_lower.endswith('.xlsm'):
        try: return pd.read_excel(file, engine='openpyxl'), None
        except Exception as e: return None, f"XLSX Error: {str(e)}"
    elif file_name_lower.endswith('.xls'):
        try: return pd.read_excel(file, engine='xlrd'), None
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
                    except Exception as e4: return None, f"Failed all formats. Last error: {str(e4)}"
    return None, "Unsupported file extension."

def extract_sort_num(val):
    nums = re.findall(r'\d+', str(val))
    return int(nums[-1]) if nums else 0

if st.button("Generate Dashboard"):
    if setup_file and (uploaded_files or repeat_files):
        st.success("Files received! Processing data...")
        try:
            setup_df = pd.read_excel(setup_file, header=None, engine='openpyxl')
            mol_name = str(setup_df.iloc[0, 1]).strip()
            study_num = str(setup_df.iloc[1, 1]).strip()
            num_subj_raw = setup_df.iloc[2, 1]
            num_subj = int(pd.to_numeric(num_subj_raw, errors='coerce')) if pd.notna(num_subj_raw) else 0
            periods_str = str(setup_df.iloc[3, 1]).strip() if pd.notna(setup_df.iloc[3, 1]) else ""
            timepoints_str = str(setup_df.iloc[4, 1]).strip() if pd.notna(setup_df.iloc[4, 1]) else ""
            
            lloq = pd.to_numeric(setup_df.iloc[5, 1], errors='coerce')
            uloq = pd.to_numeric(setup_df.iloc[6, 1], errors='coerce')
            
            nominals = {}
            std_names = []
            for index in range(8, len(setup_df)):
                cc_id = str(setup_df.iloc[index, 0]).strip().upper()
                nom_val = pd.to_numeric(setup_df.iloc[index, 1], errors='coerce')
                if cc_id not in ['NAN', 'NONE'] and pd.notna(nom_val):
                    nominals[cc_id] = nom_val
                    if 'STD' in cc_id or 'STANDARD' in cc_id: std_names.append(cc_id)
            std_names = sorted(std_names)
        except Exception as e:
            st.error(f"Failed to read Setup file. Error: {str(e)}")
            st.stop()

        summary_log = []
        excel_tabs = {}
        cc_file_data = []
        qc_file_data = []
        global_samples = {}
        repeat_samples = {}
        batch_meta_data = {}
        all_repeats_data = [] 
        
        all_files_to_process = []
        if uploaded_files:
            for f in uploaded_files: all_files_to_process.append((f, False))
        if repeat_files:
            for f in repeat_files: all_files_to_process.append((f, True))
        
        for file, is_repeat in all_files_to_process:
            file.seek(0)
            
            base_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', file.name)[:31]
            safe_sheet_name = base_sheet_name
            counter = 1
            while safe_sheet_name in excel_tabs:
                suffix = f"_{counter}"
                safe_sheet_name = f"{base_sheet_name[:31-len(suffix)]}{suffix}"
                counter += 1
            
            df, error_msg = load_data(file)
            if df is None:
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": error_msg, "Nominal Conc. Check": "", "Duplicate CC & QC ID Check": "", "Duplicate Sample Name Check": ""})
                continue
                
            df = df.dropna(how='all').dropna(how='all', axis=1).reset_index(drop=True)
            
            header_idx = None
            for idx, row in df.iterrows():
                row_string = str(row.values).lower()
                if 'sample name' in row_string or ('sample' in row_string and 'type' in row_string):
                    header_idx = idx
                    break
            
            slope, intercept, r_val = "", "", ""
            if header_idx is not None:
                for idx, row in df.iloc[:header_idx].iterrows():
                    for c_idx in range(len(row)):
                        val = str(row.values[c_idx]).strip().lower()
                        if val == 'slope' and c_idx + 1 < len(row): slope = pd.to_numeric(row.values[c_idx + 1], errors='coerce')
                        elif val == 'intercept' and c_idx + 1 < len(row): intercept = pd.to_numeric(row.values[c_idx + 1], errors='coerce')
                        elif ('correlatio' in val or val == 'r') and c_idx + 1 < len(row): r_val = pd.to_numeric(row.values[c_idx + 1], errors='coerce')
            
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
            
            analyte_area_col = next((c for c in df.columns if 'area' in c and 'ratio' not in c and c != is_area_col), None)
            
            date_col = next((c for c in df.columns if 'acq date' in c or 'acquisition date' in c or 'date/time' in c or 'date' in c), None)
            acq_date = ""
            if date_col and not df[date_col].dropna().empty:
                raw_date = str(df[date_col].dropna().iloc[0]).strip()
                try: acq_date = pd.to_datetime(raw_date).strftime('%d-%b-%Y').upper()
                except Exception: acq_date = raw_date
            if not acq_date: acq_date = pd.Timestamp.now().strftime('%d-%b-%Y').upper()
            
            if not all([name_col, type_col, is_area_col, conc_col]):
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": "Missing critical data columns.", "Nominal Conc. Check": "", "Duplicate CC & QC ID Check": "", "Duplicate Sample Name Check": ""})
                continue
            
            subj_id = "S1"
            match = re.search(r'(S\d+(?:[-_]S\d+)*)', file.name, re.IGNORECASE)
            if match: subj_id = match.group(1).upper()
            elif 'REPEAT' in file.name.upper(): subj_id = "Repeat"
            elif 'ISR' in file.name.upper(): subj_id = "ISR"
                
            ccid_str = "CC-01"
            for s_name in df[name_col]:
                if 'CC-' in str(s_name).upper():
                    cc_match = re.search(r'CC-\d+', str(s_name).upper())
                    if cc_match:
                        ccid_str = cc_match.group(0)
                        break

            file_cc_dict = {'Date of first injection': acq_date, 'SUBJECT ID': subj_id, 'CCID': ccid_str}
            for std in std_names: file_cc_dict[std] = ""
            
            lqc_list, mmqc_list, mqc_list, hqc_list = [], [], [], []
            lqc_ids, mmqc_ids, mqc_ids, hqc_ids = [], [], [], []
            
            file_dupes, file_sample_dupes = [], []
            seen_ccqc, seen_samples = set(), set()
            
            std_a_area, std_zero_is_area = None, None
            cc_pass, lqc_pass, mmqc_pass, mqc_pass, hqc_pass = 0, 0, 0, 0, 0
            out_of_range_is, above_uloq = [], []
            carryover_samples = []
            
            cc_qc_df = df[df[type_col].str.contains('Standard|QC|Quality', case=False, na=False)].copy()
            cc_qc_df[is_area_col] = pd.to_numeric(cc_qc_df[is_area_col], errors='coerce')
            avg_is_area = cc_qc_df[is_area_col].mean()
            lower_is_limit = avg_is_area * 0.4 if pd.notna(avg_is_area) else 0
            upper_is_limit = avg_is_area * 1.6 if pd.notna(avg_is_area) else float('inf')
            
            ccqc_rows = []
            sample_
