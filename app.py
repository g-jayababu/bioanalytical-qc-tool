import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

# Initialize session state for the file uploaders to allow clearing
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

def clear_files():
    st.session_state.uploader_key += 1

st.header("1. Upload Files")

col1, col2, col3, col4 = st.columns(4)
with col1:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx", "xlsm"], key=f"setup_{st.session_state.uploader_key}")
with col2:
    uploaded_files = st.file_uploader("Upload TSQ Quantis Plus Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True, key=f"main_{st.session_state.uploader_key}")
with col3:
    repeat_files = st.file_uploader("Upload Repeat Files (Optional)", type=["csv", "xls", "xlsx"], accept_multiple_files=True, key=f"repeat_{st.session_state.uploader_key}")
with col4:
    isr_files = st.file_uploader("Upload ISR Files (Optional)", type=["csv", "xls", "xlsx"], accept_multiple_files=True, key=f"isr_{st.session_state.uploader_key}")

st.button("🧹 Clear All Uploaded Files", on_click=clear_files)

def load_data(file):
    file_name_lower = file.name.lower()
    dummy_cols = list(range(100))
    if file_name_lower.endswith('.csv'):
        try: return pd.read_csv(file, names=dummy_cols, on_bad_lines='skip', na_filter=False), None
        except Exception as e: return None, f"CSV Error: {str(e)}"
    elif file_name_lower.endswith('.xlsx') or file_name_lower.endswith('.xlsm'):
        try: return pd.read_excel(file, engine='openpyxl', na_filter=False), None
        except Exception as e: return None, f"XLSX Error: {str(e)}"
    elif file_name_lower.endswith('.xls'):
        try: return pd.read_excel(file, engine='xlrd', na_filter=False), None
        except Exception as e1:
            try:
                file.seek(0)
                return pd.read_csv(file, sep='\t', names=dummy_cols, on_bad_lines='skip', na_filter=False), None
            except Exception as e2:
                try:
                    file.seek(0)
                    return pd.read_csv(file, sep='\t', encoding='utf-16', names=dummy_cols, on_bad_lines='skip', na_filter=False), None
                except Exception as e3:
                    try:
                        file.seek(0)
                        return pd.read_csv(file, sep='\t', encoding='latin1', names=dummy_cols, on_bad_lines='skip', na_filter=False), None
                    except Exception as e4: return None, f"Failed all formats. Last error: {str(e4)}"
    return None, "Unsupported file extension."

def extract_sort_num(val):
    nums = re.findall(r'\d+', str(val))
    return int(nums[-1]) if nums else 0

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def get_base_name(name):
    if not isinstance(name, str): return str(name)
    return re.sub(r'[-_](ISR\d*|R\d*|REPEAT\d*)$', '', name, flags=re.IGNORECASE).strip()

if st.button("Generate Dashboard"):
    if setup_file and uploaded_files:
        st.success("Files received! Processing data...")
        try:
            setup_df = pd.read_excel(setup_file, header=None, engine='openpyxl')
            mol_name = str(setup_df.iloc[0, 1]).strip()
            study_num = str(setup_df.iloc[1, 1]).strip()
            num_subj_raw = setup_df.iloc[2, 1]
            num_subj = int(pd.to_numeric(num_subj_raw, errors='coerce')) if pd.notna(num_subj_raw) else 0
            
            periods_str = str(setup_df.iloc[3, 1]).strip() if pd.notna(setup_df.iloc[3, 1]) else ""
            timepoints_str = str(setup_df.iloc[4, 1]).strip() if pd.notna(setup_df.iloc[4, 1]) else ""
            
            periods_list = [p.strip() for p in periods_str.split(',')] if periods_str else []
            timepoints_list = [tp.strip() for tp in timepoints_str.split(',')] if timepoints_str else []
            
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
        batch_meta_data = {}
        all_repeats_data = [] 
        repeat_samples = {}
        isr_samples = {}
        
        combined_files = [(f, "Main") for f in uploaded_files]
        if repeat_files: combined_files.extend([(f, "Repeat") for f in repeat_files])
        if isr_files: combined_files.extend([(f, "ISR") for f in isr_files])
        
        for file, file_category in combined_files:
            file.seek(0)
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', file.name)[:31]
            
            df, error_msg = load_data(file)
            if df is None:
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": error_msg, "Nominal Conc. Check": "", "Duplicate CC & QC ID Check": "", "Duplicate Sample Name Check": ""})
                continue
                
            df = df.replace(r'^\s*$', np.nan, regex=True).dropna(how='all').dropna(how='all', axis=1).reset_index(drop=True)
            df = df.fillna("")
            
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

            df = df.loc[:, ~df.columns.str.contains('^nan$|^$', case=False, na=False)]

            name_col = next((c for c in df.columns if 'sample name' in c or 'sample' in c), None)
            type_col = next((c for c in df.columns if 'sample type' in c or 'type' in c), None)
            is_area_col = next((c for c in df.columns if ('is area' in c or 'is peak area' in c or 'istd' in c) and 'ratio' not in c and 'analyte' not in c), None)
            conc_col = next((c for c in df.columns if 'calc' in c and 'conc' in c), None)
            acc_col = next((c for c in df.columns if 'accur' in c), None)
            
            analyte_area_col = next((c for c in df.columns if 'area' in c and 'ratio' not in c and c != is_area_col), None)
            date_col = next((c for c in df.columns if 'acq date' in c or 'acquisition date' in c or 'date/time' in c or 'date' in c), None)
            
            acq_date = ""
            if date_col and not df[date_col].empty:
                raw_date = str(df[date_col].iloc[0]).strip()
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
            for s_name_raw in df[name_col]:
                if 'CC-' in str(s_name_raw).upper():
                    cc_match = re.search(r'CC-\d+', str(s_name_raw).upper())
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
            cc_pass, cc_total = 0, 0
            lqc_list_objs, mmqc_list_objs, mqc_list_objs, hqc_list_objs = [], [], [], []
            failed_cc_ids = []
            
            out_of_range_is, above_uloq, predose_repeats = [], [], []
            carryover_samples = []
            
            cc_qc_df = df[df[type_col].astype(str).str.contains('Standard|QC|Quality', case=False, na=False)].copy()
            cc_qc_df[is_area_col] = pd.to_numeric(cc_qc_df[is_area_col], errors='coerce')
            avg_is_area = cc_qc_df[is_area_col].mean()
            lower_is_limit = avg_is_area * 0.4 if pd.notna(avg_is_area) else 0
            upper_is_limit = avg_is_area * 1.6 if pd.notna(avg_is_area) else float('inf')
            
            ccqc_rows = []
            sample_rows = []
            ccqc_chart = []
            sample_chart = []
            current_inj = 1
            
            ccqc_rows.append({"Sample Name": "CC and QC", "Area": "", "IS area": "", "Calc.Conc.": "", "Raw.Conc.": "", "% Accuracy": "", "failed_ccqc": False})
            ccqc_chart.append((np.nan, np.nan, np.nan, np.nan, np.nan))
            ccqc_rows.append({"Sample Name": "Sample Name", "Area": "Area", "IS area": "IS area", "Calc.Conc.": "Calc.Conc.", "Raw.Conc.": "", "% Accuracy": "% Acc", "failed_ccqc": False})
            ccqc_chart.append((np.nan, np.nan, np.nan, np.nan, np.nan))
            
            blk_a_area, blk_is_area = None, None
            zero_a_area, zero_is_area = None, None
            mph_a_area, mph_is_area = None, None
            
            for index, row in df.iterrows():
                s_name_raw = str(row[name_col]).strip()
                s_name_up = s_name_raw.upper()
                s_type_up = str(row[type_col]).strip().upper()
                
                raw_conc_val = str(row[conc_col]).strip() if conc_col else ""
                conc = pd.to_numeric(raw_conc_val, errors='coerce') if raw_conc_val != "" else np.nan
                
                raw_is_area_val = str(row[is_area_col]).strip() if is_area_col else ""
                s_is_area = pd.to_numeric(raw_is_area_val, errors='coerce') if raw_is_area_val != "" else np.nan
                
                raw_a_area_val = str(row[analyte_area_col]).strip() if analyte_area_col else ""
                a_area = pd.to_numeric(raw_a_area_val, errors='coerce') if raw_a_area_val != "" else np.nan
                
                raw_acc_val = str(row[acc_col]).strip() if acc_col else ""
                acc_val = pd.to_numeric(raw_acc_val, errors='coerce') if raw_acc_val != "" else np.nan
                
                if 'BLK' in s_name_up or ('BLANK' in s_type_up and '+IS' not in s_name_up):
                    if blk_a_area is None: blk_a_area = a_area
                    if blk_is_area is None: blk_is_area = s_is_area
                if ('ZERO' in s_name_up or '+IS' in s_name_up or 'BLANK+IS' in s_name_up) and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up:
                    if zero_a_area is None: zero_a_area = a_area
                    if zero_is_area is None: zero_is_area = s_is_area
                if 'MPH' in s_name_up or 'ESO' in s_name_up or 'RCS' in s_name_up:
                    if mph_a_area is None: mph_a_area = a_area
                    if mph_is_area is None: mph_is_area = s_is_area
                
                failed_ccqc = False
                if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up:
                    if pd.notna(acc_val):
                        if ('STANDARD' in s_type_up or 'STD' in s_name_up) and ('-A' in s_name_up or '-1' in s_name_up):
                            if not (80 <= acc_val <= 120): failed_ccqc = True
                        else:
                            if not (85 <= acc_val <= 115): failed_ccqc = True
                    else:
                        failed_ccqc = True
                
                if ('STD-A' in s_name_up or '-1' in s_name_up) and std_a_area is None: std_a_area = a_area
                if ('ZERO' in s_name_up or 'AQSTD' in s_name_up or '+IS' in s_name_up) and std_zero_is_area is None: std_zero_is_area = s_is_area
                if 'BLANK' in s_type_up or 'BLK' in s_name_up or 'MOBILE' in s_name_up: carryover_samples.append({'name': s_name_raw, 'a_area': a_area, 'is_area': s_is_area})
                
                if ('STANDARD' in s_type_up or 'STD' in s_name_up) and 'BLK' not in s_name_up and 'ZERO' not in s_name_up and 'AQSTD' not in s_name_up:
                    cc_total += 1
                    if pd.notna(acc_val):
                        if ('-A' in s_name_up or '-1' in s_name_up) and 80 <= acc_val <= 120: cc_pass += 1
                        elif ('-A' not in s_name_up and '-1' not in s_name_up) and 85 <= acc_val <= 115: cc_pass += 1
                        else: failed_cc_ids.append(s_name_raw)
                    else: failed_cc_ids.append(s_name_raw)
                
                if 'LQC' in s_name_up:
                    lqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'MMQC' in s_name_up:
                    mmqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'MQC' in s_name_up and 'MMQC' not in s_name_up:
                    mqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'HQC' in s_name_up:
                    hqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                
                if pd.notna(s_is_area) and ('STANDARD' in s_type_up or 'QC' in s_type_up or 'UNKNOWN' in s_type_up):
                    if s_is_area < lower_is_limit or s_is_area > upper_is_limit:
                        if 'BLK' not in s_name_up and 'AQSTD' not in s_name_up: out_of_range_is.append({'name': s_name_raw, 'conc': conc})
                if 'UNKNOWN' in s_type_up and pd.notna(conc):
                    if conc > uloq:
                        above_uloq.append({'name': s_name_raw, 'conc': conc})
                    
                    is_predose = 'PRE' in s_name_up or bool(re.search(r'[-_]0(\.0+)?$', s_name_up))
                    if is_predose:
                        # Extract subject to find Cmax
                        parts = re.split(r'[-_]', s_name_raw)
                        current_subj = parts[0].upper() if len(parts) > 0 else s_name_up
                        
                        # Find Cmax for this subject in the current file efficiently using pandas
                        subj_mask = df[name_col].astype(str).str.upper().str.contains(current_subj, na=False, regex=False)
                        unk_mask = df[type_col].astype(str).str.upper().str.contains('UNKNOWN', na=False)
                        subj_concs = pd.to_numeric(df.loc[subj_mask & unk_mask, conc_col], errors='coerce')
                        cmax = subj_concs.max() if not subj_concs.empty else 0
                        
                        # Evaluate multi-tier criteria
                        pd_flag = False
                        if pd.notna(lloq) and conc > lloq:
                            pd_flag = True
                        elif std_a_area is not None and pd.notna(a_area) and a_area > std_a_area:
                            pd_flag = True
                        elif pd.notna(cmax) and cmax > 0 and conc > (0.05 * cmax):
                            pd_flag = True
                            
                        if pd_flag:
                            predose_repeats.append({'name': s_name_raw, 'conc': conc})

                chart_tuple = [current_inj, np.nan, np.nan, np.nan, np.nan]
                pct_diff = np.nan
                
                if pd.notna(s_is_area) and pd.notna(avg_is_area) and avg_is_area > 0:
                    pct_diff = (s_is_area / avg_is_area) * 100
                    if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up: chart_tuple[1] = pct_diff
                    elif 'BLANK' in s_type_up or '-BLK' in s_name_up or '+IS' in s_name_up: chart_tuple[2] = pct_diff
                    else: chart_tuple[3] = pct_diff
                
                if 'UNKNOWN' in s_type_up and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up and '+IS' not in s_name_up and 'BLANK' not in s_name_up:
                    chart_tuple[4] = conc

                row_dict = {
                    "Sample Name": s_name_raw, 
                    "Area": a_area, 
                    "Raw.Area": raw_a_area_val,
                    "IS area": s_is_area, 
                    "Raw.IS": raw_is_area_val,
                    "Calc.Conc.": conc, 
                    "Raw.Conc.": raw_conc_val,
                    "failed_ccqc": failed_ccqc
                }
                
                if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up:
                    row_dict["% Accuracy"] = acc_val
                    ccqc_rows.append(row_dict)
                    ccqc_chart.append(tuple(chart_tuple))
                    if s_name_up in seen_ccqc: file_dupes.append(s_name_raw)
                    else: seen_ccqc.add(s_name_up)
                else:
                    row_dict["% Accuracy"] = pct_diff 
                    sample_rows.append(row_dict)
                    sample_chart.append(tuple(chart_tuple))
                    if 'UNKNOWN' in s_type_up and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up and '+IS' not in s_name_up and 'BLANK' not in s_name_up:
                        if s_name_up in seen_samples: 
                            file_sample_dupes.append(s_name_raw)
                        else:
                            seen_samples.add(s_name_up)
                            if file_category == "Main":
                                global_samples[s_name_raw] = (conc, safe_sheet_name)
                            elif file_category == "Repeat":
                                repeat_samples[s_name_raw] = (conc, safe_sheet_name)
                            elif file_category == "ISR":
                                isr_samples[s_name_raw] = (conc, safe_sheet_name)
                            
                current_inj += 1
                
                for std in std_names:
                    if std in s_name_up or std.replace("-", "") in s_name_up:
                        file_cc_dict[std] = conc if pd.notna(conc) else ""
                        break
                
                if 'LQC' in s_name_up: lqc_list.append(conc if pd.notna(conc) else ""); lqc_ids.append(s_name_raw)
                elif 'MMQC' in s_name_up: mmqc_list.append(conc if pd.notna(conc) else ""); mmqc_ids.append(s_name_raw)
                elif 'MQC' in s_name_up and 'MMQC' not in s_name_up: mqc_list.append(conc if pd.notna(conc) else ""); mqc_ids.append(s_name_raw)
                elif 'HQC' in s_name_up: hqc_list.append(conc if pd.notna(conc) else ""); hqc_ids.append(s_name_raw)

            file_cc_dict['Slope'] = "" if pd.isna(slope) else slope
            file_cc_dict['Intercept'] = "" if pd.isna(intercept) else intercept
            file_cc_dict['r'] = "" if pd.isna(r_val) else r_val
            cc_file_data.append(file_cc_dict)
            
            if file_category == "Main":
                for smp in out_of_range_is:
                    all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Out of Range %IS"})
                for smp in above_uloq:
                    all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Above ULOQ"})
                for smp in predose_repeats:
                    all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Predose Detectable"})
            
            for i in range(4):
                q_dict = {
                    'Date of first injection': acq_date if i==0 else "", 'SUBJECT ID': subj_id if i==0 else "",
                    'LQC, HQC ID': lqc_ids[i] if i < len(lqc_ids) else "", 'MQC, MMQC ID': mmqc_ids[i] if i < len(mmqc_ids) else (mqc_ids[i] if i < len(mqc_ids) else ""),
                    'LQC': lqc_list[i] if i < len(lqc_list) else "", 'MMQC': mmqc_list[i] if i < len(mmqc_list) else "",
                    'MQC': mqc_list[i] if i < len(mqc_list) else "", 'HQC': hqc_list[i] if i < len(hqc_list) else ""
                }
                qc_file_data.append(q_dict)
                
            avg_row = {"Sample Name": "Average IS area of CC and QC", "Area": "", "IS area": avg_is_area, "Calc.Conc.": "", "Raw.Conc.": "", "% Accuracy": "", "failed_ccqc": False}
            header_row = {"Sample Name": "Sample Name", "Area": "Area", "IS area": "IS area", "Calc.Conc.": "Calc.Conc.", "Raw.Conc.": "", "% Accuracy": "% IS", "failed_ccqc": False}
            
            final_rows = ccqc_rows + [avg_row, header_row] + sample_rows
            final_chart = ccqc_chart + [(np.nan, np.nan, np.nan, np.nan, np.nan), (np.nan, np.nan, np.nan, np.nan, np.nan)] + sample_chart
            
            batch_meta_data[safe_sheet_name] = {
                'ccid_str': ccid_str,
                'rows': final_rows, 'chart_data': final_chart,
                'slope': slope, 'intercept': intercept, 'r_val': r_val,
                'avg_is_area': avg_is_area,
                'cc_total': cc_total, 'cc_pass': cc_pass, 'failed_cc_ids': failed_cc_ids,
                'lqc_objs': lqc_list_objs, 'mmqc_objs': mmqc_list_objs, 
                'mqc_objs': mqc_list_objs, 'hqc_objs': hqc_list_objs,
                'std_a_area': std_a_area, 'std_zero_is_area': std_zero_is_area,
                'blk_a_area': blk_a_area, 'blk_is_area': blk_is_area,
                'zero_a_area': zero_a_area, 'zero_is_area': zero_is_area,
                'mph_a_area': mph_a_area, 'mph_is_area': mph_is_area,
                'carryover': carryover_samples, 'out_of_range': out_of_range_is, 
                'above_uloq': above_uloq, 'predose_repeats': predose_repeats
            }
            
            nom_check_msg = "All CC and QC concentration matching nominal value." if len(nominals) > 0 else "No nominals defined."
            dup_ccqc_msg = "All CC and QC IDs are unique." if not file_dupes else f"Duplicates found"
            dup_samp_msg = "All CC and QC sample names are unique." if not file_sample_dupes else f"Duplicates found"

            summary_log.append({
                "File Name": file.name, "Status": "SUCCESS", "Error Details": "Dashboard generated.",
                "Nominal Conc. Check": nom_check_msg, "Duplicate CC & QC ID Check": dup_ccqc_msg, "Duplicate Sample Name Check": dup_samp_msg
            })

        excel_tabs["Summary"] = pd.DataFrame(summary_log)
        
        if all_repeats_data:
            rep_df = pd.DataFrame(all_repeats_data)
            rep_df = rep_df.sort_values(by='Sample Name', key=lambda col: col.map(natural_sort_key)).reset_index(drop=True)
            excel_tabs["Repeats"] = rep_df
        else:
            excel_tabs["Repeats"] = pd.DataFrame([{"Batch Name": "-", "Sample Name": "-", "Concentration": "-", "Reason Code": "No Repeats Required"}])
            
        base_to_initial = {}
        for s_name_raw, (conc, orig_file) in global_samples.items():
            base_to_initial[get_base_name(s_name_raw).upper()] = (conc, s_name_raw, orig_file)
            
        base_to_repeat = {}
        for s_name_raw, (conc, rep_file) in repeat_samples.items():
            base_to_repeat[get_base_name(s_name_raw).upper()] = (conc, s_name_raw, rep_file)
            
        base_to_isr = {}
        for s_name_raw, (conc, isr_file) in isr_samples.items():
            base_to_isr[get_base_name(s_name_raw).upper()] = (conc, s_name_raw, isr_file)

        repeat_analysis_rows = []
        s_no = 1
        sorted_repeats = sorted(base_to_repeat.items(), key=lambda x: natural_sort_key(x[1][1]))
        for base_up, (rep_conc, rep_s_name_raw, rep_file) in sorted_repeats:
            init_conc, init_s_name_raw, orig_file = base_to_initial.get(base_up, (np.nan, "Missing", "Unknown"))
            pct_diff = ""
            if pd.notna(init_conc) and pd.notna(rep_conc) and init_conc > 0 and rep_conc > 0:
                pct_diff = (rep_conc - init_conc) / ((init_conc + rep_conc) / 2) * 100
                
            repeat_analysis_rows.append({
                "S.No": s_no, "Initial File Name": orig_file, "Repeat File Name": rep_file,
                "Sample Name": rep_s_name_raw, "Initial Concentration": init_conc,
                "Repeat Concentration": rep_conc, "% Difference": pct_diff
            })
            s_no += 1
            
        if repeat_analysis_rows: excel_tabs["Repeat Analysis"] = pd.DataFrame(repeat_analysis_rows)
        else: excel_tabs["Repeat Analysis"] = pd.DataFrame([{"S.No": "-", "Initial File Name": "-", "Repeat File Name": "-", "Sample Name": "-", "Initial Concentration": "-", "Repeat Concentration": "-", "% Difference": "-"}])
            
        isr_analysis_rows = []
        s_no_isr = 1
        sorted_isr = sorted(base_to_isr.items(), key=lambda x: natural_sort_key(x[1][1]))
        for base_up, (isr_conc, isr_s_name_raw, isr_file) in sorted_isr:
            init_conc, init_s_name_raw, orig_file = base_to_initial.get(base_up, (np.nan, "Missing", "Unknown"))
            rep_tuple = base_to_repeat.get(base_up)
            final_prior_conc = rep_tuple[0] if rep_tuple else init_conc
            
            pct_diff = ""
            if pd.notna(final_prior_conc) and pd.notna(isr_conc) and final_prior_conc > 0 and isr_conc > 0:
                pct_diff = (isr_conc - final_prior_conc) / ((final_prior_conc + isr_conc) / 2) * 100
                
            isr_analysis_rows.append({
                "S.No": s_no_isr, "Sample Name": isr_s_name_raw, "Initial Concentration": final_prior_conc,
                "ISR Concentration": isr_conc, "% Difference": pct_diff
            })
            s_no_isr += 1
            
        if isr_analysis_rows: excel_tabs["ISR Analysis"] = pd.DataFrame(isr_analysis_rows)
        else: excel_tabs["ISR Analysis"] = pd.DataFrame([{"S.No": "-", "Sample Name": "-", "Initial Concentration": "-", "ISR Concentration": "-", "% Difference": "-"}])
        
        if cc_file_data:
            cc_df_data = pd.DataFrame(cc_file_data)
            cc_df_data['Sort_Num'] = cc_df_data['CCID'].apply(extract_sort_num)
            cc_df_data = cc_df_data.sort_values(by='Sort_Num').drop('Sort_Num', axis=1)
            cc_columns = ['Date of first injection', 'SUBJECT ID', 'CCID'] + std_names + ['Slope', 'Intercept', 'r']
            row_nom_cc = {'CCID': 'Nominal'}; row_low_cc = {'CCID': 'Lower limit'}; row_up_cc = {'CCID': 'Upper limit'}
            for std in std_names:
                val = nominals.get(std, "")
                row_nom_cc[std] = val
                if val != "":
                    limit = 0.20 if ('A' in std or '1' in std) else 0.15
                    row_low_cc[std] = val * (1 - limit)
                    row_up_cc[std] = val * (1 + limit)
            cc_df = pd.concat([pd.DataFrame([row_nom_cc, row_low_cc, row_up_cc]), cc_df_data], ignore_index=True)
            excel_tabs["CC"] = cc_df[cc_columns]

        if qc_file_data:
            qc_df_data = pd.DataFrame(qc_file_data)
            qc_df_data['Original_Index'] = qc_df_data.index
            qc_df_data['Block'] = qc_df_data.index // 4
            block_sort_keys = qc_df_data.groupby('Block')['LQC, HQC ID'].first().apply(extract_sort_num)
            qc_df_data['Sort_Num'] = qc_df_data['Block'].map(block_sort_keys)
            qc_df_data = qc_df_data.sort_values(by=['Sort_Num', 'Block', 'Original_Index']).drop(['Sort_Num', 'Block', 'Original_Index'], axis=1)
            qc_columns = ['Date of first injection', 'SUBJECT ID', 'LQC, HQC ID', 'MQC, MMQC ID', 'LQC', 'MMQC', 'MQC', 'HQC']
            row_nom_qc = {'MQC, MMQC ID': 'Nominal'}; row_low_qc = {'MQC, MMQC ID': 'Lower limit'}; row_up_qc = {'MQC, MMQC ID': 'Upper limit'}
            for qc in ['LQC', 'MMQC', 'MQC', 'HQC']:
                val = nominals.get(qc, "")
                row_nom_qc[qc] = val
                if val != "":
                    limit = 0.20 if qc == 'LQC' else 0.15
                    row_low_qc[qc] = val * (1 - limit)
                    row_up_qc[qc] = val * (1 + limit)
            qc_df = pd.concat([pd.DataFrame([row_nom_qc, row_low_qc, row_up_qc]), qc_df_data], ignore_index=True)
            excel_tabs["QC"] = qc_df[qc_columns]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            for sh_name in ["Summary", "Repeats", "Repeat Analysis", "ISR Analysis", "CC", "QC"]:
                if sh_name in excel_tabs:
                    excel_tabs[sh_name].to_excel(writer, sheet_name=sh_name, index=False)
            
            workbook = writer.book
            tnr = {'font_name': 'Times New Roman', 'font_size': 10}
            green_sheet_bg = workbook.add_format({**tnr, 'bg_color': '#CDE4CE'})
            header_format = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            yellow_flag = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1})
            bold_label = workbook.add_format({**tnr, 'bold': True, 'align': 'right', 'border': 1})
            general_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter'})
            conc_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
            r_fmt = workbook.add_format({**tnr, 'num_format': '0.0000', 'align': 'center', 'valign': 'vcenter'})
            float_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            stdev_fmt = workbook.add_format({**tnr, 'num_format': '0.0000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            pct_fmt = workbook.add_format({**tnr, 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            merge_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            success_fmt = workbook.add_format({**tnr, 'bg_color': '#99FF99', 'align': 'center', 'valign': 'vcenter'})
            failed_fmt = workbook.add_format({**tnr, 'bg_color': '#FF9999', 'align': 'center', 'valign': 'vcenter'})
            data_white = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white_left = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
            data_white_float = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            data_white_int = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0'})
            data_white_pct = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            data_white_r = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000'})
            data_yellow_float = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            table_header_green = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            table_header_gray = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            failed_ccqc_string_fmt = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            failed_ccqc_pct_fmt = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            failed_ccqc_float_fmt = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})

            qa_ws = workbook.add_worksheet("QA Summary")
            qa_ws.set_column('A:B', 35, general_fmt)
            qa_ws.set_column('C:D', 35, general_fmt)
            qa_row = 0
            
            for safe_sheet_name, meta in batch_meta_data.items():
                qa_ws.merge_range(qa_row, 0, qa_row, 6, f"QA Summary - Batch: {safe_sheet_name}", table_header_green)
                qa_row += 1
                
                qa_ws.write(qa_row, 0, "LLOQ area", data_white_left)
                std_a = meta.get('std_a_area')
                avg_is = meta.get('avg_is_area')

                qa_ws.write(qa_row, 1, std_a if pd.notna(std_a) else "N/A", data_white_int if pd.notna(std_a) else data_white)
                qa_ws.merge_range(qa_row, 2, qa_row, 4, "Average IS area of accepted CC, QC samples", data_white_left)
                qa_ws.merge_range(qa_row, 5, qa_row, 6, avg_is if pd.notna(avg_is) else "N/A", data_white_int if pd.notna(avg_is) else data_white)
                qa_row += 1
                
                qa_ws.merge_range(qa_row, 0, qa_row, 6, "Acceptability of interference in Standard Blank, Standard Blank + IS and MPH/ESO/RCS", table_header_green)
                qa_row += 1
                
                def calc_interf(area, ref_area):
                    if pd.notna(area) and pd.notna(ref_area) and ref_area > 0:
                        return (area / ref_area) * 100
                    return "N/A"

                blk_a = meta.get('blk_a_area'); blk_a_pct = calc_interf(blk_a, std_a)
                blk_is = meta.get('blk_is_area'); blk_is_pct = calc_interf(blk_is, avg_is)
                zero_a = meta.get('zero_a_area'); zero_a_pct = calc_interf(zero_a, std_a)
                mph_a = meta.get('mph_a_area'); mph_a_pct = calc_interf(mph_a, std_a)
                mph_is = meta.get('mph_is_area'); mph_is_pct = calc_interf(mph_is, avg_is)
                
                interf_accepted = "Accepted"
                for pct in [blk_a_pct, zero_a_pct, mph_a_pct]:
                    if isinstance(pct, (int, float)) and pct > 20.0: interf_accepted = "Not Accepted"
                for pct in [blk_is_pct, mph_is_pct]:
                    if isinstance(pct, (int, float)) and pct > 5.0: interf_accepted = "Not Accepted"

                def write_interf_row(label, area, pct):
                    qa_ws.merge_range(qa_row, 0, qa_row, 1, label, data_white_left)
                    qa_ws.merge_range(qa_row, 2, qa_row, 3, area if pd.notna(area) else "N/A", data_white_int if pd.notna(area) else data_white)
                    qa_ws.merge_range(qa_row, 4, qa_row, 5, "% Interference", data_white)
                    qa_ws.write(qa_row, 6, pct, data_white_pct if isinstance(pct, (int, float)) else data_white)

                write_interf_row("Area Observed for Standard Blank at Analyte RT:", blk_a, blk_a_pct); qa_row += 1
                write_interf_row("Area Observed for Standard Blank at IS RT:", blk_is, blk_is_pct); qa_row += 1
                write_interf_row("Area Observed for Standard Blank+IS at Analyte RT:", zero_a, zero_a_pct); qa_row += 1
                write_interf_row("Area Observed for MPH/ESO/RCS at Analyte RT:", mph_a, mph_a_pct); qa_row += 1
                write_interf_row("Area Observed for MPH/ESO/RCS at IS RT:", mph_is, mph_is_pct); qa_row += 1

                qa_ws.merge_range(qa_row, 0, qa_row, 3, "% Interference for Standard Blank, Standard Blank+IS, MPH/ESO/RCS:", data_white_left)
                qa_ws.merge_range(qa_row, 4, qa_row, 6, interf_accepted, data_white)
                qa_row += 1
                
                qa_ws.merge_range(qa_row, 0, qa_row, 6, "Acceptability of Non-Zero calibration curve standards", table_header_green)
                qa_row += 1
                qa_ws.write(qa_row, 0, "CC ID:", data_white_left)
                qa_ws.merge_range(qa_row, 1, qa_row, 3, meta.get('ccid_str', ''), data_white)
                qa_ws.merge_range(qa_row, 4, qa_row, 5, "'r' value", data_white_left)
                qa_ws.write(qa_row, 6, meta.get('r_val', ''), data_white_r)
                qa_row += 1
                
                cc_total = meta.get('cc_total', 0); cc_pass = meta.get('cc_pass', 0)
                cc_status = "Accepted" if cc_total >= 6 and (cc_pass / cc_total) >= 0.75 else "Not Accepted"
                
                qa_ws.write(qa_row, 0, "No. of CC standards injected", data_white_left)
                qa_ws.merge_range(qa_row, 1, qa_row, 3, cc_total, data_white)
                qa_ws.merge_range(qa_row, 4, qa_row, 5, "No. of CC standard accepted", data_white_left)
                qa_ws.write(qa_row, 6, cc_pass, data_white)
                qa_row += 1
                
                qa_ws.write(qa_row, 0, "Details of Failed CC standard(s), if any", data_white_left)
                qa_ws.merge_range(qa_row, 1, qa_row, 6, ", ".join(meta.get('failed_cc_ids', [])), data_white)
                qa_row += 1
                qa_ws.write(qa_row, 0, "Calibration curve status:", data_white_left)
                qa_ws.merge_range(qa_row, 1, qa_row, 6, cc_status, data_white)
                qa_row += 1
                
                qa_ws.merge_range(qa_row, 0, qa_row, 6, "Acceptability of QC Samples", table_header_green)
                qa_row += 1
                qa_ws.merge_range(qa_row, 0, qa_row + 1, 0, "QC ID", table_header_green)
                qa_ws.merge_range(qa_row, 1, qa_row, 2, "Total No. of QC Injected", table_header_green)
                qa_ws.merge_range(qa_row, 3, qa_row, 4, "Total No. of QC Passed", table_header_green)
                qa_ws.merge_range(qa_row, 5, qa_row, 6, "Acceptance status Yes/ No", table_header_green)
                qa_row += 1
                
                qa_ws.write(qa_row, 1, "Set-1", table_header_green); qa_ws.write(qa_row, 2, "Set-2", table_header_green)
                qa_ws.write(qa_row, 3, "Set-1", table_header_green); qa_ws.write(qa_row, 4, "Set-2", table_header_green)
                qa_ws.write(qa_row, 5, "Set-1", table_header_green); qa_ws.write(qa_row, 6, "Set-2", table_header_green)
                qa_row += 1
                
                def process_qc_sets(objs):
                    mid = len(objs) // 2
                    set1 = objs[:mid] if mid > 0 else objs
                    set2 = objs[mid:] if mid > 0 else []
                    s1_t, s1_p = len(set1), sum(1 for o in set1 if o['pass'])
                    s2_t, s2_p = len(set2), sum(1 for o in set2 if o['pass'])
                    
                    s1_ok = "Yes" if s1_t > 0 and (s1_p / s1_t) >= 0.5 else ("-" if s1_t == 0 else "No")
                    s2_ok = "Yes" if s2_t > 0 and (s2_p / s2_t) >= 0.5 else ("-" if s2_t == 0 else "No")
                    
                    f_id = objs[0]['name'].split('-')[-1] if objs else ""
                    l_id = objs[-1]['name'].split('-')[-1] if objs else ""
                    return f"({f_id}-{l_id})" if f_id and l_id else "", s1_t, s2_t, s1_p, s2_p, s1_ok, s2_ok
                
                tot_s1_inj = tot_s2_inj = tot_s1_pass = tot_s2_pass = 0
                
                for qc_lvl, objs in [("LQC", meta.get('lqc_objs',[])), ("MMQC", meta.get('mmqc_objs',[])), ("MQC", meta.get('mqc_objs',[])), ("HQC", meta.get('hqc_objs',[]))]:
                    id_str, s1_t, s2_t, s1_p, s2_p, s1_ok, s2_ok = process_qc_sets(objs)
                    tot_s1_inj += s1_t; tot_s2_inj += s2_t; tot_s1_pass += s1_p; tot_s2_pass += s2_p
                    
                    qa_ws.write(qa_row, 0, f"{qc_lvl} {id_str}", data_white_left)
                    qa_ws.write(qa_row, 1, s1_t, data_white); qa_ws.write(qa_row, 2, s2_t, data_white)
                    qa_ws.write(qa_row, 3, s1_p, data_white); qa_ws.write(qa_row, 4, s2_p, data_white)
                    qa_ws.write(qa_row, 5, s1_ok, data_white); qa_ws.write(qa_row, 6, s2_ok, data_white)
                    qa_row += 1
                    
                qa_ws.write(qa_row, 0, "Total no. of QCs", data_white_left)
                qa_ws.write(qa_row, 1, tot_s1_inj, data_white); qa_ws.write(qa_row, 2, tot_s2_inj, data_white)
                qa_ws.write(qa_row, 3, tot_s1_pass, data_white); qa_ws.write(qa_row, 4, tot_s2_pass, data_white)
                
                g_s1_ok = "Yes" if tot_s1_inj > 0 and (tot_s1_pass / tot_s1_inj) >= 0.67 else ("-" if tot_s1_inj==0 else "No")
                g_s2_ok = "Yes" if tot_s2_inj > 0 and (tot_s2_pass / tot_s2_inj) >= 0.67 else ("-" if tot_s2_inj==0 else "No")
                qa_ws.write(qa_row, 5, g_s1_ok, data_white); qa_ws.write(qa_row, 6, g_s2_ok, data_white)
                qa_row += 1
                
                qa_ws.merge_range(qa_row, 0, qa_row, 6, "Repeat Sample Identification", table_header_green)
                qa_row += 1
                qa_ws.merge_range(qa_row, 0, qa_row, 1, "S.No", table_header_green)
                qa_ws.merge_range(qa_row, 2, qa_row, 4, "Sample Name", table_header_green)
                qa_ws.merge_range(qa_row, 5, qa_row, 6, "Repeat code", table_header_green)
                qa_row += 1
                
                rep_count = 1
                for smp in meta.get('out_of_range', []):
                    qa_ws.merge_range(qa_row, 0, qa_row, 1, rep_count, data_white)
                    qa_ws.merge_range(qa_row, 2, qa_row, 4, smp['name'], data_white)
                    qa_ws.merge_range(qa_row, 5, qa_row, 6, "Out of Range %IS", data_white)
                    rep_count += 1; qa_row += 1
                for smp in meta.get('above_uloq', []):
                    qa_ws.merge_range(qa_row, 0, qa_row, 1, rep_count, data_white)
                    qa_ws.merge_range(qa_row, 2, qa_row, 4, smp['name'], data_white)
                    qa_ws.merge_range(qa_row, 5, qa_row, 6, "Above ULOQ", data_white)
                    rep_count += 1; qa_row += 1
                    
                qa_row += 3 

            for sheet_name in ["Summary", "Repeats", "Repeat Analysis", "ISR Analysis", "CC", "QC"]:
                if sheet_name in excel_tabs:
                    df_tab = excel_tabs[sheet_name]
                    worksheet = writer.sheets[sheet_name]
                    for col_num, value in enumerate(df_tab.columns): worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column('A:A', 35) 
                    
                    if sheet_name == "Summary":
                        worksheet.set_column('B:Z', 25, general_fmt)
                        for row_num, row_data in enumerate(df_tab.values):
                            row_fmt = success_fmt if row_data[1] == "SUCCESS" else failed_fmt
                            for col_num, value in enumerate(row_data): worksheet.write(row_num + 1, col_num, value, row_fmt)
                    elif sheet_name == "Repeats":
                        worksheet.set_column('B:B', 30, general_fmt)
                        worksheet.set_column('C:C', 15, conc_fmt)
                        worksheet.set_column('D:D', 25, general_fmt)
                    elif sheet_name == "Repeat Analysis":
                        worksheet.set_column('A:A', 10, general_fmt)
                        worksheet.set_column('B:D', 30, general_fmt)
                        worksheet.set_column('E:F', 25, conc_fmt)
                        worksheet.set_column('G:G', 20, pct_fmt)
                    elif sheet_name == "ISR Analysis":
                        worksheet.set_column('A:A', 10, general_fmt)
                        worksheet.set_column('B:B', 30, general_fmt)
                        worksheet.set_column('C:D', 25, conc_fmt)
                        worksheet.set_column('E:E', 20, pct_fmt)
                    elif sheet_name == "CC":
                        worksheet.set_column(1, 2, 15, general_fmt)
                        start_c = 3
                        end_c = 3 + len(std_names) - 1
                        if end_c >= start_c: worksheet.set_column(start_c, end_c, 15, conc_fmt)
                        worksheet.set_column(end_c + 1, end_c + 2, 15, general_fmt)
                        worksheet.set_column(end_c + 3, end_c + 3, 15, r_fmt)
                    elif sheet_name == "QC":
                        worksheet.set_column(1, 3, 15, general_fmt)
                        worksheet.set_column(4, 7, 15, conc_fmt)

            for safe_sheet_name, meta in batch_meta_data.items():
                ws = workbook.add_worksheet(safe_sheet_name)
                final_rows = meta['rows']
                chart_data = meta['chart_data']
                max_row = len(final_rows)
                
                ws.set_column('A:AD', 15, green_sheet_bg)
                ws.set_column('A:A', 25, green_sheet_bg)
                ws.set_column('G:M', 28, green_sheet_bg)
                
                for r_idx, row_dict in enumerate(final_rows):
                    if row_dict["Sample Name"] == "CC and QC":
                        ws.merge_range(r_idx, 0, r_idx, 4, row_dict["Sample Name"], table_header_green)
                    elif row_dict["Sample Name"] == "Average IS area of CC and QC":
                        ws.merge_range(r_idx, 0, r_idx, 1, row_dict["Sample Name"], table_header_gray)
                        ws.write(r_idx, 2, row_dict["IS area"], data_white)
                        ws.write(r_idx, 3, "", data_white)
                        ws.write(r_idx, 4, "", data_white)
                    elif row_dict["Sample Name"] == "Sample Name":
                        ws.write(r_idx, 0, "Sample Name", table_header_green)
                        ws.write(r_idx, 1, "Area", table_header_green)
                        ws.write(r_idx, 2, "IS area", table_header_green)
                        ws.write(r_idx, 3, "Calc.Conc.", table_header_green)
                        ws.write(r_idx, 4, row_dict.get("% Accuracy", "% Acc"), table_header_green)
                    else:
                        ws.write(r_idx, 0, row_dict["Sample Name"], data_white)
                        
                        a_display = ""
                        if pd.notna(row_dict["Area"]):
                            a_display = int(row_dict["Area"]) if isinstance(row_dict["Area"], float) and row_dict["Area"].is_integer() else row_dict["Area"]
                        elif row_dict["Raw.Area"] != "" and row_dict["Raw.Area"].upper() != 'NAN':
                            a_display = row_dict["Raw.Area"]
                        ws.write(r_idx, 1, a_display, data_white)
                        
                        is_display = ""
                        if pd.notna(row_dict["IS area"]):
                            is_display = int(row_dict["IS area"]) if isinstance(row_dict["IS area"], float) and row_dict["IS area"].is_integer() else row_dict["IS area"]
                        elif row_dict["Raw.IS"] != "" and row_dict["Raw.IS"].upper() != 'NAN':
                            is_display = row_dict["Raw.IS"]
                            
                        if row_dict["failed_ccqc"] and is_display != "":
                            ws.write(r_idx, 2, f"{is_display}*", failed_ccqc_string_fmt)
                        else:
                            ws.write(r_idx, 2, is_display, data_white)
                            
                        c_display = ""
                        if pd.notna(row_dict["Calc.Conc."]):
                            c_display = row_dict["Calc.Conc."]
                        elif row_dict["Raw.Conc."] != "" and row_dict["Raw.Conc."].upper() != 'NAN':
                            c_display = row_dict["Raw.Conc."]
                            
                        if pd.notna(row_dict["Calc.Conc."]) and isinstance(c_display, (int, float)):
                            if c_display > uloq:
                                ws.write(r_idx, 3, c_display, data_yellow_float)
                            else:
                                ws.write(r_idx, 3, c_display, data_white_float)
                        else:
                            ws.write(r_idx, 3, c_display, data_white)
                            
                        acc = row_dict.get("% Accuracy", "")
                        if row_dict["failed_ccqc"]:
                            if pd.isna(acc) or acc == "": ws.write(r_idx, 4, "", failed_ccqc_string_fmt)
                            else: ws.write(r_idx, 4, acc, failed_ccqc_pct_fmt)
                        else:
                            if pd.isna(acc) or acc == "": ws.write(r_idx, 4, "", data_white)
                            else: ws.write(r_idx, 4, acc, data_white_pct)

                    c_tuple = chart_data[r_idx]
                    ws.write(r_idx, 26, "" if pd.isna(c_tuple[0]) else c_tuple[0])
                    ws.write(r_idx, 27, "" if pd.isna(c_tuple[1]) else c_tuple[1])
                    ws.write(r_idx, 28, "" if pd.isna(c_tuple[2]) else c_tuple[2])
                    ws.write(r_idx, 29, "" if pd.isna(c_tuple[3]) else c_tuple[3])
                    ws.write(r_idx, 30, "" if pd.isna(c_tuple[4]) else c_tuple[4])

                ws.set_column(26, 30, None, None, {'hidden': True})
                
                ws.write('G2', 'Intercept', data_white)
                ws.write('H2', meta.get('intercept', ''), data_white)
                ws.write('G3', 'Slope', data_white)
                ws.write('H3', meta.get('slope', ''), data_white)
                ws.write('G4', 'r', data_white)
                ws.write('H4', meta.get('r_val', ''), data_white_r)
                
                ws.write('G6', 'No of CC passed', table_header_gray)
                ws.write('H6', meta.get('cc_pass', ''), data_white)
                ws.write('G7', 'LQC', table_header_gray)
                ws.write('H7', meta.get('lqc_pass', ''), data_white)
                ws.write('G8', 'MMQC', table_header_gray)
                ws.write('H8', meta.get('mmqc_pass', ''), data_white)
                ws.write('G9', 'MQC', table_header_gray)
                ws.write('H9', meta.get('mqc_pass', ''), data_white)
                ws.write('G10', 'HQC', table_header_gray)
                ws.write('H10', meta.get('hqc_pass', ''), data_white)
                
                co_headers = ['Sample name', 'Analyte Area', 'IS area', '% analyte', '% IS']
                for c_idx, h in enumerate(co_headers): ws.write(12, 6 + c_idx, h, table_header_green)
                    
                co_samples = meta.get('carryover', [])
                for r_idx, smp in enumerate(co_samples[:3]):
                    r = 13 + r_idx
                    a_area = smp['a_area']
                    is_area = smp['is_area']
                    pct_a = (a_area / meta['std_a_area']) * 100 if meta.get('std_a_area') and pd.notna(a_area) and meta['std_a_area'] > 0 else np.nan
                    pct_is = (is_area / meta['std_zero_is_area']) * 100 if meta.get('std_zero_is_area') and pd.notna(is_area) and meta['std_zero_is_area'] > 0 else np.nan
                    
                    ws.write(r, 6, smp['name'], data_white)
                    ws.write(r, 7, "" if pd.isna(a_area) else a_area, data_white)
                    ws.write(r, 8, "" if pd.isna(is_area) else is_area, data_white)
                    ws.write(r, 9, "NA" if pd.isna(pct_a) else pct_a, data_white_pct if pd.notna(pct_a) else data_white)
                    ws.write(r, 10, "NA" if pd.isna(pct_is) else pct_is, data_white_pct if pd.notna(pct_is) else data_white)
                
                if max_row > 0:
                    chart1 = workbook.add_chart({'type': 'scatter'})
                    chart1.add_series({'name': 'Samples', 'categories': [safe_sheet_name, 1, 26, max_row-1, 26], 'values': [safe_sheet_name, 1, 29, max_row-1, 29], 'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#4F81BD'}, 'border': {'none': True}}, 'line': {'none': True}})
                    chart1.add_series({'name': 'Blank', 'categories': [safe_sheet_name, 1, 26, max_row-1, 26], 'values': [safe_sheet_name, 1, 28, max_row-1, 28], 'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#F79646'}, 'border': {'none': True}}, 'line': {'none': True}})
                    chart1.add_series({'name': 'CCQC', 'categories': [safe_sheet_name, 1, 26, max_row-1, 26], 'values': [safe_sheet_name, 1, 27, max_row-1, 27], 'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#C0504D'}, 'border': {'none': True}}, 'line': {'none': True}})
                    chart1.set_title({'name': 'IS Variation Plot', 'name_font': {'size': 12, 'name': 'Times New Roman'}})
                    chart1.set_x_axis({'name': 'Injection Order', 'name_font': {'size': 10, 'name': 'Times New Roman'}, 'num_font': {'name': 'Times New Roman'}, 'major_gridlines': {'visible': False}})
                    chart1.set_y_axis({'name': '% Difference from Mean IS', 'name_font': {'size': 10, 'name': 'Times New Roman'}, 'num_font': {'name': 'Times New Roman'}, 'major_gridlines': {'visible': True}, 'num_format': '0.00'})
                    chart1.set_legend({'position': 'right', 'font': {'name': 'Times New Roman', 'size': 10}})
                    chart1.set_size({'width': 450, 'height': 230})
                    chart1.show_hidden_data()
                    ws.insert_chart('G18', chart1)
                    
                    chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                    chart2.add_series({'name': 'Concentration', 'categories': [safe_sheet_name, 1, 26, max_row-1, 26], 'values': [safe_sheet_name, 1, 30, max_row-1, 30], 'marker': {'type': 'diamond', 'size': 5, 'fill': {'color': '#4F81BD'}, 'border': {'color': '#4F81BD'}}, 'line': {'color': '#4F81BD', 'width': 1.5}})
                    chart2.set_title({'name': 'Chart Title', 'name_font': {'size': 14, 'bold': True, 'name': 'Times New Roman'}})
                    chart2.set_x_axis({'major_gridlines': {'visible': False}, 'num_font': {'name': 'Times New Roman'}})
                    chart2.set_y_axis({'major_gridlines': {'visible': True}, 'num_format': '0.000', 'num_font': {'name': 'Times New Roman'}})
                    chart2.set_legend({'none': True})
                    chart2.set_size({'width': 450, 'height': 230})
                    chart2.show_hidden_data()
                    chart2.show_blanks_as('span')
                    ws.insert_chart('G30', chart2)
                    
                    or_start = 45
                    ws.write(or_start, 6, "Out of Range %IS Samples", table_header_green)
                    ws.write(or_start, 7, "Concentration", table_header_green)
                    ws.write(or_start, 8, "Above ULOQ Samples", table_header_green)
                    ws.write(or_start, 9, "Concentration", table_header_green)
                    ws.write(or_start, 10, "Predose Repeats (PD)", table_header_green)
                    ws.write(or_start, 11, "Concentration", table_header_green)
                    
                    out_is = meta.get('out_of_range', [])
                    for i, smp in enumerate(out_is):
                        ws.write(or_start + 1 + i, 6, smp['name'], data_white)
                        ws.write(or_start + 1 + i, 7, smp['conc'], data_yellow_float if pd.notna(smp['conc']) and smp['conc'] > uloq else data_white_float)
                        
                    above_u = meta.get('above_uloq', [])
                    for i, smp in enumerate(above_u):
                        ws.write(or_start + 1 + i, 8, smp['name'], data_white)
                        ws.write(or_start + 1 + i, 9, smp['conc'], data_yellow_float)

                    predose_u = meta.get('predose_repeats', [])
                    for i, smp in enumerate(predose_u):
                        ws.write(or_start + 1 + i, 10, smp['name'], data_white)
                        ws.write(or_start + 1 + i, 11, smp['conc'], data_yellow_float)

            for target_sheet in ["CC", "QC"]:
                if target_sheet in excel_tabs:
                    ws = writer.sheets[target_sheet]
                    df_target = excel_tabs[target_sheet]
                    num_rows = len(df_target)
                    stat_row = num_rows + 1 
                    
                    start_col = 3 if target_sheet == "CC" else 4
                    end_col = len(std_names) + 3 if target_sheet == "CC" else 8
                    
                    for r in range(1, num_rows + 1):
                        for c in range(len(df_target.columns)):
                            val = df_target.iloc[r-1, c]
                            col_name = df_target.columns[c]
                            
                            if target_sheet == "CC" and col_name in ['Slope', 'Intercept']:
                                fmt = data_white
                            elif target_sheet == "CC" and col_name == 'r':
                                fmt = data_white_r if isinstance(val, (int, float)) else data_white
                            else:
                                fmt = data_white_float if isinstance(val, (int, float)) else data_white
                            
                            out_val = "" if pd.isna(val) else val
                            
                            if r > 3 and start_col <= c < end_col and pd.notna(val) and val != "":
                                try:
                                    num_val = float(val)
                                    lower_limit = float(df_target.iloc[1, c])
                                    upper_limit = float(df_target.iloc[2, c])
                                    if num_val < lower_limit or num_val > upper_limit:
                                        if target_sheet == "CC":
                                            out_val = f"{num_val:.3f}*"
                                            fmt = failed_ccqc_string_fmt
                                        else:
                                            out_val = num_val
                                            fmt = failed_ccqc_float_fmt
                                except (ValueError, TypeError):
                                    pass
                            
                            ws.write(r, c, out_val, fmt)
                    
                    ws.merge_range(0, 0, 3, 0, df_target.columns[0], header_format)
                    ws.merge_range(0, 1, 3, 1, df_target.columns[1], header_format)
                    
                    if target_sheet == "CC":
                        ws.merge_range(0, 2, 3, 2, df_target.columns[2], header_format)
                        end_c = len(df_target.columns)
                        ws.merge_range(0, end_c - 3, 3, end_c - 3, df_target.columns[end_c - 3], header_format)
                        ws.merge_range(0, end_c - 2, 3, end_c - 2, df_target.columns[end_c - 2], header_format)
                        ws.merge_range(0, end_c - 1, 3, end_c - 1, df_target.columns[end_c - 1], header_format)
                    elif target_sheet == "QC":
                        for r in range(4, num_rows + 1, 4):
                            if r + 3 <= num_rows:
                                date_val, subj_val = df_target.iloc[r-1, 0], df_target.iloc[r-1, 1]
                                ws.merge_range(r, 0, r+3, 0, "" if pd.isna(date_val) else str(date_val), merge_fmt)
                                ws.merge_range(r, 1, r+3, 1, "" if pd.isna(subj_val) else str(subj_val), merge_fmt)
                    
                    label_col = 2 if target_sheet == "CC" else 1
                    ws.write(stat_row, label_col, "Mean", bold_label)
                    ws.write(stat_row + 1, label_col, "SD", bold_label)
                    ws.write(stat_row + 2, label_col, "%CV", bold_label)
                    ws.write(stat_row + 3, label_col, "%Accuracy", bold_label)
                    
                    for col_num in range(start_col, end_col):
                        col_letter = xl_col_to_name(col_num)
                        data_range = f"{col_letter}5:{col_letter}{num_rows + 1}"
                        ws.write_formula(stat_row, col_num, f"=IFERROR(AVERAGE({data_range}), \"\")", float_fmt)
                        ws.write_formula(stat_row + 1, col_num, f"=IFERROR(STDEV({data_range}), \"\")", stdev_fmt)
                        ws.write_formula(stat_row + 2, col_num, f"=IFERROR(({col_letter}{stat_row + 2}/{col_letter}{stat_row + 1})*100, \"\")", pct_fmt)
                        ws.write_formula(stat_row + 3, col_num, f"=IFERROR(({col_letter}{stat_row + 1}/{col_letter}2)*100, \"\")", pct_fmt)
                        ws.conditional_format(data_range, {'type': 'formula', 'criteria': f"=AND(ISNUMBER({col_letter}5), OR({col_letter}5<{col_letter}$3, {col_letter}5>{col_letter}$4))", 'format': yellow_flag})

            mapped_samples = {}
            if num_subj > 0 and periods_list and timepoints_list:
                for s_name, (conc, orig_file) in global_samples.items():
                    parts = re.split(r'[-_]', s_name)
                    subj_str = parts[0] if len(parts) > 0 else s_name
                    period_str = parts[1] if len(parts) > 1 else ""
                    tp_str = parts[2] if len(parts) > 2 else ""
                    
                    subj_nums = re.findall(r'\d+', subj_str)
                    subj_int = int(subj_nums[-1]) if subj_nums else 0
                    period_clean = period_str.strip().upper()
                    
                    tp_nums = re.findall(r'\d+\.\d+|\d+', tp_str)
                    tp_float = float(tp_nums[0]) if tp_nums else -1.0
                    if "PRE" in tp_str.upper(): tp_float = -1.0
                    
                    mapped_samples[(subj_int, period_clean, tp_float)] = (conc, s_name, orig_file)
            
            ws_conc = workbook.add_worksheet("Concentration Data")
            ws_conc.set_column('A:B', 15, general_fmt)
            ws_conc.set_column('C:C', 15, general_fmt)
            ws_conc.set_column('D:D', 20, conc_fmt)
            ws_conc.set_column('E:E', 25, general_fmt)
            
            conc_headers = ["Subject No.", "Period No.", "Timepoints", "Conc.(ng/mL)", "Sample Name"]
            for col_num, header in enumerate(conc_headers): ws_conc.write(0, col_num, header, table_header_green)
                
            row_idx = 1
            if num_subj > 0 and periods_list and timepoints_list:
                for subj_idx in range(1, num_subj + 1):
                    subj_formatted = f"S{subj_idx:03d}"
                    for period_val in periods_list:
                        period_clean = period_val.strip().upper()
                        for tp_val in timepoints_list:
                            tp_nums = re.findall(r'\d+\.\d+|\d+', str(tp_val))
                            tp_float = float(tp_nums[0]) if tp_nums else -1.0
                            if "PRE" in str(tp_val).upper(): tp_float = -1.0
                            
                            lookup_key = (subj_idx, period_clean, tp_float)
                            if lookup_key in mapped_samples:
                                conc, real_s_name, orig_file = mapped_samples[lookup_key]
                                display_conc = "BLQ" if (pd.isna(conc) or (pd.notna(lloq) and conc < lloq)) else conc
                                s_name_out = real_s_name
                            else:
                                display_conc = "M"
                                s_name_out = "Missing"
                                
                            ws_conc.write(row_idx, 0, subj_formatted, data_white)
                            ws_conc.write(row_idx, 1, period_val, data_white)
                            ws_conc.write(row_idx, 2, f"'{tp_val}", data_white)
                            ws_conc.write(row_idx, 3, display_conc, data_white if display_conc in ["BLQ", "M"] else data_white_float)
                            ws_conc.write(row_idx, 4, s_name_out, data_white)
                            row_idx += 1
            else:
                for s_name, (conc, orig_file) in sorted(global_samples.items(), key=lambda x: extract_sort_num(x[0])):
                    display_conc = "BLQ" if (pd.isna(conc) or (pd.notna(lloq) and conc < lloq)) else conc
                    parts = re.split(r'[-_]', s_name)
                    subj = parts[0] if len(parts) > 0 else s_name
                    period = parts[1] if len(parts) > 1 else ""
                    timep = parts[2] if len(parts) > 2 else ""
                    if subj.startswith('S') and subj[1:].isdigit(): subj = f"S{int(subj[1:]):03d}"
                        
                    ws_conc.write(row_idx, 0, subj, data_white)
                    ws_conc.write(row_idx, 1, period, data_white)
                    ws_conc.write(row_idx, 2, f"'{timep}", data_white)
                    ws_conc.write(row_idx, 3, display_conc, data_white if display_conc == "BLQ" else data_white_float)
                    ws_conc.write(row_idx, 4, s_name, data_white)
                    row_idx += 1

            if repeat_files:
                ws_rep = workbook.add_worksheet("Final Conc After Repeat")
                ws_rep.set_column('A:B', 15, general_fmt)
                ws_rep.set_column('C:C', 15, general_fmt)
                ws_rep.set_column('D:D', 20, conc_fmt)
                ws_rep.set_column('E:E', 25, general_fmt)
                ws_rep.set_column('F:F', 20, conc_fmt)
                ws_rep.set_column('G:G', 20, pct_fmt)
                
                rep_headers = ["Subject No.", "Period No.", "Timepoints", "Final Conc.(ng/mL)", "Sample Name", "Initial Conc.(ng/mL)", "% Difference"]
                for col_num, header in enumerate(rep_headers): ws_rep.write(0, col_num, header, table_header_green)
                
                row_idx = 1
                if num_subj > 0 and periods_list and timepoints_list:
                    for subj_idx in range(1, num_subj + 1):
                        subj_formatted = f"S{subj_idx:03d}"
                        for period_val in periods_list:
                            period_clean = period_val.strip().upper()
                            for tp_val in timepoints_list:
                                tp_nums = re.findall(r'\d+\.\d+|\d+', str(tp_val))
                                tp_float = float(tp_nums[0]) if tp_nums else -1.0
                                if "PRE" in str(tp_val).upper(): tp_float = -1.0
                                
                                found_match = False
                                for base_up, (initial_conc, real_s_name, orig_file) in base_to_initial.items():
                                    parts = re.split(r'[-_]', real_s_name)
                                    subj_str = parts[0] if len(parts) > 0 else real_s_name
                                    p_str = parts[1] if len(parts) > 1 else ""
                                    t_str = parts[2] if len(parts) > 2 else ""
                                    
                                    s_nums = re.findall(r'\d+', subj_str)
                                    s_int = int(s_nums[-1]) if s_nums else 0
                                    p_clean = p_str.strip().upper()
                                    
                                    t_nums = re.findall(r'\d+\.\d+|\d+', t_str)
                                    t_flt = float(t_nums[0]) if t_nums else -1.0
                                    if "PRE" in t_str.upper(): t_flt = -1.0
                                    
                                    if s_int == subj_idx and p_clean == period_clean and t_flt == tp_float:
                                        
                                        repeat_conc, repeat_s_name_raw, rep_file = base_to_repeat.get(base_up, (np.nan, None, ""))
                                        has_repeat = repeat_s_name_raw is not None
                                        
                                        final_conc = repeat_conc if has_repeat else initial_conc
                                        final_s_name = repeat_s_name_raw if has_repeat else real_s_name
                                        
                                        display_final = "BLQ" if (pd.isna(final_conc) or (pd.notna(lloq) and final_conc < lloq)) else final_conc
                                        display_initial = "BLQ" if (pd.isna(initial_conc) or (pd.notna(lloq) and initial_conc < lloq)) else initial_conc
                                        
                                        pct_diff = ""
                                        if has_repeat and pd.notna(initial_conc) and pd.notna(repeat_conc) and display_initial != "BLQ" and display_final != "BLQ" and initial_conc > 0:
                                            pct_diff = ((repeat_conc - initial_conc) / ((initial_conc + repeat_conc) / 2)) * 100
                                            
                                        ws_rep.write(row_idx, 0, subj_formatted, data_white)
                                        ws_rep.write(row_idx, 1, period_val, data_white)
                                        ws_rep.write(row_idx, 2, f"'{tp_val}", data_white)
                                        ws_rep.write(row_idx, 3, display_final, data_white if display_final in ["BLQ", "M"] else data_white_float)
                                        ws_rep.write(row_idx, 4, final_s_name, data_white)
                                        
                                        if has_repeat:
                                            ws_rep.write(row_idx, 5, display_initial, data_white if display_initial in ["BLQ", "M"] else data_white_float)
                                            ws_rep.write(row_idx, 6, pct_diff, data_white_pct if pct_diff != "" else data_white)
                                        else:
                                            ws_rep.write(row_idx, 5, "", data_white)
                                            ws_rep.write(row_idx, 6, "", data_white)
                                            
                                        found_match = True
                                        break
                                        
                                if not found_match:
                                    ws_rep.write(row_idx, 0, subj_formatted, data_white)
                                    ws_rep.write(row_idx, 1, period_val, data_white)
                                    ws_rep.write(row_idx, 2, f"'{tp_val}", data_white)
                                    ws_rep.write(row_idx, 3, "M", data_white)
                                    ws_rep.write(row_idx, 4, "Missing", data_white)
                                    ws_rep.write(row_idx, 5, "", data_white)
                                    ws_rep.write(row_idx, 6, "", data_white)
                                    
                                row_idx += 1
                else:
                    for s_name_raw, (initial_conc, orig_file) in sorted(global_samples.items(), key=lambda x: extract_sort_num(x[0])):
                        parts = re.split(r'[-_]', s_name_raw)
                        subj = parts[0] if len(parts) > 0 else s_name_raw
                        period = parts[1] if len(parts) > 1 else ""
                        timep = parts[2] if len(parts) > 2 else ""
                        if subj.startswith('S') and subj[1:].isdigit(): subj = f"S{int(subj[1:]):03d}"
                        
                        base_up = get_base_name(s_name_raw).upper()
                        repeat_conc, repeat_s_name_raw, rep_file = base_to_repeat.get(base_up, (np.nan, None, ""))
                        
                        has_repeat = repeat_s_name_raw is not None
                        final_conc = repeat_conc if has_repeat else initial_conc
                        final_s_name = repeat_s_name_raw if has_repeat else s_name_raw
                        
                        display_final = "BLQ" if (pd.isna(final_conc) or (pd.notna(lloq) and final_conc < lloq)) else final_conc
                        display_initial = "BLQ" if (pd.isna(initial_conc) or (pd.notna(lloq) and initial_conc < lloq)) else initial_conc
                        
                        pct_diff = ""
                        if has_repeat and pd.notna(initial_conc) and pd.notna(repeat_conc) and display_initial != "BLQ" and display_final != "BLQ" and initial_conc > 0:
                            pct_diff = ((repeat_conc - initial_conc) / ((initial_conc + repeat_conc) / 2)) * 100
                            
                        ws_rep.write(row_idx, 0, subj, data_white)
                        ws_rep.write(row_idx, 1, period, data_white)
                        ws_rep.write(row_idx, 2, f"'{timep}", data_white)
                        ws_rep.write(row_idx, 3, display_final, data_white if display_final == "BLQ" else data_white_float)
                        ws_rep.write(row_idx, 4, final_s_name, data_white)
                        if has_repeat:
                            ws_rep.write(row_idx, 5, display_initial, data_white if display_initial == "BLQ" else data_white_float)
                            ws_rep.write(row_idx, 6, pct_diff, data_white_pct if pct_diff != "" else data_white)
                        else:
                            ws_rep.write(row_idx, 5, "", data_white)
                            ws_rep.write(row_idx, 6, "", data_white)
                        row_idx += 1

        st.success("Dashboard successfully generated with QA Summary Tab and Automated Repeats Tracker!")
        st.download_button(
            label="Download Master Excel Dashboard",
            data=output.getvalue(),
            file_name="Master_Bioanalytical_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload BOTH the Setup file and the TSQ Export file(s).")
