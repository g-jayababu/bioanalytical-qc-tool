import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(page_title="QC Dashboard", layout="wide")

def extract_sort_num(val):
    nums = re.findall(r'\d+', str(val))
    return int(nums[-1]) if nums else 0

def extract_qc_id(name):
    parts = re.split(r'[-_]', str(name).strip())
    if parts and parts[-1].strip().isdigit():
        return parts[-1].strip()
    return str(name).strip()

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def get_base_name(name):
    if not isinstance(name, str): return str(name)
    return re.sub(r'[-_](ISR\d*|R\d*|REPEAT\d*)$', '', name, flags=re.IGNORECASE).strip()

def load_data(file):
    file_name_lower = file.name.lower()
    dummy_cols = list(range(100))
    file.seek(0)
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
                return None, f"Failed to read .xls format. Error: {str(e2)}"
    return None, "Unsupported file extension."

st.title("Bioanalytical Auto-QC Dashboard")

# UI Layout for inputs
st.header("1. Upload Files")
col1, col2 = st.columns(2)
with col1:
    setup_file = st.file_uploader("Select Setup Excel File", type=["xlsx", "xls", "xlsm"])
with col2:
    uploaded_files = st.file_uploader("Select TSQ Data Files (Main, Repeats, ISR)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

st.header("2. Settings")
col3, col4 = st.columns(2)
with col3:
    generate_cc = st.checkbox("Generate CC Sheet?", value=True)
with col4:
    generate_qc = st.checkbox("Generate QC Sheet?", value=True)

if st.button("Process Data & Generate Dashboard", type="primary"):
    if not setup_file:
        st.warning("Please upload a Setup file.")
        st.stop()
    if not uploaded_files:
        st.warning("Please upload at least one TSQ data file.")
        st.stop()

    with st.spinner("Processing Data..."):
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
            st.error(f"Failed to read Setup file.\nError: {str(e)}")
            st.stop()

        # Initialize Processing Variables
        summary_log, cc_file_data, qc_file_data, all_repeats_data = [], [], [], []
        global_samples, repeat_samples, isr_samples, batch_meta_data = {}, {}, {}, {}
        excel_tabs = {}

        for file in uploaded_files:
            filename = file.name
            
            # Categorize File
            file_category = "Main"
            if 'REPEAT' in filename.upper(): file_category = "Repeat"
            elif 'ISR' in filename.upper(): file_category = "ISR"
            
            # Extract base sheet name
            matches = re.findall(r'(S[-_]?\d+(?:[-_]+S[-_]?\d+)*)', filename, re.IGNORECASE)
            if matches: base_name = matches[-1].upper()
            elif file_category == "Repeat":
                rep_match = re.search(r'(REPEAT[-_]?\d*)', filename, re.IGNORECASE)
                base_name = rep_match.group(1).upper() if rep_match else "REPEAT"
            elif file_category == "ISR":
                isr_match = re.search(r'(ISR[-_]?\d*)', filename, re.IGNORECASE)
                base_name = isr_match.group(1).upper() if isr_match else "ISR"
            else:
                base_name = filename
                
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', base_name)[:31]
            counter = 1
            while safe_sheet_name in batch_meta_data:
                safe_sheet_name = f"{safe_sheet_name[:27]}_{counter}"
                counter += 1
            
            df, error_msg = load_data(file)
            if df is None:
                summary_log.append({"File Name": filename, "Status": "FAILED", "Error Details": error_msg})
                continue
                
            df = df.replace(r'^\s*$', np.nan, regex=True).dropna(how='all').dropna(how='all', axis=1).reset_index(drop=True)
            df = df.fillna("")
            
            # Find Headers
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
            
            acq_date = pd.Timestamp.now().strftime('%d-%b-%Y').upper()
            if date_col and not df[date_col].empty:
                raw_date = str(df[date_col].iloc[0]).strip()
                try: acq_date = pd.to_datetime(raw_date).strftime('%d-%b-%Y').upper()
                except Exception: acq_date = raw_date
            
            if not all([name_col, type_col, is_area_col, conc_col]):
                summary_log.append({"File Name": filename, "Status": "FAILED", "Error Details": "Missing critical data columns."})
                continue
            
            subj_id = "S1"
            if matches: subj_id = matches[-1].upper()
            elif file_category == "Repeat": subj_id = "Repeat"
            elif file_category == "ISR": subj_id = "ISR"
                
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
            
            ccqc_rows, sample_rows = [], []
            ccqc_chart, sample_chart = [], []
            current_inj = 1
            
            ccqc_rows.append({"Sample Name": "CC and QC", "Area": "", "IS area": "", "Calc.Conc.": "", "Raw.Conc.": "", "% Accuracy": "", "failed_ccqc": False})
            ccqc_chart.append((np.nan, np.nan, np.nan, np.nan, np.nan))
            ccqc_rows.append({"Sample Name": "Sample Name", "Area": "Area", "IS area": "IS area", "Calc.Conc.": "Calc.Conc.", "Raw.Conc.": "", "% Accuracy": "% Acc", "failed_ccqc": False})
            ccqc_chart.append((np.nan, np.nan, np.nan, np.nan, np.nan))
            
            blk_a_area, blk_is_area, zero_a_area, zero_is_area, mph_a_area, mph_is_area = None, None, None, None, None, None
            
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
                
                # Identify Blanks/Zeros
                if any(x in s_name_up for x in ['MPH', 'MOH', 'ESO', 'RCS', 'MOBILE', 'MP', 'RECON']):
                    if mph_a_area is None: mph_a_area = a_area if pd.notna(a_area) else 0.0
                    if mph_is_area is None: mph_is_area = s_is_area if pd.notna(s_is_area) else 0.0
                elif ('ZERO' in s_name_up or '+IS' in s_name_up or 'BLANK+IS' in s_name_up) and 'AQSTD' not in s_name_up:
                    if zero_a_area is None: zero_a_area = a_area if pd.notna(a_area) else 0.0
                    if zero_is_area is None: zero_is_area = s_is_area if pd.notna(s_is_area) else 0.0
                elif 'BLK' in s_name_up or 'BLANK' in s_type_up:
                    if blk_a_area is None: blk_a_area = a_area if pd.notna(a_area) else 0.0
                    if blk_is_area is None: blk_is_area = s_is_area if pd.notna(s_is_area) else 0.0
                
                failed_ccqc = False
                if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up:
                    if pd.notna(acc_val):
                        if ('STANDARD' in s_type_up or 'STD' in s_name_up) and ('-A' in s_name_up or '-1' in s_name_up):
                            if not (80 <= acc_val <= 120): failed_ccqc = True
                        else:
                            if not (85 <= acc_val <= 115): failed_ccqc = True
                    else: failed_ccqc = True
                
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
                
                if 'LQC' in s_name_up: lqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'MMQC' in s_name_up: mmqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'MQC' in s_name_up and 'MMQC' not in s_name_up: mqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                if 'HQC' in s_name_up: hqc_list_objs.append({'name': s_name_raw, 'pass': pd.notna(acc_val) and 85 <= acc_val <= 115})
                
                # Repeats Logic
                if pd.notna(s_is_area) and ('STANDARD' in s_type_up or 'QC' in s_type_up or 'UNKNOWN' in s_type_up):
                    if s_is_area < lower_is_limit or s_is_area > upper_is_limit:
                        if 'BLK' not in s_name_up and 'AQSTD' not in s_name_up: out_of_range_is.append({'name': s_name_raw, 'conc': conc})
                if 'UNKNOWN' in s_type_up and pd.notna(conc):
                    if conc > uloq: above_uloq.append({'name': s_name_raw, 'conc': conc})
                    is_predose = 'PRE' in s_name_up or bool(re.search(r'[-_]0(\.0+)?$', s_name_up))
                    if is_predose:
                        parts = re.split(r'[-_]', s_name_raw)
                        current_subj = parts[0].upper() if len(parts) > 0 else s_name_up
                        subj_mask = df[name_col].astype(str).str.upper().str.contains(current_subj, na=False, regex=False)
                        unk_mask = df[type_col].astype(str).str.upper().str.contains('UNKNOWN', na=False)
                        subj_concs = pd.to_numeric(df.loc[subj_mask & unk_mask, conc_col], errors='coerce')
                        cmax = subj_concs.max() if not subj_concs.empty else 0
                        
                        pd_flag = False
                        if pd.notna(lloq) and conc > lloq: pd_flag = True
                        elif std_a_area is not None and pd.notna(a_area) and a_area > std_a_area: pd_flag = True
                        elif pd.notna(cmax) and cmax > 0 and conc > (0.05 * cmax): pd_flag = True
                            
                        if pd_flag: predose_repeats.append({'name': s_name_raw, 'conc': conc})

                # Format rows
                chart_tuple = [current_inj, np.nan, np.nan, np.nan, np.nan]
                pct_diff = np.nan
                if pd.notna(s_is_area) and pd.notna(avg_is_area) and avg_is_area > 0:
                    pct_diff = (s_is_area / avg_is_area) * 100
                    if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up: chart_tuple[1] = pct_diff
                    elif 'BLANK' in s_type_up or '-BLK' in s_name_up or '+IS' in s_name_up: chart_tuple[2] = pct_diff
                    else: chart_tuple[3] = pct_diff
                
                if 'UNKNOWN' in s_type_up and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up and '+IS' not in s_name_up and 'BLANK' not in s_name_up:
                    chart_tuple[4] = conc

                row_dict = {"Sample Name": s_name_raw, "Area": a_area, "Raw.Area": raw_a_area_val, "IS area": s_is_area, "Raw.IS": raw_is_area_val, "Calc.Conc.": conc, "Raw.Conc.": raw_conc_val, "failed_ccqc": failed_ccqc}
                
                if 'STANDARD' in s_type_up or 'QC' in s_type_up or 'QUALITY' in s_type_up:
                    row_dict["% Accuracy"] = acc_val
                    ccqc_rows.append(row_dict)
                    ccqc_chart.append(tuple(chart_tuple))
                else:
                    row_dict["% Accuracy"] = pct_diff 
                    sample_rows.append(row_dict)
                    sample_chart.append(tuple(chart_tuple))
                    if 'UNKNOWN' in s_type_up and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up and '+IS' not in s_name_up and 'BLANK' not in s_name_up:
                        if file_category == "Main": global_samples[s_name_raw] = (conc, safe_sheet_name)
                        elif file_category == "Repeat": repeat_samples[s_name_raw] = (conc, safe_sheet_name)
                        elif file_category == "ISR": isr_samples[s_name_raw] = (conc, safe_sheet_name)
                            
                current_inj += 1
                for std in std_names:
                    if std in s_name_up or std.replace("-", "") in s_name_up:
                        file_cc_dict[std] = conc if pd.notna(conc) else ""
                        break
                
                if 'LQC' in s_name_up: lqc_list.append(conc if pd.notna(conc) else ""); lqc_ids.append(extract_qc_id(s_name_raw))
                elif 'MMQC' in s_name_up: mmqc_list.append(conc if pd.notna(conc) else ""); mmqc_ids.append(extract_qc_id(s_name_raw))
                elif 'MQC' in s_name_up and 'MMQC' not in s_name_up: mqc_list.append(conc if pd.notna(conc) else ""); mqc_ids.append(extract_qc_id(s_name_raw))
                elif 'HQC' in s_name_up: hqc_list.append(conc if pd.notna(conc) else ""); hqc_ids.append(extract_qc_id(s_name_raw))

            file_cc_dict['Slope'] = "" if pd.isna(slope) else slope
            file_cc_dict['Intercept'] = "" if pd.isna(intercept) else intercept
            file_cc_dict['r'] = "" if pd.isna(r_val) else r_val
            cc_file_data.append(file_cc_dict)
            
            if file_category == "Main":
                for smp in out_of_range_is: all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Out of Range %IS"})
                for smp in above_uloq: all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Above ULOQ"})
                for smp in predose_repeats: all_repeats_data.append({"Batch Name": safe_sheet_name, "Sample Name": smp['name'], "Concentration": smp['conc'], "Reason Code": "Predose Detectable"})
            
            for i in range(4):
                qc_file_data.append({
                    'Date of first injection': acq_date if i==0 else "", 'SUBJECT ID': subj_id if i==0 else "",
                    'LQC, HQC ID': lqc_ids[i] if i < len(lqc_ids) else "", 'MQC, MMQC ID': mmqc_ids[i] if i < len(mmqc_ids) else (mqc_ids[i] if i < len(mqc_ids) else ""),
                    'LQC': lqc_list[i] if i < len(lqc_list) else "", 'MMQC': mmqc_list[i] if i < len(mmqc_list) else "",
                    'MQC': mqc_list[i] if i < len(mqc_list) else "", 'HQC': hqc_list[i] if i < len(hqc_list) else ""
                })
                
            avg_row = {"Sample Name": "Average IS area of CC and QC", "Area": "", "IS area": avg_is_area, "Calc.Conc.": "", "Raw.Conc.": "", "% Accuracy": "", "failed_ccqc": False}
            header_row = {"Sample Name": "Sample Name", "Area": "Area", "IS area": "IS area", "Calc.Conc.": "Calc.Conc.", "Raw.Conc.": "", "% Accuracy": "% IS", "failed_ccqc": False}
            
            batch_meta_data[safe_sheet_name] = {
                'full_file_name': filename, 'ccid_str': ccid_str,
                'rows': ccqc_rows + [avg_row, header_row] + sample_rows, 'chart_data': ccqc_chart + [(np.nan, np.nan, np.nan, np.nan, np.nan), (np.nan, np.nan, np.nan, np.nan, np.nan)] + sample_chart,
                'slope': slope, 'intercept': intercept, 'r_val': r_val, 'avg_is_area': avg_is_area,
                'cc_total': cc_total, 'cc_pass': cc_pass, 'failed_cc_ids': failed_cc_ids,
                'lqc_objs': lqc_list_objs, 'mmqc_objs': mmqc_list_objs, 'mqc_objs': mqc_list_objs, 'hqc_objs': hqc_list_objs,
                'std_a_area': std_a_area, 'std_zero_is_area': std_zero_is_area,
                'blk_a_area': blk_a_area, 'blk_is_area': blk_is_area, 'zero_a_area': zero_a_area, 'zero_is_area': zero_is_area,
                'mph_a_area': mph_a_area, 'mph_is_area': mph_is_area,
                'carryover': carryover_samples, 'out_of_range': out_of_range_is, 'above_uloq': above_uloq, 'predose_repeats': predose_repeats
            }
            
            summary_log.append({"File Name": filename, "Status": "SUCCESS", "Error Details": "Dashboard generated."})

        # Prepare Global Excel Tabs
        excel_tabs["Summary"] = pd.DataFrame(summary_log)
        
        if all_repeats_data: excel_tabs["Repeats"] = pd.DataFrame(all_repeats_data).sort_values(by='Sample Name', key=lambda col: col.map(natural_sort_key)).reset_index(drop=True)
        else: excel_tabs["Repeats"] = pd.DataFrame([{"Batch Name": "-", "Sample Name": "-", "Concentration": "-", "Reason Code": "No Repeats Required"}])
            
        base_to_initial = {get_base_name(s).upper(): (c, s, f) for s, (c, f) in global_samples.items()}
        base_to_repeat = {get_base_name(s).upper(): (c, s, f) for s, (c, f) in repeat_samples.items()}
        base_to_isr = {get_base_name(s).upper(): (c, s, f) for s, (c, f) in isr_samples.items()}

        # Save to In-Memory Buffer for Streamlit Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sh_name in ["Summary", "Repeats"]:
                excel_tabs[sh_name].to_excel(writer, sheet_name=sh_name, index=False)
                
            if generate_cc and cc_file_data:
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
                cc_df[cc_columns].to_excel(writer, sheet_name="CC", index=False)

            if generate_qc and qc_file_data:
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
                qc_df[qc_columns].to_excel(writer, sheet_name="QC", index=False)

            workbook = writer.book
            tnr = {'font_name': 'Times New Roman', 'font_size': 10}
            green_sheet_bg = workbook.add_format({**tnr, 'bg_color': '#CDE4CE'})
            table_header_green = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            table_header_gray = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white_float = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            data_white_pct = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            data_yellow_float = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            failed_ccqc_string_fmt = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            failed_ccqc_pct_fmt = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#000000', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})

            for safe_sheet_name in sorted(batch_meta_data.keys(), key=natural_sort_key):
                meta = batch_meta_data[safe_sheet_name]
                ws = workbook.add_worksheet(safe_sheet_name)
                
                ws.set_column('A:AD', 15, green_sheet_bg)
                ws.set_column('A:A', 25, green_sheet_bg)
                ws.set_column('G:M', 28, green_sheet_bg)
                
                for r_idx, row_dict in enumerate(meta['rows']):
                    if row_dict["Sample Name"] == "CC and QC": ws.merge_range(r_idx, 0, r_idx, 4, row_dict["Sample Name"], table_header_green)
                    elif row_dict["Sample Name"] == "Average IS area of CC and QC":
                        ws.merge_range(r_idx, 0, r_idx, 1, row_dict["Sample Name"], table_header_gray)
                        ws.write(r_idx, 2, row_dict["IS area"], data_white)
                        ws.write(r_idx, 3, "", data_white)
                        ws.write(r_idx, 4, "", data_white)
                    elif row_dict["Sample Name"] == "Sample Name":
                        for c_i, header in enumerate(["Sample Name", "Area", "IS area", "Calc.Conc.", row_dict.get("% Accuracy", "% Acc")]):
                            ws.write(r_idx, c_i, header, table_header_green)
                    else:
                        ws.write(r_idx, 0, row_dict["Sample Name"], data_white)
                        a_display = row_dict["Area"] if pd.notna(row_dict["Area"]) else row_dict["Raw.Area"]
                        ws.write(r_idx, 1, a_display, data_white)
                        
                        is_display = row_dict["IS area"] if pd.notna(row_dict["IS area"]) else row_dict["Raw.IS"]
                        ws.write(r_idx, 2, f"{is_display}*" if row_dict["failed_ccqc"] and is_display != "" else is_display, failed_ccqc_string_fmt if row_dict["failed_ccqc"] else data_white)
                            
                        c_display = row_dict["Calc.Conc."] if pd.notna(row_dict["Calc.Conc."]) else row_dict["Raw.Conc."]
                        if pd.notna(row_dict["Calc.Conc."]) and isinstance(c_display, (int, float)):
                            ws.write(r_idx, 3, c_display, data_yellow_float if c_display > uloq else data_white_float)
                        else:
                            ws.write(r_idx, 3, c_display, data_white)
                            
                        acc = row_dict.get("% Accuracy", "")
                        if row_dict["failed_ccqc"]:
                            ws.write(r_idx, 4, "" if pd.isna(acc) or acc == "" else acc, failed_ccqc_string_fmt if pd.isna(acc) or acc == "" else failed_ccqc_pct_fmt)
                        else:
                            ws.write(r_idx, 4, "" if pd.isna(acc) or acc == "" else acc, data_white if pd.isna(acc) or acc == "" else data_white_pct)

                    c_tuple = meta['chart_data'][r_idx]
                    for offset in range(5): ws.write(r_idx, 26 + offset, "" if pd.isna(c_tuple[offset]) else c_tuple[offset])

                ws.set_column(26, 30, None, None, {'hidden': True})

                or_start = 45
                for col_i, header in enumerate(["Out of Range %IS Samples", "Concentration", "Above ULOQ Samples", "Concentration", "Predose Repeats (PD)", "Concentration"]):
                    ws.write(or_start, 6 + col_i, header, table_header_green)
                
                for i, smp in enumerate(meta.get('out_of_range', [])):
                    ws.write(or_start + 1 + i, 6, smp['name'], data_white)
                    conc_val = "" if pd.isna(smp['conc']) else smp['conc']
                    ws.write(or_start + 1 + i, 7, conc_val, data_yellow_float if pd.notna(smp['conc']) and smp['conc'] > uloq else data_white_float)
                    
                for i, smp in enumerate(meta.get('above_uloq', [])):
                    ws.write(or_start + 1 + i, 8, smp['name'], data_white)
                    conc_val = "" if pd.isna(smp['conc']) else smp['conc']
                    ws.write(or_start + 1 + i, 9, conc_val, data_yellow_float)

                for i, smp in enumerate(meta.get('predose_repeats', [])):
                    ws.write(or_start + 1 + i, 10, smp['name'], data_white)
                    conc_val = "" if pd.isna(smp['conc']) else smp['conc']
                    ws.write(or_start + 1 + i, 11, conc_val, data_yellow_float)

        st.success("Successfully processed all files!")
        
        st.download_button(
            label="Download Master Dashboard Excel File",
            data=output.getvalue(),
            file_name="Master_Bioanalytical_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
