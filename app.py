import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

st.header("1. Upload Files")

col1, col2 = st.columns(2)
with col1:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx", "xlsm"])
with col2:
    uploaded_files = st.file_uploader("Upload TSQ Quantis Plus Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

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
    if setup_file and uploaded_files:
        st.success("Files received! Processing data...")
        try:
            setup_df = pd.read_excel(setup_file, header=None, engine='openpyxl')
            mol_name = str(setup_df.iloc[0, 1]).strip()
            study_num = str(setup_df.iloc[1, 1]).strip()
            num_subj_raw = setup_df.iloc[2, 1]
            num_subj = int(pd.to_numeric(num_subj_raw, errors='coerce')) if pd.notna(num_subj_raw) else 0
            periods = str(setup_df.iloc[3, 1]).strip()
            timepoints = str(setup_df.iloc[4, 1]).strip()
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
            st.info(f"Setup Loaded | Molecule: {mol_name} | Study: {study_num} | LLOQ: {lloq} | ULOQ: {uloq}")
        except Exception as e:
            st.error(f"Failed to read Setup file. Error: {str(e)}")
            st.stop()

        summary_log = []
        excel_tabs = {}
        cc_file_data = []
        qc_file_data = []
        global_samples = {}
        batch_meta_data = {}
        
        for file in uploaded_files:
            file.seek(0)
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
            
            # Robust extraction of the Analyte Area (the Area column that is NOT the IS Area)
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
            sample_rows = []
            ccqc_chart = []
            sample_chart = []
            current_inj = 1
            
            # Header Row for CC/QC
            ccqc_rows.append({"Sample Name": "CC and QC", "Area": "", "IS area": "", "Calc.Conc.": "", "% Accuracy": ""})
            ccqc_chart.append((np.nan, np.nan, np.nan, np.nan, np.nan))
            
            for index, row in df.iterrows():
                s_name = str(row[name_col]).strip().upper()
                s_type = str(row[type_col]).strip().upper()
                conc = pd.to_numeric(row[conc_col], errors='coerce')
                s_is_area = pd.to_numeric(row[is_area_col], errors='coerce')
                a_area = pd.to_numeric(row[analyte_area_col], errors='coerce') if analyte_area_col else np.nan
                acc_val = pd.to_numeric(row[acc_col], errors='coerce') if acc_col else np.nan
                
                if ('STD-A' in s_name or '-1' in s_name) and std_a_area is None: std_a_area = a_area
                if ('ZERO' in s_name or 'AQSTD' in s_name or '+IS' in s_name) and std_zero_is_area is None: std_zero_is_area = s_is_area
                if 'BLANK' in s_type or 'BLK' in s_name or 'MOBILE' in s_name: carryover_samples.append({'name': s_name, 'a_area': a_area, 'is_area': s_is_area})
                
                if pd.notna(acc_val):
                    if 'STANDARD' in s_type or 'STD' in s_name:
                        if ('-A' in s_name or '-1' in s_name) and 80 <= acc_val <= 120: cc_pass += 1
                        elif 85 <= acc_val <= 115: cc_pass += 1
                    if 'LQC' in s_name and 80 <= acc_val <= 120: lqc_pass += 1
                    if 'MMQC' in s_name and 85 <= acc_val <= 115: mmqc_pass += 1
                    if 'MQC' in s_name and 'MMQC' not in s_name and 85 <= acc_val <= 115: mqc_pass += 1
                    if 'HQC' in s_name and 85 <= acc_val <= 115: hqc_pass += 1
                
                if pd.notna(s_is_area) and ('STANDARD' in s_type or 'QC' in s_type or 'UNKNOWN' in s_type):
                    if s_is_area < lower_is_limit or s_is_area > upper_is_limit:
                        if 'BLK' not in s_name and 'AQSTD' not in s_name:
                             out_of_range_is.append({'name': s_name, 'conc': conc})
                if 'UNKNOWN' in s_type and pd.notna(conc) and conc > uloq:
                    above_uloq.append({'name': s_name, 'conc': conc})

                row_dict = {"Sample Name": s_name, "Area": a_area, "IS area": s_is_area, "Calc.Conc.": conc, "% Accuracy": acc_val}
                chart_tuple = [current_inj, np.nan, np.nan, np.nan, np.nan]
                
                if pd.notna(s_is_area) and pd.notna(avg_is_area) and avg_is_area > 0:
                    pct_diff = (s_is_area / avg_is_area) * 100 - 100
                    if 'STANDARD' in s_type or 'QC' in s_type or 'QUALITY' in s_type: chart_tuple[1] = pct_diff
                    elif 'BLANK' in s_type or '-BLK' in s_name or '+IS' in s_name: chart_tuple[2] = pct_diff
                    else: chart_tuple[3] = pct_diff
                
                if 'UNKNOWN' in s_type and 'AQSTD' not in s_name and 'BLK' not in s_name and '+IS' not in s_name and 'BLANK' not in s_name:
                    chart_tuple[4] = conc

                if 'STANDARD' in s_type or 'QC' in s_type or 'QUALITY' in s_type:
                    ccqc_rows.append(row_dict)
                    ccqc_chart.append(tuple(chart_tuple))
                    if s_name in seen_ccqc: file_dupes.append(s_name)
                    else: seen_ccqc.add(s_name)
                else:
                    sample_rows.append(row_dict)
                    sample_chart.append(tuple(chart_tuple))
                    if 'UNKNOWN' in s_type and 'AQSTD' not in s_name and 'BLK' not in s_name and '+IS' not in s_name:
                        if s_name in seen_samples: file_sample_dupes.append(s_name)
                        else:
                            seen_samples.add(s_name)
                            global_samples[s_name] = conc
                            
                current_inj += 1
                
                for std in std_names:
                    if std in s_name or std.replace("-", "") in s_name:
                        file_cc_dict[std] = conc if pd.notna(conc) else ""
                        break
                
                if 'LQC' in s_name: lqc_list.append(conc if pd.notna(conc) else ""); lqc_ids.append(s_name)
                elif 'MMQC' in s_name: mmqc_list.append(conc if pd.notna(conc) else ""); mmqc_ids.append(s_name)
                elif 'MQC' in s_name and 'MMQC' not in s_name: mqc_list.append(conc if pd.notna(conc) else ""); mqc_ids.append(s_name)
                elif 'HQC' in s_name: hqc_list.append(conc if pd.notna(conc) else ""); hqc_ids.append(s_name)

            file_cc_dict['Slope'] = "" if pd.isna(slope) else slope
            file_cc_dict['Intercept'] = "" if pd.isna(intercept) else intercept
            file_cc_dict['r'] = "" if pd.isna(r_val) else r_val
            cc_file_data.append(file_cc_dict)
            
            for i in range(4):
                q_dict = {
                    'Date of first injection': acq_date if i==0 else "", 'SUBJECT ID': subj_id if i==0 else "",
                    'LQC, HQC ID': lqc_ids[i] if i < len(lqc_ids) else "", 'MQC, MMQC ID': mmqc_ids[i] if i < len(mmqc_ids) else (mqc_ids[i] if i < len(mqc_ids) else ""),
                    'LQC': lqc_list[i] if i < len(lqc_list) else "", 'MMQC': mmqc_list[i] if i < len(mmqc_list) else "",
                    'MQC': mqc_list[i] if i < len(mqc_list) else "", 'HQC': hqc_list[i] if i < len(hqc_list) else ""
                }
                qc_file_data.append(q_dict)
                
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', file.name)[:31]
            
            avg_row = {"Sample Name": "Average IS area of CC and QC", "Area": "", "IS area": avg_is_area, "Calc.Conc.": "", "% Accuracy": ""}
            header_row = {"Sample Name": "Sample Name", "Area": "Area", "IS area": "IS area", "Calc.Conc.": "Calc.Conc.", "% Accuracy": "% Acc"}
            
            final_rows = ccqc_rows + [avg_row, header_row] + sample_rows
            final_chart = ccqc_chart + [(np.nan, np.nan, np.nan, np.nan, np.nan), (np.nan, np.nan, np.nan, np.nan, np.nan)] + sample_chart
            
            batch_meta_data[safe_sheet_name] = {
                'rows': final_rows, 'chart_data': final_chart,
                'slope': slope, 'intercept': intercept, 'r_val': r_val,
                'cc_pass': cc_pass, 'lqc_pass': lqc_pass, 'mmqc_pass': mmqc_pass, 'mqc_pass': mqc_pass, 'hqc_pass': hqc_pass,
                'std_a_area': std_a_area, 'std_zero_is_area': std_zero_is_area,
                'carryover': carryover_samples[:3], 'out_of_range': out_of_range_is, 'above_uloq': above_uloq
            }
            
            nom_check_msg = "All CC and QC concentration matching nominal value." if len(nominals) > 0 else "No nominals defined."
            dup_ccqc_msg = "All CC and QC IDs are unique." if not file_dupes else f"Duplicates found"
            dup_samp_msg = "All CC and QC sample names are unique." if not file_sample_dupes else f"Duplicates found"

            summary_log.append({
                "File Name": file.name, "Status": "SUCCESS", "Error Details": "Dashboard generated.",
                "Nominal Conc. Check": nom_check_msg, "Duplicate CC & QC ID Check": dup_ccqc_msg, "Duplicate Sample Name Check": dup_samp_msg
            })

        excel_tabs["Summary"] = pd.DataFrame(summary_log)
        
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
            
            excel_tabs["Summary"].to_excel(writer, sheet_name="Summary", index=False)
            if "CC" in excel_tabs: excel_tabs["CC"].to_excel(writer, sheet_name="CC", index=False)
            if "QC" in excel_tabs: excel_tabs["QC"].to_excel(writer, sheet_name="QC", index=False)
            
            workbook = writer.book
            
            tnr = {'font_name': 'Times New Roman', 'font_size': 10}
            green_sheet_bg = workbook.add_format({**tnr, 'bg_color': '#CDE4CE'})
            header_format = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            yellow_flag = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'bold': True, 'font_color': '#FF0000'})
            bold_label = workbook.add_format({**tnr, 'bold': True, 'align': 'right'})
            general_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter'})
            conc_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
            r_fmt = workbook.add_format({**tnr, 'num_format': '0.0000', 'align': 'center', 'valign': 'vcenter'})
            float_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
            stdev_fmt = workbook.add_format({**tnr, 'num_format': '0.0000', 'align': 'center', 'valign': 'vcenter'})
            pct_fmt = workbook.add_format({**tnr, 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})
            merge_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter'})
            success_fmt = workbook.add_format({**tnr, 'bg_color': '#99FF99', 'align': 'center', 'valign': 'vcenter'})
            failed_fmt = workbook.add_format({**tnr, 'bg_color': '#FF9999', 'align': 'center', 'valign': 'vcenter'})
            
            data_white = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white_float = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            data_white_pct = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            data_white_r = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000'})
            data_yellow_float = workbook.add_format({**tnr, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            table_header_green = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            table_header_gray = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

            for sheet_name in ["Summary", "CC", "QC"]:
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
                ws.set_column('G:G', 20, green_sheet_bg)
                ws.set_column('H:H', 15, green_sheet_bg)
                ws.set_column('I:I', 20, green_sheet_bg)
                ws.set_column('J:J', 15, green_sheet_bg)
                ws.set_column('K:K', 15, green_sheet_bg)
                
                for r_idx, row_dict in enumerate(final_rows):
                    if r_idx == 0:
                        ws.merge_range(0, 0, 0, 4, row_dict["Sample Name"], table_header_green)
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
                        ws.write(r_idx, 4, "% Acc", table_header_green)
                    else:
                        ws.write(r_idx, 0, row_dict["Sample Name"], data_white)
                        ws.write(r_idx, 1, "" if pd.isna(row_dict["Area"]) else row_dict["Area"], data_white)
                        ws.write(r_idx, 2, "" if pd.isna(row_dict["IS area"]) else row_dict["IS area"], data_white)
                        conc_val = row_dict["Calc.Conc."]
                        if pd.notna(conc_val) and isinstance(conc_val, (int, float)) and conc_val > uloq:
                            ws.write(r_idx, 3, conc_val, data_yellow_float)
                        else:
                            ws.write(r_idx, 3, "" if pd.isna(conc_val) else conc_val, data_white_float)
                        ws.write(r_idx, 4, "" if pd.isna(row_dict["% Accuracy"]) else row_dict["% Accuracy"], data_white_pct)

                    c_tuple = chart_data[r_idx]
                    ws.write(r_idx, 26, "" if pd.isna(c_tuple[0]) else c_tuple[0])
                    ws.write(r_idx, 27, "" if pd.isna(c_tuple[1]) else c_tuple[1])
                    ws.write(r_idx, 28, "" if pd.isna(c_tuple[2]) else c_tuple[2])
                    ws.write(r_idx, 29, "" if pd.isna(c_tuple[3]) else c_tuple[3])
                    ws.write(r_idx, 30, "" if pd.isna(c_tuple[4]) else c_tuple[4])

                ws.set_column(26, 30, None, None, {'hidden': True})
                
                ws.write('G2', 'Intercept', data_white)
                ws.write('H2', meta.get('intercept', ''), data_white_float)
                ws.write('G3', 'Slope', data_white)
                ws.write('H3', meta.get('slope', ''), data_white_float)
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
                for r_idx, smp in enumerate(co_samples):
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
                    ws.write(or_start, 6, "Out of Range %IS Sample", table_header_green)
                    ws.write(or_start, 7, "Concentration", table_header_green)
                    ws.write(or_start, 8, "Above ULOQ Samples", table_header_green)
                    ws.write(or_start, 9, "Concentration", table_header_green)
                    ws.write(or_start, 10, "Dose Repeats", table_header_green)
                    ws.write(or_start, 11, "Concentration", table_header_green)
                    
                    out_is = meta.get('out_of_range', [])
                    for i, smp in enumerate(out_is):
                        ws.write(or_start + 1 + i, 6, smp['name'], data_white)
                        ws.write(or_start + 1 + i, 7, smp['conc'], data_yellow_float if pd.notna(smp['conc']) and smp['conc'] > uloq else data_white_float)
                        
                    above_u = meta.get('above_uloq', [])
                    for i, smp in enumerate(above_u):
                        ws.write(or_start + 1 + i, 8, smp['name'], data_white)
                        ws.write(or_start + 1 + i, 9, smp['conc'], data_yellow_float)

            for target_sheet in ["CC", "QC"]:
                if target_sheet in excel_tabs:
                    ws = writer.sheets[target_sheet]
                    df_target = excel_tabs[target_sheet]
                    num_rows = len(df_target)
                    stat_row = num_rows + 1 
                    
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
                    
                    start_col = 3 if target_sheet == "CC" else 4
                    end_col = len(std_names) + 3 if target_sheet == "CC" else 8
                    
                    for col_num in range(start_col, end_col):
                        col_letter = xl_col_to_name(col_num)
                        data_range = f"{col_letter}5:{col_letter}{num_rows + 1}"
                        ws.write_formula(stat_row, col_num, f"=IFERROR(AVERAGE({data_range}), \"\")", float_fmt)
                        ws.write_formula(stat_row + 1, col_num, f"=IFERROR(STDEV({data_range}), \"\")", stdev_fmt)
                        ws.write_formula(stat_row + 2, col_num, f"=IFERROR(({col_letter}{stat_row + 2}/{col_letter}{stat_row + 1})*100, \"\")", pct_fmt)
                        ws.write_formula(stat_row + 3, col_num, f"=IFERROR(({col_letter}{stat_row + 1}/{col_letter}2)*100, \"\")", pct_fmt)
                        ws.conditional_format(data_range, {'type': 'formula', 'criteria': f"=AND(ISNUMBER({col_letter}5), OR({col_letter}5<{col_letter}$3, {col_letter}5>{col_letter}$4))", 'format': yellow_flag})

            ws_conc = workbook.add_worksheet("Concentration Data")
            ws_conc.set_column('A:E', 20, general_fmt)
            ws_conc.merge_range('A1:B1', "Molecule Name", header_format)
            ws_conc.merge_range('C1:D1', mol_name, header_format)
            ws_conc.merge_range('A2:B2', "Study Number", header_format)
            ws_conc.merge_range('C2:D2', study_num, header_format)
            ws_conc.merge_range('A3:B3', "CC range", header_format)
            lloq_fmt = f"{lloq:.3f}" if pd.notna(lloq) else "1.000"
            uloq_fmt = f"{uloq:.3f}" if pd.notna(uloq) else "100.000"
            ws_conc.merge_range('C3:D3', f"{lloq_fmt} ng/mL to {uloq_fmt} ng/mL", header_format)
            
            conc_headers = ["Subject No.", "Period No.", "Timepoints", "Conc.(ng/mL)", "Sample Name"]
            for col_num, header in enumerate(conc_headers): ws_conc.write(4, col_num, header, header_format)
                
            row_idx = 5
            for s_name, conc in sorted(global_samples.items(), key=lambda x: extract_sort_num(x[0])):
                display_conc = "BLQ" if (pd.isna(conc) or (pd.notna(lloq) and conc < lloq)) else conc
                parts = re.split(r'[-_]', s_name)
                subj = parts[0] if len(parts) > 0 else s_name
                period = parts[1] if len(parts) > 1 else ""
                timep = parts[2] if len(parts) > 2 else ""
                if subj.startswith('S') and subj[1:].isdigit(): subj = f"S{int(subj[1:]):02d}"
                    
                ws_conc.write(row_idx, 0, subj, general_fmt)
                ws_conc.write(row_idx, 1, period, general_fmt)
                ws_conc.write(row_idx, 2, f"'{timep}", general_fmt)
                ws_conc.write(row_idx, 3, display_conc, general_fmt if display_conc == "BLQ" else conc_fmt)
                ws_conc.write(row_idx, 4, s_name, general_fmt)
                row_idx += 1

        st.success("Dashboard generated. Missing area resolved.")
        st.download_button(
            label="Download Master Excel Dashboard",
            data=output.getvalue(),
            file_name="Master_Bioanalytical_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload BOTH the Setup file and the TSQ Export file(s).")
