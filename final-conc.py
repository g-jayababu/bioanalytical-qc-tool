import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(layout="wide")
st.title("Simple Concentration & Repeat Tracker")

st.header("1. Upload Setup & Parameters")
col_a, col_b = st.columns(2)
with col_a:
    setup_file = st.file_uploader("Upload Setup Excel File", type=["xls", "xlsx", "xlsm"])
with col_b:
    unit = st.selectbox("Select Concentration Units", ["ng/mL", "pg/mL", "µg/mL"])

st.header("2. Upload TSQ Files")
col1, col2 = st.columns(2)
with col1:
    uploaded_files = st.file_uploader("Upload Main TSQ Export Files", type=["csv", "xls", "xlsx"], accept_multiple_files=True)
with col2:
    repeat_files = st.file_uploader("Upload Repeat Files (Optional)", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

st.header("3. ISR Selection Parameters")
col_x, col_y = st.columns(2)
with col_x:
    isr_cmax_count = st.number_input("Samples near Cmax (per subject/period)", min_value=0, value=1, step=1)
with col_y:
    isr_elim_count = st.number_input("Samples in Elimination (per subject/period)", min_value=0, value=1, step=1)

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
    """Extracts numeric portion for basic sorting"""
    nums = re.findall(r'\d+', str(val))
    return int(nums[-1]) if nums else 0

def natural_sort_key(s):
    """Splits string into text and numbers for true alphanumerical sorting"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def get_base_name(name):
    """Removes -R, -R1, -ISR, _ISR, etc. to universally match initial samples"""
    if not isinstance(name, str): return str(name)
    return re.sub(r'[-_](ISR\d*|R\d*|REPEAT\d*)$', '', name, flags=re.IGNORECASE).strip()

if st.button("Generate Final Concentration Tracker"):
    if not setup_file:
        st.warning("Please upload the Setup Excel File before generating.")
    elif uploaded_files:
        st.success("Files received! Processing concentration data...")
        
        # Parse Setup File Data
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
            lqc_val = 0.0
            for index in range(8, len(setup_df)):
                cc_id = str(setup_df.iloc[index, 0]).strip().upper()
                nom_val = pd.to_numeric(setup_df.iloc[index, 1], errors='coerce')
                if cc_id not in ['NAN', 'NONE'] and pd.notna(nom_val):
                    nominals[cc_id] = nom_val
                    if 'LQC' in cc_id: lqc_val = nom_val
            
            # Fallback LQC to 3x LLOQ if standard LQC naming wasn't found
            if lqc_val == 0.0 and pd.notna(lloq):
                lqc_val = lloq * 3
                
        except Exception as e:
            st.error(f"Failed to read Setup file. Please ensure it follows the standard format. Error: {str(e)}")
            st.stop()
            
        initial_samples = {}
        repeat_samples = {}
        
        combined_files = [(f, "Main") for f in uploaded_files]
        if repeat_files: combined_files.extend([(f, "Repeat") for f in repeat_files])
        
        for file, file_category in combined_files:
            file.seek(0)
            df, error_msg = load_data(file)
            if df is None:
                st.warning(f"Failed to load {file.name}: {error_msg}")
                continue
                
            df = df.replace(r'^\s*$', np.nan, regex=True).dropna(how='all').dropna(how='all', axis=1).reset_index(drop=True)
            df = df.fillna("")
            
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

            df = df.loc[:, ~df.columns.str.contains('^nan$|^$', case=False, na=False)]

            name_col = next((c for c in df.columns if 'sample name' in c or 'sample' in c), None)
            type_col = next((c for c in df.columns if 'sample type' in c or 'type' in c), None)
            conc_col = next((c for c in df.columns if 'calc' in c and 'conc' in c), None)
            
            if not all([name_col, type_col, conc_col]):
                st.warning(f"File {file.name} is missing critical columns (Name, Type, or Conc).")
                continue

            for index, row in df.iterrows():
                s_name_raw = str(row[name_col]).strip()
                s_name_up = s_name_raw.upper()
                s_type_up = str(row[type_col]).strip().upper()
                
                raw_conc_val = str(row[conc_col]).strip()
                conc = pd.to_numeric(raw_conc_val, errors='coerce') if raw_conc_val != "" else np.nan
                
                if 'UNKNOWN' in s_type_up and 'AQSTD' not in s_name_up and 'BLK' not in s_name_up and '+IS' not in s_name_up and 'BLANK' not in s_name_up:
                    base_up = get_base_name(s_name_raw).upper()
                    
                    if file_category == "Main":
                        initial_samples[base_up] = (conc, raw_conc_val, s_name_raw)
                    elif file_category == "Repeat":
                        repeat_samples[base_up] = (conc, raw_conc_val, s_name_raw)

        # --- PREPARE MAPPED SAMPLES FOR MATRIX ---
        mapped_initial = {}
        for base_up, (conc, raw_conc_val, s_name_raw) in initial_samples.items():
            parts = re.split(r'[-_]', s_name_raw)
            subj_str = parts[0] if len(parts) > 0 else s_name_raw
            period_str = parts[1] if len(parts) > 1 else ""
            tp_str = parts[2] if len(parts) > 2 else ""

            s_nums = re.findall(r'\d+', subj_str)
            subj_int = int(s_nums[-1]) if s_nums else 0
            period_clean = period_str.strip().upper()

            tp_nums = re.findall(r'\d+\.\d+|\d+', tp_str)
            tp_float = float(tp_nums[0]) if tp_nums else -1.0
            if "PRE" in tp_str.upper(): tp_float = -1.0

            mapped_initial[(subj_int, period_clean, tp_float)] = (conc, raw_conc_val, s_name_raw)
            
        mapped_repeat = {}
        for base_up, (conc, raw_conc_val, s_name_raw) in repeat_samples.items():
            parts = re.split(r'[-_]', s_name_raw)
            subj_str = parts[0] if len(parts) > 0 else s_name_raw
            period_str = parts[1] if len(parts) > 1 else ""
            tp_str = parts[2] if len(parts) > 2 else ""

            s_nums = re.findall(r'\d+', subj_str)
            subj_int = int(s_nums[-1]) if s_nums else 0
            period_clean = period_str.strip().upper()

            tp_nums = re.findall(r'\d+\.\d+|\d+', tp_str)
            tp_float = float(tp_nums[0]) if tp_nums else -1.0
            if "PRE" in tp_str.upper(): tp_float = -1.0

            mapped_repeat[(subj_int, period_clean, tp_float)] = (conc, raw_conc_val, s_name_raw)

        # --- MASTER DATA LIST FOR ISR SELECTION ---
        master_data_list = []

        # --- EXPORT TO EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws_rep = workbook.add_worksheet("Final Conc After Repeat")
            
            tnr = {'font_name': 'Times New Roman', 'font_size': 10}
            table_header_green = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white_float = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            data_white_pct = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            general_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter'})
            conc_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
            pct_fmt = workbook.add_format({**tnr, 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})

            ws_rep.set_column('A:B', 15, general_fmt)
            ws_rep.set_column('C:C', 15, general_fmt)
            ws_rep.set_column('D:D', 20, conc_fmt)
            ws_rep.set_column('E:E', 25, general_fmt)
            ws_rep.set_column('F:F', 20, conc_fmt)
            ws_rep.set_column('G:G', 20, pct_fmt)
            
            ws_rep.write(0, 0, "Molecule Name", data_white)
            ws_rep.merge_range(0, 1, 0, 2, mol_name.upper(), data_white)
            ws_rep.write(1, 0, "Study Number", data_white)
            ws_rep.merge_range(1, 1, 1, 2, study_num.upper(), data_white)
            ws_rep.write(2, 0, "CC range", data_white)
            cc_range_str = f"{lloq:.3f} {unit} to {uloq:.3f} {unit}"
            ws_rep.merge_range(2, 1, 2, 2, cc_range_str, data_white)
            
            rep_headers = ["Subject No.", "Period No.", "Timepoints", f"Final Conc.({unit})", "Sample Name", f"Initial Conc.({unit})", "% Difference"]
            for col_num, header in enumerate(rep_headers): 
                ws_rep.write(4, col_num, header, table_header_green)
            
            def get_display_val(conc_val, raw_str, current_lloq):
                raw_clean = str(raw_str).strip()
                if raw_clean.upper() == "M": return "M", True
                if raw_clean in ["< 0", "<0"]: return 0.0, False
                if pd.notna(conc_val) and isinstance(conc_val, (int, float)):
                    if pd.notna(current_lloq) and conc_val < current_lloq: return 0.0, False
                    return conc_val, False
                if raw_clean.upper() in ["NAN", "NONE", ""]: return "", True
                return raw_clean, True
            
            row_idx = 5
            
            if num_subj > 0 and periods_list and timepoints_list:
                for subj_idx in range(1, num_subj + 1):
                    subj_formatted = f"S{subj_idx:03d}"
                    for period_val in periods_list:
                        period_clean = period_val.strip().upper()
                        for tp_val in timepoints_list:
                            tp_nums = re.findall(r'\d+\.\d+|\d+', str(tp_val))
                            tp_float = float(tp_nums[0]) if tp_nums else -1.0
                            if "PRE" in str(tp_val).upper(): tp_float = -1.0
                            
                            timep_display = ""
                            if "PRE" in str(tp_val).upper(): timep_display = "PRE"
                            elif tp_nums: timep_display = f"{float(tp_nums[0]):05.2f}"
                            else: timep_display = str(tp_val)
                                
                            lookup_key = (subj_idx, period_clean, tp_float)
                            
                            if lookup_key in mapped_initial:
                                initial_conc, init_raw_str, init_s_name = mapped_initial[lookup_key]
                            else:
                                initial_conc, init_raw_str, init_s_name = np.nan, "M", "Missing"
                                
                            if lookup_key in mapped_repeat:
                                repeat_conc, rep_raw_str, rep_s_name = mapped_repeat[lookup_key]
                                has_repeat = True
                            else:
                                repeat_conc, rep_raw_str, rep_s_name = np.nan, "", ""
                                has_repeat = False

                            if has_repeat:
                                final_conc, final_raw_str, final_s_name = repeat_conc, rep_raw_str, rep_s_name
                            else:
                                final_conc, final_raw_str, final_s_name = initial_conc, init_raw_str, init_s_name

                            c_final_out, is_final_str = get_display_val(final_conc, final_raw_str, lloq)
                            c_final_fmt = data_white if is_final_str else data_white_float
                            c_init_out, is_init_str = get_display_val(initial_conc, init_raw_str, lloq)
                            c_init_fmt = data_white if is_init_str else data_white_float

                            pct_diff = ""
                            if has_repeat and not is_init_str and not is_final_str and isinstance(c_init_out, (int, float)) and c_init_out > 0:
                                pct_diff = ((c_final_out - c_init_out) / ((c_init_out + c_final_out) / 2)) * 100

                            ws_rep.write(row_idx, 0, subj_formatted, data_white)
                            ws_rep.write(row_idx, 1, period_val, data_white)
                            ws_rep.write_string(row_idx, 2, timep_display, data_white) 
                            ws_rep.write(row_idx, 3, c_final_out, c_final_fmt)
                            ws_rep.write(row_idx, 4, final_s_name, data_white)
                            
                            if has_repeat:
                                ws_rep.write(row_idx, 5, c_init_out, c_init_fmt)
                                ws_rep.write(row_idx, 6, pct_diff, data_white_pct if pct_diff != "" else data_white)
                            else:
                                ws_rep.write(row_idx, 5, "", data_white)
                                ws_rep.write(row_idx, 6, "", data_white)
                                
                            master_data_list.append({
                                'Subj': subj_formatted,
                                'Period': period_val,
                                'Timepoint_Float': tp_float,
                                'Timepoint_Display': timep_display,
                                'Sample Name': final_s_name,
                                'Final Conc': c_final_out if not is_final_str else np.nan,
                                'Is Repeated': has_repeat
                            })
                            row_idx += 1
            else:
                sorted_samples = sorted(initial_samples.items(), key=lambda x: natural_sort_key(x[1][2]))
                for base_up, (initial_conc, init_raw_str, init_s_name) in sorted_samples:
                    parts = re.split(r'[-_]', init_s_name)
                    subj = parts[0] if len(parts) > 0 else init_s_name
                    period = parts[1] if len(parts) > 1 else ""
                    timep = parts[2] if len(parts) > 2 else ""
                    if subj.startswith('S') and subj[1:].isdigit(): subj = f"S{int(subj[1:]):03d}"
                        
                    timep_display = ""
                    if timep:
                        if "PRE" in timep.upper(): timep_display = "PRE"
                        else:
                            t_nums = re.findall(r'\d+\.\d+|\d+', timep)
                            if t_nums: timep_display = f"{float(t_nums[0]):05.2f}"
                            else: timep_display = f"{timep}"
                    
                    repeat_tuple = repeat_samples.get(base_up)
                    has_repeat = repeat_tuple is not None
                    
                    if has_repeat:
                        repeat_conc, rep_raw_str, rep_s_name = repeat_tuple
                        final_conc, final_raw_str, final_s_name = repeat_conc, rep_raw_str, rep_s_name
                    else:
                        final_conc, final_raw_str, final_s_name = initial_conc, init_raw_str, init_s_name

                    c_final_out, is_final_str = get_display_val(final_conc, final_raw_str, lloq)
                    c_final_fmt = data_white if is_final_str else data_white_float
                    c_init_out, is_init_str = get_display_val(initial_conc, init_raw_str, lloq)
                    c_init_fmt = data_white if is_init_str else data_white_float

                    pct_diff = ""
                    if has_repeat and not is_init_str and not is_final_str and isinstance(c_init_out, (int, float)) and c_init_out > 0:
                        pct_diff = ((c_final_out - c_init_out) / ((c_init_out + c_final_out) / 2)) * 100

                    ws_rep.write(row_idx, 0, subj, data_white)
                    ws_rep.write(row_idx, 1, period, data_white)
                    ws_rep.write_string(row_idx, 2, timep_display, data_white) 
                    ws_rep.write(row_idx, 3, c_final_out, c_final_fmt)
                    ws_rep.write(row_idx, 4, final_s_name, data_white)
                    
                    if has_repeat:
                        ws_rep.write(row_idx, 5, c_init_out, c_init_fmt)
                        ws_rep.write(row_idx, 6, pct_diff, data_white_pct if pct_diff != "" else data_white)
                    else:
                        ws_rep.write(row_idx, 5, "", data_white)
                        ws_rep.write(row_idx, 6, "", data_white)
                        
                    t_nums = re.findall(r'\d+\.\d+|\d+', timep)
                    tp_float = float(t_nums[0]) if t_nums else -1.0
                    if "PRE" in timep.upper(): tp_float = -1.0
                        
                    master_data_list.append({
                        'Subj': subj,
                        'Period': period,
                        'Timepoint_Float': tp_float,
                        'Timepoint_Display': timep_display,
                        'Sample Name': final_s_name,
                        'Final Conc': c_final_out if not is_final_str else np.nan,
                        'Is Repeated': has_repeat
                    })
                    row_idx += 1

            # --- PROCESS ISR SELECTION LOGIC ---
            isr_selection_rows = []
            df_master = pd.DataFrame(master_data_list)
            
            if not df_master.empty:
                # Group by Subject and Period
                for (s, p), group in df_master.groupby(['Subj', 'Period']):
                    # Filter valid, quantifiable, non-repeated samples
                    valid_group = group[(group['Is Repeated'] == False) & (group['Final Conc'] > 0)].copy()
                    if valid_group.empty: continue
                    
                    valid_group = valid_group.sort_values(by='Timepoint_Float')
                    
                    tmax_idx = valid_group['Final Conc'].idxmax()
                    tmax_row = valid_group.loc[tmax_idx]
                    tmax_time = tmax_row['Timepoint_Float']
                    
                    # 1. Cmax Selection Phase
                    # Prioritize samples strictly < 95% of ULOQ
                    limit_95_uloq = 0.95 * uloq if pd.notna(uloq) else float('inf')
                    cmax_cands = valid_group[valid_group['Final Conc'] < limit_95_uloq].sort_values(by='Final Conc', ascending=False)
                    
                    # Fallback if no samples < 95% exist
                    if len(cmax_cands) < isr_cmax_count:
                        cmax_cands = valid_group.sort_values(by='Final Conc', ascending=False)
                        
                    selected_cmax = cmax_cands.head(isr_cmax_count)
                    selected_indices = list(selected_cmax.index)
                    
                    # 2. Elimination Selection Phase
                    elim_pool = valid_group[(valid_group['Timepoint_Float'] > tmax_time) & (~valid_group.index.isin(selected_indices))]
                    # Prioritize samples > LQC
                    elim_pref = elim_pool[elim_pool['Final Conc'] > lqc_val].sort_values(by='Timepoint_Float', ascending=False)
                    
                    if len(elim_pref) >= isr_elim_count:
                        selected_elim = elim_pref.head(isr_elim_count)
                    else:
                        elim_fallback = elim_pool.sort_values(by='Timepoint_Float', ascending=False)
                        if len(elim_fallback) >= isr_elim_count:
                            selected_elim = elim_fallback.head(isr_elim_count)
                        else:
                            rem = valid_group[~valid_group.index.isin(selected_indices)].sort_values(by='Timepoint_Float', ascending=False)
                            selected_elim = rem.head(isr_elim_count)

                    for idx in selected_cmax.index:
                        isr_selection_rows.append({
                            'Subject No.': s,
                            'Period No.': p,
                            'Timepoints': valid_group.loc[idx, 'Timepoint_Display'],
                            'Sample Name': valid_group.loc[idx, 'Sample Name'],
                            f'Final Conc.({unit})': valid_group.loc[idx, 'Final Conc'],
                            'Selection Phase': 'Near Cmax'
                        })
                    for idx in selected_elim.index:
                        isr_selection_rows.append({
                            'Subject No.': s,
                            'Period No.': p,
                            'Timepoints': valid_group.loc[idx, 'Timepoint_Display'],
                            'Sample Name': valid_group.loc[idx, 'Sample Name'],
                            f'Final Conc.({unit})': valid_group.loc[idx, 'Final Conc'],
                            'Selection Phase': 'Elimination Phase'
                        })

            # Write ISR Selection Sheet
            ws_sel = workbook.add_worksheet("ISR Selection")
            ws_sel.set_column('A:A', 10, general_fmt)
            ws_sel.set_column('B:D', 15, general_fmt)
            ws_sel.set_column('E:F', 20, general_fmt)
            ws_sel.set_column('G:G', 20, general_fmt)
            
            sel_headers = ["S.No.", "Subject No.", "Period No.", "Timepoints", "Sample Name", f"Final Conc.({unit})", "Selection Phase"]
            for col_num, header in enumerate(sel_headers): 
                ws_sel.write(0, col_num, header, table_header_green)
                
            if isr_selection_rows:
                df_isr = pd.DataFrame(isr_selection_rows)
                # Sort alphanumerically by Sample Name
                df_isr = df_isr.sort_values(by='Sample Name', key=lambda col: col.map(natural_sort_key)).reset_index(drop=True)
                
                s_no = 1
                for _, row in df_isr.iterrows():
                    ws_sel.write(s_no, 0, s_no, data_white)
                    ws_sel.write(s_no, 1, row['Subject No.'], data_white)
                    ws_sel.write(s_no, 2, row['Period No.'], data_white)
                    ws_sel.write_string(s_no, 3, row['Timepoints'], data_white)
                    ws_sel.write(s_no, 4, row['Sample Name'], data_white)
                    ws_sel.write(s_no, 5, row[f'Final Conc.({unit})'], data_white_float)
                    ws_sel.write(s_no, 6, row['Selection Phase'], data_white)
                    s_no += 1
            else:
                ws_sel.write(1, 0, "-", data_white)
                ws_sel.write(1, 1, "-", data_white)
                ws_sel.write(1, 2, "-", data_white)
                ws_sel.write(1, 3, "-", data_white)
                ws_sel.write(1, 4, "No eligible samples found.", data_white)
                ws_sel.write(1, 5, "", data_white)
                ws_sel.write(1, 6, "-", data_white)

        st.download_button(
            label="Download Final Concentration Tracker",
            data=output.getvalue(),
            file_name="Final_Concentration_Tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload BOTH the Setup file and the Main TSQ Export file(s).")
