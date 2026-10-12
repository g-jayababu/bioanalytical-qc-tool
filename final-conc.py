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
                
                # Only capture UNKNOWN samples (exclude Blanks, AQSTDs, CCs, QCs)
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

        # --- EXPORT TO EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws_rep = workbook.add_worksheet("Final Conc After Repeat")
            
            # Formatting
            tnr = {'font_name': 'Times New Roman', 'font_size': 10}
            table_header_green = workbook.add_format({**tnr, 'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            data_white_float = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.000'})
            data_white_pct = workbook.add_format({**tnr, 'bg_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            general_fmt = workbook.add_format({**tnr, 'align': 'center', 'valign': 'vcenter'})
            conc_fmt = workbook.add_format({**tnr, 'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
            pct_fmt = workbook.add_format({**tnr, 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})

            # Set Column Widths
            ws_rep.set_column('A:B', 15, general_fmt)
            ws_rep.set_column('C:C', 15, general_fmt)
            ws_rep.set_column('D:D', 20, conc_fmt)
            ws_rep.set_column('E:E', 25, general_fmt)
            ws_rep.set_column('F:F', 20, conc_fmt)
            ws_rep.set_column('G:G', 20, pct_fmt)
            
            # Write Setup Block (Molecule, Study, Range)
            ws_rep.write(0, 0, "Molecule Name", data_white)
            ws_rep.merge_range(0, 1, 0, 2, mol_name.upper(), data_white)
            
            ws_rep.write(1, 0, "Study Number", data_white)
            ws_rep.merge_range(1, 1, 1, 2, study_num.upper(), data_white)
            
            ws_rep.write(2, 0, "CC range", data_white)
            cc_range_str = f"{lloq:.3f} {unit} to {uloq:.3f} {unit}"
            ws_rep.merge_range(2, 1, 2, 2, cc_range_str, data_white)
            
            # Write Main Table Headers offset to row 4
            rep_headers = ["Subject No.", "Period No.", "Timepoints", f"Final Conc.({unit})", "Sample Name", f"Initial Conc.({unit})", "% Difference"]
            for col_num, header in enumerate(rep_headers): 
                ws_rep.write(4, col_num, header, table_header_green)
            
            # Helper to map display values with 0.000 / Missing Logic
            def get_display_val(conc_val, raw_str, current_lloq):
                raw_clean = str(raw_str).strip()
                
                if raw_clean.upper() == "M":
                    return "M", True
                
                # Manual overrides for < 0
                if raw_clean in ["< 0", "<0"]:
                    return 0.0, False  # Returns numeric 0
                    
                # Numeric evaluations against LLOQ
                if pd.notna(conc_val) and isinstance(conc_val, (int, float)):
                    if pd.notna(current_lloq) and conc_val < current_lloq:
                        return 0.0, False
                    return conc_val, False
                
                # Empty or missing states
                if raw_clean.upper() in ["NAN", "NONE", ""]:
                    return "", True
                
                # Any other generic strings from instrument
                return raw_clean, True
            
            row_idx = 5
            
            # EXHAUSTIVE MATRIX BASED ON SETUP FILE
            if num_subj > 0 and periods_list and timepoints_list:
                for subj_idx in range(1, num_subj + 1):
                    subj_formatted = f"S{subj_idx:03d}"
                    for period_val in periods_list:
                        period_clean = period_val.strip().upper()
                        for tp_val in timepoints_list:
                            
                            # Standardize timepoint for exact lookup mapping
                            tp_nums = re.findall(r'\d+\.\d+|\d+', str(tp_val))
                            tp_float = float(tp_nums[0]) if tp_nums else -1.0
                            if "PRE" in str(tp_val).upper(): tp_float = -1.0
                            
                            # Format for visual output display
                            timep_display = ""
                            if "PRE" in str(tp_val).upper():
                                timep_display = "PRE"
                            elif tp_nums:
                                timep_display = f"{float(tp_nums[0]):05.2f}"
                            else:
                                timep_display = str(tp_val)
                                
                            lookup_key = (subj_idx, period_clean, tp_float)
                            
                            # Locate Initial Sample
                            if lookup_key in mapped_initial:
                                initial_conc, init_raw_str, init_s_name = mapped_initial[lookup_key]
                            else:
                                initial_conc, init_raw_str, init_s_name = np.nan, "M", "Missing"
                                
                            # Locate Repeat Sample
                            if lookup_key in mapped_repeat:
                                repeat_conc, rep_raw_str, rep_s_name = mapped_repeat[lookup_key]
                                has_repeat = True
                            else:
                                repeat_conc, rep_raw_str, rep_s_name = np.nan, "", ""
                                has_repeat = False

                            # Process which value becomes the 'Final' value
                            if has_repeat:
                                final_conc = repeat_conc
                                final_raw_str = rep_raw_str
                                final_s_name = rep_s_name
                            else:
                                final_conc = initial_conc
                                final_raw_str = init_raw_str
                                final_s_name = init_s_name

                            # Formatting values through the logic gate
                            c_final_out, is_final_str = get_display_val(final_conc, final_raw_str, lloq)
                            c_final_fmt = data_white if is_final_str else data_white_float
                                
                            c_init_out, is_init_str = get_display_val(initial_conc, init_raw_str, lloq)
                            c_init_fmt = data_white if is_init_str else data_white_float

                            # Calculate % Difference (Suppressed if text/missing or 0.0)
                            pct_diff = ""
                            if has_repeat and not is_init_str and not is_final_str and isinstance(c_init_out, (int, float)) and c_init_out > 0:
                                pct_diff = ((c_final_out - c_init_out) / ((c_init_out + c_final_out) / 2)) * 100

                            # Write Row directly to matrix
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
                                
                            row_idx += 1
            else:
                # Fallback to Unstructured Output if Setup format failed
                sorted_samples = sorted(initial_samples.items(), key=lambda x: natural_sort_key(x[1][2]))
                for base_up, (initial_conc, init_raw_str, init_s_name) in sorted_samples:
                    parts = re.split(r'[-_]', init_s_name)
                    subj = parts[0] if len(parts) > 0 else init_s_name
                    period = parts[1] if len(parts) > 1 else ""
                    timep = parts[2] if len(parts) > 2 else ""
                    if subj.startswith('S') and subj[1:].isdigit(): 
                        subj = f"S{int(subj[1:]):03d}"
                        
                    timep_display = ""
                    if timep:
                        if "PRE" in timep.upper():
                            timep_display = "PRE"
                        else:
                            t_nums = re.findall(r'\d+\.\d+|\d+', timep)
                            if t_nums:
                                timep_display = f"{float(t_nums[0]):05.2f}"
                            else:
                                timep_display = f"{timep}"
                    
                    repeat_tuple = repeat_samples.get(base_up)
                    has_repeat = repeat_tuple is not None
                    
                    if has_repeat:
                        repeat_conc, rep_raw_str, rep_s_name = repeat_tuple
                        final_conc = repeat_conc
                        final_raw_str = rep_raw_str
                        final_s_name = rep_s_name
                    else:
                        repeat_conc = np.nan
                        final_conc = initial_conc
                        final_raw_str = init_raw_str
                        final_s_name = init_s_name

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
                        
                    row_idx += 1

        st.download_button(
            label="Download Final Concentration Tracker",
            data=output.getvalue(),
            file_name="Final_Concentration_Tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload BOTH the Setup file and the Main TSQ Export file(s).")
