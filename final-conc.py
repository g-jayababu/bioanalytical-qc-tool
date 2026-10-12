import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(layout="wide")
st.title("Simple Concentration & Repeat Tracker")

st.header("Upload Files")

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
    if uploaded_files:
        st.success("Files received! Processing concentration data...")
        
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
            
            # Write Headers
            rep_headers = ["Subject No.", "Period No.", "Timepoints", "Final Conc.(ng/mL)", "Sample Name", "Initial Conc.(ng/mL)", "% Difference"]
            for col_num, header in enumerate(rep_headers): 
                ws_rep.write(0, col_num, header, table_header_green)
            
            # Sort initial samples alphanumerically by their raw sample name
            sorted_samples = sorted(initial_samples.items(), key=lambda x: natural_sort_key(x[1][2]))
            
            row_idx = 1
            for base_up, (initial_conc, init_raw_str, init_s_name) in sorted_samples:
                
                # Parse Subject, Period, Timepoint from Sample Name if possible
                parts = re.split(r'[-_]', init_s_name)
                subj = parts[0] if len(parts) > 0 else init_s_name
                period = parts[1] if len(parts) > 1 else ""
                timep = parts[2] if len(parts) > 2 else ""
                if subj.startswith('S') and subj[1:].isdigit(): 
                    subj = f"S{int(subj[1:]):03d}"
                    
                # Apply 00.00 hrs padding to timepoints
                timep_display = ""
                if timep:
                    if "PRE" in timep.upper():
                        timep_display = "'PRE"
                    else:
                        t_nums = re.findall(r'\d+\.\d+|\d+', timep)
                        if t_nums:
                            timep_display = f"'{float(t_nums[0]):05.2f} hrs"
                        else:
                            timep_display = f"'{timep}"
                
                # Check for Repeat
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

                # Display Logic for Final Conc
                if pd.notna(final_conc) and isinstance(final_conc, (int, float)):
                    c_final_out = final_conc
                    c_final_fmt = data_white_float
                else:
                    c_final_out = final_raw_str
                    c_final_fmt = data_white
                    
                # Display Logic for Initial Conc
                if pd.notna(initial_conc) and isinstance(initial_conc, (int, float)):
                    c_init_out = initial_conc
                    c_init_fmt = data_white_float
                else:
                    c_init_out = init_raw_str
                    c_init_fmt = data_white

                # Calculate % Difference
                pct_diff = ""
                if has_repeat and pd.notna(initial_conc) and pd.notna(repeat_conc) and initial_conc > 0:
                    pct_diff = ((repeat_conc - initial_conc) / ((initial_conc + repeat_conc) / 2)) * 100

                # Write Row
                ws_rep.write(row_idx, 0, subj, data_white)
                ws_rep.write(row_idx, 1, period, data_white)
                ws_rep.write(row_idx, 2, timep_display, data_white)
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
        st.error("Please upload at least Export file.")
