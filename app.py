import streamlit as st
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(layout="wide")
st.title("Bioanalytical Auto-QC Dashboard")

# --- 1. SETUP VARIABLES (Replacing the 'Setup' tab in VBA) ---
st.sidebar.header("Setup Variables")
mol_name = st.sidebar.text_input("Molecule Name", "Analyte_X")
study_num = st.sidebar.text_input("Study Number", "S-25-911")
num_subj = st.sidebar.number_input("Number of Subjects", min_value=1, value=24)
periods = st.sidebar.text_input("Periods (comma separated)", "1,2")
timepoints = st.sidebar.text_input("Timepoints (comma separated)", "0.00,0.50,1.00,2.00,4.00,8.00,12.00,24.00")
lloq = st.sidebar.number_input("LLOQ (ng/mL)", value=1.00)
uloq = st.sidebar.number_input("ULOQ (ng/mL)", value=100.00)

st.header("1. Upload Instrument Data")
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
    if uploaded_files:
        st.success("Files received! Processing data...")
        
        # --- GLOBAL DICTIONARIES (Mirroring global variables in VBA) ---
        global_ccqc_ids = {}
        global_samples = {}
        summary_log = []
        
        # Dictionary to hold the final DataFrames for each tab
        excel_tabs = {}
        
        for file in uploaded_files:
            file.seek(0)
            df = load_data(file)
            
            if df is None:
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": "Could not read file."})
                continue
                
            # Clean data
            df = df.dropna(how='all')
            
            # Identify critical columns (similar to VBA find loop)
            # Make columns lowercase for easier searching
            df.columns = [str(c).lower().strip() for c in df.columns]
            
            name_col = next((c for c in df.columns if 'sample name' in c or 'sample' in c), None)
            type_col = next((c for c in df.columns if 'sample type' in c or 'type' in c), None)
            is_area_col = next((c for c in df.columns if 'is area' in c or 'istd' in c), None)
            conc_col = next((c for c in df.columns if 'calc' in c and 'conc' in c), None)
            acc_col = next((c for c in df.columns if 'accur' in c), None)
            
            if not all([name_col, type_col, is_area_col, conc_col, acc_col]):
                summary_log.append({"File Name": file.name, "Status": "FAILED", "Error Details": "Missing critical columns (Sample Name, Type, IS Area, Calc Conc, Accuracy)."})
                continue
                
            # --- START PROCESSING (Mirroring ProcessSingleDashboard) ---
            file_mismatches = []
            file_dupes = []
            file_sample_dupes = []
            
            # Calculate Average IS Area for CC and QC
            cc_qc_df = df[df[type_col].str.contains('Standard|QC', case=False, na=False)].copy()
            
            # Convert IS Area to numeric, forcing errors to NaN
            cc_qc_df[is_area_col] = pd.to_numeric(cc_qc_df[is_area_col], errors='coerce')
            avg_is_area = cc_qc_df[is_area_col].mean()
            
            lower_is_limit = avg_is_area * 0.4
            upper_is_limit = avg_is_area * 1.6
            
            # Check for Duplicate CC/QC IDs and Out of Range IS
            out_of_range_is = []
            above_uloq = []
            
            for index, row in df.iterrows():
                s_name = str(row[name_col]).strip().upper()
                s_type = str(row[type_col]).strip()
                s_conc = pd.to_numeric(row[conc_col], errors='coerce')
                s_is_area = pd.to_numeric(row[is_area_col], errors='coerce')
                
                # Check IS variance
                if pd.notna(s_is_area) and ('Standard' in s_type or 'QC' in s_type or 'Unknown' in s_type):
                    if s_is_area < lower_is_limit or s_is_area > upper_is_limit:
                        # Exclude mobile phase/blanks from this specific check if needed
                        if 'BLK' not in s_name and 'AQSTD' not in s_name:
                             out_of_range_is.append({"Sample Name": s_name, "Concentration": s_conc, "IS Area": s_is_area})
                
                # Check ULOQ for unknowns
                if 'Unknown' in s_type and pd.notna(s_conc):
                    if s_conc > uloq:
                         above_uloq.append({"Sample Name": s_name, "Concentration": s_conc})
                
                # Duplicate Checks
                if 'Standard' in s_type or 'QC' in s_type:
                    if s_name in global_ccqc_ids:
                        file_dupes.append(f"Duplicate CC/QC ID: {s_name} (Found in {global_ccqc_ids[s_name]['file']})")
                    else:
                        global_ccqc_ids[s_name] = {'file': file.name, 'conc': s_conc}
                
                elif 'Unknown' in s_type and 'AQSTD' not in s_name:
                    if s_name in global_samples:
                        file_sample_dupes.append(f"Duplicate Sample Name: {s_name} (Found in {global_samples[s_name]['file']})")
                    else:
                        global_samples[s_name] = {'file': file.name, 'conc': s_conc}

            # Build the specific tab DataFrame for this file (Mirroring wsDest)
            # For simplicity in this translation, we will build a clean version of the data
            clean_df = df[[name_col, type_col, is_area_col, conc_col, acc_col]].copy()
            clean_df.columns = ["Sample Name", "Sample Type", "IS Area", "Calc. Conc.", "% Accuracy"]
            
            excel_tabs[file.name[:31]] = clean_df # Excel limits sheet names to 31 chars
            
            # Log the success
            summary_log.append({
                "File Name": file.name, 
                "Status": "SUCCESS", 
                "Error Details": "Dashboard generated.",
                "Duplicate CC & QC ID Check": ", ".join(file_dupes) if file_dupes else "All CC and QC IDs are unique.",
                "Duplicate Sample Name Check": ", ".join(file_sample_dupes) if file_sample_dupes else "All CC and QC sample names are unique.",
                "Out of Range IS Count": len(out_of_range_is),
                "Above ULOQ Count": len(above_uloq)
            })
            
        # --- BUILD FINAL SUMMARY TABS ---
        
        # 1. Summary Log
        excel_tabs["Summary"] = pd.DataFrame(summary_log)
        
        # 2. CC Tab (Aggregated across files)
        # Note: True VBA equivalent requires extracting nominals and doing complex pivots. 
        # This provides a structured list of all Calibrators found.
        cc_data = []
        for s_name, data in global_ccqc_ids.items():
            if 'CC' in s_name or 'STD' in s_name or 'STANDARD' in s_name:
                cc_data.append({"CCID": s_name, "Source File": data['file'], "Concentration": data['conc']})
        if cc_data:
            excel_tabs["CC"] = pd.DataFrame(cc_data)

        # 3. QC Tab (Aggregated)
        qc_data = []
        for s_name, data in global_ccqc_ids.items():
            if 'QC' in s_name:
                qc_data.append({"QC ID": s_name, "Source File": data['file'], "Concentration": data['conc']})
        if qc_data:
            excel_tabs["QC"] = pd.DataFrame(qc_data)

        # --- GENERATE EXCEL FILE FOR DOWNLOAD ---
        output = io.BytesIO()
        # Use XlsxWriter to format the output
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Reorder tabs: Summary first, then individual files, then CC/QC
            if "Summary" in excel_tabs:
                excel_tabs["Summary"].to_excel(writer, sheet_name="Summary", index=False)
                del excel_tabs["Summary"]
                
            for sheet_name, df_tab in excel_tabs.items():
                if sheet_name not in ["CC", "QC"]:
                     df_tab.to_excel(writer, sheet_name=sheet_name, index=False)
                     
            if "CC" in excel_tabs: excel_tabs["CC"].to_excel(writer, sheet_name="CC", index=False)
            if "QC" in excel_tabs: excel_tabs["QC"].to_excel(writer, sheet_name="QC", index=False)
            
            # Formatting (e.g., coloring the Summary tab headers)
            workbook  = writer.book
            header_format = workbook.add_format({'bold': True, 'bg_color': '#C6E0B4', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                # Apply header format to the first row
                for col_num, value in enumerate(writer.sheets[sheet_name].table[0]):
                    worksheet.write(0, col_num, value, header_format)
                worksheet.set_column('A:F', 25) # Widen columns

        st.success("Dashboard successfully generated!")
        st.download_button(
            label="Download Master Excel Dashboard",
            data=output.getvalue(),
            file_name="Master_Bioanalytical_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload at least one TSQ export file.")
