import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Malaysia Automated Payroll System", page_icon="🇲🇾", layout="wide")

st.title("🇲🇾 Malaysia Web Payroll Engine")
st.markdown("""
**Drop & Calculate System**: Drop your monthly attendance & overtime Excel sheet. 
The engine extracts all staff members, merges OT hours, and generates an **auto-calculating Excel report** with live formulas for custom salary packages.
""")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Drop Attendance Excel File Here", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        all_sheets = xls.sheet_names
        filtered_sheets = [s for s in all_sheets if 'AGENCY' not in s.upper() and 'SHEET1' not in s.upper()]
        if not filtered_sheets:
            filtered_sheets = all_sheets
            
        default_index = 0
        for i, s in enumerate(filtered_sheets):
            if 'ATTENDANCE (2)' in s.upper():
                default_index = i
                break
                
        selected_sheet = st.sidebar.selectbox("Select Attendance Sheet", filtered_sheets, index=default_index)
        
        # Read selected sheet with header row 8
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=8)
        df.columns = [str(c).strip() for c in df.columns]
        
        name_col = 'STAFF / CREW NAME' if 'STAFF / CREW NAME' in df.columns else next((c for c in df.columns if 'NAME' in c.upper() and c != 'RESIGN/TERMINATE'), df.columns[2])
        pos_col = 'POSITION' if 'POSITION' in df.columns else None
        ic_col = 'STAFF / CREW' if 'STAFF / CREW' in df.columns else None
        total_col = 'TOTAL' if 'TOTAL' in df.columns else None
        
        df_clean = df.dropna(subset=[name_col]).copy()
        
        # READ OVERTIME SHEET FOR OT HOURS
        ot_data = {}
        if 'OVERTIME' in xls.sheet_names:
            try:
                ot_df = pd.read_excel(uploaded_file, sheet_name='OVERTIME', header=5)
                ot_df.columns = [str(c).strip() for c in ot_df.columns]
                ot_name_col = 'NAME' if 'NAME' in ot_df.columns else ot_df.columns[2]
                
                for _, ot_row in ot_df.iterrows():
                    ot_name = str(ot_row.get(ot_name_col, '')).strip().upper()
                    if ot_name and ot_name != 'NAN' and 'NAME' not in ot_name:
                        try:
                            val1 = ot_row.iloc[-3]
                            val2 = ot_row.iloc[-2]
                            val3 = ot_row.iloc[-1]
                            
                            h1 = float(val1) if pd.notna(val1) else 0.0
                            h2 = float(val2) if pd.notna(val2) else 0.0
                            h3 = float(val3) if pd.notna(val3) else 0.0
                            ot_data[ot_name] = {'ot_1.5': h1, 'ot_2.0': h2, 'ot_3.0': h3}
                        except:
                            pass
            except Exception as e:
                pass

        default_base = st.sidebar.number_input("Default Base Salary for Crew (RM)", value=2200.0, step=100.0)
        manager_base = st.sidebar.number_input("Default Base Salary for Manager (RM)", value=4500.0, step=100.0)
        
        records = []
        for idx, row in df_clean.iterrows():
            name = str(row.get(name_col, '')).strip()
            
            if not name or name.lower() == 'nan' or 'NAME' in name.upper() or 'NO.' in name.upper() or name.isdigit() or len(name) < 2:
                continue
                
            position = str(row.get(pos_col, 'Staff')).strip() if pos_col and pos_col in df.columns else 'Staff'
            ic_no = str(row.get(ic_col, '-')).strip() if ic_col and ic_col in df.columns else '-'
            
            raw_days = row.get(total_col, 26) if total_col and total_col in df.columns else 26
            try:
                days_worked = float(raw_days)
            except:
                days_worked = 26.0
                
            is_manager = "MANAGER" in position.upper()
            base = manager_base if is_manager else default_base
            
            # Match Overtime hours by staff name
            ot_info = ot_data.get(name.upper(), {'ot_1.5': 0.0, 'ot_2.0': 0.0, 'ot_3.0': 0.0})
            
            records.append({
                'Staff Name': name,
                'IC / ID': ic_no if ic_no.lower() != 'nan' else '-',
                'Position': position if position.lower() != 'nan' else 'Staff',
                'Days Worked': days_worked,
                'Base Salary (RM)': base,
                'OT Hours (1.5x)': ot_info['ot_1.5'],
                'OT Hours (2.0x)': ot_info['ot_2.0'],
                'OT Hours (3.0x)': ot_info['ot_3.0']
            })
            
        base_df = pd.DataFrame(records)
        
        if len(base_df) == 0:
            st.warning("No staff rows extracted. Please verify sheet selection.")
        else:
            st.success(f"Success! Extracted **{len(base_df)}** staff records with live Excel formulas enabled from `{selected_sheet}`.")
            
            # Build workbook with live Excel formulas for easy editing
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Payroll Summary"
            
            headers = [
                'Staff Name', 'IC / ID', 'Position', 'Days Worked', 'Base Salary (RM)', 
                'Earned Basic (RM)', 'EPF Employee (RM)', 'SOCSO Employee (RM)', 
                'SOCSO 24 Jam perlindungan (RM)', 'EIS Employee (RM)', 'PCB Tax (RM)', 
                'Total Deductions (RM)', 'Basic Salary After Deductions (RM)', 
                'OT Hours (1.5x)', 'OT Hours (2.0x)', 'OT Hours (3.0x)', 'Total OT Pay (RM)', 'Net Salary Payable (RM)'
            ]
            ws.append(headers)
            
            for i, r in base_df.iterrows():
                row_num = i + 2  # Excel row index starts at 2
                # Column letters:
                # D = Days Worked, E = Base Salary, F = Earned Basic
                # G = EPF, H = SOCSO, I = SOCSO 24H, J = EIS, K = PCB
                # L = Total Deductions, M = Basic After Deductions
                # N = OT 1.5x, O = OT 2.0x, P = OT 3.0x, Q = Total OT Pay, R = Net Salary
                
                row_formulas = [
                    r['Staff Name'],
                    r['IC / ID'],
                    r['Position'],
                    r['Days Worked'],
                    r['Base Salary (RM)'],
                    f'=IF(D{row_num}>26, (E{row_num}/26)*26, (E{row_num}/26)*D{row_num})',  # Earned Basic
                    f'=F{row_num}*0.11',  # EPF
                    f'=MIN(F{row_num}, 5000)*0.005',  # SOCSO
                    f'=F{row_num}*0.0075',  # SOCSO 24H
                    f'=MIN(F{row_num}, 5000)*0.002',  # EIS
                    f'=IF(F{row_num}>3100, (F{row_num}-3100)*0.03, 0)',  # PCB
                    f'=SUM(G{row_num}:K{row_num})',  # Total Deductions
                    f'=F{row_num}-L{row_num}',  # Basic After Deductions
                    r['OT Hours (1.5x)'],
                    r['OT Hours (2.0x)'],
                    r['OT Hours (3.0x)'],
                    f'=(N{row_num}*(E{row_num}/26/8)*1.5) + (O{row_num}*(E{row_num}/26/8)*2.0) + (P{row_num}*(E{row_num}/26/8)*3.0)',  # Total OT Pay
                    f'=M{row_num}+Q{row_num}'  # Net Salary Payable
                ]
                ws.append(row_formulas)
                
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Auto-Calculating Editable Payroll Report (.xlsx)",
                data=processed_data,
                file_name=f"Editable_Payroll_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
