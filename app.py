import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Malaysia Automated Payroll System", page_icon="🇲🇾", layout="wide")

st.title("🇲🇾 Malaysia Web Payroll Engine")
st.markdown("""
**Drop & Calculate System**: Drop your monthly attendance & overtime Excel sheet. 
The engine auto-detects attendance and overtime sheets, merges OT hours, and calculates EPF, SOCSO, EIS, PCB, and Net Pay.
""")

# --- MALAYSIAN STATUTORY & OT RULES ---
def compute_epf(salary):
    ee_epf = round(salary * 0.11, 2)
    er_epf = round(salary * 0.13 if salary <= 5000 else salary * 0.12, 2)
    return ee_epf, er_epf

def compute_socso_eis(salary):
    capped_wages = min(salary, 5000.0)
    socso_ee = round(capped_wages * 0.005, 2)
    socso_er = round(capped_wages * 0.0175, 2)
    
    eis_ee = round(min(salary, 5000.0) * 0.002, 2)
    eis_er = round(min(salary, 5000.0) * 0.002, 2)
    return socso_ee, socso_er, eis_ee, eis_er

def compute_pcb(taxable_gross):
    if taxable_gross <= 3100:
        return 0.0
    return round((taxable_gross - 3100) * 0.03, 2)

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
        
        # 1. READ ATTENDANCE SHEET
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
        header_row = 8
        for idx, row in raw_df.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            row_str = " ".join(row_vals)
            if ('NAME' in row_str or 'STAFF' in row_str) and ('NO.' in row_str or 'POSITION' in row_str):
                header_row = idx
                break
        
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        
        name_col = next((c for c in df.columns if 'NAME' in c.upper()), df.columns[2] if len(df.columns) > 2 else None)
        pos_col = next((c for c in df.columns if 'POSITION' in c.upper()), None)
        ic_col = next((c for c in df.columns if 'I/C' in c.upper() or ('STAFF' in c.upper() and c != name_col)), None)
        total_col = next((c for c in df.columns if c.upper() == 'TOTAL'), None)
        
        df_clean = df.dropna(subset=[name_col]).copy()
        
        # 2. READ OVERTIME SHEET IF AVAILABLE
        ot_data = {}
        if 'OVERTIME' in xls.sheet_names:
            try:
                ot_df = pd.read_excel(uploaded_file, sheet_name='OVERTIME', header=5)
                ot_df.columns = [str(c).strip() for c in ot_df.columns]
                ot_name_col = next((c for c in ot_df.columns if 'NAME' in c.upper()), ot_df.columns[2])
                
                # Last 3 columns usually represent OT hours (Normal 1.5x, Rest Day 2.0x, PH 3.5x etc.)
                col_ot1 = ot_df.columns[-3]
                col_ot2 = ot_df.columns[-2]
                col_ot3 = ot_df.columns[-1]
                
                for _, ot_row in ot_df.iterrows():
                    ot_name = str(ot_row.get(ot_name_col, '')).strip().upper()
                    if ot_name and ot_name != 'NAN' and 'NAME' not in ot_name:
                        try:
                            h1 = float(ot_row.get(col_ot1, 0) or 0)
                            h2 = float(ot_row.get(col_ot2, 0) or 0)
                            h3 = float(ot_row.get(col_ot3, 0) or 0)
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
            
            if not name or name.lower() == 'nan' or 'NAME' in name.upper() or 'NO.' in name.upper() or len(name) < 2:
                continue
                
            position = str(row.get(pos_col, 'Staff')).strip() if pos_col else 'Staff'
            ic_no = str(row.get(ic_col, '-')).strip() if ic_col else '-'
            
            raw_days = row.get(total_col, 26) if total_col else 26
            try:
                days_worked = float(raw_days)
            except:
                days_worked = 26.0
                
            is_manager = "MANAGER" in position.upper()
            base = manager_base if is_manager else default_base
            
            # Earned Basic Pay based on attendance days (standard 26 days)
            earned_basic = round((base / 26.0) * min(days_worked, 26.0), 2)
            
            # Hourly rate calculation (Base / 26 days / 8 hours)
            hourly_rate = base / 26.0 / 8.0
            
            # Fetch Overtime Hours
            ot_info = ot_data.get(name.upper(), {'ot_1.5': 0.0, 'ot_2.0': 0.0, 'ot_3.0': 0.0})
            ot_pay = round((ot_info['ot_1.5'] * hourly_rate * 1.5) + 
                           (ot_info['ot_2.0'] * hourly_rate * 2.0) + 
                           (ot_info['ot_3.0'] * hourly_rate * 3.0), 2)
            
            gross_salary = round(earned_basic + ot_pay, 2)
            
            # Statutory Deductions
            epf_ee, epf_er = compute_epf(gross_salary)
            socso_ee, socso_er, eis_ee, eis_er = compute_socso_eis(gross_salary)
            pcb = compute_pcb(gross_salary)
            
            total_deductions = round(epf_ee + socso_ee + eis_ee + pcb, 2)
            net_pay = round(gross_salary - total_deductions, 2)
            
            records.append({
                'Staff Name': name,
                'IC / ID': ic_no if ic_no.lower() != 'nan' else '-',
                'Position': position if position.lower() != 'nan' else 'Staff',
                'Days Worked': days_worked,
                'Base Salary (RM)': base,
                'Earned Basic (RM)': earned_basic,
                'OT Hours (1.5x)': ot_info['ot_1.5'],
                'OT Hours (2.0x)': ot_info['ot_2.0'],
                'OT Hours (3.0x)': ot_info['ot_3.0'],
                'OT Pay (RM)': ot_pay,
                'Gross Salary (RM)': gross_salary,
                'EPF Employee (RM)': epf_ee,
                'SOCSO Employee (RM)': socso_ee,
                'EIS Employee (RM)': eis_ee,
                'PCB Tax (RM)': pcb,
                'Total Deductions (RM)': total_deductions,
                'Net Salary Payable (RM)': net_pay,
                'EPF Employer (RM)': epf_er,
                'SOCSO Employer (RM)': socso_er,
                'EIS Employer (RM)': eis_er
            })
            
        payroll_df = pd.DataFrame(records)
        
        if len(payroll_df) == 0:
            st.warning("No staff rows extracted. Please verify sheet selection.")
        else:
            st.success(f"Success! Accurately extracted and computed **{len(payroll_df)}** staff records with Overtime included from `{selected_sheet}`.")
            
            st.subheader("📊 Summary Payroll Breakdown")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Staff Count", f"{len(payroll_df)}")
            col2.metric("Total Gross Payout", f"RM {payroll_df['Gross Salary (RM)'].sum():,.2f}")
            col3.metric("Total Employee Deductions", f"RM {payroll_df['Total Deductions (RM)'].sum():,.2f}")
            col4.metric("Total Net Payable", f"RM {payroll_df['Net Salary Payable (RM)'].sum():,.2f}")
            
            st.subheader("📋 Calculated Payroll Data")
            st.dataframe(payroll_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                payroll_df.to_excel(writer, index=False, sheet_name='Payroll Summary')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Department Payroll Excel Report (.xlsx)",
                data=processed_data,
                file_name=f"Payroll_Summary_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
