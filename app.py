import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Malaysia Automated Payroll System", page_icon="🇲🇾", layout="wide")

st.title("🇲🇾 Malaysia Web Payroll Engine")
st.markdown("""
**Drop & Calculate System**: Drop your monthly attendance & overtime Excel sheet. 
The engine extracts all staff members, applies statutory deductions on Basic Salary (including SOCSO 24 Jam Perlindungan 0.75%), and adds Overtime Pay to generate Net Take-Home Pay.
""")

# --- MALAYSIAN STATUTORY RULES ---
def compute_epf(salary):
    ee_epf = round(salary * 0.11, 2)
    er_epf = round(salary * 0.13 if salary <= 5000 else salary * 0.12, 2)
    return ee_epf, er_epf

def compute_socso_eis(salary):
    capped_wages = min(salary, 5000.0)
    socso_ee = round(capped_wages * 0.005, 2)
    socso_er = round(capped_wages * 0.0175, 2)
    
    # SOCSO 24 Jam Perlindungan (0.75% deducted from Basic)
    socso_24h_ee = round(salary * 0.0075, 2)
    
    eis_ee = round(min(salary, 5000.0) * 0.002, 2)
    eis_er = round(min(salary, 5000.0) * 0.002, 2)
    return socso_ee, socso_er, socso_24h_ee, eis_ee, eis_er

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
            
            earned_basic = round((base / 26.0) * min(days_worked, 26.0), 2)
            hourly_rate = base / 26.0 / 8.0
            
            # Match Overtime hours by staff name
            ot_info = ot_data.get(name.upper(), {'ot_1.5': 0.0, 'ot_2.0': 0.0, 'ot_3.0': 0.0})
            ot_pay = round((ot_info['ot_1.5'] * hourly_rate * 1.5) + 
                           (ot_info['ot_2.0'] * hourly_rate * 2.0) + 
                           (ot_info['ot_3.0'] * hourly_rate * 3.0), 2)
            
            # Deductions on Earned Basic Salary
            epf_ee, epf_er = compute_epf(earned_basic)
            socso_ee, socso_er, socso_24h_ee, eis_ee, eis_er = compute_socso_eis(earned_basic)
            pcb = compute_pcb(earned_basic)
            
            total_deductions = round(epf_ee + socso_ee + socso_24h_ee + eis_ee + pcb, 2)
            
            # Gross Basic After Deductions
            basic_after_deductions = round(earned_basic - total_deductions, 2)
            
            # Net Pay = Gross Salary After Deductions + OT Pay
            net_pay = round(basic_after_deductions + ot_pay, 2)
            
            records.append({
                'Staff Name': name,
                'IC / ID': ic_no if ic_no.lower() != 'nan' else '-',
                'Position': position if position.lower() != 'nan' else 'Staff',
                'Days Worked': days_worked,
                'Base Salary (RM)': base,
                'Earned Basic (RM)': earned_basic,
                'EPF Employee (RM)': epf_ee,
                'SOCSO Employee (RM)': socso_ee,
                'SOCSO 24 Jam perlindungan (RM)': socso_24h_ee,
                'EIS Employee (RM)': eis_ee,
                'PCB Tax (RM)': pcb,
                'Total Deductions (RM)': total_deductions,
                'Basic Salary After Deductions (RM)': basic_after_deductions,
                'OT Hours (1.5x)': ot_info['ot_1.5'],
                'OT Hours (2.0x)': ot_info['ot_2.0'],
                'OT Hours (3.0x)': ot_info['ot_3.0'],
                'Total OT Pay (RM)': ot_pay,
                'Net Salary Payable (RM)': net_pay,
                'EPF Employer (RM)': epf_er,
                'SOCSO Employer (RM)': socso_er,
                'EIS Employer (RM)': eis_er
            })
            
        payroll_df = pd.DataFrame(records)
        
        if len(payroll_df) == 0:
            st.warning("No staff rows extracted. Please verify sheet selection.")
        else:
            st.success(f"Success! Accurately extracted **{len(payroll_df)}** staff records with SOCSO 24 Jam & OT Pay included from `{selected_sheet}`.")
            
            st.subheader("📊 Summary Payroll Breakdown")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Staff Count", f"{len(payroll_df)}")
            col2.metric("Total Basic Earned", f"RM {payroll_df['Earned Basic (RM)'].sum():,.2f}")
            col3.metric("Total OT Payout", f"RM {payroll_df['Total OT Pay (RM)'].sum():,.2f}")
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
