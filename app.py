import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="Malaysia Automated Payroll System", page_icon="🇲🇾", layout="wide")

st.title("🇲🇾 Malaysia Web Payroll Engine")
st.markdown("""
**Drop & Calculate System**: Drop your monthly attendance Excel sheet. 
The engine auto-detects rows, extracts employee details, and calculates EPF, SOCSO, EIS, and PCB/Net Pay.
""")

# --- MALAYSIAN STATUTORY RULES (EPF, SOCSO, EIS, PCB) ---
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
        target_sheet = 'ATTENDANCE (2)' if 'ATTENDANCE (2)' in xls.sheet_names else xls.sheet_names[0]
        
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=7)
        
        # Auto-detect name column dynamically
        name_col = None
        for col in df.columns:
            col_str = str(col).strip().upper()
            if 'NAME' in col_str or 'STAFF' in col_str:
                name_col = col
                break
                
        if name_col is None:
            name_col = df.columns[2] if len(df.columns) > 2 else df.columns[1]

        df_clean = df.dropna(subset=[name_col]).copy()
        
        st.success(f"File loaded successfully! Detected **{len(df_clean)}** staff/crew rows.")
        
        default_base = st.sidebar.number_input("Default Base Salary for Crew (RM)", value=2200.0, step=100.0)
        manager_base = st.sidebar.number_input("Default Base Salary for Manager (RM)", value=4500.0, step=100.0)
        
        records = []
        for idx, row in df_clean.iterrows():
            name = str(row.get(name_col, 'Unknown'))
            if 'NAME' in name.upper() or 'NO.' in name.upper():
                continue
                
            ic_no = str(row.get('I/C', row.get('STAFF / CREW', '-')))
            position = str(row.get('POSITION', 'Staff'))
            
            total_days = row.get('TOTAL', 26)
            try:
                days_worked = float(total_days)
            except:
                days_worked = 26.0
                
            is_manager = "MANAGER" in position.upper()
            base = manager_base if is_manager else default_base
            
            earned_basic = round((base / 26.0) * min(days_worked, 26.0), 2)
            
            epf_ee, epf_er = compute_epf(earned_basic)
            socso_ee, socso_er, eis_ee, eis_er = compute_socso_eis(earned_basic)
            pcb = compute_pcb(earned_basic)
            
            total_deductions = round(epf_ee + socso_ee + eis_ee + pcb, 2)
            net_pay = round(earned_basic - total_deductions, 2)
            
            records.append({
                'Staff Name': name,
                'IC / ID': ic_no,
                'Position': position,
                'Days Worked': days_worked,
                'Base Salary (RM)': base,
                'Earned Basic (RM)': earned_basic,
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
        
        st.subheader("📊 Summary Payroll Breakdown")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Staff Count", f"{len(payroll_df)}")
        col2.metric("Total Gross Payout", f"RM {payroll_df['Earned Basic (RM)'].sum():,.2f}")
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
            file_name="MY_Payroll_Summary_September_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
