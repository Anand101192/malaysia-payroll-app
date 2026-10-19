# Read sheet with header at row 7 (index 7)
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=7)
        
        # Auto-detect name column dynamically
        name_col = None
        for col in df.columns:
            col_str = str(col).strip().upper()
            if 'NAME' in col_str or 'STAFF' in col_str:
                name_col = col
                break
                
        if name_col is None:
            # Fallback to column index 2 if header search fails
            name_col = df.columns[2] if len(df.columns) > 2 else df.columns[1]

        # Filter valid employee rows
        df_clean = df.dropna(subset=[name_col]).copy()
        
        st.success(f"File loaded successfully! Detected **{len(df_clean)}** staff/crew rows using column `{name_col}`.")
        
        # Default Base Salary settings
        default_base = st.sidebar.number_input("Default Base Salary for Crew (RM)", value=2200.0, step=100.0)
        manager_base = st.sidebar.number_input("Default Base Salary for Manager (RM)", value=4500.0, step=100.0)
        
        records = []
        for idx, row in df_clean.iterrows():
            name = str(row.get(name_col, 'Unknown'))
            # Skip if header row values leaked into data
            if 'NAME' in name.upper() or 'NO.' in name.upper():
                continue
                
            ic_no = str(row.get('I/C', row.get('STAFF / CREW', '-')))
            position = str(row.get('POSITION', 'Staff'))
            
            # Days present extracted from "TOTAL" column
            total_days = row.get('TOTAL', 26)
            try:
                days_worked = float(total_days)
            except:
                days_worked = 26.0
                
            is_manager = "MANAGER" in position.upper()
            base = manager_base if is_manager else default_base
            
            earned_basic = round((base / 26.0) * min(days_worked, 26.0), 2)
            
            # Calculations
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
