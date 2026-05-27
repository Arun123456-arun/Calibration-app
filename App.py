
import streamlit as st
import pandas as pd
import datetime
from datetime import timedelta
import io

# Setup full width page layout
st.set_page_config(page_title="CCBSA Calibration Portal", layout="wide", page_icon="📊")

st.title("📊 CCBSA Calibration Master System — Pretoria")
st.markdown("Upload your master Excel sheet to automatically process tracking segments and generate live dashboard matrices.")

# --- STEP 1: EXCEL FILE UPLOADER ENGINE ---
st.subheader("📁 1. Load System Master Database")
uploaded_file = st.file_uploader("Drag and drop or select your 'Pretoria - Updated Calibration Template 24.03.2026.xlsx' file here:", type=["xlsx", "xls"])

# Define structural layout components
departments = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

if uploaded_file is not None:
    try:
        # Read the file. If there are multiple sheets, we default to the first one or let the user choose
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
       
        # Select target data sheet
        selected_sheet = st.selectbox("Select the sheet containing raw instrument records:", sheet_names, index=0)
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
       
        st.success("✅ Excel file loaded successfully!")
       
        # Map or confirm columns (Fallback to index names if custom layout headers match your layout)
        st.markdown("### 🔍 System Column Mapping Verification")
        col_list = raw_df.columns.tolist()
       
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            id_col = st.selectbox("Equipment ID / Serial Column", col_list, index=0 if len(col_list) > 0 else 0)
        with c2:
            name_col = st.selectbox("Equipment Name / Description", col_list, index=1 if len(col_list) > 1 else 0)
        with c3:
            dept_col = st.selectbox("Department Category Column", col_list, index=2 if len(col_list) > 2 else 0)
        with c4:
            due_col = st.selectbox("Next Calibration Due Date Column", col_list, index=3 if len(col_list) > 3 else 0)

        # Uniform processing dataset
        process_df = raw_df[[id_col, name_col, dept_col, due_col]].copy()
        process_df.columns = ["ID", "Name", "Department", "Due Date"]
       
        # Ensure correct date conversion types
        process_df["Due Date"] = pd.to_datetime(process_df["Due Date"]).dt.date
       
        # Initialize internal storage tracking parameters for active modifications
        if 'scrapped_items' not in st.session_state:
            st.session_state.scrapped_items = set()

        # Remove items flagged as scrapped or removed by the operator
        process_df = process_df[~process_df["ID"].isin(st.session_state.scrapped_items)]

        # --- STEP 2: MATHEMATICAL DEADLINE ENGINE ---
        today = datetime.date.today()

        def compute_dashboard_bucket(row):
            due = row["Due Date"]
            if pd.isnull(due):
                return "VALID"
           
            delta_days = (due - today).days
           
            if delta_days < 0:
                return "OVERDUE"
            elif delta_days <= 30:
                return "Due in Next 30 days"
            elif delta_days <= 91:  # 3 Months
                return "Due in Next 3 Months"
            elif delta_days <= 182: # 6 Months
                return "Due in Next 6 Months"
            else:
                return "VALID"

        process_df["Time Segment"] = process_df.apply(compute_dashboard_bucket, axis=1)

        # --- STEP 3: ACTION CONTROLLER ---
        st.markdown("---")
        st.subheader("⚙️ 2. Manage Equipment Status (Scrap / Remove Items)")
       
        active_ids = process_df["ID"].unique().tolist()
        if active_ids:
            col_sel, col_btn = st.columns([3, 1])
            with col_sel:
                target_id = st.selectbox("Search and select Equipment ID to scrap/remove from dashboard:", active_ids)
            with col_btn:
                st.markdown("<br>", unsafe_allowed_html=True)
                if st.button("🔴 Confirm Removal / Scrap Item", use_container_width=True):
                    st.session_state.scrapped_items.add(target_id)
                    st.toast(f"Item {target_id} moved to Out-of-Scope status.")
                    st.rerun()

        # --- STEP 4: LIVE CALIBRATION SUMMARY MATRIX ---
        st.markdown("---")
        st.subheader("📊 3. Live Calibration Dashboard Summary")
       
        matrix_records = []
        for d in departments:
            dept_mask = process_df[process_df["Department"] == d]
           
            overdue_cnt = len(dept_mask[dept_mask["Time Segment"] == "OVERDUE"])
            days30_cnt = len(dept_mask[dept_mask["Time Segment"] == "Due in Next 30 days"])
            months3_cnt = len(dept_mask[dept_mask["Time Segment"] == "Due in Next 3 Months"])
            months6_cnt = len(dept_mask[dept_mask["Time Segment"] == "Due in Next 6 Months"])
            valid_cnt = len(dept_mask[dept_mask["Time Segment"] == "VALID"])
            total_cnt = len(dept_mask)
           
            matrix_records.append({
                "Departments": d,
                "OVERDUE": overdue_cnt,
                "Due in Next 30 days": days30_cnt,
                "Due in Next 3 Months": months3_cnt,
                "Due in Next 6 Months": months6_cnt,
                "VALID": valid_cnt,
                "TOTAL": total_cnt
            })
           
        summary_matrix_df = pd.DataFrame(matrix_records)
       
        # Display the live matrix view grid matching your report design
        st.dataframe(summary_matrix_df, use_container_width=True, hide_index=True)
       
        # Display metrics metrics cards
        c_ov, c_30, c_3m, c_6m = st.columns(4)
        c_ov.metric("Total OVERDUE Items", summary_matrix_df["OVERDUE"].sum(), delta="- Action Needed", delta_color="inverse")
        c_30.metric("Due in 30 Days", summary_matrix_df["Due in Next 30 days"].sum())
        c_3m.metric("Due in 3 Months", summary_matrix_df["Due in Next 3 Months"].sum())
        c_6m.metric("Total Tracked Assets", summary_matrix_df["TOTAL"].sum())

        # --- STEP 5: DETAIL VIEWER ---
        st.markdown("---")
        st.subheader("📋 4. Asset Register Detailed Rows")
        st.dataframe(process_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Error parsing file: {str(e)}. Please check that your file structural format is correct.")
else:
    st.info("💡 Please upload your calibration excel workbook above to initialize the reporting program dashboards.")


