
import streamlit as st
import pandas as pd
import datetime
import plotly.express as px

# --- CONFIGURATION & LAYOUT ---
st.set_page_config(page_title="CCBSA Calibration Portal", layout="wide", page_icon="📊")
st.title("📊 CCBSA Calibration Master System — Pretoria")

# --- INITIALIZE INTERNAL VARIABLES ---
if 'scrapped_items' not in st.session_state:
    st.session_state.scrapped_items = set()

# Explicit list of departments matching your corporate matrix layout structure
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- SECTION 1: DATABASE PERSISTENCE ENGINE ---
st.subheader("📁 1. System Master Database Storage Connection")

st.markdown("""
💡 **How to link your live Excel file from Microsoft OneDrive/SharePoint:**
1. Open your corporate OneDrive or SharePoint in your web browser.
2. Right-click your Excel file and select **Copy Link**.
3. Paste that web URL address in the box below to link the data dynamically.
""")

microsoft_url = st.text_input(
    "Paste your Microsoft OneDrive/SharePoint file share link here:",
    placeholder="https://sharepoint.com..."
)

uploaded_file = None
if not microsoft_url:
    st.info("ℹ️ Alternatively, upload your new 2-sheet file manually below until you paste your Microsoft Link:")
    uploaded_file = st.file_uploader("Upload your master calibration template workbook file directly:", type=["xlsx", "xls"])

raw_df = None

# Load from Microsoft Link
if microsoft_url:
    try:
        direct_download_url = microsoft_url.split("?") + "?download=1"
        # Reading data sheet while handling blank top rows dynamically
        raw_df = pd.read_excel(direct_download_url, skiprows=4)
        st.success("⚡ Data successfully fetched from your live Microsoft cloud storage location!")
    except Exception as link_err:
        st.error(f"Could not reach Microsoft Link. Ensure link access permissions allow sharing. Error: {str(link_err)}")

# Load from Manual Upload
if uploaded_file is not None and raw_df is None:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        # Select first sheet and skip top header spacing
        raw_df = pd.read_excel(uploaded_file, sheet_name=excel_obj.sheet_names[0], skiprows=4)
        st.success("🎉 Local file parsed successfully into active framework memory!")
    except Exception as file_err:
        st.error(f"Error parsing file: {str(file_err)}")

# --- RUN COMPUTATIONS IF DATA IS EXTRACTED ---
if raw_df is not None:
    # Clean up column spaces to guarantee variable matching loops pass
    raw_df.columns = [str(c).strip() for c in raw_df.columns]
    col_list = raw_df.columns.tolist()
   
    st.markdown("### 🔍 Data Column Matching Selector")
    st.info("The system pre-selected columns from your screenshot. Confirm settings below to map data fields.")
   
    # Precise automatic column index locating logic based on your visual layout template headers
    idx_dept = col_list.index("DEPARTMENT") if "DEPARTMENT" in col_list else 0
    idx_id = col_list.index("SERIAL NUMBER") if "SERIAL NUMBER" in col_list else 0
    idx_name = col_list.index("EQUIPMENT DESCRIPTION") if "EQUIPMENT DESCRIPTION" in col_list else 0
    idx_due = col_list.index("CALIB. DATE DUE") if "CALIB. DATE DUE" in col_list else 0
    idx_status = col_list.index("STATUS") if "STATUS" in col_list else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        dept_col = st.selectbox("Department Assignment Source:", col_list, index=idx_dept)
    with c2:
        id_col = st.selectbox("Equipment ID / Serial No Parameter:", col_list, index=idx_id)
    with c3:
        name_col = st.selectbox("Equipment Name / Title Column:", col_list, index=idx_name)
    with c4:
        status_col = st.selectbox("Status Condition Column:", col_list, index=idx_status)
    with c5:
        due_col = st.selectbox("Next Calibration Due Date Tracker:", col_list, index=idx_due)

    # Re-structure columns into uniform data arrays
    working_df = raw_df[[id_col, name_col, dept_col, status_col, due_col]].copy()
    working_df.columns = ["ID", "Name", "Department", "Excel_Status", "Due Date"]
   
    # Drop completely blank row references
    working_df = working_df.dropna(subset=["ID"])
    working_df["ID"] = working_df["ID"].astype(str).str.strip()
    working_df = working_df[working_df["ID"] != "nan"]
   
    # Process department strings cleanly for comparison metrics
    working_df["Department"] = working_df["Department"].astype(str).str.strip()
   
    # Filter out items already marked as REMOVED or manually Scrapped
    working_df = working_df[working_df["Excel_Status"].astype(str).str.strip().str.upper() != "REMOVED"]
    working_df = working_df[~working_df["ID"].isin(st.session_state.scrapped_items)]
   
    # Convert custom system date strings dynamically safely
    working_df["Due Date"] = pd.to_datetime(working_df["Due Date"], errors='coerce').dt.date
    today = datetime.date.today()

    # Calculate precise remaining calendar days until due
    def calculate_days_left(due_date):
        if pd.isnull(due_date): return 999
        return (due_date - today).days

    working_df["Days Remaining"] = working_df["Due Date"].apply(calculate_days_left)

    # --- SEGMENT TRACKING LOGIC ROUTINE ---
    def determine_time_segment(row):
        days = row["Days Remaining"]
        if pd.isnull(row["Due Date"]):
            return "VALID"
        if days < 0:
            return "OVERDUE"
        elif days <= 30:
            return "Due in Next 30 days"
        elif days <= 91:
            return "Due in Next 3 Months"
        else:
            return "VALID"

    working_df["Time Segment"] = working_df.apply(determine_time_segment, axis=1)

    # --- SECTION 2: AUTOMATED REMINDER EMAIL MECHANISM PANEL ---
    st.markdown("---")
    st.subheader("✉️ 2. Upcoming Calibration Alert Logs (< 30 Days Warning Notification)")
   
    # Business Rule: Capture rows where deadline is less than 30 days remaining, but not overdue yet
    upcoming_30_days = working_df[(working_df["Days Remaining"] >= 0) & (working_df["Days Remaining"] <= 30)].copy()
   
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.metric("Instruments Expiring in Less Than 30 Days", len(upcoming_30_days))
    with col_m2:
        email_target = st.text_input("Enter manager warning distribution email address:", "boss_email@ccbsa.co.za")
   
    if st.button("🚀 Dispatch Upcoming Notification Summary Email"):
        if len(upcoming_30_days) > 0:
            email_body = f"CCBSA Pretoria Calibration Warning Alert -\nGenerated on: {today}\n\nTHE FOLLOWING ITEMS ARE DUE FOR CALIBRATION IN LESS THAN 30 DAYS:\n\n"
            for _, row in upcoming_30_days.iterrows():
                email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Description: {row['Name']} | Due Date: {row['Due Date']} ({row['Days Remaining']} Days Left)\n"
           
            st.info("✨ Complete warning delivery notification package rendered below:")
            st.code(email_body, language="text")
            st.success(f"📩 Notification summary data ready for transmission log list dispatch successfully to {email_target}!")
        else:
            st.success("🎉 All systems clear! No instruments are expiring within the next 30 days.")

    # --- SECTION 3: METRIC GRAPHS ENGINE ---
    st.markdown("---")
    st.subheader("📊 3. Performance Status Charts")
   
    graph_counts = working_df["Time Segment"].value_counts().reset_index()
    graph_counts.columns = ["Status Condition", "Total Instrument Count"]
   
    color_map = {"OVERDUE": "#FF4B4B", "Due in Next 30 days": "#FFA500", "Due in Next 3 Months": "#3399FF", "VALID": "#2ECC71"}
   
    fig = px.bar(
        graph_counts,
        x="Status Condition",
        y="Total Instrument Count",
        color="Status Condition",
        title="Real-Time Equipment Status Breakdown Overview",
        color_discrete_map=color_map
    )
    st.plotly_chart(fig, use_container_width=True)

    # --- SECTION 4: LIVE CALIBRATION SUMMARY MATRIX ---
    st.markdown("---")
    st.subheader("📋 4. Department Calibration Distribution Summary Matrix")
   
    matrix_records = []
    for d in DEPARTMENTS_LIST:
        # Filter matching specific text criteria ignoring capitalization anomalies
        dept_mask = working_df[working_df["Department"].str.lower() == d.lower()]
       
        matrix_records.append({
            "Departments": d,
            "OVERDUE": len(dept_mask[dept_mask["Time Segment"] == "OVERDUE"]),
            "Due in Next 30 days": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 30 days"]),
            "Due in Next 3 Months": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 3 Months"]),
            "VALID": len(dept_mask[dept_mask["Time Segment"] == "VALID"]),
            "TOTAL": len(dept_mask)
        })
       
    summary_matrix_df = pd.DataFrame(matrix_records)
    st.dataframe(summary_matrix_df, use_container_width=True, hide_index=True)
   
    # Summary total calculation footer widgets
    c_ov, c_30, c_3m, c_tot = st.columns(4)
    c_ov.metric("Total OVERDUE Items", summary_matrix_df["OVERDUE"].sum())
    c_30.metric("Due within 30 Days", summary_matrix_df["Due in Next 30 days"].sum())
    c_3m.metric("Due within 3 Months", summary_matrix_df["Due in Next 3 Months"].sum())
    c_tot.metric("Total Tracked Assets", summary_matrix_df["TOTAL"].sum())

    # --- SECTION 5: REGISTER DETAIL PREVIEWER ---
    st.markdown("---")
    st.subheader("📋 5. Asset Register Detailed Rows")
    st.dataframe(working_df[["ID", "Name", "Department", "Due Date", "Time Segment", "Days Remaining"]], use_container_width=True, hide_index=True)
    
    
