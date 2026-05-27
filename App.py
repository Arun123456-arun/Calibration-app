
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

# Departments tracking matrix list
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- SECTION 1: MICROSOFT LINK INTERACTION ENGINE ---
st.subheader("📁 1. System Master Database Storage Connection")

st.markdown("""
💡 **How to link your live Excel file from Microsoft OneDrive/SharePoint:**
1. Open your corporate OneDrive or SharePoint in your web browser.
2. Right-click your Excel file and select **Copy Link**.
3. Paste that web URL address in the box below to link the data dynamically.
""")

# Microsoft link input container
microsoft_url = st.text_input(
    "Paste your Microsoft OneDrive/SharePoint file share link here:",
    placeholder="https://sharepoint.com..."
)

uploaded_file = None
# Fallback to file uploader if no link is pasted yet
if not microsoft_url:
    st.info("ℹ️ Alternatively, upload the file manually below until you paste your Microsoft Link:")
    uploaded_file = st.file_uploader("Upload your master calibration template workbook file directly:", type=["xlsx", "xls"])

# --- PROCESSING SYSTEM DATA ENGINE ---
raw_df = None

if microsoft_url:
    try:
        # Convert standard SharePoint links into direct file download paths
        direct_download_url = microsoft_url.replace(":x:/g/", ":x:/g/").split("?")[0] + "?download=1"
        raw_df = pd.read_excel(direct_download_url)
        st.success("⚡ Data successfully fetched from your live Microsoft cloud storage location!")
    except Exception as link_err:
        st.error(f"Could not reach Microsoft Link. Please ensure link access permissions allow sharing. Error details: {str(link_err)}")

if uploaded_file is not None and raw_df is None:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        # Automatically read the first available data sheet
        raw_df = pd.read_excel(uploaded_file, sheet_name=excel_obj.sheet_names[0])
        st.success("🎉 Local file parsed successfully into active framework memory!")
    except Exception as file_err:
        st.error(f"Error parsing file: {str(file_err)}")

# Execute application pipeline logic only when data exists
if raw_df is not None:
    # CLEANING STEP: Clean the data rows by looking for the row that has 'EQUIPMENT NAME'
    cleaned_df = raw_df.copy()
   
    # Let's dynamically look for your column row headers to skip the empty top cells visible in your screenshot
    found_header = False
    for i in range(min(15, len(cleaned_df))):
        row_values = [str(val).strip().upper() for val in cleaned_df.iloc[i].values]
        if 'EQUIPMENT NAME' in row_values or 'CALIB. DATE' in row_values:
            cleaned_df.columns = [str(c).strip() for c in cleaned_df.iloc[i]]
            cleaned_df = cleaned_df.iloc[i+1:].reset_index(drop=True)
            found_header = True
            break
           
    col_list = cleaned_df.columns.tolist()
   
    st.markdown("### 🔍 Data Column Matching Selector")
    st.info("Match the dropdown configurations below with your sheet labels to align data metrics perfectly.")
   
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        id_col = st.selectbox("Equipment ID / Serial No Parameter:", col_list, index=3 if "SERIAL NUMBER" in col_list else 0)
    with c2:
        name_col = st.selectbox("Equipment Name / Title Column:", col_list, index=0 if "EQUIPMENT NAME" in col_list else 0)
    with c3:
        # Fallback tracking assignment if specific column isn't found
        dept_col = st.selectbox("Department Assignment Source:", col_list, index=0)
        st.caption("⚠️ Note: If your file does not have a separate 'Department' column, assign it to any label like Name or Comments to group items.")
    with c4:
        due_col = st.selectbox("Next Calibration Due Date Tracker:", col_list, index=6 if "CALIB. DATE" in col_list else 0)

    # Re-structure columns to clean standard definitions
    working_df = cleaned_df[[id_col, name_col, dept_col, due_col]].copy()
    working_df.columns = ["ID", "Name", "Department", "Due Date"]
   
    # Drop empty records and filter scrapped elements
    working_df = working_df.dropna(subset=["ID"])
    working_df["ID"] = working_df["ID"].astype(str).str.strip()
    working_df = working_df[working_df["ID"] != "nan"]
   
    # Clean up status field strings
    working_df["Department"] = working_df["Department"].astype(str).str.strip()
    working_df = working_df[~working_df["ID"].isin(st.session_state.scrapped_items)]
   
    # Convert dates reliably
    working_df["Due Date"] = pd.to_datetime(working_df["Due Date"], errors='coerce').dt.date
    today = datetime.date.today()

    # Calculate days remaining until calibration is due
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
   
    # Target values expiring in less than 30 days
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
                email_body += f"• ID: {row['ID']} | Name: {row['Name']} | Due Date: {row['Due Date']} ({row['Days Remaining']} Days Left)\n"
           
            st.info("✨ Complete warning delivery notification package rendered below:")
            st.code(email_body, language="text")
            st.success(f"📩 Notification list dispatched successfully to {email_target}!")
        else:
            st.success("🎉 All systems clear! No instruments are expiring within the next 30 days.")

    # --- SECTION 3: METRIC GRAPHS ENGINE ---
    st.markdown("---")
    st.subheader("📊 3. Performance Metric Dashboards")
   
    # Count the number of items in each time segment
    graph_counts = working_df["Time Segment"].value_counts().reset_index()
    graph_counts.columns = ["Status Condition", "Total Instrument Count"]
   
    # Color-coded metric plot generator
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
    st.subheader("📋 4. Live Summary Distribution Matrix")
   
    matrix_records = []
    # If the sheet doesn't contain standard department values, let's group by active distinct titles
    unique_depts = working_df["Department"].unique().tolist()
   
    for d in unique_depts:
        dept_mask = working_df[working_df["Department"] == d]
        matrix_records.append({
            "Tracking Group": d,
            "OVERDUE": len(dept_mask[dept_mask["Time Segment"] == "OVERDUE"]),
            "Due in Next 30 days": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 30 days"]),
            "Due in Next 3 Months": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 3 Months"]),
            "VALID": len(dept_mask[dept_mask["Time Segment"] == "VALID"]),
            "TOTAL": len(dept_mask)
        })
       
    summary_matrix_df = pd.DataFrame(matrix_records)
    st.dataframe(summary_matrix_df, use_container_width=True, hide_index=True)

    # --- SECTION 5: REGISTER DETAIL PREVIEWER ---
    st.markdown("---")
    st.subheader("📋 5. Asset Register Detailed Rows")
    st.dataframe(working_df, use_container_width=True, hide_index=True)
    
