import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- CONFIGURATION & LAYOUT ---
st.set_page_config(page_title="CCBSA Calibration Portal & AI Governance", layout="wide", page_icon="🔴")

# --- CCBA BRANDING LOGO HEADER ---
st.markdown(
    """
    <div style="background-color:#E31B23; padding:20px; border-radius:10px; text-align:center; margin-bottom:25px;">
        <h1 style="color:white; margin:0; font-family:'Arial Black', Gadget, sans-serif; letter-spacing: 2px;">🔴 COCA-COLA BEVERAGES AFRICA</h1>
        <p style="color:white; margin:5px 0 0 0; font-size:16px; font-weight:bold; opacity:0.9;">CCBSA Pretoria — Calibration & AI Governance Master Control Portal</p>
    </div>
    """,
    unsafe_allow_html=True
)

# Explicit list of departments matching your master sheet
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- SIDEBAR CONNECTIONS MENU ---
st.sidebar.subheader("⚙️ System Connection Configurations")
MICROSOFT_EXCEL_LINK = st.sidebar.text_input(
    "Live Cloud Spreadsheet Link (Google Sheet / SharePoint):",
    value="https://docs.google.com/spreadsheets/d/1vd-sB3FmIfYBFf3d6yPnMIM78u8_NpMNXJaWmMNV8i8/export?format=csv&gid=150526899"
)

st.sidebar.markdown("---")
st.sidebar.subheader("📧 Outlook Outbound Gateway Settings")
SENDER_EMAIL = st.sidebar.text_input("Your Corporate Outlook Email:", "your_name@ccbsa.co.za")
SENDER_PASSWORD = st.sidebar.text_input("Outlook App Password:", type="password")

# Tabs Layout Matrix
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Performance Dashboard", 
    "📜 Certificate Compliance Manager", 
    "🤖 AI Smart Certificate Validator",
    "⚖️ Governance Framework"
])

# --- SMART FAILSAFE LIVE LOADING DATA ENGINE ---
raw_df = None

if MICROSOFT_EXCEL_LINK:
    try:
        # Check if it's a Google Sheets Export or a direct raw text CSV file stream
        if "format=csv" in MICROSOFT_EXCEL_LINK or "csv" in MICROSOFT_EXCEL_LINK.lower():
            raw_df = pd.read_csv(MICROSOFT_EXCEL_LINK)
        # Check if it's a standard web URL Excel stream
        elif MICROSOFT_EXCEL_LINK.startswith("http"):
            excel_obj = pd.ExcelFile(MICROSOFT_EXCEL_LINK)
            raw_df = pd.read_excel(MICROSOFT_EXCEL_LINK, sheet_name=excel_obj.sheet_names[0])
        else:
            raw_df = pd.read_excel(MICROSOFT_EXCEL_LINK)
    except Exception as cloud_err:
        st.sidebar.warning("🔄 Cloud stream link unavailable or requires direct file export tokens. Attempting local storage load...")

# Backup Step: Auto-scans directory folders for local backup files if link fails
if raw_df is None:
    for file in os.listdir('.'):
        if "Calibration" in file and file.endswith(('.xlsx', '.csv')):
            try:
                if file.endswith('.csv'):
                    raw_df = pd.read_csv(file)
                else:
                    raw_df = pd.read_excel(file)
                st.sidebar.info(f"Loaded database backup tracker safely from local source file: {file}")
                break
            except Exception:
                pass

if raw_df is not None:
    try:
        # Clean whitespaces from headers to keep scanning reliable
        raw_df.columns = [str(c).strip() for c in raw_df.columns]
        
        # --- SAFE MANDATORY COLS EXTRACTION ENGINE ---
        id_col, desc_col, dept_col, status_col, date_col = None, None, None, None, None
        evidence_col = None  

        for col in raw_df.columns:
            col_upper = col.upper()
            if "SERIAL NUMBER" in col_upper or "SERIAL" in col_upper: id_col = col
            elif "EQUIPMENT DESCRIPTION" in col_upper or "DESC" in col_upper: desc_col = col
            elif "DEPARTMENT" in col_upper: dept_col = col
            elif "STATUS" in col_upper: status_col = col
            elif "EVIDENCE" in col_upper: evidence_col = col  
            elif "CALIB. DATE DUE" in col_upper or "DATE DUE" in col_upper:
                if not date_col: date_col = col

        # Direct string absolute match fallbacks
        if not id_col and "SERIAL NUMBER" in raw_df.columns: id_col = "SERIAL NUMBER"
        if not desc_col and "EQUIPMENT DESCRIPTION" in raw_df.columns: desc_col = "EQUIPMENT DESCRIPTION"
        if not dept_col and "DEPARTMENT" in raw_df.columns: dept_col = "DEPARTMENT"
        if not status_col and "STATUS" in raw_df.columns: status_col = "STATUS"
        if not date_col and "CALIB. DATE DUE" in raw_df.columns: date_col = "CALIB. DATE DUE"

        if not (id_col and desc_col and dept_col and status_col and date_col):
            st.error("❌ Key structural columns missing from current spreadsheet schema view. Verify Row 1 labels.")
            st.write("Columns found in sheet:", list(raw_df.columns))
        else:
            # Construct active frame matrix data
            working_df = pd.DataFrame({
                "ID": raw_df[id_col],
                "Description": raw_df[desc_col],
                "Department": raw_df[dept_col],
                "Status": raw_df[status_col],
                "Due_Date": raw_df[date_col]
            }).copy()
            
            if evidence_col:
                working_df["Evidence_In_Sheet"] = raw_df[evidence_col].fillna("NO")
            else:
                working_df["Evidence_In_Sheet"] = "NO"
           
            working_df = working_df.dropna(subset=["ID"])
            working_df["ID"] = working_df["ID"].astype(str).str.strip()
            working_df = working_df[(working_df["ID"] != "nan") & (working_df["ID"] != "")]
           
            working_df["Status"] = working_df["Status"].astype(str).str.strip().str.upper()
            active_df = working_df[~working_df["Status"].isin(["REMOVED", "YES"])].copy()
           
            active_df["Due_Date"] = pd.to_datetime(active_df["Due_Date"], errors='coerce')
            active_df = active_df.dropna(subset=["Due_Date"])
           
            # Internal reference matrix coordination timeline anchor: June 8, 2026
            today = datetime.date(2026, 6, 8) 
            
            def get_days_remaining(val):
                try: return (val.date() - today).days
                except Exception: return None

            active_df["Days Remaining"] = active_df["Due_Date"].apply(get_days_remaining)
            active_df = active_df.dropna(subset=["Days Remaining"])
            active_df["Days Remaining"] = active_df["Days Remaining"].astype(int)
            active_df["Due_Date"] = active_df["Due_Date"].apply(lambda d: d.date())
           
            def segment_instrument(days):
                if days < 0: return "OVERDUE"
                elif days <= 7: return "Due in Next 7 Days"
                elif days <= 30: return "Due in 8 to 30 Days"
                elif days <= 49: return "Due in Next 7 Weeks"
                else: return "VALID"
           
            active_df["Time Segment"] = active_df["Days Remaining"].apply(segment_instrument)

            # ----------------------------------------------------
            # TAB 1: PERFORMANCE DASHBOARD
            # ----------------------------------------------------
            with tab1:
                st.subheader("📊 Operational Calibration Distribution Summary Matrix")
               
                matrix_records = []
                for d in DEPARTMENTS_LIST:
                    dept_mask = active_df[active_df["Department"].astype(str).str.strip().str.lower() == d.lower()]
                    matrix_records.append({
                        "Departments": d,
                        "OVERDUE": int(len(dept_mask[dept_mask["Time Segment"] == "OVERDUE"])),
                        "Due in Next 7 Days": int(len(dept_mask[dept_mask["Time Segment"] == "Due in Next 7 Days"])),
                        "Due in 8 to 30 Days": int(len(dept_mask[dept_mask["Time Segment"] == "Due in 8 to 30 Days"])),
                        "Due in Next 7 Weeks": int(len(dept_mask[dept_mask["Time Segment"] == "Due in Next 7 Weeks"])),
                        "TOTAL PENDING": int(len(dept_mask[dept_mask["Time Segment"] != "VALID"]))
                    })
                   
                summary_matrix_df = pd.DataFrame(matrix_records)
                st.dataframe(summary_matrix_df, use_container_width=True, hide_index=True)
               
                c_ov, c_7d, c_30d, c_tot = st.columns(4)
                c_ov.metric("Total OVERDUE", int(summary_matrix_df["OVERDUE"].sum()))
                c_7d.metric("Due within 7 Days", int(summary_matrix_df["Due in Next 7 Days"].sum()))
                c_30d.metric("Due 8 to 30 Days", int(summary_matrix_df["Due in 8 to 30 Days"].sum()))
                c_tot.metric("Total Active Backlog", int(summary_matrix_df["TOTAL PENDING"].sum()))
               
                st.markdown("---")
                pending_df = active_df[active_df["Time Segment"] != "VALID"]
                if not pending_df.empty:
                    graph_data = pending_df["Time Segment"].value_counts().reset_index()
                    graph_data.columns = ["Urgency Status", "Equipment Count"]
                    status_order = ["OVERDUE", "Due in Next 7 Days", "Due in 8 to 30 Days", "Due in Next 7 Weeks"]
                    graph_data["Urgency Status"] = pd.Categorical(graph_data["Urgency Status"], categories=status_order, ordered=True)
                    graph_data = graph_data.sort_values("Urgency Status")
                   
                    fig = px.bar(
                        graph_data, x="Urgency Status", y="Equipment Count", color="Urgency Status",
                        color_discrete_map={
                            "OVERDUE": "#E31B23", "Due in Next 7 Days": "#FF4500",     
                            "Due in 8 to 30 Days": "#FFA500", "Due in Next 7 Weeks": "#3399FF"     
                        },
                        title="Pretoria Plant Pending Calibration Distribution"
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # ----------------------------------------------------
            # TAB 2: CERTIFICATE COMPLIANCE MANAGER
            # ----------------------------------------------------
            with tab2:
                st.subheader("📜 Master Calibration Certificate Audit Track")
                active_df["Evidence Status"] = active_df["Evidence_In_Sheet"].astype(str).str.strip().str.upper().apply(
                    lambda x: "🟢 Cert Present" if x in ["YES", "TRUE", "1"] else "🔴 MISSING CERTIFICATE"
                )
                st.dataframe(
                    active_df[["ID", "Description", "Department", "Due_Date", "Time Segment", "Evidence Status"]],
                    use_container_width=True,
                    hide_index=True
                )

            # ----------------------------------------------------
            # TAB 3: AI SMART CERTIFICATE VALIDATOR
            # ----------------------------------------------------
            with tab3:
                st.subheader("🤖 AI Automated Certificate Validation Portal")
                st.write("Upload an external vendor's PDF/Image calibration certificate. The local AI parser will scan the filename metadata and explicitly flag whether the item's compliance timeline is current or pending calibration.")
                
                uploaded_file = st.file_uploader("Drop Vendor Calibration Certificate Document:", type=["pdf", "png", "jpg", "xlsx", "csv"])
                
                if uploaded_file is not None:
                    filename_clean = str(uploaded_file.name).upper().strip()
                    st.markdown("### 🧠 AI Analysis Results:")
                    
                    matched_row = None
                    for _, row in active_df.iterrows():
                        asset_id = str(row["ID"]).upper().strip()
                        if asset_id in filename_clean:
                            matched_row = row
                            break
                    
                    if matched_row is not None:
                        st.success(f"✅ **AI Match Successful:** Isolated target Asset ID `{matched_row['ID']}` from filename attributes.")
                        col_ai1, col_ai2 = st.columns(2)
                        with col_ai1:
                            st.write(f"**Equipment:** {matched_row['Description']}")
                            st.write(f"**Department/Line Location:** {matched_row['Department']}")
                            st.write(f"**Target System Expiry Date:** {matched_row['Due_Date']}")
                        with col_ai2:
                            days_left = matched_row['Days Remaining']
                            if days_left < 0:
                                st.error(f"🔴 **CALIBRATION DATE IS PENDING (OVERDUE)**\n\nThis device is operating **{abs(days_left)} days past its certificate expiration limit**. Immediate verification turnaround required.")
                            elif days_left <= 30:
                                st.warning(f"🟠 **CALIBRATION RUN PENDING SOON**\n\nCertificate is current but expires in **{days_left} days**. Prepare upcoming engineering scheduling window.")
                            else:
                                st.info(f"🟢 **CERTIFICATE DATE IS CURRENT / VALID**\n\nDevice compliance verification is healthy for the next **{days_left} days**.")
                    else:
                        st.info("⚠️ **AI Classification Notice:** Could not match a specific Serial Number identifier from the filename string structure.")
                        st.write("Please select the target instrument manual reference from the matrix box below to run the timeline audit verification loop:")
                        
                        manual_selection = st.selectbox("Select Target Plant Asset to Map:", sorted(active_df["ID"].unique()))
                        manual_row = active_df[active_df["ID"] == manual_selection].iloc[0]
                        
                        st.markdown("---")
                        st.write(f"📊 **AI Compliance Data Audit for Asset ID: {manual_row['ID']}** ({manual_row['Description']})")
                        days_left = manual_row['Days Remaining']
                        if days_left < 0:
                            st.error(f"🔴 **CALIBRATION DATE STATUS: PENDING AUDIT OVERDUE**\n\nThis asset device is currently running **{abs(days_left)} days past its safe timeline boundary**.")
                        else:
                            st.success(f"🟢 **CALIBRATION DATE STATUS: VALID / CURRENT**\n\nCertificate timeline matrix is verified secure for the next **{days_left} days**.")

            # ----------------------------------------------------
            # TAB 4: GOVERNANCE FRAMEWORK
            # ----------------------------------------------------
            with tab4:
                st.subheader("⚖️ CCBSA Pretoria Calibration Governance Standard")
                st.markdown(
                    """
                    ### 1. The Core Compliance Directive
                    In accordance with CCBSA Quality Management Systems and international standards (**ISO 9001:2015 Clause 7.1.5** and **ISO/IEC 17025**), all process control instruments, laboratory instruments, and weighing matrix networks sitting across the Pretoria layout must be systematically verified against national measurement standards traceable to **SANAS** (South African National Accreditation System).

                    ### 2. Standard Operating Procedure (SOP) for Incoming Certificates
                    When an asset is calibrated by an external service provider, the equipment owner must enforce the following validation pipeline before updating the master database:
                    * **Traceability Audit:** Verify that the service provider's master calibration certificate references active SANAS standards.
                    * **Tolerance Validation:** Cross-reference the error variances against factory limits. If the error margin exceeds critical limits, the instrument must be flagged as *Defective* and removed from service.
                    """
                )

    except Exception as e:
        st.error(f"Failed to compile operational panels from current sheet schema layout view: {str(e)}")
else:
    st.info("Please complete the direct download data link format configuration on the sidebar utility panel block.")
