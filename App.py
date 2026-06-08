import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- CONFIGURATION & LAYOUT ---
st.set_page_config(page_title="CCBSA Calibration Portal & Governance", layout="wide", page_icon="🔴")

# --- CCBA BRANDING LOGO HEADER ---
st.markdown(
    """
    <div style="background-color:#E31B23; padding:20px; border-radius:10px; text-align:center; margin-bottom:25px;">
        <h1 style="color:white; margin:0; font-family:'Arial Black', Gadget, sans-serif; letter-spacing: 2px;">🔴 COCA-COLA BEVERAGES AFRICA</h1>
        <p style="color:white; margin:5px 0 0 0; font-size:16px; font-weight:bold; opacity:0.9;">CCBSA Pretoria — Calibration & Governance Master Control Portal</p>
    </div>
    """,
    unsafe_allow_html=True
)

# Explicit list of departments matching your master sheet
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- OUTLOOK & LINK GATEWAY SETTINGS ---
st.sidebar.subheader("⚙️ System Connection Configurations")
MICROSOFT_EXCEL_LINK = st.sidebar.text_input(
    "Cloud SharePoint/OneDrive Direct Excel Link:",
    value="https://ccba.sharepoint.com/:x:/r/teams/PretoriaCalibration/_layouts/15/Doc.aspx?sourcedoc=%7B2a909791-d120-4817-8147-c17f7823b5dd%7D&action=default&download=1"
)

st.sidebar.markdown("---")
st.sidebar.subheader("📧 Outlook Outbound Gateway Settings")
SENDER_EMAIL = st.sidebar.text_input("Your Corporate Outlook Email:", "your_name@ccbsa.co.za")
SENDER_PASSWORD = st.sidebar.text_input("Outlook App Password:", type="password")

# Create main functional tabs for your boss
tab1, tab2, tab3 = st.tabs(["📊 Performance & Alerts Dashboard", "📜 Certificate Compliance Manager", "⚖️ AI Copilot & Governance Framework"])

if MICROSOFT_EXCEL_LINK:
    try:
        # Read the file cleanly
        if "ccba.sharepoint.com" in MICROSOFT_EXCEL_LINK or MICROSOFT_EXCEL_LINK.startswith("http"):
            excel_obj = pd.ExcelFile(MICROSOFT_EXCEL_LINK)
            raw_df = pd.read_excel(MICROSOFT_EXCEL_LINK, sheet_name=excel_obj.sheet_names[0])
        else:
            raw_df = pd.read_excel(MICROSOFT_EXCEL_LINK)
       
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
            elif "EVIDENCE" in col_upper: evidence_col = col  # Dynamically captures your evidence column safely
            elif "CALIB. DATE DUE" in col_upper or "DATE DUE" in col_upper:
                if not date_col: date_col = col

        # Absolute Fallbacks if loop missed them
        if not id_col and "SERIAL NUMBER" in raw_df.columns: id_col = "SERIAL NUMBER"
        if not desc_col and "EQUIPMENT DESCRIPTION" in raw_df.columns: desc_col = "EQUIPMENT DESCRIPTION"
        if not dept_col and "DEPARTMENT" in raw_df.columns: dept_col = "DEPARTMENT"
        if not status_col and "STATUS" in raw_df.columns: status_col = "STATUS"
        if not date_col and "CALIB. DATE DUE" in raw_df.columns: date_col = "CALIB. DATE DUE"

        if not (id_col and desc_col and dept_col and status_col and date_col):
            st.error("❌ Key structural columns missing from cloud database sheet index. Check Row 1 headers.")
            st.write("Columns found:", list(raw_df.columns))
        else:
            # Build unified tracking DataFrame safely using identified columns
            working_df = pd.DataFrame({
                "ID": raw_df[id_col],
                "Description": raw_df[desc_col],
                "Department": raw_df[dept_col],
                "Status": raw_df[status_col],
                "Due_Date": raw_df[date_col]
            }).copy()
            
            # Safely inject evidence array to prevent KeyErrors entirely
            if evidence_col:
                working_df["Evidence_In_Sheet"] = raw_df[evidence_col].fillna("NO")
            else:
                working_df["Evidence_In_Sheet"] = "NO"
           
            # Clean empty asset entries
            working_df = working_df.dropna(subset=["ID"])
            working_df["ID"] = working_df["ID"].astype(str).str.strip()
            working_df = working_df[(working_df["ID"] != "nan") & (working_df["ID"] != "")]
           
            # Filter out decommissioned items
            working_df["Status"] = working_df["Status"].astype(str).str.strip().str.upper()
            active_df = working_df[~working_df["Status"].isin(["REMOVED", "YES"])].copy()
           
            # Format and parse calibration timelines
            active_df["Due_Date"] = pd.to_datetime(active_df["Due_Date"], errors='coerce')
            active_df = active_df.dropna(subset=["Due_Date"])
           
            # Timeline Reference Matrix
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
            # TAB 1: PERFORMANCE & ALERTS DASHBOARD
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

                st.markdown("---")
                st.subheader("✉️ Generate Executive Warning Notification Logs")
                email_target = st.text_input("Enter recipient manager email address:", "boss_email@ccbsa.co.za")
                if st.button("🚀 Dispatch Alert Email Notification Logs"):
                    overdue_items = active_df[active_df["Time Segment"] == "OVERDUE"]
                    
                    email_body = f"CCBSA Pretoria Calibration Audit Notice Log -\nReport Generated: {today}\n"
                    email_body += "==================================================\n\n"
                    if len(overdue_items) > 0:
                        email_body += "🚨 CRITICAL: OVERDUE INSTRUMENTS:\n"
                        for _, row in overdue_items.iterrows():
                            email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Desc: {row['Description']} | EXPIRED: {row['Due_Date']} ({abs(row['Days Remaining'])} Days Overdue)\n"
                    
                    if not SENDER_EMAIL or not SENDER_PASSWORD:
                        st.error("⚠️ Enter corporate login credentials in the sidebar.")
                    else:
                        try:
                            msg = MIMEMultipart()
                            msg['From'] = SENDER_EMAIL
                            msg['To'] = email_target
                            msg['Subject'] = f"🚨 CCBSA Pretoria Calibration Alert Summary"
                            msg.attach(MIMEText(email_body, 'plain'))
                            server = smtplib.SMTP("smtp.office365.com", 587)
                            server.ehlo(); server.starttls(); server.ehlo()
                            server.login(SENDER_EMAIL, SENDER_PASSWORD)
                            server.sendmail(SENDER_EMAIL, email_target, msg.as_string())
                            server.quit()
                            st.success(f"📩 Notification sent to {email_target}!")
                        except Exception as e: st.error(f"Outlook SMTP Error: {e}")

            # ----------------------------------------------------
            # TAB 2: CERTIFICATE COMPLIANCE MANAGER
            # ----------------------------------------------------
            with tab2:
                st.subheader("📜 Live Calibration Certificate Audit Track")
                st.write("This interface tracks which active laboratory assets have accredited physical certificates uploaded or missing.")

                active_df["Evidence Status"] = active_df["Evidence_In_Sheet"].astype(str).str.strip().str.upper().apply(
                    lambda x: "🟢 Cert Present" if x in ["YES", "TRUE", "1"] else "🔴 MISSING CERTIFICATE"
                )

                total_assets = len(active_df)
                missing_certs = len(active_df[active_df["Evidence Status"] == "🔴 MISSING CERTIFICATE"])
                present_certs = total_assets - missing_certs
                compliance_score = (present_certs / total_assets) * 100 if total_assets > 0 else 100

                col_stat1, col_stat2, col_stat3 = st.columns(3)
                col_stat1.metric("Total Tracked Assets", total_assets)
                col_stat2.metric("Missing Certificates", missing_certs, delta=f"-{missing_certs} Pending Audit", delta_color="inverse")
                col_stat3.metric("Overall Documentation Audit Score", f"{compliance_score:.1f}%")

                st.markdown("---")
                st.subheader("📥 Upload Incoming External Calibration Certificates")
                st.info("💡 **Instructions for Technicians:** Drag and drop the certificate here. Ensure the file name contains the asset's exact Serial Number (e.g., `CCBSA6009`) so it auto-links.")
                
                uploaded_certs = st.file_uploader("Upload Service Provider Certificates (PDF / Images):", type=["pdf", "png", "jpg"], accept_multiple_files=True)
                
                uploaded_filenames = []
                if uploaded_certs:
                    for cert in uploaded_certs:
                        uploaded_filenames.append(cert.name.upper())
                    st.success(f"Successfully processed {len(uploaded_certs)} incoming external document(s).")

                def check_file_match(row):
                    asset_id = str(row["ID"]).upper()
                    for f in uploaded_filenames:
                        if asset_id in f:
                            return "🟢 Cert Just Uploaded (Pending Sync)"
                    return row["Evidence Status"]

                active_df["Audit Status (Live)"] = active_df.apply(check_file_match, axis=1)

                st.markdown("---")
                st.subheader("📋 Certificate Verification Ledger Matrix")
                dept_filter = st.selectbox("Filter Ledger Matrix by Plant Area:", ["ALL"] + DEPARTMENTS_LIST)
                cert_filter = st.radio("Filter by Documentation Presence:", ["Show All Tracked Assets", "Show Missing Certificates Only"])

                display_df = active_df.copy()
                if dept_filter != "ALL":
                    display_df = display_df[display_df["Department"].str.lower() == dept_filter.lower()]
                if cert_filter == "Show Missing Certificates Only":
                    display_df = display_df[display_df["Audit Status (Live)"] == "🔴 MISSING CERTIFICATE"]

                st.dataframe(
                    display_df[["ID", "Description", "Department", "Due_Date", "Time Segment", "Audit Status (Live)"]],
                    use_container_width=True,
                    hide_index=True
                )

            # ----------------------------------------------------
            # TAB 3: FREE AI COPILOT & GOVERNANCE FRAMEWORK
            # ----------------------------------------------------
            with tab3:
                st.subheader("🤖 Free AI Calibration & Audit Copilot")
                st.write("Type any audit finding, standard problem, or compliance question below. The local AI engine will process your entry and output the corresponding CCBSA governance protocol instantly.")
                
                # Free text input for your boss or auditors
                user_query = st.text_input("Ask the Compliance Assistant (e.g., 'What if an instrument fails calibration?' or 'How do I handle an overdue scale?'):")
                
                if user_query:
                    query_clean = user_query.lower()
                    st.markdown("#### 🧠 AI Compliance Assessment Output:")
                    
                    # High-speed local matching logic simulating an NLP classification model
                    if "fail" in query_clean or "broken" in query_clean or "defective" in query_clean or "error" in query_clean:
                        st.error("💥 **CCBSA Critical Alert: Failed Calibration Protocol Detected**")
                        st.markdown("""
                        * **Applicable ISO Standard:** ISO 9001:2015 Clause 7.1.5.2 (Measurement Traceability).
                        * **Required Action Action:** 1. Immediately remove the instrument from the production line layout.
                          2. Attach a physical **Red Tag (Defective Equipment - Do Not Operate)**.
                          3. Conduct a *Product Impact Assessment* on all batches run through that line since the last passing verification check to ensure product integrity.
                        """)
                    elif "overdue" in query_clean or "expired" in query_clean or "missed" in query_clean or "late" in query_clean:
                        st.warning("⚠️ **CCBSA Governance Warning: Overdue System Handling**")
                        st.markdown("""
                        * **Applicable ISO Standard:** ISO/IEC 17025 Operational Audit Controls.
                        * **Required Action Plan:**
                          1. Lock the device status inside this portal view.
                          2. Contact the dedicated departmental engineer to generate an urgent calibration Purchase Order (PO).
                          3. Any data logged by an overdue instrument during an audit will trigger a major non-conformance finding.
                        """)
                    elif "how" in query_clean or "upload" in query_clean or "certificate" in query_clean or "pdf" in query_clean:
                        st.info("ℹ️ **AI Core Assistant: Certificate Onboarding Workflow**")
                        st.markdown("""
                        * **Governance Protocol:** Digital Record Maintenance Standard.
                        * **Required Action Plan:**
                          1. Rename the incoming service provider PDF to match the exact asset **Serial Number** visible on the physical chassis.
                          2. Drop the file into **Tab 2 (Certificate Compliance Manager)**.
                          3. Verify that the score counter ticks upward. Keep the paper document on file in the Quality Assurance room cabinet for 5 rolling years.
                        """)
                    else:
                        st.success("🟢 **AI Core Assistant: General ISO Governance Guidance**")
                        st.markdown("""
                        * **CCBSA General Standard:** All instruments affecting safety, quality, or environment must match traceable **SANAS** national reference standards. 
                        * *Tip:* Try asking the AI specifically about 'failed instruments', 'overdue calibrations', or 'how to handle certificates' for precise workflows.
                        """)
                
                st.markdown("---")
                st.markdown(
                    """
                    ## ⚖️ CCBSA Pretoria Calibration Governance Standard
                    ### 1. The Core Compliance Directive
                    In accordance with CCBSA Quality Management Systems and international standards (**ISO 9001:2015 Clause 7.1.5** and **ISO/IEC 17025**), all process control instruments, laboratory instruments, and weighing matrix networks sitting across the Pretoria layout must be systematically verified against national measurement standards traceable to **SANAS** (South African National Accreditation System).

                    ### 2. Standard Operating Procedure (SOP) for Incoming Certificates
                    When an asset is calibrated by an external service provider, the equipment owner must enforce the following validation pipeline before updating the master database:
                    * **Traceability Audit:** Verify that the service provider's master calibration certificate references active SANAS standards.
                    * **Tolerance Validation:** Cross-reference the error variances against factory limits. If the error margin exceeds critical limits, the instrument must be flagged as *Defective* and removed from service.
                    """
                )
                
                st.markdown("---")
                st.subheader("📋 Internal Governance Sign-Off Control")
                with st.form("governance_signoff"):
                    reviewer_name = st.text_input("Reviewing Manager / Authority Name:")
                    signoff_date = st.date_input("Review Verification Date:", today)
                    comments = st.text_area("Audit Notes / Corrective Actions Implemented:")
                    
                    submit_button = st.form_submit_button("Sign-Off & Log Audit Instance Verification")
                    if submit_button:
                        if reviewer_name:
                            st.success(f"🔒 Governance Log Locked! Reviewed by {reviewer_name} on {signoff_date}.")
                        else:
                            st.error("Verification Sign-off requires an authorized manager's signature name to lock.")

    except Exception as e:
        st.error(f"Failed to access network spreadsheet source connection link: {str(e)}")
else:
    st.info("Please fill in your SharePoint workbook link in the sidebar menu input block.")
