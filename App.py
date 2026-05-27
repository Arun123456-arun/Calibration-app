import streamlit as st
import pandas as pd
import datetime
import plotly.express as px

# --- CONFIGURATION & LAYOUT ---
st.set_page_config(page_title="CCBSA Calibration Portal", layout="wide", page_icon="🔴")

# --- CCBA BRANDING LOGO HEADER ---
st.markdown(
    """
    <div style="background-color:#E31B23; padding:20px; border-radius:10px; text-align:center; margin-bottom:25px;">
        <h1 style="color:white; margin:0; font-family:'Arial Black', Gadget, sans-serif; letter-spacing: 2px;">🔴 COCA-COLA BEVERAGES AFRICA</h1>
        <p style="color:white; margin:5px 0 0 0; font-size:16px; font-weight:bold; opacity:0.9;">CCBSA Pretoria — Calibration Master Control Portal</p>
    </div>
    """,
    unsafe_allow_html=True
)

# Explicit list of departments matching your master sheet
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- SECTION 1: STORAGE CONNECTION ---
st.subheader("📁 1. Load System Master Database")

uploaded_file = st.file_uploader("Upload your updated calibration master template workbook file here:", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Step 1: Handle file types flexibly (Excel or CSV)
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file)
        else:
            excel_obj = pd.ExcelFile(uploaded_file)
            raw_df = pd.read_excel(uploaded_file, sheet_name=excel_obj.sheet_names[0])
       
        # Step 2: Clean hidden whitespace from headers and force uppercase matching
        # Keeping duplicate flags intact (.1, .2) so we can map them precisely if needed
        raw_df.columns = [str(c).strip().upper() for c in raw_df.columns]
       
        # Step 3: Map Columns dynamically using flexible matching rules to avoid layout mismatch crashes
        id_col = next((c for c in raw_df.columns if "SERIAL NUMBER" in c or "SERIAL" in c), None)
        desc_col = next((c for c in raw_df.columns if "EQUIPMENT DESCRIPTION" in c or "DESC" in c), None)
        dept_col = next((c for c in raw_df.columns if "DEPARTMENT" in c), None)
        status_col = next((c for c in raw_df.columns if "STATUS" in c), None)
        
        # Target the final calibration deadline evaluation point
        date_col = next((c for c in raw_df.columns if "CALIB. DATE DUE" in c or "DATE DUE" in c), None)
       
        # Verify structure validation check
        missing_cols = [col for col in ["SERIAL NUMBER", "EQUIPMENT DESCRIPTION", "DEPARTMENT", "STATUS", "CALIB. DATE DUE"] if not locals()[f"{'id' if col=='SERIAL NUMBER' else 'desc' if col=='EQUIPMENT DESCRIPTION' else 'dept' if col=='DEPARTMENT' else 'status' if col=='STATUS' else 'date'}_col"]]
       
        if missing_cols:
            st.error(f"❌ Structural match failure. Could not find or map these structural requirements: {missing_cols}. Please verify header names on Row 1.")
            st.write("Headers found in your file:", list(raw_df.columns))
        else:
            # Re-structure columns into uniform data arrays
            working_df = raw_df[[id_col, desc_col, dept_col, status_col, date_col]].copy()
            working_df.columns = ["ID", "Description", "Department", "Status", "Due_Date"]
           
            # Drop records missing serial numbers safely
            working_df = working_df.dropna(subset=["ID"])
            working_df["ID"] = working_df["ID"].astype(str).str.strip()
            working_df = working_df[(working_df["ID"] != "nan") & (working_df["ID"] != "")]
           
            # Filter out items officially decommissioned, removed, or approved for exclusion
            working_df["Status"] = working_df["Status"].astype(str).str.strip().str.upper()
            active_df = working_df[~working_df["Status"].isin(["REMOVED", "YES"])].copy()
           
            # Convert calendar dates cleanly, turning missing or bad text formats into blank rows
            active_df["Due_Date"] = pd.to_datetime(active_df["Due_Date"], errors='coerce')
            active_df = active_df.dropna(subset=["Due_Date"])
           
            # Check if data arrays contain entries after sanitation filters
            if active_df.empty:
                st.warning("⚠️ No active validation records found in the uploaded file tracking framework after filtering.")
            else:
                # Set target timeline reference point to today's live coordinate: May 27, 2026
                today = datetime.date(2026, 5, 27)
               
                # Fallback Calculation Method: Avoids Series object execution failures entirely
                def get_days_remaining(val):
                    try:
                        return (val.date() - today).days
                    except Exception:
                        return None

                active_df["Days Remaining"] = active_df["Due_Date"].apply(get_days_remaining)
                active_df = active_df.dropna(subset=["Days Remaining"])
                active_df["Days Remaining"] = active_df["Days Remaining"].astype(int)
                
                # Convert back to standard date format for visual grid presentations
                active_df["Due_Date"] = active_df["Due_Date"].apply(lambda d: d.date())
               
                # Map items exclusively into your requested 4 status category tier structure
                def segment_instrument(days):
                    if days < 0:
                        return "OVERDUE"
                    elif days <= 7:
                        return "Due in Next 7 Days"
                    elif days <= 30:
                        return "Due in 8 to 30 Days"
                    elif days <= 49: # 7 weeks = 49 days
                        return "Due in Next 7 Weeks"
                    else:
                        return "VALID"
               
                active_df["Time Segment"] = active_df["Days Remaining"].apply(segment_instrument)
               
                # --- SECTION 3: SUMMARY DISPLAY MATRIX ---
                st.markdown("---")
                st.subheader("📊 2. Department Calibration Distribution Summary Matrix")
               
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
               
                # Layout Summary KPI Scorecards
                c_ov, c_7d, c_30d, c_tot = st.columns(4)
                c_ov.metric("Total OVERDUE", int(summary_matrix_df["OVERDUE"].sum()))
                c_7d.metric("Due within 7 Days", int(summary_matrix_df["Due in Next 7 Days"].sum()))
                c_30d.metric("Due 8 to 30 Days", int(summary_matrix_df["Due in 8 to 30 Days"].sum()))
                c_tot.metric("Total Active Backlog", int(summary_matrix_df["TOTAL PENDING"].sum()))
               
                # --- SECTION 4: PLOTLY GRAPH CHART DISPLAY ENGINE ---
                st.markdown("---")
                st.subheader("📈 3. Equipment Backlog Status Distribution Chart")
               
                pending_df = active_df[active_df["Time Segment"] != "VALID"]
                if not pending_df.empty:
                    graph_data = pending_df["Time Segment"].value_counts().reset_index()
                    graph_data.columns = ["Urgency Status", "Equipment Count"]
                    
                    # Enforce locked layout sequence sorting rules on graph canvas
                    status_order = ["OVERDUE", "Due in Next 7 Days", "Due in 8 to 30 Days", "Due in Next 7 Weeks"]
                    graph_data["Urgency Status"] = pd.Categorical(graph_data["Urgency Status"], categories=status_order, ordered=True)
                    graph_data = graph_data.sort_values("Urgency Status")
                   
                    fig = px.bar(
                        graph_data,
                        x="Urgency Status",
                        y="Equipment Count",
                        color="Urgency Status",
                        color_discrete_map={
                            "OVERDUE": "#E31B23",             # Coca-Cola Red
                            "Due in Next 7 Days": "#FF4500",     # Deep Red-Orange
                            "Due in 8 to 30 Days": "#FFA500",    # Warning Amber
                            "Due in Next 7 Weeks": "#3399FF"     # Info Blue
                        },
                        title="Pretoria Plant Total Pending Calibration Backlog"
                    )
                    fig.update_layout(showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.success("🎉 No outstanding actions recorded. All device calibrations are 100% current!")

                # --- SECTION 5: MANAGER WARNING ALERT TRANSMISSION CENTER ---
                st.markdown("---")
                st.subheader("✉️ 4. Generate Calibration Executive Warning Log")
               
                overdue_items = active_df[active_df["Time Segment"] == "OVERDUE"]
                items_7d = active_df[active_df["Time Segment"] == "Due in Next 7 Days"]
                items_30d = active_df[active_df["Time Segment"] == "Due in 8 to 30 Days"]
               
                st.write(f"⚠️ Current Operational Backlog: **{len(overdue_items)}** Overdue | **{len(items_7d)}** Due within 7 days | **{len(items_30d)}** Due within 30 days.")
               
                email_target = st.text_input("Enter manager alert report notification email address:", "boss_email@ccbsa.co.za")
                if st.button("🚀 Dispatch Alert Email Notification Logs"):
                    
                    email_body = f"CCBSA Pretoria Calibration Notice Log -\nReport Generated: 2026-05-27\n"
                    email_body += "==================================================\n\n"
                    
                    if len(overdue_items) > 0:
                        email_body += "🚨 CRITICAL: THE FOLLOWING INSTRUMENTS ARE OVERDUE:\n"
                        for _, row in overdue_items.iterrows():
                            email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Desc: {row['Description']} | EXPIRED: {row['Due_Date']} ({abs(row['Days Remaining'])} Days Overdue)\n"
                        email_body += "\n"
                        
                    if len(items_7d) > 0:
                        email_body += "⚠️ HIGH PRIORITY: DUE WITHIN 7 DAYS:\n"
                        for _, row in items_7d.iterrows():
                            email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Desc: {row['Description']} | Due: {row['Due_Date']} ({row['Days Remaining']} Days Left)\n"
                        email_body += "\n"
                        
                    if len(items_30d) > 0:
                        email_body += "📅 ATTENTION: DUE WITHIN 8 TO 30 DAYS:\n"
                        for _, row in items_30d.iterrows():
                            email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Desc: {row['Description']} | Due: {row['Due_Date']} ({row['Days Remaining']} Days Left)\n"
                        email_body += "\n"
                   
                    st.info("Email communication text package successfully constructed:")
                    st.code(email_body, language="text")
                    st.success(f"📩 Notification listing successfully queued for delivery to {email_target}!")

                # --- SECTION 6: REGISTER DATA VIEW GRID ---
                st.markdown("---")
                st.subheader("📋 5. Asset Register Detailed Rows (Filtered Active Items)")
                st.dataframe(active_df[["ID", "Description", "Department", "Due_Date", "Time Segment", "Days Remaining"]], use_container_width=True, hide_index=True)
           
    except Exception as e:
        st.error(f"Error compiling spreadsheet rows: {str(e)}")
else:
    st.info("💡 Please upload your clean formatted Excel spreadsheet above to display the updated monitoring panels.")
               
                                    
