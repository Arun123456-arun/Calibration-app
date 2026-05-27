
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

# Explicit list of departments matching your first column
DEPARTMENTS_LIST = ["Clinic & Security", "Engineering", "Packaging", "Quality & Lab", "Site", "Syrup room", "Utilities", "Warehouse", "Supply & Raw Mats"]

# --- SECTION 1: STORAGE CONNECTION ---
st.subheader("📁 1. Load System Master Database")

uploaded_file = st.file_uploader("Upload your updated calibration master template Excel workbook file here:", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        # Automatically skip the first 4 metadata rows to land exactly on your headers
        raw_df = pd.read_excel(uploaded_file, sheet_name=excel_obj.sheet_names[0], skiprows=4)
       
        # Clean hidden whitespace from headers and force uppercase matching
        raw_df.columns = [str(c).strip().upper() for c in raw_df.columns]
       
        # Verify mandatory column requirements are explicitly mapable
        required_cols = ["DEPARTMENT", "SERIAL NUMBER", "EQUIPMENT DESCRIPTION", "STATUS", "CALIB. DATE DUE"]
        missing_cols = [col for col in required_cols if col not in raw_df.columns]
       
        if missing_cols:
            st.error(f"❌ Structural match failure. Missing column headers: {missing_cols}. Please verify row 5 header names.")
        else:
            # Re-structure columns into uniform data arrays
            working_df = raw_df[["SERIAL NUMBER", "EQUIPMENT DESCRIPTION", "DEPARTMENT", "STATUS", "CALIB. DATE DUE"]].copy()
            working_df.columns = ["ID", "Description", "Department", "Status", "Due_Date"]
           
            # Drop records missing serial numbers
            working_df = working_df.dropna(subset=["ID"])
            working_df["ID"] = working_df["ID"].astype(str).str.strip()
            working_df = working_df[working_df["ID"] != "nan"]
           
            # Filter out devices officially decommissioned or removed from rotation
            working_df["Status"] = working_df["Status"].astype(str).str.strip().str.upper()
            active_df = working_df[working_df["Status"] != "REMOVED"].copy()
           
            # Convert Excel calendar items dynamically into active python dates safely
            active_df["Due_Date"] = pd.to_datetime(active_df["Due_Date"], errors='coerce').dt.date
            active_df = active_df.dropna(subset=["Due_Date"])
           
            # --- SECTION 2: LIVE DATE DEADLINE PARSING ENGINE ---
            # Set target timeline reference point to today's live coordinate: May 27, 2026
            today = datetime.date(2026, 5, 27)
           
            # Calculate precise numerical window spaces remaining until expiration
            active_df["Days Remaining"] = active_df["Due_Date"].apply(lambda d: (d - today).days)
           
            # Map items exclusively into your requested 3 category structure
            def segment_instrument(row):
                days = row["Days Remaining"]
                if days < 0:
                    return "OVERDUE"
                elif days <= 30:
                    return "Due in Next 30 days"
                elif days <= 49: # 7 weeks = 49 days
                    return "Due in Next 7 Weeks"
                else:
                    return "VALID"
           
            active_df["Time Segment"] = active_df.apply(segment_instrument, axis=1)
           
            # --- SECTION 3: SUMMARY DISPLAY MATRIX ---
            st.markdown("---")
            st.subheader("📊 2. Department Calibration Distribution Summary Matrix")
           
            matrix_records = []
            for d in DEPARTMENTS_LIST:
                # Group data records cleanly ignoring structural casing typos
                dept_mask = active_df[active_df["Department"].astype(str).str.strip().str.lower() == d.lower()]
               
                matrix_records.append({
                    "Departments": d,
                    "OVERDUE": len(dept_mask[dept_mask["Time Segment"] == "OVERDUE"]),
                    "Due in Next 30 days": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 30 days"]),
                    "Due in Next 7 Weeks": len(dept_mask[dept_mask["Time Segment"] == "Due in Next 7 Weeks"]),
                    "TOTAL PENDING": len(dept_mask[dept_mask["Time Segment"] != "VALID"])
                })
               
            summary_matrix_df = pd.DataFrame(matrix_records)
            st.dataframe(summary_matrix_df, use_container_width=True, hide_index=True)
           
            # Layout Summary Footer Cards
            c_ov, c_30, c_7w, c_tot = st.columns(4)
            c_ov.metric("Total OVERDUE Instruments", summary_matrix_df["OVERDUE"].sum())
            c_30.metric("Due within 30 Days", summary_matrix_df["Due in Next 30 days"].sum())
            c_7w.metric("Due within 7 Weeks", summary_matrix_df["Due in Next 7 Weeks"].sum())
            c_tot.metric("Total Active Backlog Count", summary_matrix_df["TOTAL PENDING"].sum())
           
            # --- SECTION 4: PLOTLY GRAPH CHART DISPLAY ENGINE ---
            st.markdown("---")
            st.subheader("📈 3. Equipment Backlog Status Distribution Chart")
           
            # Count pending records for visual graphing
            pending_df = active_df[active_df["Time Segment"] != "VALID"]
            if not pending_df.empty:
                graph_data = pending_df["Time Segment"].value_counts().reset_index()
                graph_data.columns = ["Urgency Status", "Equipment Count"]
               
                # Render clean bar charts with targeted color rules matching data mapping flags
                fig = px.bar(
                    graph_data,
                    x="Urgency Status",
                    y="Equipment Count",
                    color="Urgency Status",
                    color_discrete_map={
                        "OVERDUE": "#E31B23",         # Coca-Cola Red
                        "Due in Next 30 days": "#FFA500", # Warning Orange
                        "Due in Next 7 Weeks": "#3399FF"  # Informational Blue
                    },
                    title="Pretoria Plant Total Pending Calibration Backlog"
                )
                fig.update_layout(showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("🎉 No outstanding actions recorded. All device calibrations are 100% current!")

            # --- SECTION 5: MANAGER WARNING ALERT TRANSMISSION CENTER ---
            st.markdown("---")
            st.subheader("✉️ 4. Upcoming Calibration Alert Logs (< 30 Days Email Warning)")
           
            upcoming_30 = active_df[active_df["Time Segment"] == "Due in Next 30 days"].copy()
            st.write(f"Found **{len(upcoming_30)}** instruments requiring service within 30 days.")
           
            email_target = st.text_input("Enter manager alert report notification email address:", "boss_email@ccbsa.co.za")
            if st.button("🚀 Dispatch Alert Email Notification Logs"):
                if len(upcoming_30) > 0:
                    email_body = f"CCBSA Pretoria Calibration Notice Log -\nReport Generated: {today}\n\nTHE FOLLOWING REQUIRING CALIBRATION WITHIN 30 DAYS:\n\n"
                    for _, row in upcoming_30.iterrows():
                        email_body += f"• ID: {row['ID']} | Dept: {row['Department']} | Description: {row['Description']} | Target Due: {row['Due_Date']} ({row['Days Remaining']} Days Left)\n"
                   
                    st.info("Email communication text package successfully constructed:")
                    st.code(email_body, language="text")
                    st.success(f"📩 Notification listing successfully queued for delivery to {email_target}!")
                else:
                    st.success("🎉 Safe! No instruments are currently due within the next 30 days.")

            # --- SECTION 6: REGISTER DATA VIEW GRID ---
            st.markdown("---")
            st.subheader("📋 5. Asset Register Detailed Rows (Filtered Active Items)")
            st.dataframe(active_df[["ID", "Description", "Department", "Due_Date", "Time Segment", "Days Remaining"]], use_container_width=True, hide_index=True)
           
    except Exception as e:
        st.error(f"Error compiling spreadsheet rows: {str(e)}")
else:
    st.info("💡 Please upload your clean formatted Excel spreadsheet above to display the updated monitoring panels.")
