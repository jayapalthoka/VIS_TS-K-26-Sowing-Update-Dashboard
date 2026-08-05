import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import plotly.graph_objects as go
import gspread
import io
import os

from datetime import datetime
from google.oauth2.service_account import Credentials
from streamlit_autorefresh import st_autorefresh

# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(
    page_title="VIS TS Kharif-26 Dashboard",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Auto Refresh Every 60 Seconds
st_autorefresh(interval=60000, key="dashboard_refresh")

# ==========================================================
# CUSTOM CSS
# ==========================================================

st.markdown("""
<style>

.main{
    background:#F4F7FC;
}

.block-container{
    padding-top:1rem;
    padding-bottom:1rem;
}

div[data-testid="metric-container"]{
    background:white;
    border-radius:12px;
    padding:15px;
    box-shadow:0 3px 10px rgba(0,0,0,.08);
    border-left:5px solid #1f77b4;
}

section[data-testid="stSidebar"]{
    background:white;
}

thead tr th{
    background:#12355B !important;
    color:white !important;
    text-align:center !important;
}

tbody td{
    text-align:center;
}

</style>
""", unsafe_allow_html=True)
# ==========================================
# DASHBOARD HEADER
# ==========================================

col1, col2 = st.columns([1,6])

with col1:
    # st.image("logo.png", width=90)
    pass

with col2:
    st.title("🌾 TS Kharif-26 Monitoring Dashboard")
    st.caption(
        f"Last Updated : {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    )

st.divider()

# ==========================================================
# GOOGLE SHEETS CONFIGURATION
# ==========================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

#JSON_PATH = os.path.join(BASE_DIR, "service_account.json")

SPREADSHEET_ID = "1XNzggS3DKn8ucffUx1QKBlrDXVLigLg3psR2WHdyR9I"

WORKSHEET_NAME = "Overall Farmer List - Kharif-26"

# ==========================================================
# LOAD DATA
# ==========================================================

@st.cache_data(ttl=60)
def load_data():

    creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)

    client = gspread.authorize(creds)

    worksheet = (
        client
        .open_by_key(SPREADSHEET_ID)
        .worksheet(WORKSHEET_NAME)
    )

    df = pd.DataFrame(
        worksheet.get_all_records()
    )

    # Clean Column Names
    df.columns = (
        df.columns
        .str.strip()
        .str.replace("\n"," ",regex=False)
        .str.replace(r"\s+"," ",regex=True)
    )

    # Numeric Columns
    numeric_columns = [
        "K/R-25 Hectares",
        "Kharif-26 Hectares",
        "Kharif-26 Acreage"
    ]

    for col in numeric_columns:

        if col in df.columns:

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)

    # Date Columns
    if "Kharif-26 DSR/Nursery Sowing Date" in df.columns:

        df["Kharif-26 DSR/Nursery Sowing Date"] = pd.to_datetime(
            df["Kharif-26 DSR/Nursery Sowing Date"],
            errors="coerce",
            dayfirst=True
        )

    # Active / Inactive
    STATUS_COL = "Kharif-26 Plot Status (Active / Inactive )"

    if STATUS_COL in df.columns:

        df[STATUS_COL] = (
            df[STATUS_COL]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    return df

# ==========================================================
# READ DATA
# ==========================================================

try:

     with st.spinner("Loading Dashboard..."):
        df = load_data()

except Exception as e:

    st.error(f"Unable to load Google Sheet.\n\n{e}")

    st.stop()
    # ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.image(
    "https://img.icons8.com/color/96/rice-bowl.png",
    width=80
)

st.sidebar.title("Dashboard Filters")

# ----------------------------
# Filters
# ----------------------------

study_filter = st.sidebar.multiselect(
    "Study",
    sorted(df["Study2"].dropna().unique())
)

cluster_filter = st.sidebar.multiselect(
    "Cluster",
    sorted(df["Cluster"].dropna().unique())
)

tahsil_filter = st.sidebar.multiselect(
    "Tahsil",
    sorted(df["Tahsil"].dropna().unique())
)

fo_filter = st.sidebar.multiselect(
    "Field Officer",
    sorted(df["Field Officer"].dropna().unique())
)

cp_filter = st.sidebar.multiselect(
    "Cultivation Practice",
    sorted(
        df["Kharif-26 Cultivation Practice"]
        .dropna()
        .astype(str)
        .unique()
    )
    
)

# ==========================================================
# APPLY FILTERS
# ==========================================================

filtered_df = df.copy()

if study_filter:
    filtered_df = filtered_df[
        filtered_df["Study2"].isin(study_filter)
    ]

if cluster_filter:
    filtered_df = filtered_df[
        filtered_df["Cluster"].isin(cluster_filter)
    ]

if tahsil_filter:
    filtered_df = filtered_df[
        filtered_df["Tahsil"].isin(tahsil_filter)
    ]

if fo_filter:
    filtered_df = filtered_df[
        filtered_df["Field Officer"].isin(fo_filter)
    ]

if cp_filter:
    filtered_df = filtered_df[
        filtered_df["Kharif-26 Cultivation Practice"].isin(cp_filter)
    ]

# ==========================================================
# ACTIVE DATA
# ==========================================================

STATUS_COL = "Kharif-26 Plot Status (Active / Inactive )"

active_df = filtered_df[
    filtered_df[STATUS_COL] == "ACTIVE"
]

# ==========================================================
# KPI CALCULATIONS
# ==========================================================

target_farmers = filtered_df["Farmer Code"].nunique()

target_plots = filtered_df["Plot Codes"].nunique()

target_hectares = round(
    filtered_df["K/R-25 Hectares"].sum(),
    2
)

achieved_plots = active_df["Plot Codes"].nunique()

achieved_hectares = round(
    active_df["Kharif-26 Hectares"].sum(),
    2
)

pending_plots = target_plots - achieved_plots

pending_hectares = round(
    target_hectares - achieved_hectares,
    2
)

achievement_pct = round(
    (achieved_hectares / target_hectares * 100)
    if target_hectares > 0 else 0,
    2
)

pending_pct = round(
    100 - achievement_pct,
    2
)

# ==========================================================
# REUSABLE SUMMARY FUNCTION
# ==========================================================

def cultivation_summary(group_column):

    target = (
        filtered_df
        .groupby(group_column)
        .agg(
            Target_Farmers=("Farmer Code", "nunique"),
            Target_Plots=("Plot Codes", "nunique"),
            Target_Hectares=("K/R-25 Hectares", "sum")
        )
    )

    achieved = (
        active_df
        .groupby(group_column)
        .agg(
            Achieved_Plots=("Plot Codes", "nunique"),
            Achieved_Hectares=("Kharif-26 Hectares", "sum")
        )
    )

    summary = (
        target
        .join(achieved, how="left")
        .fillna(0)
    )

    summary["Pending_Plots"] = (
        summary["Target_Plots"] -
        summary["Achieved_Plots"]
    )

    summary["Pending_Hectares"] = (
        summary["Target_Hectares"] -
        summary["Achieved_Hectares"]
    )

    summary["Achieved %"] = (
        summary["Achieved_Hectares"] /
        summary["Target_Hectares"] * 100
    ).fillna(0).round(2)

    summary["Pending %"] = (
        100 - summary["Achieved %"]
    ).round(2)

    summary = summary.reset_index()

    return summary

# ==========================================================
# DOWNLOAD FUNCTION
# ==========================================================
def cultivation_practice_summary(group_column):

    target = (
        filtered_df.groupby(group_column)
        .agg(
            Target_Ha=("K/R-25 Hectares", "sum")
        )
    )

    achieved = (
        active_df.groupby(group_column)
        .agg(
            Achieved_Ha=("Kharif-26 Hectares", "sum")
        )
    )

    summary = target.join(achieved, how="left").fillna(0)

    dry = (
        active_df[
            active_df["Kharif-26 Cultivation Practice"] == "Dry DSR"
        ]
        .groupby(group_column)["Kharif-26 Hectares"]
        .sum()
    )

    wet = (
        active_df[
            active_df["Kharif-26 Cultivation Practice"] == "WET DSR"
        ]
        .groupby(group_column)["Kharif-26 Hectares"]
        .sum()
    )

    tpr = (
        active_df[
            active_df["Kharif-26 Cultivation Practice"] == "Transplanting+AWD"
        ]
        .groupby(group_column)["Kharif-26 Hectares"]
        .sum()
    )

    summary["Dry DSR (Ha)"] = dry
    summary["WET DSR (Ha)"] = wet
    summary["TPR+AWD (Ha)"] = tpr

    summary = summary.fillna(0)

    summary["Achieved %"] = (
        summary["Achieved_Ha"] / summary["Target_Ha"] * 100
    ).fillna(0).round(2)

    summary["Dry DSR %"] = (
        summary["Dry DSR (Ha)"] / summary["Achieved_Ha"] * 100
    ).fillna(0).round(2)

    summary["WET DSR %"] = (
        summary["WET DSR (Ha)"] / summary["Achieved_Ha"] * 100
    ).fillna(0).round(2)

    summary["TPR+AWD %"] = (
        summary["TPR+AWD (Ha)"] / summary["Achieved_Ha"] * 100
    ).fillna(0).round(2)

    return summary.reset_index()
@st.cache_data
def convert_excel(data):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        data.to_excel(
            writer,
            sheet_name="Filtered Data",
            index=False
        )

    return output.getvalue()
@st.cache_data
def convert_dashboard_excel(
    study_summary,
    cluster_summary,
    study_cp,
    cluster_cp,
    raw_data
):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        study_summary.to_excel(
            writer,
            sheet_name="Study Summary",
            index=False
        )

        cluster_summary.to_excel(
            writer,
            sheet_name="Cluster Summary",
            index=False
        )

        study_cp.to_excel(
            writer,
            sheet_name="Study Cultivation",
            index=False
        )

        cluster_cp.to_excel(
            writer,
            sheet_name="Cluster Cultivation",
            index=False
        )

        raw_data.to_excel(
            writer,
            sheet_name="Raw Data",
            index=False
        )

    return output.getvalue()

# ==========================================================
# TABS
# ==========================================================

overview_tab, study_tab, fo_tab, timeline_tab, data_tab = st.tabs(
    [
        "📊 Overview",
        "🧭 Study & Cluster",
        "👨‍🌾 Field Officer",
        "📅 Timeline",
        "📋 Data"
    ]
)
# ==========================================================
# OVERVIEW TAB
# ==========================================================

with overview_tab:

    st.subheader("📊 Dashboard Overview")

    # ======================================================
    # KPI CARDS
    # ======================================================

    k1, k2, k3, k4, k5, k6 = st.columns(6)

    k1.metric(
        "👨‍🌾 Farmers",
        f"{target_farmers:,}"
    )

    k2.metric(
        "🌾 Target Plots",
        f"{target_plots:,}"
    )

    k3.metric(
        "🌱 Target Ha",
        f"{target_hectares:,.2f}"
    )

    k4.metric(
        "✅ Achieved Ha",
        f"{achieved_hectares:,.2f}",
        f"{achievement_pct:.2f}%"
    )

    k5.metric(
        "⏳ Pending Ha",
        f"{pending_hectares:,.2f}",
        f"{pending_pct:.2f}%"
    )

    k6.metric(
        "📈 Achievement",
        f"{achievement_pct:.2f}%"
    )

    st.divider()

    # ======================================================
    # DONUT CHARTS
    # ======================================================

    c1, c2 = st.columns(2)

    # ======================================================
    # TARGET vs ACHIEVED HECTARES
    # ======================================================

    with c1:

        donut = pd.DataFrame({
            "Status": ["Achieved", "Pending"],
            "Hectares": [
                achieved_hectares,
                pending_hectares
            ]
        })

        fig = px.pie(
            donut,
            names="Status",
            values="Hectares",
            hole=0.65,
            color="Status",
            color_discrete_map={
                "Achieved": "#2E7D32",
                "Pending": "#EF6C00"
            }
        )

        fig.update_traces(
            textposition="inside",
            textinfo="percent+label+value",
            texttemplate="%{label}<br>%{value:.2f} Ha<br>%{percent}",
            textfont_size=15
        )

        fig.update_layout(
            title="<b>Target vs Achieved Hectares</b>",
            height=430,
            legend_title=""
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # ======================================================
    # CULTIVATION PRACTICE
    # ======================================================

    with c2:

        cp = (
            filtered_df["Kharif-26 Cultivation Practice"]
           .replace([0, "0", ""], "Pending")
    .fillna("Pending")
    .value_counts()
    .reset_index()
        )

        cp.columns = [
            "Cultivation Practice",
            "Plots"
        ]

        fig = px.pie(
    cp,
    names="Cultivation Practice",
    values="Plots",
    hole=0.65,
    color="Cultivation Practice",
    color_discrete_map={
    "Transplanting+AWD": "#EF9A9A",   # Light Red
    "WET DSR": "#66BB6A",             # Medium Green
    "Dry DSR": "#1B5E20",             # Dark Green
    "Pending": "#D32F2F"              # Red
}
)

        fig.update_traces(
            textposition="inside",
            textinfo="percent+label",
            textfont_size=15
        )

        fig.update_layout(
            title="<b>Cultivation Practice</b>",
            height=430,
            legend_title=""
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    st.divider()
    # ======================================================
    # OVERALL PROGRESS
    # ======================================================
    st.subheader("Overall Cultivation Progress")

    st.progress(achievement_pct / 100)

    left, right = st.columns(2)

    with left:

        st.success(
            f"✅ Achieved : {achieved_hectares:,.2f} Ha ({achievement_pct:.2f}%)"
        )

    with right:

        st.error(
            f"❌ Pending : {pending_hectares:,.2f} Ha ({pending_pct:.2f}%)"
        )
      # ==========================================================
# STUDY & CLUSTER TAB
# ==========================================================

with study_tab:

    st.subheader("🧭 Study & Cluster Cultivation Progress")

    # ======================================================
    # STUDY SUMMARY
    # ======================================================

    st.markdown("### 📊 Study-wise Progress")

    study_summary = cultivation_summary("Study2")

    styled_study = (
        study_summary.style
        .format({
            "Target_Hectares": "{:.2f}",
            "Achieved_Hectares": "{:.2f}",
            "Pending_Hectares": "{:.2f}",
            "Achieved %": "{:.2f}%",
            "Pending %": "{:.2f}%"
        })
        .map(
            lambda _: "color:green;font-weight:bold;",
            subset=["Achieved %"]
        )
        .map(
            lambda _: "color:red;font-weight:bold;",
            subset=["Pending %"]
        )
    )

    st.dataframe(
        styled_study,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇ Download Study Summary",
        data=convert_excel(study_summary),
        file_name="Study_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()
    st.markdown("### 🌾 Study-wise Cultivation Summary")

    study_cp = cultivation_practice_summary("Study2")

    st.dataframe(
        study_cp,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇ Download Study Cultivation Summary",
        data=convert_excel(study_cp),
        file_name="Study_Cultivation_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    st.markdown("### 🌾 Cluster-wise Cultivation Summary")

    cluster_cp = cultivation_practice_summary("Cluster")

    st.dataframe(
        cluster_cp,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇ Download Cluster Cultivation Summary",
        data=convert_excel(cluster_cp),
        file_name="Cluster_Cultivation_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()
    

    # ======================================================
    # CLUSTER SUMMARY
    # ======================================================

    st.markdown("### 📊 Cluster-wise Progress")

    cluster_summary = cultivation_summary("Cluster")

    styled_cluster = (
        cluster_summary.style
        .format({
            "Target_Hectares": "{:.2f}",
            "Achieved_Hectares": "{:.2f}",
            "Pending_Hectares": "{:.2f}",
            "Achieved %": "{:.2f}%",
            "Pending %": "{:.2f}%"
        })
        .map(
            lambda _: "color:green;font-weight:bold;",
            subset=["Achieved %"]
        )
        .map(
            lambda _: "color:red;font-weight:bold;",
            subset=["Pending %"]
        )
    )

    st.dataframe(
        styled_cluster,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇ Download Cluster Summary",
        data=convert_excel(cluster_summary),
        file_name="Cluster_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # ==========================================================
# FIELD OFFICER TAB
# ==========================================================

with fo_tab:

    st.subheader("👨‍🌾 Field Officer Performance")

    fo_summary = cultivation_summary("Field Officer")

    styled_fo = (
        fo_summary.style
        .format({
            "Target_Hectares": "{:.2f}",
            "Achieved_Hectares": "{:.2f}",
            "Pending_Hectares": "{:.2f}",
            "Achieved %": "{:.2f}%",
            "Pending %": "{:.2f}%"
        })
        .map(
            lambda _: "color:green;font-weight:bold;",
            subset=["Achieved %"]
        )
        .map(
            lambda _: "color:red;font-weight:bold;",
            subset=["Pending %"]
        )
    )

    st.dataframe(
        styled_fo,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # ======================================================
    # FIELD OFFICER CHART
    # ======================================================

    fig = px.bar(
        fo_summary,
        x="Field Officer",
        y="Achieved %",
        text="Achieved %",
        color="Achieved %",
        color_continuous_scale="Greens",
        title="Field Officer Achievement %"
    )

    fig.update_traces(
        texttemplate="%{text:.2f}%",
        textposition="outside"
    )

    fig.update_layout(
        yaxis_title="Achievement %",
        xaxis_title="Field Officer",
        height=500
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.divider()

    st.download_button(
        "⬇ Download Field Officer Summary",
        data=convert_excel(fo_summary),
        file_name="Field_Officer_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # ==========================================================
# TIMELINE TAB
# ==========================================================

with timeline_tab:

    st.subheader("📅 Sowing Timeline")

    timeline_df = filtered_df.copy()

    date_col = "Kharif-26 DSR/Nursery Sowing Date"

    timeline_df = timeline_df[
        timeline_df[date_col].notna()
    ]

    if timeline_df.empty:

        st.warning("No sowing date data available.")

    else:

        # --------------------------------------------------
        # Daily Summary
        # --------------------------------------------------

        daily = (
            
            timeline_df
            .groupby(date_col)
            .agg(
                Plots=("Plot Codes", "nunique"),
                Hectares=("Kharif-26 Hectares", "sum")
            )
            .reset_index()
            .sort_values(date_col)
        )

        # --------------------------------------------------
        # KPIs
        # --------------------------------------------------

        k1, k2 = st.columns(2)

        k1.metric(
            "🌾 Total Sowing Plots",
            f"{daily['Plots'].sum():,}"
        )

        k2.metric(
            "🌱 Total Sown Hectares",
            f"{daily['Hectares'].sum():,.2f}"
        )

        st.divider()

        # --------------------------------------------------
        # Daily Plot Trend
        # --------------------------------------------------

        fig = px.line(
            daily,
            x=date_col,
            y="Plots",
            markers=True,
            title="Daily Sowing Progress (Plots)"
        )

        fig.update_layout(
            height=450,
            xaxis_title="Sowing Date",
            yaxis_title="Plots"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.divider()

        # --------------------------------------------------
        # Daily Hectares Trend
        # --------------------------------------------------

        fig = px.bar(
            daily,
            x=date_col,
            y="Hectares",
            text="Hectares",
            title="Daily Sowing Hectares"
        )

        fig.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside"
        )

        fig.update_layout(
            height=450,
            xaxis_title="Sowing Date",
            yaxis_title="Hectares"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.divider()

        st.dataframe(
            daily,
            use_container_width=True,
            hide_index=True
        )
        # ==========================================================
# DATA TAB
# ==========================================================

with data_tab:

    st.subheader("📋 Filtered Data")

    st.info(f"Showing **{len(filtered_df):,}** records")

    st.divider()

    # -----------------------------------------
    # Column Selector
    # -----------------------------------------

    selected_columns = st.multiselect(
        "Select Columns to Display",
        options=filtered_df.columns.tolist(),
        default=filtered_df.columns.tolist()
    )

    display_df = filtered_df[selected_columns]

    # -----------------------------------------
    # Search
    # -----------------------------------------

    search = st.text_input(
        "🔍 Search (works across all selected columns)"
    )

    if search:

        mask = display_df.astype(str).apply(
            lambda x: x.str.contains(search, case=False, na=False)
        ).any(axis=1)

        display_df = display_df[mask]

    # -----------------------------------------
    # Display Data
    # -----------------------------------------

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        height=600
    )

    st.divider()

    # -----------------------------------------
    # Download Button
    # -----------------------------------------

    st.download_button(
        label="⬇ Download Filtered Data",
        data=convert_excel(display_df),
        file_name="Filtered_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.divider()

st.caption(
    "Developed by Jayapal Thoka"
)