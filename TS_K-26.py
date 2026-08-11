import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import plotly.graph_objects as go
import gspread
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from zoneinfo import ZoneInfo

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
    padding-top:0.8rem;
    padding-bottom:1rem;
    padding-left:2rem;
    padding-right:2rem;
}

.kpi-card{
    background:#FFFFFF;
    border:1px solid #E2E8E3;
    border-radius:14px;
    padding:14px 16px 12px 16px;
    margin:0 0 8px 0;
    box-shadow:0 2px 8px rgba(0,0,0,.06);
}

.kpi-title{
    font-size:1.15rem;
    font-weight:700;
    color:#243024;
    margin-bottom:2px;
}

.kpi-percent{
    font-size:2rem;
    font-weight:800;
    color:#2E7D32;
    line-height:1.1;
    margin:4px 0 7px 0;
}

.kpi-progress{
    height:9px;
    width:100%;
    background:#E8F1E8;
    border-radius:10px;
    overflow:hidden;
    margin-bottom:9px;
}

.kpi-progress-fill{
    height:100%;
    background:#2E7D32;
    border-radius:10px;
}

.kpi-row{
    display:flex;
    justify-content:space-between;
    align-items:center;
    font-size:.88rem;
    color:#777;
    padding:3px 0;
}

.kpi-row b{
    color:#263238;
}

.kpi-row.achieved b{
    color:#2E7D32;
}

.kpi-row.pending b{
    color:#D32F2F;
}

[data-testid="stVerticalBlock"] > div:has(> .kpi-card){
    gap:0 !important;
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

from datetime import datetime
from zoneinfo import ZoneInfo

# Get current India Standard Time
india_time = datetime.now(ZoneInfo("Asia/Kolkata"))

st.markdown(
    "<h1 style='margin:0; padding:0;'>🌾 TS Kharif-26 Monitoring Dashboard</h1>",
    unsafe_allow_html=True
)
st.caption(
    f"Last Updated : {india_time.strftime('%d-%b-%Y %I:%M %p')}"
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
def _style_excel_sheet(ws, table_ranges=None, freeze_panes="A2"):
    """Apply a clean professional style to an Excel worksheet."""

    # Professional dashboard colors
    dark_green = "1B5E20"
    green = "2E7D32"
    light_green = "E8F5E9"
    very_light_green = "F6FBF6"
    white = "FFFFFF"
    dark_text = "263238"
    grey = "6B7280"
    light_grey = "E5E7EB"
    red = "D32F2F"
    orange = "EF6C00"

    thin = Side(style="thin", color=light_grey)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = freeze_panes
    ws.sheet_view.showGridLines = False

    # Default row height
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 20

    # Style section headings / normal headers
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue

            cell.font = Font(
                name="Calibri",
                size=10,
                color=dark_text
            )
            cell.alignment = Alignment(
                vertical="center",
                horizontal="left",
                wrap_text=True
            )

            # Section title rows
            if (
                isinstance(cell.value, str)
                and cell.column == 1
                and cell.value in {
                    "STUDY-WISE PROGRESS",
                    "STUDY-WISE CULTIVATION",
                    "CLUSTER-WISE PROGRESS",
                    "CLUSTER-WISE CULTIVATION"
                }
            ):
                cell.font = Font(
                    name="Calibri",
                    size=12,
                    bold=True,
                    color=white
                )
                cell.fill = PatternFill(
                    "solid",
                    fgColor=green
                )
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )
                ws.row_dimensions[cell.row].height = 24

    # Detect table header rows and style them
    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if any(v in values for v in [
            "Target_Farmers",
            "Target_Plots",
            "Target_Hectares",
            "Target_Ha",
            "Field Officer",
            "Kharif-26 DSR/Nursery Sowing Date",
            "Study2",
            "Cluster"
        ]):
            for cell in row:
                if cell.value is not None:
                    cell.font = Font(
                        name="Calibri",
                        size=10,
                        bold=True,
                        color=white
                    )
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=dark_green
                    )
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True
                    )
                    cell.border = border
            ws.row_dimensions[row[0].row].height = 30

    # Apply borders, alignment and number formats
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue

            cell.border = border

            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            # Percentage columns
            header = ws.cell(1, cell.column).value
            if isinstance(header, str) and "%" in header:
                cell.number_format = '0.00"%"'

            # Ha / numeric hectare columns
            if isinstance(header, str) and (
                "Hectares" in header
                or header.endswith("_Ha")
                or "Ha)" in header
            ):
                cell.number_format = '#,##0.00'

    # Green/red emphasis for achievement and pending percentages
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in ("Achieved %", "Pending %"):
                for data_cell in ws[cell.row + 1]:
                    if data_cell.value is not None:
                        if cell.value == "Achieved %":
                            data_cell.font = Font(
                                name="Calibri",
                                size=10,
                                bold=True,
                                color=green
                            )
                        else:
                            data_cell.font = Font(
                                name="Calibri",
                                size=10,
                                bold=True,
                                color=red
                            )

    # Alternating row fill for data areas
    for row in ws.iter_rows():
        if row[0].row <= 1:
            continue
        if row[0].row % 2 == 0:
            for cell in row:
                if cell.value is not None and cell.fill.fill_type is None:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=very_light_green
                    )

    # Add Excel tables where valid ranges were supplied
    if table_ranges:
        for idx, ref in enumerate(table_ranges, start=1):
            try:
                table = Table(
                    displayName=f"ReportTable{idx}",
                    ref=ref
                )
                style = TableStyleInfo(
                    name="TableStyleMedium4",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                table.tableStyleInfo = style
                ws.add_table(table)
            except Exception:
                # Styling should never stop the dashboard from exporting.
                pass

    # Smart column widths
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_length = 0

        for cell in col_cells:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[col_letter].width = min(
            max(max_length + 2, 11),
            30
        )

    # Keep very wide raw-data columns manageable
    if ws.title == "Raw Data":
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = min(
                ws.column_dimensions[get_column_letter(col)].width,
                24
            )

    # Conditional formatting for percentage columns
    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col).value
        if isinstance(header, str) and header in ("Achieved %", "Pending %"):
            letter = get_column_letter(col)
            if ws.max_row > 1:
                ws.conditional_formatting.add(
                    f"{letter}2:{letter}{ws.max_row}",
                    ColorScaleRule(
                        start_type="min",
                        start_color="FEE2E2",
                        mid_type="percentile",
                        mid_value=50,
                        mid_color="FEF3C7",
                        end_type="max",
                        end_color="DCFCE7"
                    )
                )


def _style_study_cluster_sheet(ws):
    """Format the combined Study & Cluster worksheet with clear sections."""

    dark_green = "1B5E20"
    green = "2E7D32"
    white = "FFFFFF"
    light_green = "E8F5E9"
    light_red = "FFEBEE"
    border_color = "DDE5DD"

    thin = Side(style="thin", color=border_color)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    section_names = {
        "STUDY-WISE PROGRESS",
        "STUDY-WISE CULTIVATION",
        "CLUSTER-WISE PROGRESS",
        "CLUSTER-WISE CULTIVATION"
    }

    section_rows = []
    header_rows = []

    for r in range(1, ws.max_row + 1):
        value = ws.cell(r, 1).value

        if value in section_names:
            section_rows.append(r)

            ws.merge_cells(
                start_row=r,
                start_column=1,
                end_row=r,
                end_column=max(1, ws.max_column)
            )

            cell = ws.cell(r, 1)
            cell.fill = PatternFill("solid", fgColor=green)
            cell.font = Font(
                name="Calibri",
                size=12,
                bold=True,
                color=white
            )
            cell.alignment = Alignment(
                horizontal="left",
                vertical="center"
            )
            ws.row_dimensions[r].height = 26

            continue

        # Identify header immediately after a section title
        if r > 1 and (r - 1) in section_rows:
            header_rows.append(r)

    # Style headers and data
    for r in range(1, ws.max_row + 1):
        if r in section_rows:
            continue

        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)

            if cell.value is None:
                continue

            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )
            cell.font = Font(
                name="Calibri",
                size=10,
                color="263238"
            )
            cell.alignment = Alignment(
                vertical="center",
                horizontal="right" if isinstance(cell.value, (int, float)) else "left",
                wrap_text=True
            )

    for r in header_rows:
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if cell.value is not None:
                cell.fill = PatternFill(
                    "solid",
                    fgColor=dark_green
                )
                cell.font = Font(
                    name="Calibri",
                    size=10,
                    bold=True,
                    color=white
                )
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )
                cell.border = Border(
                    left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin
                )
        ws.row_dimensions[r].height = 32

    # Highlight achievement/pending percentage columns
    for r in header_rows:
        for c in range(1, ws.max_column + 1):
            header = ws.cell(r, c).value
            if header in ("Achieved %", "Pending %"):
                for rr in range(r + 1, ws.max_row + 1):
                    # Stop at the next section
                    if rr in section_rows:
                        break
                    cell = ws.cell(rr, c)
                    if cell.value is not None:
                        cell.number_format = '0.00"%"'
                        cell.font = Font(
                            name="Calibri",
                            size=10,
                            bold=True,
                            color=green if header == "Achieved %" else "D32F2F"
                        )

    # Number formats throughout the combined report
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            header_values = []
            for hr in header_rows:
                if hr < r:
                    header_values.append(ws.cell(hr, c).value)

            latest_header = header_values[-1] if header_values else None

            if isinstance(cell.value, (int, float)):
                if latest_header == "Achieved %":
                    cell.number_format = '0.00"%"'
                elif latest_header == "Pending %":
                    cell.number_format = '0.00"%"'
                elif latest_header and (
                    "Hectares" in str(latest_header)
                    or str(latest_header).endswith("_Ha")
                ):
                    cell.number_format = '#,##0.00'

    # Alternate data rows within each section
    for r in range(1, ws.max_row + 1):
        if r in section_rows or r in header_rows:
            continue
        if r % 2 == 0:
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value is not None:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="F7FBF7"
                    )

    # Column widths
    widths = {
        1: 26, 2: 16, 3: 16, 4: 18,
        5: 18, 6: 18, 7: 18, 8: 18,
        9: 18, 10: 18, 11: 18
    }

    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(c, 18)


def convert_dashboard_excel(
    study_summary,
    cluster_summary,
    study_cp,
    cluster_cp,
    fo_summary,
    timeline_data,
    raw_data
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ONE clean worksheet for all Study & Cluster information
        sheet_name = "Study & Cluster"

        start_row = 0

        pd.DataFrame([["STUDY-WISE PROGRESS"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False,
            header=False
        )
        start_row += 1

        study_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False
        )
        start_row += len(study_summary) + 3

        pd.DataFrame([["STUDY-WISE CULTIVATION"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False,
            header=False
        )
        start_row += 1

        study_cp.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False
        )
        start_row += len(study_cp) + 3

        pd.DataFrame([["CLUSTER-WISE PROGRESS"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False,
            header=False
        )
        start_row += 1

        cluster_summary.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False
        )
        start_row += len(cluster_summary) + 3

        pd.DataFrame([["CLUSTER-WISE CULTIVATION"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False,
            header=False
        )
        start_row += 1

        cluster_cp.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row,
            index=False
        )

        fo_summary.to_excel(
            writer,
            sheet_name="Field Officer",
            index=False
        )

        timeline_data.to_excel(
            writer,
            sheet_name="Timeline",
            index=False
        )

        raw_data.to_excel(
            writer,
            sheet_name="Raw Data",
            index=False
        )

    # ======================================================
    # PROFESSIONAL EXCEL FORMATTING
    # ======================================================

    wb = load_workbook(io.BytesIO(output.getvalue()))

    # Combined Study & Cluster sheet
    _style_study_cluster_sheet(
        wb["Study & Cluster"]
    )

    # Other sheets
    for sheet_name in ["Field Officer", "Timeline", "Raw Data"]:
        ws = wb[sheet_name]

        if sheet_name == "Raw Data":
            _style_excel_sheet(
                ws,
                table_ranges=None,
                freeze_panes="A2"
            )
        else:
            _style_excel_sheet(
                ws,
                table_ranges=None,
                freeze_panes="A2"
            )

    # Field Officer: make achievement percentage visually stronger
    ws = wb["Field Officer"]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Achieved %":
                for r in range(cell.row + 1, ws.max_row + 1):
                    ws.cell(r, cell.column).number_format = '0.00"%"'
                    ws.cell(r, cell.column).font = Font(
                        name="Calibri",
                        size=10,
                        bold=True,
                        color="2E7D32"
                    )

    # Timeline number formats
    ws = wb["Timeline"]
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)):
                if cell.column == 3:
                    cell.number_format = '#,##0.00'
    ws.freeze_panes = "A2"

    # Workbook metadata
    wb.properties.title = "TS Kharif-26 Overall Dashboard Report"
    wb.properties.subject = "TS Kharif-26 Sowing Progress"
    wb.properties.creator = "VIS Dashboard"

    final_output = io.BytesIO()
    wb.save(final_output)

    return final_output.getvalue()


# ==========================================================
# MASTER DOWNLOAD DATA INITIALIZATION
# ==========================================================

daily = pd.DataFrame()
# ==========================================================
# MASTER DOWNLOAD DATA
# ==========================================================

# Study Summary
master_study_summary = cultivation_summary("Study2")

# Cluster Summary
master_cluster_summary = cultivation_summary("Cluster")

# Study Cultivation
master_study_cp = cultivation_practice_summary("Study2")

# Cluster Cultivation
master_cluster_cp = cultivation_practice_summary("Cluster")

# Field Officer
master_fo_summary = cultivation_summary("Field Officer")


# Timeline data
master_timeline_df = filtered_df.copy()

date_col = "Kharif-26 DSR/Nursery Sowing Date"

master_timeline_df = master_timeline_df[
    master_timeline_df[date_col].notna()
]

if not master_timeline_df.empty:

    master_daily = (
        master_timeline_df
        .groupby(date_col)
        .agg(
            Plots=("Plot Codes", "nunique"),
            Hectares=("Kharif-26 Hectares", "sum")
        )
        .reset_index()
        .sort_values(date_col)
    )

else:

    master_daily = pd.DataFrame(
        columns=[
            date_col,
            "Plots",
            "Hectares"
        ]
    )
    # ==========================================================
# TOP RIGHT MASTER DOWNLOAD
# ==========================================================

top_left, top_right = st.columns([9, 1.5])

with top_right:

    master_excel = convert_dashboard_excel(
        study_summary=master_study_summary,
        cluster_summary=master_cluster_summary,
        study_cp=master_study_cp,
        cluster_cp=master_cluster_cp,
        fo_summary=master_fo_summary,
        timeline_data=master_daily,
        raw_data=filtered_df
    )

    st.download_button(
        label="📥 Download Overall Data",
        data=master_excel,
        file_name="TS_Kharif_26_Overall_Master.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="master_download_top",
        use_container_width=True
    )
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
def summary_card(title, icon, target, achieved, unit="", decimals=0):

    pending = max(target - achieved, 0)
    percent = (achieved / target * 100) if target > 0 else 0
    percent = min(percent, 100)

    if decimals == 0:
        target_text = f"{target:,.0f}"
        achieved_text = f"{achieved:,.0f}"
        pending_text = f"{pending:,.0f}"
    else:
        target_text = f"{target:,.2f}{unit}"
        achieved_text = f"{achieved:,.2f}{unit}"
        pending_text = f"{pending:,.2f}{unit}"

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{icon} {title}</div>
            <div class="kpi-percent">{percent:.1f}%</div>
            <div class="kpi-progress">
                <div class="kpi-progress-fill" style="width:{percent:.1f}%"></div>
            </div>
            <div class="kpi-row">
                <span>🎯 Target</span><b>{target_text}</b>
            </div>
            <div class="kpi-row achieved">
                <span>✅ Achieved</span><b>{achieved_text}</b>
            </div>
            <div class="kpi-row pending">
                <span>⏳ Pending</span><b>{pending_text}</b>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
# ==========================================================
# OVERVIEW TAB
# ==========================================================
with overview_tab:

    st.subheader("📊 Dashboard Overview")

    c1, c2, c3 = st.columns(3)

    achieved_farmers = active_df["Farmer Code"].nunique()

    with c1:
        summary_card(
            title="Farmers",
            icon="👨‍🌾",
            target=target_farmers,
            achieved=achieved_farmers
        )
    with c2:
                summary_card(
                    title="Plots",
                    icon="📍",
                    target=target_plots,
                    achieved=achieved_plots
                )

    with c3:
        summary_card(
            title="Area (Ha)",
            icon="🌱",
            target=target_hectares,
            achieved=achieved_hectares,
            unit=" Ha",
            decimals=2
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
            height=330,
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
            height=330,
            legend_title=""
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

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
            height=360,
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
            height=360,
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


# ==========================================================
# OVERALL MASTER DOWNLOAD
# ==========================================================

st.divider()

st.caption(
    "Developed by Jayapal Thoka"
)