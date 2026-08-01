import io
import re
import numpy as np
import openpyxl
import pandas as pd
import streamlit as st

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="KPI Data Entry System", page_icon="📊", layout="wide"
)

st.title("📊 KPI Data Entry System")
st.caption(
    "Multi-sheet KPI data entry with strict periodicity validation, text evidence tracking, and comments."
)

# ============================================================
# CONSTANTS
# ============================================================

MONTH_COLUMNS = [str(i) for i in range(1, 13)]
EXPORT_PASSWORD = "KPi#Secure2026!Lock"

# ============================================================
# COLUMN ALIASES
# ============================================================

YEAR_ALIASES = ["Year", "year", "السنة"]
LOCATION_ALIASES = [
    "Location",
    "location",
    "Location Code",
    "location_code",
    "الموقع",
]
FREQUENCY_ALIASES = [
    "Periodicity",
    "periodicity",
    "Freq",
    "frequency",
    "Frequency",
    "التكرار",
]
KPI_CODE_ALIASES = ["KPI Code", "kpi_code", "Code", "code"]
KPI_NAME_ALIASES = ["KPI Name", "KPI Name (AR)", "kpi_name", "name"]
EVIDENCE_ALIASES = [
    "Evidence Status",
    "evidence_status",
    "Evidence",
    "evodence",
    "حالة الدليل",
    "Status",
    "status",
]
COMMENT_ALIASES = ["Comments", "Comment", "comments", "ملاحظات"]

# ============================================================
# HELPER FUNCTIONS
# ============================================================


def clean_header(col):
    if col is None:
        return ""
    col = str(col)
    col = col.replace("\xa0", " ")
    col = re.sub(r"[\n\r\t]", " ", col)
    return col.strip()


def find_column(df, aliases):
    for c in df.columns:
        if c in aliases:
            return c
    return None


def get_value(row, aliases, default=None):
    for c in aliases:
        if c in row.index:
            if pd.notna(row[c]):
                return row[c]
    return default


def allowed_months(periodicity):
    val = str(periodicity).strip().upper()

    if val in ["1", "M", "MONTHLY"]:
        return MONTH_COLUMNS
    elif val in ["2", "Q", "QUARTERLY"]:
        return ["3", "6", "9", "12"]
    elif val in ["3", "S", "SA", "SEMI ANNUAL", "SEMI-ANNUAL"]:
        return ["6", "12"]
    elif val in ["4", "A", "ANNUAL", "ANNUALLY", "Y", "YEARLY"]:
        return ["12"]

    try:
        p = int(float(val))
        if p == 1:
            return MONTH_COLUMNS
        elif p == 2:
            return ["3", "6", "9", "12"]
        elif p == 3:
            return ["6", "12"]
        elif p == 4:
            return ["12"]
    except Exception:
        pass

    return MONTH_COLUMNS


def clean_invalid_entries(df):
    df = df.copy()
    for m in MONTH_COLUMNS:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors="coerce").astype("float64")

    for idx in df.index:
        row = df.loc[idx]
        periodicity = get_value(row, FREQUENCY_ALIASES, "M")
        valid_months = allowed_months(periodicity)

        for month in MONTH_COLUMNS:
            if month not in df.columns:
                continue
            if month not in valid_months:
                df.loc[idx, month] = np.nan

    return df


def load_sheets(uploaded_file):
    excel = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheets = {}

    for sheet in excel.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
        df.columns = [clean_header(c) for c in df.columns]

        for col in df.columns:
            col_str = str(col).strip()
            if col_str in MONTH_COLUMNS:
                df[col] = pd.to_numeric(df[col], errors="coerce").astype(
                    "float64"
                )
            elif col_str in EVIDENCE_ALIASES or col_str in COMMENT_ALIASES:
                df[col] = df[col].fillna("").astype(str)

        df = clean_invalid_entries(df)
        sheets[sheet] = df

    return sheets


# ============================================================
# MAIN APPLICATION
# ============================================================

uploaded_file = st.file_uploader(
    "Upload KPI Excel Workbook", type=["xlsx", "xlsm"]
)

if uploaded_file:
    if (
        "kpi_sheets" not in st.session_state
        or st.session_state.get("filename") != uploaded_file.name
    ):
        st.session_state["kpi_sheets"] = load_sheets(uploaded_file)
        st.session_state["filename"] = uploaded_file.name

    sheets = st.session_state["kpi_sheets"]
    st.success("Workbook loaded successfully")

    available_sheets = list(sheets.keys())
    selected_sheet = st.selectbox("Select Entry Sheet", available_sheets)

    df = sheets[selected_sheet].copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = clean_invalid_entries(df)

    # Filters
    st.subheader("🔎 Filters")
    c1, c2 = st.columns(2)

    with c1:
        year_col = find_column(df, YEAR_ALIASES)
        if year_col:
            years = sorted(df[year_col].dropna().astype(str).unique())
            selected_year = st.selectbox("Year", ["All"] + years)
        else:
            selected_year = "All"

    with c2:
        location_col = find_column(df, LOCATION_ALIASES)
        if location_col:
            locations = sorted(df[location_col].dropna().astype(str).unique())
            selected_location = st.selectbox("Location", ["All"] + locations)
        else:
            selected_location = "All"

    # Apply Filters
    filtered_df = df.copy()
    if selected_year != "All" and year_col:
        filtered_df = filtered_df[
            filtered_df[year_col].astype(str) == selected_year
        ]
    if selected_location != "All" and location_col:
        filtered_df = filtered_df[
            filtered_df[location_col].astype(str) == selected_location
        ]

    # Data Entry Section
    st.subheader("📝 KPI Data Entry")

    periodicity_col = find_column(filtered_df, FREQUENCY_ALIASES)

    # Categorize into tabs by Periodicity to enforce strict cell locking
    tabs = st.tabs(["Monthly", "Quarterly", "Semi-Annual", "Annual", "All Records"])

    tab_configs = [
        ("Monthly", ["1", "M", "MONTHLY", 1], MONTH_COLUMNS),
        ("Quarterly", ["2", "Q", "QUARTERLY", 2], ["3", "6", "9", "12"]),
        ("Semi-Annual", ["3", "S", "SA", "SEMI ANNUAL", "SEMI-ANNUAL", 3], ["6", "12"]),
        ("Annual", ["4", "A", "ANNUAL", "ANNUALLY", "Y", "YEARLY", 4], ["12"]),
    ]

    for i, (tab_name, match_vals, valid_m) in enumerate(tab_configs):
        with tabs[i]:
            if periodicity_col:
                tab_df = filtered_df[
                    filtered_df[periodicity_col].astype(str).str.upper().isin([str(v).upper() for v in match_vals])
                ].copy()
            else:
                tab_df = filtered_df.copy()

            if tab_df.empty:
                st.info(f"No {tab_name} KPIs found.")
                continue

            # Lock all month columns except valid ones for this tab
            disabled_cols = []
            for col in tab_df.columns:
                if col in MONTH_COLUMNS and col not in valid_m:
                    disabled_cols.append(col)
                elif col not in MONTH_COLUMNS and col not in EVIDENCE_ALIASES and col not in COMMENT_ALIASES:
                    disabled_cols.append(col)

            column_config = {}
            for col in tab_df.columns:
                if col in MONTH_COLUMNS:
                    column_config[col] = st.column_config.NumberColumn(col, format="%g")
                elif col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
                    column_config[col] = st.column_config.TextColumn(col)

            edited_tab_df = st.data_editor(
                tab_df,
                column_config=column_config,
                disabled=disabled_cols,
                use_container_width=True,
                hide_index=True,
                key=f"editor_{selected_sheet}_{tab_name}",
            )

            if st.button(f"💾 Save Changes ({tab_name})", key=f"save_{tab_name}"):
                cleaned_tab_df = clean_invalid_entries(edited_tab_df)
                code_col = find_column(df, KPI_CODE_ALIASES)
                
                if code_col:
                    for idx, row in cleaned_tab_df.iterrows():
                        kpi_code = str(row[code_col])
                        mask = df[code_col].astype(str) == kpi_code
                        if mask.any():
                            for col in tab_df.columns:
                                if col in MONTH_COLUMNS or col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
                                    df.loc[mask, col] = row[col]

                st.session_state["kpi_sheets"][selected_sheet] = clean_invalid_entries(df)
                st.success(f"Saved {tab_name} entries successfully!")
                st.rerun()

    with tabs[4]:
        st.dataframe(filtered_df, use_container_width=True)
