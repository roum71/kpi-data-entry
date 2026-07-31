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


# ============================================================
# PERIODICITY RULES
# ============================================================


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


# ============================================================
# SILENT AUTO-CLEAR VALIDATION ENGINE
# ============================================================


def clean_invalid_entries(df):
    df = df.copy()

    # Enforce float64 for month columns to ensure Streamlit NumberColumn compatibility
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

            # Clear values entered into disallowed months for this row's periodicity
            if month not in valid_months:
                df.loc[idx, month] = np.nan

    return df


# ============================================================
# EXCEL PROTECTION
# ============================================================


def protect_workbook(workbook):
    for ws in workbook.worksheets:
        ws.protection.sheet = True
        ws.protection.password = EXPORT_PASSWORD
        ws.protection.enable()


# ============================================================
# LOAD WORKBOOK
# ============================================================


def load_sheets(uploaded_file):
    excel = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheets = {}

    for sheet in excel.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
        df.columns = [clean_header(c) for c in df.columns]

        # Explicitly setup dtypes
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

    # ========================================================
    # SELECT ENTRY SHEET
    # ========================================================

    available_sheets = list(sheets.keys())
    selected_sheet = st.selectbox("Select Entry Sheet", available_sheets)

    df = sheets[selected_sheet].copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Clean existing data upon loading view
    df = clean_invalid_entries(df)

    # ========================================================
    # FILTERS
    # ========================================================

    st.subheader("🔎 Filters")
    c1, c2, c3 = st.columns(3)

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

    with c3:
        periodicity_col = find_column(df, FREQUENCY_ALIASES)
        if periodicity_col:
            frequencies = sorted(
                df[periodicity_col].dropna().astype(str).unique()
            )
            selected_periodicity = st.selectbox(
                "Periodicity", ["All"] + frequencies
            )
        else:
            selected_periodicity = "All"

    # ========================================================
    # APPLY FILTERS
    # ========================================================

    filtered_df = df.copy()

    if selected_year != "All" and year_col:
        filtered_df = filtered_df[
            filtered_df[year_col].astype(str) == selected_year
        ]

    if selected_location != "All" and location_col:
        filtered_df = filtered_df[
            filtered_df[location_col].astype(str) == selected_location
        ]

    if selected_periodicity != "All" and periodicity_col:
        filtered_df = filtered_df[
            filtered_df[periodicity_col].astype(str) == selected_periodicity
        ]

    # ========================================================
    # DYNAMIC COLUMN VISIBILITY & DATA EDITOR
    # ========================================================

    st.subheader("📝 KPI Data Entry")

    valid_display_months = allowed_months(selected_periodicity)

    cols_to_display = []
    for col in filtered_df.columns:
        if col in MONTH_COLUMNS:
            if selected_periodicity == "All" or col in valid_display_months:
                cols_to_display.append(col)
        else:
            cols_to_display.append(col)

    filtered_df = filtered_df[cols_to_display]
    filtered_df = clean_invalid_entries(filtered_df)

    # Define editable columns
    editable_columns = []
    for col in filtered_df.columns:
        if col in MONTH_COLUMNS:
            editable_columns.append(col)
        elif col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
            editable_columns.append(col)

    disabled_columns = [
        col for col in filtered_df.columns if col not in editable_columns
    ]

    # Explicit column configurations matching underlying DataFrame dtypes
    column_config = {}
    for col in filtered_df.columns:
        if col in MONTH_COLUMNS:
            column_config[col] = st.column_config.NumberColumn(
                col, format="%g"
            )
        elif col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
            column_config[col] = st.column_config.TextColumn(
                col, help="Editable text field"
            )

    # Render data editor with clean static key per sheet
    edited_df = st.data_editor(
        filtered_df,
        column_config=column_config,
        disabled=disabled_columns,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=f"editor_{selected_sheet}",
    )

    # Clean the edited inputs immediately
    cleaned_df = clean_invalid_entries(edited_df)

    # ========================================================
    # SAVE CHANGES BACK TO SESSION STATE
    # ========================================================

    if st.button(f"💾 Save Changes - {selected_sheet}"):
        original = sheets[selected_sheet].copy()

        code_col = find_column(original, KPI_CODE_ALIASES)

        if code_col:
            for idx, row in cleaned_df.iterrows():
                kpi_code = str(row[code_col])
                mask = original[code_col].astype(str) == kpi_code

                if mask.any():
                    for col in editable_columns:
                        if col in cleaned_df.columns:
                            val = row[col]
                            original.loc[mask, col] = (
                                np.nan if pd.isna(val) else val
                            )
        else:
            original = cleaned_df.copy()

        # Final sanitization pass before updating session state
        original = clean_invalid_entries(original)
        st.session_state["kpi_sheets"][selected_sheet] = original

        st.success(f"Changes saved successfully for sheet '{selected_sheet}'!")
        st.rerun()

# ============================================================
# EXPORT WORKBOOK
# ============================================================

st.markdown("---")
st.subheader("📥 Export Updated Workbook")


def export_workbook(all_sheets):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in all_sheets.items():
            cleaned_sheet_data = clean_invalid_entries(data)
            cleaned_sheet_data.to_excel(
                writer, sheet_name=sheet_name, index=False
            )

        protect_workbook(writer.book)

    output.seek(0)
    return output


if uploaded_file:
    export_file = export_workbook(st.session_state["kpi_sheets"])
    new_name = (
        uploaded_file.name.replace(".xlsx", "").replace(".xlsm", "")
        + "_Updated.xlsx"
    )

    st.download_button(
        label="⬇️ Download Updated KPI Workbook",
        data=export_file,
        file_name=new_name,
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

else:
    st.info("Upload a workbook to enable export.")
