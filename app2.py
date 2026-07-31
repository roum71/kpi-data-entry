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
    "Two-sheet KPI data entry with validation, evidence tracking, and comments."
)

# ============================================================
# CONSTANTS
# ============================================================

MONTH_COLUMNS = [str(i) for i in range(1, 13)]

ENTRY_SHEETS = ["Entry", "Entry_2"]

# Password protection
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

EVIDENCE_ALIASES = ["Evidence Status", "evidence_status", "Evidence", "حالة الدليل"]

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
    try:
        p = int(float(periodicity))
    except Exception:
        return MONTH_COLUMNS

    if p == 1:
        # Monthly
        return MONTH_COLUMNS

    elif p == 2:
        # Quarterly
        return ["3", "6", "9", "12"]

    elif p == 3:
        # Semi Annual
        return ["6", "12"]

    elif p == 4:
        # Annual
        return ["12"]

    return MONTH_COLUMNS


# ============================================================
# VALIDATION ENGINE
# ============================================================


def validate_entries(df):
    df = df.copy()
    errors = []

    for idx, row in df.iterrows():
        periodicity = get_value(row, FREQUENCY_ALIASES, 1)
        kpi_code = get_value(row, KPI_CODE_ALIASES, "")
        valid_months = allowed_months(periodicity)

        for month in MONTH_COLUMNS:
            if month not in df.columns:
                continue

            value = row[month]

            # Lock invalid months
            if month not in valid_months:
                if pd.notna(value) and str(value).strip() != "":
                    errors.append({
                        "KPI": kpi_code,
                        "Month": month,
                        "Error": (
                            f"Month {month} not allowed for periodicity"
                            f" {periodicity}"
                        ),
                    })

                df.loc[idx, month] = np.nan
                continue

            # Empty value
            if pd.isna(value) or str(value).strip() == "":
                continue

            # Numeric validation
            try:
                df.loc[idx, month] = float(value)
            except Exception:
                errors.append({
                    "KPI": kpi_code,
                    "Month": month,
                    "Error": "Only numeric values are allowed",
                })
                df.loc[idx, month] = np.nan

    return df, errors


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
        df, _ = validate_entries(df)
        sheets[sheet] = df

    return sheets


# ============================================================
# MAIN APPLICATION
# ============================================================

uploaded_file = st.file_uploader(
    "Upload KPI Excel Workbook", type=["xlsx", "xlsm"]
)


if uploaded_file:
    # Load workbook once
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

    # Ensure header column names are strictly strings
    df.columns = [str(c).strip() for c in df.columns]

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
    # DATA EDITOR CONFIGURATION
    # ========================================================

    st.subheader("📝 KPI Data Entry")

    editable_columns = []

    for col in filtered_df.columns:
        if col in MONTH_COLUMNS:
            editable_columns.append(col)
        elif col in EVIDENCE_ALIASES:
            editable_columns.append(col)
        elif col in COMMENT_ALIASES:
            editable_columns.append(col)

    # Convert month columns explicitly to float64 for numeric entry
    for month in MONTH_COLUMNS:
        if month in filtered_df.columns:
            filtered_df[month] = pd.to_numeric(
                filtered_df[month], errors="coerce"
            )

    # Prepare disabled columns list for native Streamlit disabling
    disabled_columns = [
        col for col in filtered_df.columns if col not in editable_columns
    ]

    # Only define explicit column config for month columns
    column_config = {}
    for month in MONTH_COLUMNS:
        if month in filtered_df.columns:
            column_config[month] = st.column_config.NumberColumn(
                month, format="%g"
            )

    edited_df = st.data_editor(
        filtered_df,
        column_config=column_config,
        disabled=disabled_columns,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=f"editor_{selected_sheet}",
    )

    # ========================================================
    # VALIDATE AFTER EDIT
    # ========================================================

    validated_df, errors = validate_entries(edited_df)

    if errors:
        st.warning(f"{len(errors)} invalid entries detected")

        with st.expander("View Errors"):
            st.dataframe(pd.DataFrame(errors), use_container_width=True)

    # ========================================================
    # SAVE CHANGES BACK TO SHEET
    # ========================================================

    if st.button(f"💾 Save Changes - {selected_sheet}"):
        original = sheets[selected_sheet].copy()

        code_col = find_column(original, KPI_CODE_ALIASES)

        if code_col:
            for _, row in validated_df.iterrows():
                kpi_code = str(row[code_col])
                mask = original[code_col].astype(str) == kpi_code

                for col in editable_columns:
                    if col in validated_df.columns:
                        original.loc[mask, col] = row[col]
        else:
            original = validated_df.copy()

        st.session_state["kpi_sheets"][selected_sheet] = original

        st.success("Changes saved successfully")

# ============================================================
# EXPORT UPDATED WORKBOOK
# ============================================================

st.markdown("---")

st.subheader("📥 Export Updated Workbook")


def export_workbook(all_sheets):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in all_sheets.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

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
