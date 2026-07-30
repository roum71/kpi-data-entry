import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import openpyxl

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="KPI Data Entry System",
    page_icon="📊",
    layout="wide"
)

st.title("📊 KPI Data Entry System")
st.write("Upload your KPI Excel file, edit month entries with automated validation, and export your cleaned workbook.")

# ============================================================
# CONSTANTS & COLUMN ALIASES
# ============================================================
ALL_MONTH_COLUMNS = [str(i) for i in range(1, 13)]

FREQ_ALIASES = ["Freq", "measurement_frequency", "freq", "frequency", "التردد"]
UNIT_ALIASES = ["Unit", "unit_id", "unit", "Unit ID", "unit_code"]
UNIT_NAME_ALIASES = ["unit_name_ar", "Unit Name (AR)", "unit_name", "الوحدة"]
NEG_ALIASES  = ["allow_negative_values", "Allow Negative", "allow_negative", "allow_neg"]
CODE_ALIASES = ["kpi_code", "KPI Code", "code"]
NAME_ALIASES = ["kpi_name_ar", "KPI Name (AR)", "kpi_name", "name_ar"]

SHEET_NAME = "Entry"

# Strong backend password for exported Excel file protection
STRONG_ADMIN_PASSWORD = "KPi#Secure2026!Lock"

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def clean_header(header_name):
    """Cleans column headers to match standard field names."""
    if header_name is None:
        return ""
    header_str = str(header_name)
    header_str = header_str.replace('\xa0', ' ')
    header_str = re.sub(r'[\r\n\t]', ' ', header_str)
    return header_str.strip()

def get_col_val(row, aliases, default=None):
    """Safely reads a column value using any of its supported header aliases."""
    for name in aliases:
        if name in row and pd.notna(row[name]):
            return row[name]
    return default

def get_col_name(df, aliases):
    """Returns the actual matching column name present in a DataFrame."""
    for col in df.columns:
        if col in aliases:
            return col
    return None

def get_valid_months(frequency):
    """
    Frequency Mapping:
    1 = Monthly      -> All months [1..12]
    2 = Quarterly    -> [3, 6, 9, 12]
    3 = Semi-Annual  -> [6, 12]
    4 = Annual       -> [12]
    """
    try:
        freq = int(float(frequency))
    except (ValueError, TypeError):
        return list(range(1, 13))

    if freq == 1:
        return list(range(1, 13))
    elif freq == 2:
        return [3, 6, 9, 12]
    elif freq == 3:
        return [6, 12]
    elif freq == 4:
        return [12]
    return list(range(1, 13))

def apply_excel_protection(workbook, password=STRONG_ADMIN_PASSWORD):
    """Applies strict password protection to all sheets in an openpyxl Workbook."""
    for ws in workbook.worksheets:
        ws.protection.password = password
        ws.protection.sheet = True
        ws.protection.enable()
        
        ws.protection.formatCells = False
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.protection.insertColumns = False
        ws.protection.insertRows = False
        ws.protection.insertHyperlinks = False
        ws.protection.deleteColumns = False
        ws.protection.deleteRows = False
        ws.protection.selectLockedCells = True
        ws.protection.selectUnlockedCells = True

def validate_and_clean_data(df):
    """
    Validates rules, enforces locked months (wiping disallowed entries),
    and checks allow_negative_values controls.
    """
    cleaned_df = df.copy()
    errors = []
    cleared_count = 0

    for index, row in cleaned_df.iterrows():
        frequency = get_col_val(row, FREQ_ALIASES)
        kpi_code = get_col_val(row, CODE_ALIASES, "")
        kpi_name = get_col_val(row, NAME_ALIASES, "")

        allow_neg_val = get_col_val(row, NEG_ALIASES, 0)
        try:
            is_neg_allowed = int(float(allow_neg_val)) == 1
        except (ValueError, TypeError):
            is_neg_allowed = False

        valid_months = get_valid_months(frequency)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

            # 1. STRICT MONTH LOCKING BY FREQUENCY
            if month_num not in valid_months:
                if pd.notna(value) and str(value).strip() != "" and str(value).strip().lower() != "none":
                    errors.append({
                        "kpi_code": kpi_code,
                        "kpi_name_ar": kpi_name,
                        "Month": month_col,
                        "Value": value,
                        "Error": f"Month {month_col} is LOCKED for Frequency {frequency}. Value cleared."
                    })
                    cleaned_df.loc[index, month_col] = np.nan
                    cleared_count += 1
                else:
                    cleaned_df.loc[index, month_col] = np.nan
                continue

            # Skip empty cells
            if pd.isna(value) or str(value).strip() == "" or str(value).strip().lower() == "none":
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # 2. NUMERIC & NEGATIVE CONTROL VALIDATION
            try:
                val_float = float(value)

                # Check Negative Entry Control Rule
                if val_float < 0.0 and not is_neg_allowed:
                    errors.append({
                        "kpi_code": kpi_code,
                        "kpi_name_ar": kpi_name,
                        "Month": month_col,
                        "Value": value,
                        "Error": "Negative values NOT allowed for this KPI. Value cleared."
                    })
                    cleaned_df.loc[index, month_col] = np.nan
                    cleared_count += 1
                    continue

                cleaned_df.loc[index, month_col] = val_float

            except (ValueError, TypeError):
                errors.append({
                    "kpi_code": kpi_code,
                    "kpi_name_ar": kpi_name,
                    "Month": month_col,
                    "Value": value,
                    "Error": "Non-numeric value entered. Value cleared."
                })
                cleaned_df.loc[index, month_col] = np.nan
                cleared_count += 1

    return cleaned_df, errors, cleared_count

def generate_unformatted_excel(df, uploaded_file, target_sheet_name):
    """Exports cleaned dataset to .xlsx directly as protected read-only worksheets."""
    output_xlsx = io.BytesIO()
    excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=target_sheet_name, index=False)
        
        for sheet in excel_file.sheet_names:
            if sheet != target_sheet_name:
                other_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                other_df.to_excel(writer, sheet_name=sheet, index=False)
        
        apply_excel_protection(writer.book)

    output_xlsx.seek(0)
    return output_xlsx

def generate_long_format_excel(df):
    """
    Converts wide-format month data (1-12 columns) into long unpivoted format.
    Includes Date column. NOT password protected.
    """
    present_month_cols = [col for col in ALL_MONTH_COLUMNS if col in df.columns]
    id_vars = [col for col in df.columns if col not in ALL_MONTH_COLUMNS]

    long_df = pd.melt(
        df,
        id_vars=id_vars,
        value_vars=present_month_cols,
        var_name="Month",
        value_name="Value"
    )

    valid_rows = []
    freq_col = get_col_name(df, FREQ_ALIASES)

    for idx, row in long_df.iterrows():
        frequency = row[freq_col] if freq_col and pd.notna(row[freq_col]) else 1
        valid_months = get_valid_months(frequency)
        valid_rows.append(int(row["Month"]) in valid_months)

    long_df = long_df[valid_rows].reset_index(drop=True)
    long_df["Month"] = long_df["Month"].astype(int)

    if "year" in long_df.columns:
        def build_date(r):
            try:
                yr = int(float(r["year"]))
                mn = int(r["Month"])
                return f"{yr:04d}-{mn:02d}-01"
            except (ValueError, TypeError):
                return np.nan
        long_df["Date"] = long_df.apply(build_date, axis=1)
    else:
        long_df["Date"] = np.nan

    output_xlsx = io.BytesIO()
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        long_df.to_excel(writer, sheet_name="Long_Format_Data", index=False)

    output_xlsx.seek(0)
    return output_xlsx

# ============================================================
# MAIN APPLICATION INTERFACE
# ============================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsm", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        target_sheet = SHEET_NAME if SHEET_NAME in excel_file.sheet_names else excel_file.sheet_names[0]

        if "master_df" not in st.session_state or st.session_state.get("file_name") != uploaded_file.name:
            df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, engine="openpyxl")
            df_raw.columns = [clean_header(col) for col in df_raw.columns]
            
            cleaned_master, _, _ = validate_and_clean_data(df_raw)
            st.session_state["master_df"] = cleaned_master
            st.session_state["file_name"] = uploaded_file.name
            st.session_state["editor_version"] = 0

        df = st.session_state["master_df"]
        st.success(f"✅ Sheet '{target_sheet}' loaded successfully.")

        # Filter Section
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            years = sorted(df["year"].dropna().astype(str).unique().tolist()) if "year" in df.columns else []
            selected_year = st.selectbox("Select Year to Edit", years) if years else None

        with col_f2:
            freq_options = ["All Frequencies", "1 - Monthly", "2 - Quarterly", "3 - Semi-Annual", "4 - Annual"]
            selected_freq_str = st.selectbox("Filter by Frequency", freq_options)

        working_df = df.copy()
        if selected_year:
            working_df = working_df[working_df["year"].astype(str) == str(selected_year)]

        selected_freq_num = None
        if selected_freq_str != "All Frequencies":
            selected_freq_num = int(selected_freq_str.split(" - ")[0])
            freq_col_name = get_col_name(working_df, FREQ_ALIASES)
            if freq_col_name:
                working_df = working_df[working_df[freq_col_name].astype(str).str.startswith(str(selected_freq_num))]

        st.subheader("📝 Edit KPI Data")

        if selected_freq_num is not None:
            allowed_months = get_valid_months(selected_freq_num)
        else:
            allowed_months = list(range(1, 13))

        allowed_month_cols = [str(m) for m in allowed_months]
        hidden_month_cols = [m for m in ALL_MONTH_COLUMNS if m in working_df.columns and m not in allowed_month_cols]
        display_df = working_df.drop(columns=hidden_month_cols, errors="ignore")

        column_configs = {}
        for col in display_df.columns:
            if col not in ALL_MONTH_COLUMNS:
                column_configs[col] = st.column_config.TextColumn(disabled=True)
            else:
                column_configs[col] = st.column_config.NumberColumn(label=f"Month {col}", disabled=False)

        editor_key = f"kpi_editor_v_{st.session_state['editor_version']}"
        
        edited_df = st.data_editor(
            display_df,
            column_config=column_configs,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=editor_key
        )

        cleaned_edited_df, edit_errors, cleared_count = validate_and_clean_data(edited_df)

        if cleared_count > 0 or not cleaned_edited_df.equals(display_df):
            code_col = get_col_name(df, CODE_ALIASES)
            if code_col:
                for idx, row in cleaned_edited_df.iterrows():
                    kpi_code_val = row[code_col]
                    mask = df[code_col] == kpi_code_val
                    if selected_year and "year" in df.columns:
                        mask = mask & (df["year"].astype(str) == str(selected_year))
                    
                    for m in allowed_month_cols:
                        if m in df.columns:
                            df.loc[mask, m] = row[m]

            st.session_state["master_df"] = df

            if cleared_count > 0:
                st.session_state["editor_version"] += 1
                st.warning(f"⚠️ {cleared_count} invalid or disallowed entry(ies) were wiped.")
                st.rerun()

        # ============================================================
        # EXPORT & DOWNLOAD SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Export Options")

        col1, col2 = st.columns(2)

        # 1. Wide Format
        with col1:
            st.markdown("### 1. Wide Format (Protected)")
            st.caption("🔒 Updates original file structure with password protection.")
            output_xlsx_data = generate_unformatted_excel(df, uploaded_file, target_sheet)
            st.download_button(
                label=f"⬇️ Download ({uploaded_file.name})",
                data=output_xlsx_data,
                file_name=uploaded_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        # 2. Long Format
        with col2:
            st.markdown("### 2. Long Format (Unprotected)")
            st.caption("🔓 Unpivoted month columns with Date field. No password.")
            output_long_data = generate_long_format_excel(df)
            long_filename = uploaded_file.name.rsplit('.', 1)[0] + "_LongFormat.xlsx"
            st.download_button(
                label="⬇️ Download Long Format (.xlsx)",
                data=output_long_data,
                file_name=long_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:
        st.error("An error occurred while processing the Excel workbook.")
        st.exception(e)
else:
    st.info("👆 Please upload your Excel file to get started.")
