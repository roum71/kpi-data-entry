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
st.write("Upload your Excel file (`Actual10.xlsm`), edit KPI entries with automated month-locking, and export your cleaned workbook.")

# ============================================================
# CONSTANTS & COLUMN ALIASES
# ============================================================
MONTH_COLUMNS = [str(i) for i in range(1, 13)]

FREQ_ALIASES = ["Freq", "measurement_frequency", "freq", "frequency", "التردد"]
UNIT_ALIASES = ["Unit", "unit_id", "unit", "Unit ID", "unit_code"]
NEG_ALIASES  = ["allow_negative_values", "Allow Negative", "allow_negative", "allow_neg"]
CODE_ALIASES = ["kpi_code", "KPI Code", "code"]
NAME_ALIASES = ["kpi_name_ar", "KPI Name (AR)", "kpi_name", "name_ar"]

SHEET_NAME = "Entry"

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
    1 = Monthly     -> All months [1..12]
    2 = Quarterly   -> [3, 6, 9, 12]
    3 = Semi-Annual -> [6, 12]
    4 = Annual      -> [12]
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

def is_percentage_unit(unit_id):
    """Checks if Unit ID is a percentage type (10, 11, 12)."""
    try:
        u_id = int(float(unit_id))
        return u_id in [10, 11, 12]
    except (ValueError, TypeError):
        return False

def validate_and_clean_data(df, kpi_unit_map=None):
    """
    Validates rules, enforces locked months (wiping disallowed entries), 
    and handles percentage standardizations (80 -> 0.8).
    """
    if kpi_unit_map is None:
        kpi_unit_map = {}

    cleaned_df = df.copy()
    errors = []
    cleared_count = 0

    for index, row in cleaned_df.iterrows():
        frequency = get_col_val(row, FREQ_ALIASES)
        unit_id = get_col_val(row, UNIT_ALIASES)
        kpi_code = get_col_val(row, CODE_ALIASES, "")
        kpi_name = get_col_val(row, NAME_ALIASES, "")

        # Fallback to map if unit_id missing on row
        if (unit_id is None or pd.isna(unit_id)) and kpi_code in kpi_unit_map:
            unit_id = kpi_unit_map[kpi_code]

        allow_negative = get_col_val(row, NEG_ALIASES, 0)
        valid_months = get_valid_months(frequency)
        is_pct = is_percentage_unit(unit_id)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

            # 1. STRICT MONTH LOCKING
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

            # 2. NUMERIC & PERCENTAGE CONVERSION LOGIC
            try:
                val_float = float(value)

                if is_pct:
                    # Converts 80 -> 0.8 (80%)
                    if 1.0 < val_float <= 100.0:
                        val_float = val_float / 100.0

                    if val_float < 0.0 or val_float > 1.0:
                        errors.append({
                            "kpi_code": kpi_code,
                            "kpi_name_ar": kpi_name,
                            "Month": month_col,
                            "Value": value,
                            "Error": "Percentage must be between 0% and 100% (or 0 and 1). Value cleared."
                        })
                        cleaned_df.loc[index, month_col] = np.nan
                        cleared_count += 1
                    else:
                        cleaned_df.loc[index, month_col] = val_float

                else:
                    min_val = -1000000000.0 if str(allow_negative).strip() == "1" else 0.0
                    if val_float < min_val:
                        errors.append({
                            "kpi_code": kpi_code,
                            "kpi_name_ar": kpi_name,
                            "Month": month_col,
                            "Value": value,
                            "Error": f"Negative value not allowed (Min: {min_val}). Value cleared."
                        })
                        cleaned_df.loc[index, month_col] = np.nan
                        cleared_count += 1
                    else:
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

def generate_formatted_excel(df, uploaded_file, target_sheet_name, kpi_unit_map):
    """
    Exports dataset to .xlsx, ensures numbers are strictly stored as float values,
    and applies openpyxl number_format (0.00% or #,##0.00).
    """
    output_xlsx = io.BytesIO()
    excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=target_sheet_name, index=False)
        for sheet in excel_file.sheet_names:
            if sheet != target_sheet_name:
                other_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                other_df.to_excel(writer, sheet_name=sheet, index=False)

    output_xlsx.seek(0)

    # Apply openpyxl cell formatting with strict float casting
    wb = openpyxl.load_workbook(output_xlsx)
    ws = wb[target_sheet_name]

    headers = [cell.value for cell in ws[1]]
    code_col_idx = None
    unit_col_idx = None

    for alias in CODE_ALIASES:
        if alias in headers:
            code_col_idx = headers.index(alias) + 1
            break

    for alias in UNIT_ALIASES:
        if alias in headers:
            unit_col_idx = headers.index(alias) + 1
            break

    for row_idx in range(2, ws.max_row + 1):
        unit_id = ws.cell(row=row_idx, column=unit_col_idx).value if unit_col_idx else None
        
        # Fallback to map lookup
        if (unit_id is None or pd.isna(unit_id)) and code_col_idx:
            kpi_code_val = ws.cell(row=row_idx, column=code_col_idx).value
            unit_id = kpi_unit_map.get(kpi_code_val, None)

        is_pct = is_percentage_unit(unit_id)

        for m_str in MONTH_COLUMNS:
            if m_str in headers:
                col_idx = headers.index(m_str) + 1
                cell = ws.cell(row=row_idx, column=col_idx)

                if cell.value is not None and str(cell.value).strip() != "" and str(cell.value).lower() != "none":
                    try:
                        num_val = float(cell.value)
                        cell.value = num_val  # Store as numeric float
                        
                        if is_pct:
                            cell.number_format = '0.00%'  # Formats 0.8 as 80.00%
                        else:
                            cell.number_format = '#,##0.00' # Standard numeric format
                    except (ValueError, TypeError):
                        pass

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# ============================================================
# MAIN APPLICATION INTERFACE
# ============================================================
uploaded_file = st.file_uploader("Upload Excel File (`Actual10.xlsm`)", type=["xlsm", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        target_sheet = SHEET_NAME if SHEET_NAME in excel_file.sheet_names else excel_file.sheet_names[0]

        # Extract KPI Unit Map from 'Units' sheet if available
        kpi_unit_map = {}
        if "Units" in excel_file.sheet_names:
            try:
                units_df = pd.read_excel(uploaded_file, sheet_name="Units", engine="openpyxl")
                units_df.columns = [clean_header(c) for c in units_df.columns]
                c_col = get_col_name(units_df, CODE_ALIASES)
                u_col = get_col_name(units_df, UNIT_ALIASES)
                if c_col and u_col:
                    kpi_unit_map = dict(zip(units_df[c_col], units_df[u_col]))
            except Exception:
                pass

        # Load file into Session State
        if "master_df" not in st.session_state or st.session_state.get("file_name") != uploaded_file.name:
            df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, engine="openpyxl")
            df_raw.columns = [clean_header(col) for col in df_raw.columns]
            
            cleaned_master, _, _ = validate_and_clean_data(df_raw, kpi_unit_map)
            st.session_state["master_df"] = cleaned_master
            st.session_state["file_name"] = uploaded_file.name
            st.session_state["editor_version"] = 0

        df = st.session_state["master_df"]
        st.success(f"✅ Sheet '{target_sheet}' loaded successfully.")

        # Filter Section: Year & Frequency
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            years = sorted(df["year"].dropna().astype(str).unique().tolist()) if "year" in df.columns else []
            selected_year = st.selectbox("Select Year to Edit", years) if years else None

        with col_f2:
            freq_options = ["All Frequencies", "1 - Monthly", "2 - Quarterly", "3 - Semi-Annual", "4 - Annual"]
            selected_freq_str = st.selectbox("Filter by Frequency (Locks non-applicable months)", freq_options)

        # Apply Filters
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
        st.caption("🔒 Invalid month cells are automatically disabled or wiped instantly if entered.")

        # Dynamic Column Config for Editor
        column_configs = {}
        for alias in ["kpi_code", "KPI Code", "location_id", "year", "Freq", "measurement_frequency", "KPI Name (AR)", "kpi_name_ar", "Unit", "unit_id"]:
            if alias in working_df.columns:
                column_configs[alias] = st.column_config.TextColumn(disabled=True)

        # Set UI Month Disabling Rules
        if selected_freq_num is not None:
            allowed_months = get_valid_months(selected_freq_num)
        else:
            allowed_months = list(range(1, 13))

        for m in MONTH_COLUMNS:
            m_num = int(m)
            is_disabled = m_num not in allowed_months

            column_configs[m] = st.column_config.NumberColumn(
                label=f"Month {m}",
                format="%.2f",
                disabled=is_disabled,
                help="Locked" if is_disabled else "Editable"
            )

        # Render Data Editor with unique version key
        editor_key = f"kpi_editor_v_{st.session_state['editor_version']}"
        
        edited_df = st.data_editor(
            working_df,
            column_config=column_configs,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=editor_key
        )

        # Validate edits submitted through editor
        cleaned_edited_df, edit_errors, cleared_count = validate_and_clean_data(edited_df, kpi_unit_map)

        # Update state and force rerun if changes/clears occurred
        if cleared_count > 0 or not cleaned_edited_df.equals(working_df):
            code_col = get_col_name(df, CODE_ALIASES)

            if code_col:
                for idx, row in cleaned_edited_df.iterrows():
                    kpi_code_val = row[code_col]
                    mask = df[code_col] == kpi_code_val
                    if selected_year and "year" in df.columns:
                        mask = mask & (df["year"].astype(str) == str(selected_year))
                    for m in MONTH_COLUMNS:
                        df.loc[mask, m] = row[m]

            st.session_state["master_df"] = df

            if cleared_count > 0:
                st.session_state["editor_version"] += 1
                st.warning(f"⚠️ {cleared_count} invalid or locked month entry(ies) were wiped.")
                st.rerun()

        if edit_errors:
            with st.expander("View Wiped Entries Log"):
                st.dataframe(pd.DataFrame(edit_errors), use_container_width=True)

        # ============================================================
        # EXPORT / DOWNLOAD SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Choose Your Download Format")

        col1, col2 = st.columns(2)

        # OPTION 1: Clean Excel (.xlsx) with Formatted Cells & All Sheets Preserved
        with col1:
            st.markdown("### **Option 1: Complete Workbook (`.xlsx`)**")
            st.write("Preserves all sheets (`Entry`, `Units`, etc.) and formats percentages as `0.00%` directly in Excel.")

            output_xlsx_data = generate_formatted_excel(df, uploaded_file, target_sheet, kpi_unit_map)
            download_name_xlsx = uploaded_file.name.rsplit('.', 1)[0] + "_updated.xlsx"

            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=output_xlsx_data,
                file_name=download_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        # OPTION 2: Entry Sheet CSV for Power Query
        with col2:
            st.markdown("### **Option 2: Data Table Only (`.csv`)**")
            st.write("Exports only the cleaned `Entry` sheet with UTF-8 encoding (Arabic language safe). Best for **Power Query**.")

            output_csv = df.to_csv(index=False).encode('utf-8-sig')

            st.download_button(
                label="⬇️ Download CSV (.csv)",
                data=output_csv,
                file_name="Actual10_Entry_updated.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error("An error occurred while processing the Excel workbook.")
        st.exception(e)
else:
    st.info("👆 Please upload your `Actual10.xlsm` file to get started.")
