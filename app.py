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
# CONSTANTS & CONFIGURATION
# ============================================================
MONTH_COLUMNS = [str(i) for i in range(1, 13)]

REQUIRED_COLUMNS = [
    "sub_objective_id",
    "kpi_code",
    "location_id",
    "year",
    "measurement_frequency",
    "kpi_name_ar",
    "unit_id"
] + MONTH_COLUMNS

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

def get_valid_months(frequency):
    """
    Frequency Mapping (VBA Logic):
    1 = Monthly     -> All months [1..12]
    2 = Quarterly   -> [3, 6, 9, 12]
    3 = Semi-Annual -> [6, 12]
    4 = Annual      -> [12]
    """
    try:
        freq = int(float(frequency))
    except (ValueError, TypeError):
        return []

    if freq == 1:
        return list(range(1, 13))
    elif freq == 2:
        return [3, 6, 9, 12]
    elif freq == 3:
        return [6, 12]
    elif freq == 4:
        return [12]
    return []

def is_percentage_unit(unit_id):
    """Checks if Unit ID is a percentage type (10, 11, 12)."""
    try:
        u_id = int(float(unit_id))
        return u_id in [10, 11, 12]
    except (ValueError, TypeError):
        return False

def validate_and_clean_data(df):
    """
    Validates, converts percentages (e.g., 80 -> 0.8), and wipes invalid inputs/locked months.
    """
    cleaned_df = df.copy()
    errors = []

    for index, row in cleaned_df.iterrows():
        frequency = row.get("measurement_frequency", None)
        unit_id = row.get("unit_id", None)
        allow_negative = row.get("allow_negative_values", 0)

        valid_months = get_valid_months(frequency)
        is_pct = is_percentage_unit(unit_id)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

            # 1. Strict Locking Rule for Disallowed Months
            if month_num not in valid_months:
                if not pd.isna(value) and str(value).strip() != "":
                    errors.append({
                        "kpi_code": row.get("kpi_code", ""),
                        "kpi_name_ar": row.get("kpi_name_ar", ""),
                        "Month": month_col,
                        "Value": value,
                        "Error": f"Month {month_col} is locked for frequency {frequency}. Entry cleared."
                    })
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # Leave empty cells untouched
            if pd.isna(value) or str(value).strip() == "":
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # 2. Numeric & Percentage Parsing Logic
            try:
                val_float = float(value)

                if is_pct:
                    # Accepts either 80 or 0.8 -> stores as 0.8 (80%)
                    if 1.0 < val_float <= 100.0:
                        val_float = val_float / 100.0

                    if val_float < 0.0 or val_float > 1.0:
                        errors.append({
                            "kpi_code": row.get("kpi_code", ""),
                            "kpi_name_ar": row.get("kpi_name_ar", ""),
                            "Month": month_col,
                            "Value": value,
                            "Error": "Percentage must be between 0% and 100% (or 0 and 1). Cleared."
                        })
                        cleaned_df.loc[index, month_col] = np.nan
                    else:
                        cleaned_df.loc[index, month_col] = val_float

                else:
                    min_val = -1000000000.0 if str(allow_negative).strip() == "1" else 0.0
                    if val_float < min_val:
                        errors.append({
                            "kpi_code": row.get("kpi_code", ""),
                            "kpi_name_ar": row.get("kpi_name_ar", ""),
                            "Month": month_col,
                            "Value": value,
                            "Error": f"Negative value not allowed (Min: {min_val}). Cleared."
                        })
                        cleaned_df.loc[index, month_col] = np.nan
                    else:
                        cleaned_df.loc[index, month_col] = val_float

            except (ValueError, TypeError):
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": "Value must be numeric. Cleared."
                })
                cleaned_df.loc[index, month_col] = np.nan

    return cleaned_df, errors

def generate_formatted_excel(df, uploaded_file, target_sheet_name):
    """
    Exports to .xlsx and applies openpyxl formatting rules (0.00% or #,##0.00).
    Preserves all original sheets in the workbook.
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

    # Format cell display styles via openpyxl
    wb = openpyxl.load_workbook(output_xlsx)
    ws = wb[target_sheet_name]

    headers = [cell.value for cell in ws[1]]
    unit_col_idx = headers.index("unit_id") + 1 if "unit_id" in headers else None

    for row_idx in range(2, ws.max_row + 1):
        unit_id = ws.cell(row=row_idx, column=unit_col_idx).value if unit_col_idx else None
        is_pct = is_percentage_unit(unit_id)

        for m_str in MONTH_COLUMNS:
            if m_str in headers:
                col_idx = headers.index(m_str) + 1
                cell = ws.cell(row=row_idx, column=col_idx)

                if cell.value is not None and cell.value != "":
                    if is_pct:
                        cell.number_format = '0.00%'
                    else:
                        cell.number_format = '#,##0.00'

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

        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, engine="openpyxl")
        df.columns = [clean_header(col) for col in df.columns]

        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            st.error(f"Missing required columns in sheet '{target_sheet}': {missing_cols}")
            st.stop()

        st.success(f"✅ Sheet '{target_sheet}' loaded successfully.")

        # Filter by Year
        years = sorted(df["year"].dropna().astype(str).unique().tolist())
        selected_year = st.selectbox("Select Year to Edit", years) if years else None
        
        if selected_year:
            working_df = df[df["year"].astype(str) == str(selected_year)].copy()
        else:
            working_df = df.copy()

        # Clean/Clear initial data state
        working_df, initial_errors = validate_and_clean_data(working_df)

        st.subheader("📝 Edit KPI Data")
        st.caption("🔒 Invalid month cells for specified KPI frequencies are automatically disabled or wiped.")

        # Dynamically set Column Configuration for Streamlit UI
        column_configs = {
            "kpi_code": st.column_config.TextColumn("KPI Code", disabled=True),
            "kpi_name_ar": st.column_config.TextColumn("KPI Name (AR)", disabled=True),
            "measurement_frequency": st.column_config.NumberColumn("Freq", disabled=True),
            "unit_id": st.column_config.NumberColumn("Unit", disabled=True)
        }

        # Apply UI Month Disabling based on frequency
        # Determine global allowed months across filtered KPIs for display
        active_frequencies = working_df["measurement_frequency"].dropna().astype(int).unique().tolist()
        
        for m in MONTH_COLUMNS:
            m_num = int(m)
            # If selected row set has only 1 frequency (e.g. Annual), completely disable non-valid columns in UI
            if len(active_frequencies) == 1:
                valid_m = get_valid_months(active_frequencies[0])
                is_disabled = m_num not in valid_m
            else:
                is_disabled = False

            column_configs[m] = st.column_config.NumberColumn(
                label=f"Month {m}",
                format="%.2f",
                disabled=is_disabled,
                help="Enter numbers (or percentages like 80 or 0.8)"
            )

        edited_df = st.data_editor(
            working_df,
            column_config=column_configs,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed"
        )

        # Validate entries submitted through the editor
        cleaned_edited_df, edit_errors = validate_and_clean_data(edited_df)

        if edit_errors:
            st.warning(f"⚠️ {len(edit_errors)} invalid entries detected and automatically cleared.")
            with st.expander("View Cleared Invalid Entries Log"):
                st.dataframe(pd.DataFrame(edit_errors), use_container_width=True)

        # Merge changes back to master dataframe
        if selected_year:
            for index, row in cleaned_edited_df.iterrows():
                kpi_code = row["kpi_code"]
                mask = (df["year"].astype(str) == str(selected_year)) & (df["kpi_code"] == kpi_code)
                if mask.any():
                    for month in MONTH_COLUMNS:
                        df.loc[mask, month] = row[month]
        else:
            df = cleaned_edited_df

        # Final validation pass on full dataset
        df, _ = validate_and_clean_data(df)

        # ============================================================
        # EXPORT / DOWNLOAD SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Choose Your Download Format")

        col1, col2 = st.columns(2)

        # OPTION 1: Clean Excel (.xlsx) with Formatted Cells & All Sheets
        with col1:
            st.markdown("### **Option 1: Complete Workbook (`.xlsx`)**")
            st.write("Preserves all sheets (`Entry`, `Units`, etc.) and formats percentages as `0.00%` directly in Excel.")

            output_xlsx_data = generate_formatted_excel(df, uploaded_file, target_sheet)
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
