import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="KPI Data Entry System",
    page_icon="📊",
    layout="wide"
)

st.title("📊 KPI Data Entry System")
st.write("Upload your Excel file (`Actual10.xlsm`), edit your KPI entries, and download the updated workbook.")

# ============================================================
# MONTH & COLUMN CONFIGURATION
# ============================================================
MONTH_COLUMNS = [str(i) for i in range(1, 13)]

REQUIRED_COLUMNS = [
    "sub_objective_id",
    "kpi_code",
    "location_id",
    "year",
    "measurement_frequency",
    "kpi_name_ar"
] + MONTH_COLUMNS

SHEET_NAME = "Entry"

# ============================================================
# HELPER FUNCTIONS (Aligned with VBA Logic)
# ============================================================
def clean_header(header_name):
    """Cleans column headers to match Power Query standards."""
    if header_name is None:
        return ""
    header_str = str(header_name)
    header_str = header_str.replace('\xa0', ' ')
    header_str = re.sub(r'[\r\n\t]', ' ', header_str)
    return header_str.strip()

def get_valid_months(frequency):
    """
    Frequency mapping matching VBA ColorByFrequency logic:
    1 = Monthly     -> [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    2 = Quarterly   -> [3, 6, 9, 12]
    3 = Semi-Annual -> [6, 12]
    4 = Annual      -> [12]
    """
    try:
        freq = int(float(frequency))
    except (ValueError, TypeError):
        return []

    if freq == 1:      # Monthly
        return list(range(1, 13))
    elif freq == 2:    # Quarterly
        return [3, 6, 9, 12]
    elif freq == 3:    # Semi-Annual
        return [6, 12]
    elif freq == 4:    # Annual
        return [12]
    
    return []

def is_percentage_unit(unit_id):
    """Matches VBA IsPercentageUnit function (Unit IDs 10, 11, 12)."""
    try:
        u_id = int(float(unit_id))
        return u_id in [10, 11, 12]
    except (ValueError, TypeError):
        return False

def validate_and_clean_data(df):
    """
    Validates data against frequency, negative value allowance, and percentage limits.
    Matches VBA ApplyFormatAndValidation rules.
    Invalid inputs are automatically CLEARED (set to NaN) to protect saved file integrity.
    """
    cleaned_df = df.copy()
    errors = []

    for index, row in cleaned_df.iterrows():
        frequency = row.get("measurement_frequency", None)
        unit_id = row.get("unit_id", None)
        allow_negative = row.get("allow_negative_values", 0)

        valid_months = get_valid_months(frequency)
        is_pct = is_percentage_unit(unit_id)

        # Min value rule from VBA: if allow_negative == 1 -> -1 Billion, else 0
        min_val = -1000000000.0 if str(allow_negative).strip() == "1" else 0.0
        max_val = 1.0 if is_pct else 1000000000.0

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

            # Leave empty cells untouched
            if pd.isna(value) or str(value).strip() == "":
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # 1. Check Periodicity/Frequency Rule
            if month_num not in valid_months:
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": f"Month {month_col} not allowed for frequency {frequency} (CLEARED)"
                })
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # 2. Check Numeric & Range Rule
            try:
                val_float = float(value)
                
                if val_float < min_val or val_float > max_val:
                    err_msg = "Percentage must be between 0 and 1" if is_pct else f"Value must be >= {min_val}"
                    errors.append({
                        "kpi_code": row.get("kpi_code", ""),
                        "kpi_name_ar": row.get("kpi_name_ar", ""),
                        "Month": month_col,
                        "Value": value,
                        "Error": f"{err_msg} (CLEARED)"
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
                    "Error": "Value must be numeric (CLEARED)"
                })
                cleaned_df.loc[index, month_col] = np.nan

    return cleaned_df, errors

# ============================================================
# FILE UPLOAD SECTION
# ============================================================
uploaded_file = st.file_uploader(
    "Choose your KPI Excel File (`Actual10.xlsm`)", 
    type=["xlsm", "xlsx", "xls"]
)

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

        st.success(f"✅ Loaded sheet '{target_sheet}' successfully.")

        # ============================================================
        # DATA EDITING SECTION
        # ============================================================
        st.subheader("📝 Edit KPI Data")

        years = sorted(df["year"].dropna().astype(str).unique().tolist())
        if years:
            selected_year = st.selectbox("Select Year to Edit", years)
            working_df = df[df["year"].astype(str) == str(selected_year)].copy()
        else:
            selected_year = None
            working_df = df.copy()

        # Formatting monthly display columns
        column_configs = {}
        for m in MONTH_COLUMNS:
            column_configs[m] = st.column_config.NumberColumn(
                label=f"Month {m}",
                format="%.2f",
                help="Enter numeric values"
            )

        edited_df = st.data_editor(
            working_df,
            column_config=column_configs,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed"
        )

        # Validate and clean edits
        cleaned_edited_df, edit_errors = validate_and_clean_data(edited_df)

        if edit_errors:
            st.warning(f"⚠️ {len(edit_errors)} invalid entries detected. They have been automatically cleared to protect data integrity.")
            with st.expander("View cleared invalid entries"):
                st.dataframe(pd.DataFrame(edit_errors), use_container_width=True)
        else:
            st.success("✅ All current entries are valid.")

        # Merge cleaned edits back into master dataset
        if selected_year:
            for index, row in cleaned_edited_df.iterrows():
                kpi_code = row["kpi_code"]
                mask = (df["year"].astype(str) == str(selected_year)) & (df["kpi_code"] == kpi_code)
                if mask.any():
                    for month in MONTH_COLUMNS:
                        df.loc[mask, month] = row[month]
        else:
            df = cleaned_edited_df

        # Final cleaning pass on master dataframe before export
        df, _ = validate_and_clean_data(df)

        # ============================================================
        # DOWNLOAD SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Choose Your Download Format")

        col1, col2 = st.columns(2)

        # OPTION 1: Full Excel Workbook (.xlsx)
        with col1:
            st.markdown("### **Option 1: Complete Workbook (`.xlsx`)**")
            st.write("Preserves all sheets (`Entry`, `Units`, `extended`, `Locker`) in a single clean Excel file.")

            output_xlsx = io.BytesIO()
            with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
                # Write updated Entry sheet
                df.to_excel(writer, sheet_name=target_sheet, index=False)
                
                # Copy all other reference sheets untouched
                for sheet in excel_file.sheet_names:
                    if sheet != target_sheet:
                        other_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                        other_df.to_excel(writer, sheet_name=sheet, index=False)

            output_xlsx.seek(0)
            download_name_xlsx = uploaded_file.name.rsplit('.', 1)[0] + "_updated.xlsx"

            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=output_xlsx,
                file_name=download_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        # OPTION 2: Entry Sheet Only (.csv)
        with col2:
            st.markdown("### **Option 2: Data Table Only (`.csv`)**")
            st.write("Exports only the cleaned `Entry` sheet. Best for importing directly into **Power Query** or databases.")

            # utf-8-sig ensures Arabic characters read correctly in Excel/Power Query
            output_csv = df.to_csv(index=False).encode('utf-8-sig')

            st.download_button(
                label="⬇️ Download CSV (.csv)",
                data=output_csv,
                file_name="Actual10_Entry_updated.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error("Error reading Excel file.")
        st.exception(e)
else:
    st.info("👆 Please upload your `Actual10.xlsm` file to get started.")
