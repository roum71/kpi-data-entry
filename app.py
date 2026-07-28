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
# HELPER FUNCTIONS
# ============================================================
def clean_header(header_name):
    if header_name is None:
        return ""
    header_str = str(header_name)
    header_str = header_str.replace('\xa0', ' ')
    header_str = re.sub(r'[\r\n\t]', ' ', header_str)
    return header_str.strip()

def get_valid_months(frequency):
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

def is_numeric(value):
    if pd.isna(value) or str(value).strip() == "":
        return True
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False

def validate_and_clean_data(df):
    """
    Validates data against frequency rules and numeric rules.
    If an error is found, the value is cleared (set to None/NaN) 
    so it is NEVER saved in the output file.
    """
    cleaned_df = df.copy()
    errors = []

    for index, row in cleaned_df.iterrows():
        frequency = row.get("measurement_frequency", None)
        valid_months = get_valid_months(frequency)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

            # If empty/blank, leave as is
            if pd.isna(value) or str(value).strip() == "":
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # Check Periodicity Rule
            if month_num not in valid_months:
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": f"Month {month_col} not allowed for frequency {frequency} (CLEARED)"
                })
                # CLEAR THE INVALID VALUE
                cleaned_df.loc[index, month_col] = np.nan
                continue

            # Check Numeric Rule
            if not is_numeric(value):
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": "Value must be numeric (CLEARED)"
                })
                # CLEAR THE INVALID VALUE
                cleaned_df.loc[index, month_col] = np.nan
            else:
                cleaned_df.loc[index, month_col] = float(value)

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

        edited_df = st.data_editor(
            working_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed"
        )

        # Validate and clean edits
        cleaned_edited_df, edit_errors = validate_and_clean_data(edited_df)

        if edit_errors:
            st.warning(f"⚠️ {len(edit_errors)} invalid entries detected. They have been automatically cleared and will NOT be saved.")
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

        # Final pass to ensure whole dataframe is clean before download
        df, _ = validate_and_clean_data(df)

        # ============================================================
        # DOWNLOAD SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Export Options")

        col1, col2 = st.columns(2)

        # Option 1: Download as updated .xlsm file
        with col1:
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=target_sheet, index=False)
                for sheet in excel_file.sheet_names:
                    if sheet != target_sheet:
                        other_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                        other_df.to_excel(writer, sheet_name=sheet, index=False)

            output_excel.seek(0)

            st.download_button(
                label="⬇️ Download Updated Excel (.xlsm)",
                data=output_excel,
                file_name=f"Updated_{uploaded_file.name}",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                type="primary",
                use_container_width=True
            )

        # Option 2: Download as CSV file (Entry sheet only)
        with col2:
            output_csv = df.to_csv(index=False).encode('utf-8-sig') # utf-8-sig keeps Arabic characters readable in Excel

            st.download_button(
                label="⬇️ Download Entry Sheet (.csv)",
                data=output_csv,
                file_name="Actual10_Entry.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error("Error reading Excel file.")
        st.exception(e)
else:
    st.info("👆 Please upload your `Actual10.xlsm` file to get started.")
