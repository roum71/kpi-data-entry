import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import openpyxl
import requests
import msal

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="KPI Data Entry System",
    page_icon="📊",
    layout="wide"
)

st.title("📊 KPI Data Entry System")
st.write("Upload your KPI Excel file, edit month entries with automated validation, and export or save to SharePoint.")

# ============================================================
# CONSTANTS & COLUMN ALIASES
# ============================================================
MONTH_COLUMNS = [str(i) for i in range(1, 13)]

FREQ_ALIASES = ["Freq", "measurement_frequency", "freq", "frequency", "التردد"]
UNIT_ALIASES = ["Unit", "unit_id", "unit", "Unit ID", "unit_code"]
UNIT_NAME_ALIASES = ["unit_name_ar", "Unit Name (AR)", "unit_name", "الوحدة"]
NEG_ALIASES  = ["allow_negative_values", "Allow Negative", "allow_negative", "allow_neg"]
CODE_ALIASES = ["kpi_code", "KPI Code", "code"]
NAME_ALIASES = ["kpi_name_ar", "KPI Name (AR)", "kpi_name", "name_ar"]

SHEET_NAME = "Entry"

# ============================================================
# DELEGATED SHAREPOINT AUTHENTICATION (NO ADMIN CONSENT NEEDED)
# ============================================================
def get_user_access_token():
    """Authenticates using Microsoft Device Flow (No Admin Consent required)."""
    cfg = st.secrets["sharepoint"]
    client_id = cfg["client_id"]
    tenant_id = cfg["tenant_id"]
    
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/Files.ReadWrite"]

    app = msal.PublicClientApplication(client_id, authority=authority)
    
    # Check if user already logged in during this session
    if "token_cache" in st.session_state:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]

    # If not logged in, initiate device code login
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" in flow:
        st.warning(f"🔑 **SharePoint Login Required:** Go to [{flow['verification_uri']}]({flow['verification_uri']}) and enter code: **`{flow['user_code']}`**")
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            st.session_state["token_cache"] = result
            return result["access_token"]
        else:
            raise Exception("Failed to authenticate user.")
    else:
        raise Exception("Could not start device login flow.")

def upload_to_sharepoint_as_user(file_bytes, file_name):
    """Uploads file buffer to SharePoint using the user's logged-in session."""
    token = get_user_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    cfg = st.secrets["sharepoint"]

    # 1. Fetch SharePoint Site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{cfg['hostname']}:{cfg['site_path']}"
    site_res = requests.get(site_url, headers=headers)
    site_res.raise_for_status()
    site_id = site_res.json()["id"]

    # 2. Upload file to target folder
    clean_folder_path = cfg["folder_path"].strip("/")
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{clean_folder_path}/{file_name}:/content"

    upload_res = requests.put(upload_url, headers=headers, data=file_bytes.getvalue())
    upload_res.raise_for_status()
    return upload_res.json()

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

def get_col_val(row, aliases, default=None):
    for name in aliases:
        if name in row and pd.notna(row[name]):
            return row[name]
    return default

def get_col_name(df, aliases):
    for col in df.columns:
        if col in aliases:
            return col
    return None

def get_valid_months(frequency):
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

def is_percentage_unit(unit_id, unit_name=None):
    if unit_name and "%" in str(unit_name):
        return True
    try:
        u_id = int(float(unit_id))
        return u_id in [10, 11, 12]
    except (ValueError, TypeError):
        return False

def merge_units_metadata(df, uploaded_file):
    df_merged = df.copy()
    try:
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        if "Units" in excel_file.sheet_names:
            units_df = pd.read_excel(uploaded_file, sheet_name="Units", engine="openpyxl")
            units_df.columns = [clean_header(c) for c in units_df.columns]
            
            code_col = get_col_name(units_df, CODE_ALIASES)
            unit_col = get_col_name(units_df, UNIT_ALIASES)
            unit_name_col = get_col_name(units_df, UNIT_NAME_ALIASES)

            if code_col:
                main_code_col = get_col_name(df_merged, CODE_ALIASES)
                if main_code_col:
                    clean_codes = units_df[code_col].astype(str).str.strip()

                    if unit_col:
                        unit_map = dict(zip(clean_codes, units_df[unit_col]))
                        main_unit_col = get_col_name(df_merged, UNIT_ALIASES)
                        mapped_units = df_merged[main_code_col].astype(str).str.strip().map(unit_map)
                        if main_unit_col:
                            df_merged[main_unit_col] = df_merged[main_unit_col].fillna(mapped_units)
                        else:
                            df_merged["unit_id"] = mapped_units

                    if unit_name_col:
                        name_map = dict(zip(clean_codes, units_df[unit_name_col]))
                        main_name_col = get_col_name(df_merged, UNIT_NAME_ALIASES)
                        mapped_names = df_merged[main_code_col].astype(str).str.strip().map(name_map)
                        if main_name_col:
                            df_merged[main_name_col] = df_merged[main_name_col].fillna(mapped_names)
                        else:
                            df_merged["unit_name_ar"] = mapped_names
    except Exception:
        pass
        
    return df_merged

def validate_and_clean_data(df, uploaded_file=None):
    cleaned_df = df.copy()
    
    if uploaded_file is not None:
        cleaned_df = merge_units_metadata(cleaned_df, uploaded_file)

    errors = []
    cleared_count = 0

    for index, row in cleaned_df.iterrows():
        frequency = get_col_val(row, FREQ_ALIASES)
        unit_id = get_col_val(row, UNIT_ALIASES)
        unit_name = get_col_val(row, UNIT_NAME_ALIASES, "")
        kpi_code = get_col_val(row, CODE_ALIASES, "")
        kpi_name = get_col_val(row, NAME_ALIASES, "")

        allow_neg_val = get_col_val(row, NEG_ALIASES, 0)
        try:
            is_neg_allowed = int(float(allow_neg_val)) == 1
        except (ValueError, TypeError):
            is_neg_allowed = False

        valid_months = get_valid_months(frequency)
        is_pct = is_percentage_unit(unit_id, unit_name)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in cleaned_df.columns:
                continue

            value = row[month_col]

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

            if pd.isna(value) or str(value).strip() == "" or str(value).strip().lower() == "none":
                cleaned_df.loc[index, month_col] = np.nan
                continue

            try:
                val_float = float(value)

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

                if is_pct:
                    if 1.0 < val_float <= 100.0:
                        val_float = val_float / 100.0

                    if (val_float < 0.0 and not is_neg_allowed) or val_float > 1.0:
                        errors.append({
                            "kpi_code": kpi_code,
                            "kpi_name_ar": kpi_name,
                            "Month": month_col,
                            "Value": value,
                            "Error": "Percentage out of range. Value cleared."
                        })
                        cleaned_df.loc[index, month_col] = np.nan
                        cleared_count += 1
                    else:
                        cleaned_df.loc[index, month_col] = val_float
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

def generate_unformatted_excel(df, uploaded_file, target_sheet_name):
    output_xlsx = io.BytesIO()
    excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
    df_to_export = merge_units_metadata(df, uploaded_file)

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df_to_export.to_excel(writer, sheet_name=target_sheet_name, index=False)
        for sheet in excel_file.sheet_names:
            if sheet != target_sheet_name:
                other_df = pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
                other_df.to_excel(writer, sheet_name=sheet, index=False)

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
            
            cleaned_master, _, _ = validate_and_clean_data(df_raw, uploaded_file)
            st.session_state["master_df"] = cleaned_master
            st.session_state["file_name"] = uploaded_file.name
            st.session_state["editor_version"] = 0

        df = st.session_state["master_df"]
        st.success(f"✅ Sheet '{target_sheet}' loaded successfully.")

        # Filters
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            years = sorted(df["year"].dropna().astype(str).unique().tolist()) if "year" in df.columns else []
            selected_year = st.selectbox("Select Year to Edit", years) if years else None

        with col_f2:
            freq_options = ["All Frequencies", "1 - Monthly", "2 - Quarterly", "3 - Semi-Annual", "4 - Annual"]
            selected_freq_str = st.selectbox("Filter by Frequency (Locks non-applicable months)", freq_options)

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
        st.caption("🔒 Non-month columns are locked. Negative values are blocked unless allow_negative_values = 1.")

        column_configs = {}
        for col in working_df.columns:
            if col not in MONTH_COLUMNS:
                column_configs[col] = st.column_config.TextColumn(disabled=True)

        allowed_months = get_valid_months(selected_freq_num) if selected_freq_num is not None else list(range(1, 13))

        for m in MONTH_COLUMNS:
            if m in working_df.columns:
                m_num = int(m)
                is_disabled = m_num not in allowed_months

                column_configs[m] = st.column_config.NumberColumn(
                    label=f"Month {m}",
                    disabled=is_disabled,
                    help="Locked for Frequency" if is_disabled else "Editable Month Entry"
                )

        editor_key = f"kpi_editor_v_{st.session_state['editor_version']}"
        
        edited_df = st.data_editor(
            working_df,
            column_config=column_configs,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=editor_key
        )

        cleaned_edited_df, edit_errors, cleared_count = validate_and_clean_data(edited_df, uploaded_file)

        if cleared_count > 0 or not cleaned_edited_df.equals(working_df):
            code_col = get_col_name(df, CODE_ALIASES)

            if code_col:
                for idx, row in cleaned_edited_df.iterrows():
                    kpi_code_val = row[code_col]
                    mask = df[code_col] == kpi_code_val
                    if selected_year and "year" in df.columns:
                        mask = mask & (df["year"].astype(str) == str(selected_year))
                    for m in MONTH_COLUMNS:
                        if m in df.columns:
                            df.loc[mask, m] = row[m]

            st.session_state["master_df"] = df

            if cleared_count > 0:
                st.session_state["editor_version"] += 1
                st.warning(f"⚠️ {cleared_count} invalid or disallowed entry(ies) were wiped.")
                st.rerun()

        if edit_errors:
            with st.expander("View Wiped Entries Log"):
                st.dataframe(pd.DataFrame(edit_errors), use_container_width=True)

        # ============================================================
        # EXPORT & SHAREPOINT SAVE SECTION
        # ============================================================
        st.markdown("---")
        st.subheader("💾 Export Options & Cloud Save")

        col1, col2, col3 = st.columns(3)

        output_xlsx_data = generate_unformatted_excel(df, uploaded_file, target_sheet)
        download_name_xlsx = uploaded_file.name.rsplit('.', 1)[0] + "_updated.xlsx"

        # 1. Excel Download
        with col1:
            st.markdown("### **Option 1: Complete Workbook**")
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=output_xlsx_data,
                file_name=download_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        # 2. CSV Download
        with col2:
            st.markdown("### **Option 2: Data Table Only**")
            output_csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="⬇️ Download CSV (.csv)",
                data=output_csv,
                file_name="Entry_updated.csv",
                mime="text/csv",
                use_container_width=True
            )

        # 3. Save to SharePoint
        with col3:
            st.markdown("### **Option 3: Save to SharePoint**")
            if st.button("☁️ Save to SharePoint", type="secondary", use_container_width=True):
                with st.spinner("Connecting to SharePoint..."):
                    try:
                        res = upload_to_sharepoint_as_user(output_xlsx_data, download_name_xlsx)
                        st.success(f"✅ Saved directly to SharePoint!\n**File:** `{download_name_xlsx}`")
                    except Exception as sp_err:
                        st.error("❌ Failed to upload file to SharePoint.")
                        st.exception(sp_err)

    except Exception as e:
        st.error("An error occurred while processing the Excel workbook.")
        st.exception(e)
else:
    st.info("👆 Please upload your Excel file to get started.")
