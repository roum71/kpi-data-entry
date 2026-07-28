import streamlit as st
import pandas as pd
import requests
import msal
import numpy as np
import re

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="KPI Data Entry",
    page_icon="📊",
    layout="wide"
)
st.title("📊 KPI Data Entry System")

# ============================================================
# CONFIGURATION
# ============================================================
# Configured based on your Power Query M code setup:
# Folder: HR Dashboard
# File: Actual9.xlsm
# Table/Sheet Name: Entry

TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
SHAREPOINT_HOSTNAME = st.secrets.get("SHAREPOINT_HOSTNAME", "egarakae-my.sharepoint.com")
SITE_PATH = st.secrets.get("SITE_PATH", "/personal/kareem_b_ec_rak_ae")

# Matches Power Query file & table targets:
FILE_PATH = st.secrets.get("FILE_PATH", "HR Dashboard/Actual9.xlsm")
SHEET_NAME = st.secrets.get("SHEET_NAME", "Entry")

GRAPH_URL = "https://graph.microsoft.com/v1.0"

# ============================================================
# MONTH CONFIGURATION (1 to 12)
# ============================================================
MONTH_COLUMNS = [str(i) for i in range(1, 13)]

# ============================================================
# REQUIRED COLUMNS
# ============================================================
REQUIRED_COLUMNS = [
    "sub_objective_id",
    "kpi_code",
    "location_id",
    "year",
    "measurement_frequency",
    "kpi_name_ar"
] + MONTH_COLUMNS

# ============================================================
# MICROSOFT GRAPH AUTHENTICATION
# ============================================================
@st.cache_resource
def get_access_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise Exception("Could not authenticate with Microsoft Graph: " + str(result))
    return result["access_token"]

@st.cache_resource
def get_site_id(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()["id"]

@st.cache_resource
def get_drive_id(access_token, site_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/sites/{site_id}/drives"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    drives = response.json()["value"]
    
    for drive in drives:
        if drive["name"].lower() in ["documents", "shared documents"]:
            return drive["id"]
    return drives[0]["id"]

def get_file_item(access_token, drive_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    encoded_path = requests.utils.quote(FILE_PATH)
    url = f"{GRAPH_URL}/drives/{drive_id}/root:/{encoded_path}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# ============================================================
# CLEAN HEADER HELPER (Matches Power Query Text.Clean & Trim)
# ============================================================
def clean_header(header_name):
    if header_name is None:
        return ""
    header_str = str(header_name)
    # Replace non-breaking spaces (ASCII 160) and clean control characters
    header_str = header_str.replace('\xa0', ' ')
    header_str = re.sub(r'[\r\n\t]', ' ', header_str)
    return header_str.strip()

# ============================================================
# READ EXCEL WORKSHEET / TABLE
# ============================================================
def read_excel_sheet(access_token, drive_id):
    file_item = get_file_item(access_token, drive_id)
    item_id = file_item["id"]
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Try reading as Excel Table first (Power Query Target Table = "Entry")
    table_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
        f"tables/{SHEET_NAME}/range(valuesOnly=true)"
    )
    response = requests.get(table_url, headers=headers)
    
    # If table endpoint fails, fall back to worksheet endpoint
    if response.status_code != 200:
        worksheet_url = (
            f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
            f"worksheets/{SHEET_NAME}/usedRange(valuesOnly=true)"
        )
        response = requests.get(worksheet_url, headers=headers)
    
    response.raise_for_status()
    values = response.json().get("values", [])
    
    if not values:
        return pd.DataFrame()
    
    # Apply Power Query header cleaning logic
    headers_row = [clean_header(col) for col in values[0]]
    data_rows = values[1:]
    
    df = pd.DataFrame(data_rows, columns=headers_row)
    return df

# ============================================================
# UPDATE EXCEL WORKSHEET / TABLE
# ============================================================
def update_excel_sheet(access_token, drive_id, df):
    file_item = get_file_item(access_token, drive_id)
    item_id = file_item["id"]
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    output_df = df.copy()
    output_df = output_df.replace({np.nan: None, pd.NaT: None})
    values = [output_df.columns.tolist()] + output_df.values.tolist()
    
    # Check used range
    range_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
        f"worksheets/{SHEET_NAME}/usedRange"
    )
    response = requests.get(range_url, headers=headers)
    
    # Fallback if worksheet fails (in case it's strictly a named table)
    if response.status_code != 200:
        range_url = (
            f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
            f"tables/{SHEET_NAME}/range"
        )
        response = requests.get(range_url, headers=headers)
        
    response.raise_for_status()
    used_range = response.json()["address"]
    
    clear_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
        f"worksheets/{SHEET_NAME}/range(address='{used_range}')/clear"
    )
    requests.post(clear_url, headers=headers, json={"applyTo": "Contents"})
    
    number_of_rows = len(values)
    number_of_columns = len(values[0])
    
    def excel_column_name(number):
        result = ""
        while number > 0:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    last_column = excel_column_name(number_of_columns)
    target_range = f"A1:{last_column}{number_of_rows}"
    
    write_url = (
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/workbook/"
        f"worksheets/{SHEET_NAME}/range(address='{target_range}')"
    )
    response = requests.patch(write_url, headers=headers, json={"values": values})
    response.raise_for_status()
    return True

# ============================================================
# PERIODICITY / FREQUENCY RULES
# ============================================================
def get_valid_months(frequency):
    """
    1: Monthly   -> Months [1..12]
    2: Semi-Annual -> Months [6, 12]
    3: Quarterly -> Months [3, 6, 9, 12]
    4: Annually  -> Month [12]
    """
    try:
        freq = int(float(frequency))
    except (ValueError, TypeError):
        return []

    if freq == 1:
        return list(range(1, 13))
    elif freq == 2:
        return [6, 12]
    elif freq == 3:
        return [3, 6, 9, 12]
    elif freq == 4:
        return [12]
    
    return []

# ============================================================
# NUMERIC VALIDATION
# ============================================================
def is_numeric(value):
    if pd.isna(value) or str(value).strip() == "":
        return True
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False

# ============================================================
# VALIDATE DATA
# ============================================================
def validate_data(df):
    df = df.copy()
    errors = []

    for index, row in df.iterrows():
        frequency = row.get("measurement_frequency", None)
        valid_months = get_valid_months(frequency)

        for month_num in range(1, 13):
            month_col = str(month_num)
            if month_col not in df.columns:
                continue

            value = row[month_col]

            if pd.isna(value) or str(value).strip() == "":
                continue

            # Check Periodicity Rule
            if month_num not in valid_months:
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": f"Month {month_col} not allowed for frequency {frequency}"
                })
                df.loc[index, month_col] = np.nan
                continue

            # Check Numeric Rule
            if not is_numeric(value):
                errors.append({
                    "kpi_code": row.get("kpi_code", ""),
                    "kpi_name_ar": row.get("kpi_name_ar", ""),
                    "Month": month_col,
                    "Value": value,
                    "Error": "Value must be numeric"
                })
                df.loc[index, month_col] = np.nan
            else:
                df.loc[index, month_col] = float(value)

    return df, errors

# ============================================================
# MAIN APPLICATION
# ============================================================
try:
    with st.spinner("Connecting to SharePoint..."):
        access_token = get_access_token()
        site_id = get_site_id(access_token)
        drive_id = get_drive_id(access_token, site_id)
    st.success("✅ Connected to SharePoint successfully.")
except Exception as e:
    st.error("❌ Could not connect to SharePoint.")
    st.exception(e)
    st.stop()

# ============================================================
# LOAD DATA
# ============================================================
if st.button("🔄 Load KPI Data (Actual9.xlsm)", type="primary"):
    try:
        with st.spinner("Reading 'Entry' table from Actual9.xlsm..."):
            df = read_excel_sheet(access_token, drive_id)

        if df.empty:
            st.warning("The Excel table 'Entry' is empty.")
            st.stop()

        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            st.error("Missing required columns:")
            st.write(missing_columns)
            st.stop()

        df, errors = validate_data(df)

        if errors:
            st.warning(f"{len(errors)} invalid values were found and cleared.")
            with st.expander("View validation errors"):
                st.dataframe(pd.DataFrame(errors), use_container_width=True)
        else:
            st.success("No invalid values found.")

        st.session_state["kpi_df"] = df

    except Exception as e:
        st.error("Error loading KPI data.")
        st.exception(e)

# ============================================================
# DISPLAY & EDIT DATA
# ============================================================
if "kpi_df" in st.session_state:
    df = st.session_state["kpi_df"]

    st.subheader("📊 KPI Data (Actual9.xlsm - Entry Table)")

    # Select Year Filter
    years = sorted(df["year"].dropna().astype(str).unique().tolist())
    if years:
        selected_year = st.selectbox("Select Year", years)
        working_df = df[df["year"].astype(str) == str(selected_year)].copy()
    else:
        working_df = df.copy()

    # Interactive Data Editor
    edited_df = st.data_editor(
        working_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed"
    )

    # Validate User Entries
    validated_df, errors = validate_data(edited_df)

    if errors:
        st.error(f"{len(errors)} invalid entries detected and cleared.")
        with st.expander("View invalid entries"):
            st.dataframe(pd.DataFrame(errors), use_container_width=True)
    else:
        st.success("✅ All entered values are valid.")

    # Save Back to SharePoint
    if st.button("💾 Save Changes to SharePoint", type="primary", use_container_width=True):
        final_df, _ = validate_data(validated_df)

        for index, row in final_df.iterrows():
            kpi_code = row["kpi_code"]
            mask = (df["year"].astype(str) == str(selected_year)) & (df["kpi_code"] == kpi_code)

            if mask.any():
                for month in MONTH_COLUMNS:
                    df.loc[mask, month] = row[month]

        try:
            with st.spinner("Saving data to SharePoint Excel..."):
                update_excel_sheet(access_token, drive_id, df)

            st.session_state["kpi_df"] = df
            st.success("✅ KPI data saved successfully to SharePoint Excel.")
        except Exception as e:
            st.error("❌ Error saving data to Excel.")
            st.exception(e)

    st.subheader("📋 Current KPI Data")
    st.dataframe(df, use_container_width=True, hide_index=True)