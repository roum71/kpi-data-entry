import base64
import json
import re
import gspread
import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from streamlit_oauth import OAuth2Component

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="KPI Data Entry System", page_icon="📊", layout="wide"
)

# ============================================================
# OAUTH & GOOGLE SHEETS CONSTANTS
# ============================================================

CLIENT_ID = st.secrets["oauth2"]["client_id"]
CLIENT_SECRET = st.secrets["oauth2"]["client_secret"]
AUTHORIZE_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
REVOKE_ENDPOINT = "https://oauth2.googleapis.com/revoke"
REDIRECT_URI = "https://kpientry.streamlit.app/"

SPREADSHEET_ID = "1Bsb7EiQ4-az1W12Xrmd6HVg9I4naB-JpBJQBkkdWYTs"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MONTH_COLUMNS = [str(i) for i in range(1, 13)]
YEAR_ALIASES = ["Year", "year", "السنة"]
LOCATION_ALIASES = ["Location", "location", "Location Code", "location_code", "الموقع"]
FREQUENCY_ALIASES = ["Periodicity", "periodicity", "Freq", "frequency", "Frequency", "التكرار"]
KPI_CODE_ALIASES = ["KPI Code", "kpi_code", "Code", "code"]
KPI_NAME_ALIASES = ["KPI Name", "KPI Name (AR)", "kpi_name", "name"]
EVIDENCE_ALIASES = ["Evidence Status", "evidence_status", "Evidence", "evodence", "حالة الدليل", "Status", "status"]
COMMENT_ALIASES = ["Comments", "Comment", "comments", "ملاحظات"]

# ============================================================
# HELPER FUNCTIONS & CACHED GOOGLE CONNECTION
# ============================================================

@st.cache_resource
def get_gspread_client():
    info = dict(st.secrets["gcp_service_account"])
    
    # Extract and normalize the private key string to handle all TOML escaping variations
    pk = str(info["private_key"])
    pk = pk.replace("\\n", "\n")
    
    if "BEGIN PRIVATE KEY" in pk and not pk.startswith("-----BEGIN PRIVATE KEY-----"):
        parts = pk.split("-----BEGIN PRIVATE KEY-----")
        pk = "-----BEGIN PRIVATE KEY-----" + parts[-1]
    if "END PRIVATE KEY" in pk and not pk.endswith("-----END PRIVATE KEY-----"):
        parts = pk.split("-----END PRIVATE KEY-----")
        pk = parts[0] + "-----END PRIVATE KEY-----"

    credentials = service_account.Credentials.from_service_account_info(
        {
            "type": info.get("type", "service_account"),
            "project_id": info.get("project_id"),
            "private_key_id": info.get("private_key_id"),
            "private_key": pk.strip(),
            "client_email": info.get("client_email"),
            "client_id": info.get("client_id"),
            "token_uri": info.get("token_uri", "https://oauth2.googleapis.com/token"),
        },
        scopes=SCOPES
    )
    
    return gspread.authorize(credentials)

def clean_header(col):
    if col is None:
        return ""
    col = str(col).replace("\xa0", " ")
    return re.sub(r"[\n\r\t]", " ", col).strip()

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
    return MONTH_COLUMNS

def find_column(df, aliases):
    for c in df.columns:
        if c in aliases:
            return c
    return None

def get_value(row, aliases, default=None):
    for c in aliases:
        if c in row.index and pd.notna(row[c]):
            return row[c]
    return default

def clean_invalid_entries(df):
    df = df.copy()
    for m in MONTH_COLUMNS:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors="coerce")

    for idx in df.index:
        row = df.loc[idx]
        periodicity = get_value(row, FREQUENCY_ALIASES, "M")
        valid_months = allowed_months(periodicity)

        for month in MONTH_COLUMNS:
            if month in df.columns and month not in valid_months:
                df.loc[idx, month] = np.nan

    return df

@st.cache_data(ttl=60)
def load_google_sheet_data_cached(spreadsheet_id):
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    sheets = {}

    for ws in sh.worksheets():
        if ws.title == "User_Permissions":
            continue

        data = ws.get_all_values()
        if not data:
            continue

        headers = [clean_header(c) for c in data[0]]
        df = pd.DataFrame(data[1:], columns=headers)

        for col in df.columns:
            col_str = str(col).strip()
            if col_str in MONTH_COLUMNS:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            elif col_str in EVIDENCE_ALIASES or col_str in COMMENT_ALIASES:
                df[col] = df[col].fillna("").astype(str)

        df = clean_invalid_entries(df)
        sheets[ws.title] = df

    return sheets

def save_sheet_to_google(spreadsheet_id, sheet_name, df):
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)

    clean_df = df.copy().fillna("")
    full_data = [clean_df.columns.values.tolist()] + clean_df.values.tolist()

    ws.clear()
    ws.update("A1", full_data)
    st.cache_data.clear()

# ============================================================
# GOOGLE OAUTH AUTHENTICATION STEP
# ============================================================

if "token" not in st.session_state:
    st.session_state["token"] = None

if not st.session_state["token"]:
    st.title("🔒 KPI Data Entry System")
    st.caption("Multi-sheet KPI data entry with secure access control.")
    st.info("Please sign in using your Google account to access your KPIs.")

    oauth2 = OAuth2Component(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        authorize_endpoint=AUTHORIZE_ENDPOINT,
        token_endpoint=TOKEN_ENDPOINT,
        revoke_token_endpoint=REVOKE_ENDPOINT,
    )

    try:
        result = oauth2.authorize_button(
            name="Log in with Google",
            icon="https://www.google.com/favicon.ico",
            redirect_uri=REDIRECT_URI,
            scope="openid email profile",
            key="google_auth",
            use_container_width=True,
        )

        if result and "token" in result:
            st.session_state["token"] = result["token"]
            st.query_params.clear()
            st.rerun()

    except Exception as e:
        st.error(f"Authentication Error: {e}")
        st.session_state["token"] = None

    st.stop()

# ============================================================
# EXTRACT USER INFO
# ============================================================

def parse_jwt(token_str):
    try:
        parts = token_str.split(".")
        if len(parts) > 1:
            payload = parts[1]
            padded = payload + "=" * (-len(payload) % 4)
            return json.loads(base64.urlsafe_b64decode(padded))
    except Exception:
        pass
    return {}

token_data = st.session_state["token"]
id_token = token_data.get("id_token", "") if isinstance(token_data, dict) else ""
claims = parse_jwt(id_token)

CURRENT_USER_EMAIL = claims.get("email", "").strip().lower()
if not CURRENT_USER_EMAIL and isinstance(token_data, dict):
    CURRENT_USER_EMAIL = token_data.get("email", "").strip().lower()

if not CURRENT_USER_EMAIL:
    st.error("⚠️ Failed to verify Google Email.")
    if st.button("Retry Sign In"):
        st.session_state["token"] = None
        st.query_params.clear()
        st.rerun()
    st.stop()

# Sidebar User Info
st.sidebar.markdown(f"👤 **User:** `{CURRENT_USER_EMAIL}`")
if st.sidebar.button("🚪 Log Out"):
    st.session_state["token"] = None
    st.query_params.clear()
    st.rerun()

# ============================================================
# FETCH PERMISSIONS DIRECTLY
# ============================================================

def load_user_permissions(spreadsheet_id, user_email):
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet("User_Permissions")
    perm_df = pd.DataFrame(ws.get_all_records())

    if perm_df.empty:
        return [], "none"

    user_row = perm_df[
        perm_df["email"].astype(str).str.strip().str.lower() == user_email.strip().lower()
    ]

    if user_row.empty:
        return [], "none"

    locs_raw = str(user_row.iloc[0]["locations"])
    role = str(user_row.iloc[0].get("role", "user")).strip().lower()

    if locs_raw.strip().upper() == "ALL" or role == "admin":
        return ["ALL"], "admin"

    allowed_locs = [l.strip() for l in locs_raw.split(",") if l.strip()]
    return allowed_locs, role

allowed_locations, user_role = load_user_permissions(SPREADSHEET_ID, CURRENT_USER_EMAIL)

if not allowed_locations:
    st.error(f"⛔ Access Denied for `{CURRENT_USER_EMAIL}`.")
    st.warning("Your email is not authorized in the `User_Permissions` Google Sheet tab.")
    st.stop()

st.sidebar.markdown(f"🔑 **Role:** `{user_role.upper()}`")
st.sidebar.markdown(f"📍 **Locations:** `{', '.join(allowed_locations)}`")

# ============================================================
# MAIN APPLICATION PAGE
# ============================================================

st.title("📊 KPI Data Entry System")
st.caption("Multi-sheet KPI data entry with live Google Sheets sync.")

with st.spinner("Loading KPI Sheets from Google..."):
    sheets = load_google_sheet_data_cached(SPREADSHEET_ID)

if "kpi_sheets" not in st.session_state:
    st.session_state["kpi_sheets"] = sheets

available_sheets = list(sheets.keys())
selected_sheet = st.selectbox("Select Entry Sheet", available_sheets)

df = st.session_state["kpi_sheets"][selected_sheet].copy()
df.columns = [str(c).strip() for c in df.columns]

location_col = find_column(df, LOCATION_ALIASES)
if not location_col:
    st.error("Sheet missing valid Location column!")
    st.stop()

if "ALL" in allowed_locations:
    authorized_df = df.copy()
else:
    authorized_df = df[df[location_col].astype(str).str.strip().isin(allowed_locations)].copy()

if authorized_df.empty:
    st.warning("No KPI rows match your assigned location permissions.")
    st.stop()

# ============================================================
# DATA ENTRY & SAVE
# ============================================================

st.subheader("🔎 Filters")
c1, c2, c3 = st.columns(3)

with c1:
    year_col = find_column(authorized_df, YEAR_ALIASES)
    years = sorted(authorized_df[year_col].dropna().astype(str).unique()) if year_col else []
    selected_year = st.selectbox("Year", ["All"] + years) if year_col else "All"

with c2:
    locations = sorted(authorized_df[location_col].dropna().astype(str).unique())
    selected_location = st.selectbox("Location", locations)

with c3:
    periodicity_col = find_column(authorized_df, FREQUENCY_ALIASES)
    frequencies = sorted(authorized_df[periodicity_col].dropna().astype(str).unique()) if periodicity_col else []
    selected_periodicity = st.selectbox("Periodicity", ["All"] + frequencies) if periodicity_col else "All"

filtered_df = authorized_df[authorized_df[location_col].astype(str) == selected_location].copy()

if selected_year != "All" and year_col:
    filtered_df = filtered_df[filtered_df[year_col].astype(str) == selected_year]

if selected_periodicity != "All" and periodicity_col:
    filtered_df = filtered_df[filtered_df[periodicity_col].astype(str) == selected_periodicity]

st.subheader("📝 KPI Data Entry")

if selected_periodicity == "All":
    st.info("🔒 **View-Only Mode**: Select a specific **Periodicity** above to enable data entry.")
    cols_to_display = list(filtered_df.columns)
    disabled_columns = list(filtered_df.columns)
    editable_columns = []
else:
    valid_display_months = allowed_months(selected_periodicity)
    cols_to_display = [
        col for col in filtered_df.columns
        if col not in MONTH_COLUMNS or col in valid_display_months
    ]
    editable_columns = [
        col for col in cols_to_display
        if col in MONTH_COLUMNS or col in EVIDENCE_ALIASES or col in COMMENT_ALIASES
    ]
    disabled_columns = [col for col in cols_to_display if col not in editable_columns]

filtered_df = filtered_df[cols_to_display]
filtered_df = clean_invalid_entries(filtered_df)

column_config = {}
for col in filtered_df.columns:
    if col in MONTH_COLUMNS:
        column_config[col] = st.column_config.NumberColumn(col, format="%g", help="Numeric KPI value only")
    elif col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
        column_config[col] = st.column_config.TextColumn(col, help="Editable text field")

edited_df = st.data_editor(
    filtered_df,
    column_config=column_config,
    disabled=disabled_columns,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key=f"editor_{selected_sheet}_{selected_location}_{selected_periodicity}",
)

cleaned_df = clean_invalid_entries(edited_df)

if selected_periodicity != "All":
    if st.button(f"💾 Save Changes for Location {selected_location}", use_container_width=True):
        if "ALL" not in allowed_locations and selected_location not in allowed_locations:
            st.error("🚨 Security Alert: Unauthorized location save attempt!")
            st.stop()

        master_df = st.session_state["kpi_sheets"][selected_sheet].copy()
        code_col = find_column(master_df, KPI_CODE_ALIASES)

        if code_col:
            for idx, row in cleaned_df.iterrows():
                kpi_code = str(row[code_col])
                mask = (master_df[code_col].astype(str) == kpi_code) & (master_df[location_col].astype(str) == selected_location)

                if mask.any():
                    row_periodicity = get_value(row, FREQUENCY_ALIASES, "M")
                    valid_m_for_row = allowed_months(row_periodicity)

                    for col in editable_columns:
                        if col in cleaned_df.columns:
                            val = row[col]
                            if col in MONTH_COLUMNS and col not in valid_m_for_row:
                                val = np.nan
                            master_df.loc[mask, col] = np.nan if pd.isna(val) else val

        master_df = clean_invalid_entries(master_df)
        st.session_state["kpi_sheets"][selected_sheet] = master_df

        with st.spinner("Saving changes directly to Google Sheets..."):
            save_sheet_to_google(SPREADSHEET_ID, selected_sheet, master_df)

        st.success(f"Successfully saved updates for Location `{selected_location}`!")
        st.rerun()
