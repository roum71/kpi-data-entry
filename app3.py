import io
import re
import gspread
import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="KPI Data Entry System", page_icon="📊", layout="wide"
)

# ============================================================
# GOOGLE OAUTH AUTHENTICATION (OPTION 1)
# ============================================================
# Handle both stable (st.user) and experimental (st.experimental_user) APIs safely
user_info = getattr(st, "user", None) or getattr(
    st, "experimental_user", None
)

if not user_info or not user_info.is_logged_in:
    st.title("🔒 KPI Data Entry System")
    st.caption("Multi-sheet KPI data entry with secure access control.")
    st.info("Please sign in using your Google account to access your KPIs.")
    if st.button("🔑 Log in with Google", use_container_width=True):
        st.login("google")
    st.stop()

# Retrieve verified Google Email from the session
CURRENT_USER_EMAIL = user_info.email

# Sidebar User Info & Logout
st.sidebar.markdown(f"👤 **Logged in as:**\n`{CURRENT_USER_EMAIL}`")
if st.sidebar.button("🚪 Log Out"):
    st.logout()
    st.rerun()



st.title("📊 KPI Data Entry System")
st.caption(
    "Multi-sheet KPI data entry with strict periodicity validation and live Google Sheets sync."
)

# ============================================================
# CONSTANTS & CONFIGURATION
# ============================================================

MONTH_COLUMNS = [str(i) for i in range(1, 13)]

# Hardcoded Google Sheet ID
SPREADSHEET_ID = "1Bsb7EiQ4-az1W12Xrmd6HVg9I4naB-JpBJQBkkdWYTs"

# Required OAuth Scopes for Google Sheets API
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

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
# GOOGLE SHEETS CONNECTION & USER ACCESS CONTROL
# ============================================================


@st.cache_resource
def get_gspread_client():
    """Authenticates using Streamlit Secrets for Google Service Account."""
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(credentials)


def load_user_permissions(spreadsheet_id, user_email):
    """Fetches allowed location codes for the logged-in email from 'User_Permissions' tab."""
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)

    try:
        ws = sh.worksheet("User_Permissions")
        perm_df = pd.DataFrame(ws.get_all_records())

        # Normalize email strings
        user_row = perm_df[
            perm_df["email"].astype(str).str.strip().str.lower()
            == user_email.strip().lower()
        ]

        if user_row.empty:
            return [], "none"

        locs_raw = str(user_row.iloc[0]["locations"])
        role = str(user_row.iloc[0].get("role", "user")).strip().lower()

        if locs_raw.strip().upper() == "ALL" or role == "admin":
            return ["ALL"], "admin"

        allowed_locs = [
            l.strip() for l in locs_raw.split(",") if l.strip()
        ]
        return allowed_locs, role

    except Exception as e:
        st.error(f"Error checking User_Permissions tab: {e}")
        return [], "none"


def load_google_sheet_data(spreadsheet_id):
    """Loads worksheets from Google Sheets (skips internal admin tabs)."""
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    sheets = {}

    for ws in sh.worksheets():
        if ws.title == "User_Permissions":
            continue  # Do not load administrative permissions table into memory

        data = ws.get_all_values()
        if not data:
            continue

        headers = [clean_header(c) for c in data[0]]
        df = pd.DataFrame(data[1:], columns=headers)

        for col in df.columns:
            col_str = str(col).strip()
            if col_str in MONTH_COLUMNS:
                df[col] = pd.to_numeric(df[col], errors="coerce").astype(
                    "float64"
                )
            elif col_str in EVIDENCE_ALIASES or col_str in COMMENT_ALIASES:
                df[col] = df[col].fillna("").astype(str)

        df = clean_invalid_entries(df)
        sheets[ws.title] = df

    return sheets


def save_sheet_to_google(spreadsheet_id, sheet_name, df):
    """Writes the updated DataFrame directly back to Google Sheet."""
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)

    clean_df = df.copy().fillna("")
    full_data = [clean_df.columns.values.tolist()] + clean_df.values.tolist()

    ws.clear()
    ws.update("A1", full_data)


# ============================================================
# HELPER FUNCTIONS
# ============================================================


def clean_header(col):
    if col is None:
        return ""
    col = str(col).replace("\xa0", " ")
    return re.sub(r"[\n\r\t]", " ", col).strip()


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


def clean_invalid_entries(df):
    df = df.copy()

    for m in MONTH_COLUMNS:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors="coerce").astype("float64")

    for idx in df.index:
        row = df.loc[idx]
        periodicity = get_value(row, FREQUENCY_ALIASES, "M")
        valid_months = allowed_months(periodicity)

        for month in MONTH_COLUMNS:
            if month in df.columns and month not in valid_months:
                df.loc[idx, month] = np.nan

    return df


# ============================================================
# ACCESS CONTROL CHECK
# ============================================================

allowed_locations, user_role = load_user_permissions(
    SPREADSHEET_ID, CURRENT_USER_EMAIL
)

if not allowed_locations:
    st.error(
        f"⛔ Access Denied: User `{CURRENT_USER_EMAIL}` is not authorized to access any location."
    )
    st.info(
        "Please ensure your email is registered in the `User_Permissions` worksheet tab."
    )
    st.stop()

st.sidebar.markdown(f"🔑 **Role:** `{user_role.upper()}`")
st.sidebar.markdown(f"📍 **Locations:** `{', '.join(allowed_locations)}`")

# ============================================================
# MAIN APPLICATION LOGIC
# ============================================================

if "kpi_sheets" not in st.session_state:
    with st.spinner("Connecting to Google Sheets..."):
        st.session_state["kpi_sheets"] = load_google_sheet_data(SPREADSHEET_ID)

sheets = st.session_state["kpi_sheets"]
available_sheets = list(sheets.keys())
selected_sheet = st.selectbox("Select Entry Sheet", available_sheets)

df = sheets[selected_sheet].copy()
df.columns = [str(c).strip() for c in df.columns]

location_col = find_column(df, LOCATION_ALIASES)
if not location_col:
    st.error("Sheet does not contain a valid Location column!")
    st.stop()

# HARD SERVER-SIDE FILTER BY PERMISSIONS
if "ALL" in allowed_locations:
    authorized_df = df.copy()
else:
    authorized_df = df[
        df[location_col].astype(str).str.strip().isin(allowed_locations)
    ].copy()

if authorized_df.empty:
    st.warning("No KPI rows match your assigned location permissions.")
    st.stop()

# ============================================================
# FILTERS
# ============================================================

st.subheader("🔎 Filters")
c1, c2, c3 = st.columns(3)

with c1:
    year_col = find_column(authorized_df, YEAR_ALIASES)
    if year_col:
        years = sorted(authorized_df[year_col].dropna().astype(str).unique())
        selected_year = st.selectbox("Year", ["All"] + years)
    else:
        selected_year = "All"

with c2:
    locations = sorted(
        authorized_df[location_col].dropna().astype(str).unique()
    )
    selected_location = st.selectbox("Location", locations)

with c3:
    periodicity_col = find_column(authorized_df, FREQUENCY_ALIASES)
    if periodicity_col:
        frequencies = sorted(
            authorized_df[periodicity_col].dropna().astype(str).unique()
        )
        selected_periodicity = st.selectbox(
            "Periodicity", ["All"] + frequencies
        )
    else:
        selected_periodicity = "All"

# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = authorized_df[
    authorized_df[location_col].astype(str) == selected_location
].copy()

if selected_year != "All" and year_col:
    filtered_df = filtered_df[
        filtered_df[year_col].astype(str) == selected_year
    ]

if selected_periodicity != "All" and periodicity_col:
    filtered_df = filtered_df[
        filtered_df[periodicity_col].astype(str) == selected_periodicity
    ]

# ============================================================
# DATA EDITOR
# ============================================================

st.subheader("📝 KPI Data Entry")

if selected_periodicity == "All":
    st.info(
        "🔒 **View-Only Mode**: Select a specific **Periodicity** above to enable data entry."
    )
    cols_to_display = list(filtered_df.columns)
    disabled_columns = list(filtered_df.columns)
    editable_columns = []
else:
    valid_display_months = allowed_months(selected_periodicity)
    cols_to_display = [
        col
        for col in filtered_df.columns
        if col not in MONTH_COLUMNS or col in valid_display_months
    ]

    editable_columns = [
        col
        for col in cols_to_display
        if col in MONTH_COLUMNS
        or col in EVIDENCE_ALIASES
        or col in COMMENT_ALIASES
    ]
    disabled_columns = [
        col for col in cols_to_display if col not in editable_columns
    ]

filtered_df = filtered_df[cols_to_display]
filtered_df = clean_invalid_entries(filtered_df)

column_config = {}
for col in filtered_df.columns:
    if col in MONTH_COLUMNS:
        column_config[col] = st.column_config.NumberColumn(
            col, format="%g", help="Numeric KPI value only"
        )
    elif col in EVIDENCE_ALIASES or col in COMMENT_ALIASES:
        column_config[col] = st.column_config.TextColumn(
            col, help="Editable text field"
        )

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

# ============================================================
# SECURE SAVE LOGIC
# ============================================================

if selected_periodicity != "All":
    if st.button(
        f"💾 Save Changes for Location {selected_location}",
        use_container_width=True,
    ):
        # Security Re-check
        if (
            "ALL" not in allowed_locations
            and selected_location not in allowed_locations
        ):
            st.error("🚨 Security Alert: Unauthorized location save attempt!")
            st.stop()

        master_df = sheets[selected_sheet].copy()
        code_col = find_column(master_df, KPI_CODE_ALIASES)

        if code_col:
            for idx, row in cleaned_df.iterrows():
                kpi_code = str(row[code_col])
                # Dual Key Match (KPI Code + Location)
                mask = (master_df[code_col].astype(str) == kpi_code) & (
                    master_df[location_col].astype(str) == selected_location
                )

                if mask.any():
                    row_periodicity = get_value(row, FREQUENCY_ALIASES, "M")
                    valid_m_for_row = allowed_months(row_periodicity)

                    for col in editable_columns:
                        if col in cleaned_df.columns:
                            val = row[col]
                            if (
                                col in MONTH_COLUMNS
                                and col not in valid_m_for_row
                            ):
                                val = np.nan

                            master_df.loc[mask, col] = (
                                np.nan if pd.isna(val) else val
                            )

        master_df = clean_invalid_entries(master_df)
        st.session_state["kpi_sheets"][selected_sheet] = master_df

        with st.spinner("Saving changes directly to Google Sheets..."):
            save_sheet_to_google(SPREADSHEET_ID, selected_sheet, master_df)

        st.success(
            f"Successfully saved updates for Location `{selected_location}`!"
        )
        st.rerun()
