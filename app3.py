import io
import re
import gspread
import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from streamlit_oauth import OAuth2Component

# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="KPI Data Entry System", page_icon="📊", layout="wide"
)

# ============================================================
# GOOGLE OAUTH CONFIGURATION (VIA STREAMLIT-OAUTH)
# ============================================================

CLIENT_ID = st.secrets["oauth2"]["client_id"]
CLIENT_SECRET = st.secrets["oauth2"]["client_secret"]
AUTHORIZE_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
REVOKE_ENDPOINT = "https://oauth2.googleapis.com/revoke"

oauth2 = OAuth2Component(
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET,
    authorize_endpoint=AUTHORIZE_ENDPOINT,
    token_endpoint=TOKEN_ENDPOINT,
    revoke_endpoint=REVOKE_ENDPOINT,
)

if "token" not in st.session_state:
    st.session_state["token"] = None

if not st.session_state["token"]:
    st.title("🔒 KPI Data Entry System")
    st.caption("Multi-sheet KPI data entry with secure access control.")
    st.info("Please sign in using your Google account to access your KPIs.")
    
    # Render Google Login button
    result = oauth2.authorize_button(
        name="Log in with Google",
        icon="https://www.google.com/favicon.ico",
        redirect_uri="https://kpientry.streamlit.app/",
        scope="openid email profile",
        key="google",
        use_container_width=True,
    )
    
    if result and "token" in result:
        st.session_state["token"] = result["token"]
        st.rerun()
    st.stop()

# Decode/Extract user email from ID token or user info endpoint
import json
import base64

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

id_token = st.session_state["token"].get("id_token", "")
claims = parse_jwt(id_token)
CURRENT_USER_EMAIL = claims.get("email", "")

if not CURRENT_USER_EMAIL:
    st.error("Failed to retrieve verified email from token. Please log out and try again.")
    if st.button("Clear Session"):
        st.session_state["token"] = None
        st.rerun()
    st.stop()

# Sidebar User Info & Logout
st.sidebar.markdown(f"👤 **Logged in as:**\n`{CURRENT_USER_EMAIL}`")
if st.sidebar.button("🚪 Log Out"):
    st.session_state["token"] = None
    st.rerun()

st.title("📊 KPI Data Entry System")
st.caption(
    "Multi-sheet KPI data entry with strict periodicity validation and live Google Sheets sync."
)
