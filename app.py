"""
NC Registration System – Streamlit Web App
Storage backend: Google Sheets (permanent, survives all restarts/redeployments).
Local Excel fallback when Google Sheets is not configured.
"""

import io
import json
import datetime

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# Constants  (exact values from original Tkinter app)
# ─────────────────────────────────────────────
FILE_NAME = "NC_Registration.xlsx"

RESPONSIBILITY_MAP = {
    "SMT": "Santosh",
    "QA": "Vidhyashree",
    "Production": "Karthik",
    "Maintenance": "Basavaraj",
    "Engineering": "Sujith",
    "Stores": "Prashanth",
}

CUSTOMERS = ["Bosch", "Continental", "Valeo", "Cummins", "Delphi"]
FUNCTIONS = ["SMT", "QA", "Production", "Maintenance", "Engineering", "Stores"]
STAGES    = ["Incoming", "SPI", "Placement", "Reflow", "AOI", "Testing", "Packing"]
NC_CATEGORIES   = ["Major", "Minor"]
SIXM_CATEGORIES = ["Man", "Machine", "Method", "Material", "Measurement", "Mother Nature"]
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

COLUMNS = [
    "NC Number", "Date", "Month", "Week", "Customer", "Model",
    "Function", "Stage", "Observation", "NC Category", "6M Category",
    "Sub Category", "Responsibility", "NC Raised By", "Action Plan",
]


# ─────────────────────────────────────────────
# Google Sheets helpers
# ─────────────────────────────────────────────

def _gsheets_configured() -> bool:
    try:
        creds = st.secrets.get("gcp_service_account", {})
        sheet = st.secrets.get("SHEET_ID", "")
        return bool(creds and sheet)
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    """Return an authorised gspread client (cached across reruns)."""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def _get_worksheet():
    """Return the first worksheet of the configured Google Sheet."""
    client   = _get_gspread_client()
    sheet_id = st.secrets["SHEET_ID"]
    sh       = client.open_by_key(sheet_id)
    return sh.sheet1


# ── Read from Google Sheets ──────────────────
def load_from_gsheets() -> pd.DataFrame:
    try:
        ws      = _get_worksheet()
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=COLUMNS)
        df = pd.DataFrame(records)
        # Ensure all expected columns exist
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUMNS]
    except Exception as e:
        st.warning(f"⚠️ Could not read Google Sheets: {e}")
        return pd.DataFrame(columns=COLUMNS)


# ── Write one row to Google Sheets ───────────
def append_to_gsheets(row: dict) -> bool:
    try:
        ws = _get_worksheet()
        # Ensure header row exists
        existing = ws.get_all_values()
        if not existing:
            ws.append_row(COLUMNS, value_input_option="RAW")
        ws.append_row([str(row.get(c, "")) for c in COLUMNS],
                      value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"❌ Google Sheets save error: {e}")
        return False


# ── Load dataframe (Google Sheets → local fallback) ──
def load_dataframe() -> pd.DataFrame:
    if _gsheets_configured():
        return load_from_gsheets()
    # Local Excel fallback
    try:
        return pd.read_excel(FILE_NAME)
    except Exception:
        return pd.DataFrame(columns=COLUMNS)


# ── Save dataframe (Google Sheets → local fallback) ──
def save_row(row: dict, df_all: pd.DataFrame) -> bool:
    if _gsheets_configured():
        return append_to_gsheets(row)
    # Local Excel fallback
    try:
        df_all.to_excel(FILE_NAME, index=False, engine="openpyxl")
        return True
    except Exception as e:
        st.error(f"❌ Local save error: {e}")
        return False


# ── Duplicate NC Number check ────────────────
def nc_number_exists(nc_no: str, df: pd.DataFrame) -> bool:
    if df.empty or "NC Number" not in df.columns:
        return False
    return nc_no.strip() in df["NC Number"].astype(str).str.strip().tolist()


# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="NC Registration System",
    page_icon="📋",
    layout="centered",
)

# ─────────────────────────────────────────────
# Custom CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', Arial, sans-serif; }

.nc-header {
    background: linear-gradient(135deg, #1a5276 0%, #2e86c1 100%);
    color: white;
    padding: 18px 24px;
    border-radius: 10px;
    margin-bottom: 24px;
    text-align: center;
}
.nc-header h1 { margin: 0; font-size: 1.7rem; letter-spacing: 1px; }
.nc-header p  { margin: 4px 0 0; font-size: 0.9rem; opacity: 0.85; }

.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1a5276;
    border-bottom: 2px solid #2e86c1;
    padding-bottom: 4px;
    margin: 18px 0 10px;
}

.storage-badge {
    display: inline-block;
    padding: 3px 12px;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-bottom: 10px;
}
.badge-sheets { background:#d5f5e3; color:#1e8449; border:1px solid #a9dfbf; }
.badge-local  { background:#fef9e7; color:#9a7d0a; border:1px solid #f9e79f; }

.stDataFrame { border-radius: 8px; overflow: hidden; }
.stAlert     { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────
defaults = {
    "success_msg":    "",
    "error_msgs":     [],
    "confirm_reset":  False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="nc-header">
  <h1>📋 NC Registration System</h1>
  <p>Non-Conformance Registration &amp; Tracking</p>
</div>
""", unsafe_allow_html=True)

# Storage indicator
if _gsheets_configured():
    st.markdown('<span class="storage-badge badge-sheets">🟢 Google Sheets – Permanent Storage Active</span>',
                unsafe_allow_html=True)
else:
    st.markdown('<span class="storage-badge badge-local">🟡 Local Excel Storage (configure Google Sheets for cloud deployment)</span>',
                unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Success / error banners
# ─────────────────────────────────────────────
if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = ""

for err in st.session_state.error_msgs:
    st.error(err)
st.session_state.error_msgs = []

# ─────────────────────────────────────────────
# Load data
# ─────────────────────────────────────────────
df_existing = load_dataframe()

# ─────────────────────────────────────────────
# Default values for editable fields
# ─────────────────────────────────────────────
now = datetime.datetime.now()


# ─────────────────────────────────────────────
# SECTION 1 – Registration Form
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📝 Registration Details</div>', unsafe_allow_html=True)

with st.form(key="nc_form", clear_on_submit=True):

    # ── Row 1: NC Number + Date ──────────────
    r1a, r1b = st.columns(2)
    with r1a:
        nc_number = st.text_input(
            "NC Number ✱",
            value=f"NC-{now.year}-001",
            placeholder="e.g. NC-2026-001",
            help="Enter manually. Format: NC-YYYY-NNN",
        )
    with r1b:
        date_picked = st.date_input(
            "Date ✱",
            value=now.date(),
            help="Select the registration date",
        )

    # ── Row 2: Month + Week ─────────────────
    r2a, r2b = st.columns(2)
    with r2a:
        month = st.selectbox(
            "Month ✱",
            options=MONTHS,
            index=now.month - 1,
            help="Select the month",
        )
    with r2b:
        week = st.text_input(
            "Week ✱",
            value=str(now.isocalendar()[1]),
            placeholder="e.g. 47",
            help="Enter ISO week number",
        )

    st.markdown("---")

    # ── Row 3: Customer + Model ──────────────
    r3a, r3b = st.columns(2)
    with r3a:
        customer = st.selectbox(
            "Customer ✱",
            options=["-- Select --"] + CUSTOMERS,
            help="Required",
        )
    with r3b:
        model = st.text_input("Model", placeholder="Enter model number / name")

    # ── Row 4: Function + Stage ──────────────
    r4a, r4b = st.columns(2)
    with r4a:
        function = st.selectbox(
            "Function ✱",
            options=["-- Select --"] + FUNCTIONS,
        )
    with r4b:
        stage = st.selectbox(
            "Stage ✱",
            options=["-- Select --"] + STAGES,
        )

    # ── Row 5: NC Category + 6M Category ────
    r5a, r5b = st.columns(2)
    with r5a:
        nc_category = st.selectbox(
            "NC Category ✱",
            options=["-- Select --"] + NC_CATEGORIES,
        )
    with r5b:
        sixm_category = st.selectbox(
            "6M Category",
            options=["-- Select --"] + SIXM_CATEGORIES,
        )

    # ── Row 6: Sub Category + NC Raised By ──
    r6a, r6b = st.columns(2)
    with r6a:
        sub_category = st.text_input("Sub Category", placeholder="Enter sub-category")
    with r6b:
        raised_by = st.text_input(
            "NC Raised By ✱",
            placeholder="Name of person raising NC",
        )

    # ── Responsibility (manual) ──────────────
    suggested = RESPONSIBILITY_MAP.get(
        function if function != "-- Select --" else "", ""
    )
    responsibility = st.text_input(
        "Responsibility ✱",
        value=suggested,
        placeholder="Enter responsible person's name",
        help="Auto-suggested from Function – you can edit it freely.",
    )

    # ── Observation ──────────────────────────
    st.markdown("**Observation ✱**")
    observation = st.text_area(
        "Observation",
        placeholder="Describe the non-conformance observed...",
        height=120,
        label_visibility="collapsed",
    )

    # ── Action Plan ──────────────────────────
    st.markdown("**Action Plan**")
    action_plan = st.text_area(
        "Action Plan",
        placeholder="Describe the corrective / preventive action planned...",
        height=120,
        label_visibility="collapsed",
    )

    st.markdown("---")
    submitted = st.form_submit_button("💾  Save Registration", use_container_width=True)


# ─────────────────────────────────────────────
# Form submission handler
# ─────────────────────────────────────────────
if submitted:
    errors = []

    # Required field validation
    if not nc_number.strip():
        errors.append("**NC Number** cannot be empty.")
    if not week.strip():
        errors.append("**Week** cannot be empty.")
    if customer == "-- Select --":
        errors.append("Please select a **Customer**.")
    if function == "-- Select --":
        errors.append("Please select a **Function**.")
    if stage == "-- Select --":
        errors.append("Please select a **Stage**.")
    if nc_category == "-- Select --":
        errors.append("Please select an **NC Category**.")
    if not responsibility.strip():
        errors.append("**Responsibility** cannot be empty.")
    if not observation.strip():
        errors.append("**Observation** cannot be empty.")
    if not raised_by.strip():
        errors.append("**NC Raised By** cannot be empty.")

    # Duplicate NC Number check
    if nc_number.strip() and nc_number_exists(nc_number.strip(), df_existing):
        errors.append(
            f"**NC Number `{nc_number.strip()}` already exists.** "
            "Please use a different NC Number to prevent duplicates."
        )

    if errors:
        st.session_state.error_msgs = errors
        st.rerun()
    else:
        formatted_date = date_picked.strftime("%d-%m-%Y")
        new_row = {
            "NC Number":    nc_number.strip(),
            "Date":         formatted_date,
            "Month":        month,
            "Week":         week.strip(),
            "Customer":     customer,
            "Model":        model.strip(),
            "Function":     function,
            "Stage":        stage,
            "Observation":  observation.strip(),
            "NC Category":  nc_category,
            "6M Category":  sixm_category if sixm_category != "-- Select --" else "",
            "Sub Category": sub_category.strip(),
            "Responsibility": responsibility.strip(),
            "NC Raised By": raised_by.strip(),
            "Action Plan":  action_plan.strip(),
        }

        df_new   = pd.DataFrame([new_row])
        df_final = pd.concat([df_existing, df_new], ignore_index=True)

        ok = save_row(new_row, df_final)

        if ok:
            st.session_state.success_msg = (
                f"✅ **{nc_number.strip()}** saved successfully!  "
                f"(Customer: {customer} | Function: {function} | "
                f"Responsibility: {responsibility.strip()})"
            )
            # Clear cache so records table refreshes
            _get_gspread_client.clear() if _gsheets_configured() else None
            st.rerun()


# ─────────────────────────────────────────────
# SECTION 2 – Registered Records
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Registered Records</div>', unsafe_allow_html=True)

df_display = load_dataframe()

if df_display.empty:
    st.info("No records saved yet. Fill in the form above and click **Save Registration**.")
else:
    # Summary metrics
    total      = len(df_display)
    customers_ = df_display["Customer"].nunique() if "Customer" in df_display.columns else 0
    functions_ = df_display["Function"].nunique() if "Function" in df_display.columns else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Total NCs", total)
    m2.metric("Unique Customers", customers_)
    m3.metric("Functions Used", functions_)

    # Records table
    st.dataframe(df_display, use_container_width=True, hide_index=True)

    # ── Download Excel ───────────────────────
    buf = io.BytesIO()
    df_display.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        label="⬇️  Download All Records (Excel)",
        data=buf.getvalue(),
        file_name=FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ─────────────────────────────────────────────
# SECTION 3 – Reset Form
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">🔄 Reset Form</div>', unsafe_allow_html=True)

if not st.session_state.confirm_reset:
    if st.button("🗑️  Reset / Clear Form", use_container_width=True):
        st.session_state.confirm_reset = True
        st.rerun()
else:
    st.warning(
        "⚠️ **Are you sure?** This will clear all form fields. "
        "Saved records in Google Sheets / Excel are NOT affected."
    )
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("✅  Yes, Reset Form", use_container_width=True):
            st.session_state.confirm_reset = False
            st.session_state.success_msg   = "Form has been reset."
            st.rerun()
    with col_no:
        if st.button("❌  Cancel", use_container_width=True):
            st.session_state.confirm_reset = False
            st.rerun()

# ─────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────
st.markdown("---")
storage_label = "Google Sheets" if _gsheets_configured() else "Local Excel"
st.caption(f"NC Registration System • Built with Streamlit • Storage: {storage_label}")
