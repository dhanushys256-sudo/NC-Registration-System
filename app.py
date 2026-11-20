"""
NC Registration System – Streamlit Web App
Converted from the original Tkinter desktop application.
Persistent storage: Excel file committed to the same GitHub repo via GitHub API,
so all records survive app restarts on Streamlit Community Cloud.
"""

import io
import os
import base64
import datetime

import pandas as pd
import requests
import streamlit as st

# ─────────────────────────────────────────────
# Constants (exact values from original app)
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

STAGES = ["Incoming", "SPI", "Placement", "Reflow", "AOI", "Testing", "Packing"]

NC_CATEGORIES = ["Major", "Minor"]

SIXM_CATEGORIES = ["Man", "Machine", "Method", "Material", "Measurement", "Mother Nature"]

EXCEL_COLUMNS = [
    "NC Number", "Date", "Month", "Week", "Customer", "Model",
    "Function", "Stage", "Observation", "NC Category", "6M Category",
    "Sub Category", "Responsibility", "NC Raised By", "Action Plan",
]

YEAR = 2026


# ─────────────────────────────────────────────
# GitHub Persistent Storage helpers
# ─────────────────────────────────────────────

def _gh_headers():
    token = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", ""))
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}


def _gh_api_url():
    repo  = st.secrets.get("GITHUB_REPO", os.getenv("GITHUB_REPO", ""))
    path  = st.secrets.get("GITHUB_FILE_PATH", os.getenv("GITHUB_FILE_PATH", FILE_NAME))
    return f"https://api.github.com/repos/{repo}/contents/{path}"


def _use_github() -> bool:
    """Return True only when both token and repo are configured."""
    token = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", ""))
    repo  = st.secrets.get("GITHUB_REPO",  os.getenv("GITHUB_REPO",  ""))
    return bool(token and repo)


# ── Read ──────────────────────────────────────
def load_dataframe() -> pd.DataFrame:
    empty = pd.DataFrame(columns=EXCEL_COLUMNS)

    if _use_github():
        try:
            resp = requests.get(_gh_api_url(), headers=_gh_headers(), timeout=10)
            if resp.status_code == 200:
                content = base64.b64decode(resp.json()["content"])
                return pd.read_excel(io.BytesIO(content))
            # 404 means the file doesn't exist yet → return empty frame
            return empty
        except Exception:
            st.warning("⚠️ Could not reach GitHub. Showing local data only.")

    # Local fallback (works on your PC or Streamlit Cloud without GitHub secrets)
    if os.path.exists(FILE_NAME):
        return pd.read_excel(FILE_NAME)
    return empty


# ── Write ─────────────────────────────────────
def save_dataframe(df: pd.DataFrame):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    excel_bytes = buffer.getvalue()

    if _use_github():
        try:
            # Get current SHA (needed for update)
            sha = None
            resp = requests.get(_gh_api_url(), headers=_gh_headers(), timeout=10)
            if resp.status_code == 200:
                sha = resp.json()["sha"]

            payload = {
                "message": f"NC Registration update – {datetime.datetime.utcnow().isoformat()}Z",
                "content": base64.b64encode(excel_bytes).decode(),
            }
            if sha:
                payload["sha"] = sha

            put = requests.put(_gh_api_url(), headers=_gh_headers(),
                               json=payload, timeout=15)
            if put.status_code not in (200, 201):
                st.error(f"GitHub save failed: {put.status_code} – {put.json().get('message')}")
                return False
        except Exception as e:
            st.error(f"GitHub save error: {e}")
            return False

    # Always write locally (even when GitHub is configured, for speed)
    with open(FILE_NAME, "wb") as f:
        f.write(excel_bytes)

    return True


# ─────────────────────────────────────────────
# NC Number generation (duplicate-safe)
# ─────────────────────────────────────────────
def generate_nc_number(df: pd.DataFrame) -> str:
    if df.empty or "NC Number" not in df.columns:
        return f"NC-{YEAR}-001"

    existing = df["NC Number"].astype(str).tolist()
    count = 1
    while f"NC-{YEAR}-{count:03d}" in existing:
        count += 1
    return f"NC-{YEAR}-{count:03d}"


# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="NC Registration System",
    page_icon="📋",
    layout="centered",
)

# ─────────────────────────────────────────────
# Custom CSS – professional, clean, responsive
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] { font-family: 'Segoe UI', Arial, sans-serif; }

/* ── Header banner ── */
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

/* ── Auto-filled info boxes ── */
.info-box {
    background: #eaf4fb;
    border-left: 4px solid #2e86c1;
    border-radius: 6px;
    padding: 8px 14px;
    font-size: 0.95rem;
    color: #1a5276;
    font-weight: 600;
    margin-bottom: 2px;
}

/* ── Responsibility badge ── */
.resp-badge {
    display: inline-block;
    background: #d5f5e3;
    color: #1e8449;
    border: 1px solid #a9dfbf;
    border-radius: 20px;
    padding: 4px 14px;
    font-weight: 700;
    font-size: 0.95rem;
}

/* ── Section divider ── */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1a5276;
    border-bottom: 2px solid #2e86c1;
    padding-bottom: 4px;
    margin: 18px 0 10px;
}

/* ── Save button ── */
div[data-testid="stButton"] > button {
    background-color: #1e8449;
    color: white;
    font-size: 1rem;
    font-weight: 700;
    border-radius: 8px;
    padding: 10px 32px;
    border: none;
    width: 100%;
    transition: background 0.2s;
}
div[data-testid="stButton"] > button:hover {
    background-color: #145a32;
}

/* ── Records table ── */
.stDataFrame { border-radius: 8px; overflow: hidden; }

/* ── Success / Error alerts ── */
.stAlert { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Session-state initialisation
# ─────────────────────────────────────────────
if "form_submitted" not in st.session_state:
    st.session_state.form_submitted = False

if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""


# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="nc-header">
  <h1>📋 NC Registration System</h1>
  <p>Non-Conformance Registration &amp; Tracking</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Load existing data
# ─────────────────────────────────────────────
df_existing = load_dataframe()

# Auto-computed values
now = datetime.datetime.now()
auto_nc_no   = generate_nc_number(df_existing)
auto_date    = now.strftime("%d-%m-%Y")
auto_month   = now.strftime("%B")
auto_week    = str(now.isocalendar()[1])

# ─────────────────────────────────────────────
# Success banner (shown after save)
# ─────────────────────────────────────────────
if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = ""

# ─────────────────────────────────────────────
# ── SECTION 1: Auto-filled Information ───────
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">🔢 Auto-filled Information</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown(f'<div class="info-box">NC Number &nbsp;|&nbsp; {auto_nc_no}</div>', unsafe_allow_html=True)
    st.caption("Auto-generated • read-only")

    st.markdown(f'<div class="info-box">Month &nbsp;|&nbsp; {auto_month}</div>', unsafe_allow_html=True)
    st.caption("Auto-detected • read-only")

with col2:
    st.markdown(f'<div class="info-box">Date &nbsp;|&nbsp; {auto_date}</div>', unsafe_allow_html=True)
    st.caption("Today's date • read-only")

    st.markdown(f'<div class="info-box">Week &nbsp;|&nbsp; {auto_week}</div>', unsafe_allow_html=True)
    st.caption("ISO week number • read-only")

# ─────────────────────────────────────────────
# ── SECTION 2: Registration Form ─────────────
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📝 Registration Details</div>', unsafe_allow_html=True)

with st.form(key="nc_form", clear_on_submit=True):

    col_a, col_b = st.columns(2)

    with col_a:
        customer = st.selectbox(
            "Customer ✱",
            options=["-- Select --"] + CUSTOMERS,
            help="Required field",
        )

        function = st.selectbox(
            "Function ✱",
            options=["-- Select --"] + FUNCTIONS,
            help="Automatically assigns Responsibility",
        )

        stage = st.selectbox(
            "Stage ✱",
            options=["-- Select --"] + STAGES,
        )

        nc_category = st.selectbox(
            "NC Category ✱",
            options=["-- Select --"] + NC_CATEGORIES,
        )

        sixm_category = st.selectbox(
            "6M Category",
            options=["-- Select --"] + SIXM_CATEGORIES,
        )

    with col_b:
        model = st.text_input("Model", placeholder="Enter model number / name")

        sub_category = st.text_input(
            "Sub Category",
            placeholder="Enter sub-category",
        )

        raised_by = st.text_input(
            "NC Raised By ✱",
            placeholder="Enter name of person raising NC",
        )

        # Responsibility – auto from Function
        responsibility = RESPONSIBILITY_MAP.get(
            function if function != "-- Select --" else "", ""
        )
        st.markdown("**Responsibility** (auto-assigned)")
        if responsibility:
            st.markdown(f'<span class="resp-badge">👤 {responsibility}</span>',
                        unsafe_allow_html=True)
        else:
            st.caption("_Select a Function above to auto-assign_")

    st.markdown("**Observation ✱**")
    observation = st.text_area(
        "Observation",
        placeholder="Describe the non-conformance observed...",
        height=120,
        label_visibility="collapsed",
    )

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
    # ── Validation ──────────────────────────
    errors = []
    if customer == "-- Select --":
        errors.append("Please select a **Customer**.")
    if function == "-- Select --":
        errors.append("Please select a **Function**.")
    if stage == "-- Select --":
        errors.append("Please select a **Stage**.")
    if nc_category == "-- Select --":
        errors.append("Please select an **NC Category**.")
    if not observation.strip():
        errors.append("**Observation** cannot be empty.")
    if not raised_by.strip():
        errors.append("**NC Raised By** cannot be empty.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        # ── Build new row ────────────────────
        new_row = {
            "NC Number":    auto_nc_no,
            "Date":         auto_date,
            "Month":        auto_month,
            "Week":         auto_week,
            "Customer":     customer,
            "Model":        model.strip(),
            "Function":     function,
            "Stage":        stage,
            "Observation":  observation.strip(),
            "NC Category":  nc_category,
            "6M Category":  sixm_category if sixm_category != "-- Select --" else "",
            "Sub Category": sub_category.strip(),
            "Responsibility": responsibility,
            "NC Raised By": raised_by.strip(),
            "Action Plan":  action_plan.strip(),
        }

        df_new    = pd.DataFrame([new_row])
        df_final  = pd.concat([df_existing, df_new], ignore_index=True)

        ok = save_dataframe(df_final)

        if ok:
            st.session_state.success_msg = (
                f"✅  **{auto_nc_no}** saved successfully! "
                f"(Customer: {customer} | Function: {function})"
            )
            st.rerun()
        else:
            st.error("❌ Save failed. Please check your connection and try again.")

# ─────────────────────────────────────────────
# ── SECTION 3: Registered Records ────────────
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Registered Records</div>', unsafe_allow_html=True)

df_display = load_dataframe()

if df_display.empty:
    st.info("No records saved yet. Fill in the form above and click **Save Registration**.")
else:
    total      = len(df_display)
    customers_ = df_display["Customer"].nunique() if "Customer" in df_display.columns else 0
    functions_ = df_display["Function"].nunique() if "Function" in df_display.columns else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Total NCs", total)
    m2.metric("Customers", customers_)
    m3.metric("Functions", functions_)

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
    )

    # ── Excel download ───────────────────────
    buffer = io.BytesIO()
    df_display.to_excel(buffer, index=False, engine="openpyxl")
    st.download_button(
        label="⬇️  Download All Records (Excel)",
        data=buffer.getvalue(),
        file_name=FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ─────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────
st.markdown("---")
st.caption("NC Registration System • Built with Streamlit • Data stored in GitHub Repository")
