"""
NEN Accelerate — Portfolio Success Intelligence Dashboard (Frontend)
Part 2 of 2: Reads Knowledge Repository → Shows Portfolio Dashboard + Venture Cards
"""
import streamlit as st
import pandas as pd
import os, json, re, tempfile, shutil
from pathlib import Path
from anthropic import Anthropic

# ── load .env ──────────────────────────────────────────
def load_env_file():
    env_path = Path(__file__).parent / ".env"
    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, val = line.partition("=")
                    os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))
load_env_file()

ENV_API_KEY       = os.environ.get("ANTHROPIC_API_KEY", "")
ENV_CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "")
ENV_TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "")
ENV_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
ENV_PASSWORD      = os.environ.get("APP_PASSWORD", "nen2026")

SP_FOLDER        = "04. Advisors/2026/Portfolio Success Dashboard"
COMMON_FOLDER    = f"{SP_FOLDER}/Common Documents"
REPO_FOLDER      = f"{COMMON_FOLDER}/Knowledge Repository"
DASHBOARD_FILE   = "0. Journey_Accelerate_Portfolio Dashboard.xlsx"

SIGNALS_REPO_PATH  = f"{REPO_FOLDER}/signals_repository.json"
FEEDBACK_REPO_PATH = f"{REPO_FOLDER}/feedback_repository.json"
JOURNEY_REPO_PATH  = f"{REPO_FOLDER}/journey_repository.json"
PEOPLE_REPO_PATH   = f"{REPO_FOLDER}/people_repository.json"
CALL_INTEL_REPO_PATH = f"{REPO_FOLDER}/call_intelligence_repository.json"

st.set_page_config(
    page_title="Portfolio Success Intelligence",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS — mirrors HTML showcase exactly ───────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* ══ RESET & BASE ══════════════════════════════════════ */
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
  font-size: 14px !important;
  color: #1E293B;
}

/* ══ KILL STREAMLIT DEFAULT SPACING ════════════════════
   This is the main cause of "spread out" look.
   Streamlit adds gap/margin to every element wrapper.
   We override all of it here.
══════════════════════════════════════════════════════ */

/* Remove gap between all stacked elements */
div[data-testid="stVerticalBlock"] > div {
  gap: 0 !important;
}
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
  gap: 0 !important;
}

/* Tighten every block element's bottom margin */
div[data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; padding-bottom: 0 !important; }
div[data-testid="stMarkdownContainer"] p { margin-bottom: 4px !important; margin-top: 0 !important; }
div[data-testid="stMarkdownContainer"] h1 { margin-bottom: 2px !important; margin-top: 0 !important; }
div[data-testid="stMarkdownContainer"] h2 { margin-bottom: 4px !important; margin-top: 0 !important; }
div[data-testid="stMarkdownContainer"] h3 { margin-bottom: 4px !important; margin-top: 0 !important; }

/* Column gaps */
div[data-testid="stHorizontalBlock"] {
  gap: 12px !important;
  align-items: stretch !important;
}

/* Element containers — remove padding top/bottom */
div[class*="element-container"] {
  margin-bottom: 6px !important;
  margin-top: 0 !important;
}

/* Metric containers — tighten */
div[data-testid="metric-container"] {
  padding: 12px 10px !important;
  margin: 0 !important;
}
div[data-testid="stMetricValue"] { line-height: 1.1 !important; }
div[data-testid="stMetricLabel"] { line-height: 1.2 !important; margin-top: 3px !important; }

/* Tabs — no extra padding above */
div[data-testid="stTabs"] { margin-top: 0 !important; }
.stTabs [data-baseweb="tab-panel"] { padding-top: 12px !important; }

/* Buttons — no extra wrapper margin */
div[data-testid="stButton"] { margin: 0 !important; }
div[data-testid="stButton"] > button { width: auto !important; }

/* Selectbox / input — tighten */
div[data-testid="stSelectbox"] { margin-bottom: 6px !important; }
div[data-testid="stTextInput"]  { margin-bottom: 6px !important; }

/* Caption — tight */
div[data-testid="stCaptionContainer"] { margin: 0 0 4px 0 !important; padding: 0 !important; }

/* Alert boxes */
div[data-testid="stAlert"] { margin: 6px 0 !important; padding: 8px 12px !important; }

/* Expander — tighten summary */
details > summary { padding: 8px 12px !important; }
.stExpander { margin-bottom: 6px !important; }

/* Progress bar */
div[data-testid="stProgressBar"] { margin: 4px 0 !important; }

/* Divider */
div[data-testid="stDecoration"] { margin: 8px 0 !important; }

/* Dialog */
div[data-testid="stModal"] { padding: 20px !important; }

/* Remove Streamlit footer */
footer { display: none !important; }
#MainMenu { display: none !important; }
.viewerBadge_container__1QSob { display: none !important; }

/* ══ PAGE BACKGROUND & CONTAINER ═══════════════════════ */
.stApp { background: #F1F5F9 !important; }

/* Hide sidebar completely */
section[data-testid="stSidebar"],
section[data-testid="stSidebarNav"],
div[data-testid="collapsedControl"] { display: none !important; width: 0 !important; }

/* Hide Streamlit top header bar */
header[data-testid="stHeader"] { display: none !important; height: 0 !important; }

/* Remove top padding that Streamlit adds for the header */
.main > div:first-child { padding-top: 0 !important; }

/* Main content area — compact matching HTML showcase */
.main .block-container {
  padding: 16px 32px 24px 32px !important;
  max-width: 1200px !important;
  margin: 0 auto !important;
}

/* Full width when sidebar is gone */
.main { margin-left: 0 !important; }

/* Remove Streamlit's own top padding for the header */
.main > div { padding-top: 0 !important; }
section.main > div.block-container { padding-top: 12px !important; }

/* ══ PAGE COVER BANNER ══════════════════════════════════ */
.cover-banner {
  background: #4F46E5; border-radius: 12px;
  padding: 28px 32px; margin-bottom: 28px;
}
.cover-banner h1 { color: #fff !important; font-size: 1.5rem !important;
  font-weight: 600 !important; margin-bottom: 4px !important; }
.cover-banner p  { color: #C7D2FE !important; font-size: 0.82rem !important;
  line-height: 1.6 !important; margin: 0 !important; }
.cover-banner .meta { color: #A5B4FC; font-size: 0.78rem; margin-top: 8px; }

/* ══ TYPOGRAPHY ═════════════════════════════════════════ */
h1 { font-size: 1.15rem !important; font-weight: 600 !important;
  color: #1E293B !important; margin-bottom: 2px !important; line-height: 1.3 !important; }
h2 { font-size: 1rem !important; font-weight: 600 !important; color: #1E293B !important; }
h3 { font-size: 0.9rem !important; font-weight: 600 !important; color: #1E293B !important; }
p  { color: #475569 !important; line-height: 1.6 !important; font-size: 0.87rem !important; }
.stCaption, .stCaption p { color: #64748B !important; font-size: 0.78rem !important; }
label, .stSelectbox label, .stTextInput label {
  font-size: 0.78rem !important; font-weight: 600 !important;
  color: #64748B !important; text-transform: uppercase !important;
  letter-spacing: 0.05em !important;
}

/* ══ TABS ════════════════════════════════════════════════ */
.stTabs [data-baseweb="tab-list"] {
  gap: 3px !important;
  background: #ffffff !important;
  padding: 5px 6px !important;
  border-radius: 10px !important;
  border: 1px solid #E2E8F0 !important;
  margin-bottom: 20px !important;
  flex-wrap: wrap !important;
}
.stTabs [data-baseweb="tab"] {
  padding: 6px 12px !important;
  border-radius: 8px !important;
  font-size: 0.78rem !important;
  font-weight: 500 !important;
  color: #64748B !important;
  border: none !important;
  background: transparent !important;
  white-space: nowrap !important;
}
.stTabs [data-baseweb="tab"]:hover {
  background: #EEF2FF !important;
  color: #4F46E5 !important;
}
.stTabs [aria-selected="true"] {
  background: #4F46E5 !important;
  color: #ffffff !important;
  border-radius: 8px !important;
}
.stTabs [aria-selected="true"] *,
.stTabs [aria-selected="true"] p,
.stTabs [aria-selected="true"] span,
.stTabs [aria-selected="true"] div { color: #ffffff !important; }
/* Hide the default underline indicator */
.stTabs [data-baseweb="tab-highlight"] { display: none !important; }
.stTabs [data-baseweb="tab-border"]    { display: none !important; }

/* ══ METRICS ═════════════════════════════════════════════ */
div[data-testid="metric-container"] {
  background: #EEF2FF !important;
  border-radius: 10px !important;
  padding: 14px 12px !important;
  border: none !important;
  text-align: center !important;
  box-shadow: none !important;
}
div[data-testid="stMetricValue"] {
  font-size: 1.6rem !important; font-weight: 700 !important;
  color: #4F46E5 !important; justify-content: center !important;
}
div[data-testid="stMetricLabel"] {
  font-size: 0.75rem !important; color: #64748B !important;
  justify-content: center !important; text-align: center !important;
}
div[data-testid="stMetricDelta"] { justify-content: center !important; }
/* Metric colour variants (applied via wrapper divs) */
.metric-green div[data-testid="metric-container"] { background: #DCFCE7 !important; }
.metric-green div[data-testid="stMetricValue"]    { color: #166534 !important; }
.metric-amber div[data-testid="metric-container"] { background: #FEF9C3 !important; }
.metric-amber div[data-testid="stMetricValue"]    { color: #D97706 !important; }
.metric-red   div[data-testid="metric-container"] { background: #FEE2E2 !important; }
.metric-red   div[data-testid="stMetricValue"]    { color: #DC2626 !important; }

/* ══ BUTTONS ═════════════════════════════════════════════ */
.stButton > button {
  background: #4F46E5 !important; color: #ffffff !important;
  border: none !important; border-radius: 8px !important;
  font-weight: 500 !important; font-size: 0.83rem !important;
  padding: 7px 18px !important; transition: background .15s !important;
  letter-spacing: 0.01em !important;
}
.stButton > button:hover  { background: #4338CA !important; color: #ffffff !important; }
.stButton > button:active { background: #3730A3 !important; color: #ffffff !important; }
.stButton > button p,
.stButton > button span,
.stButton > button div   { color: #ffffff !important; }

/* ══ INPUTS ══════════════════════════════════════════════ */
div[data-baseweb="select"] > div {
  border-radius: 8px !important; border-color: #E2E8F0 !important;
  font-size: 0.82rem !important; background: #fff !important;
  min-height: 36px !important;
}
div[data-baseweb="input"] > div input {
  border-radius: 8px !important; border-color: #E2E8F0 !important;
  font-size: 0.82rem !important; background: #fff !important;
  padding: 6px 12px !important;
}
div[data-baseweb="select"]:focus-within > div,
div[data-baseweb="input"]:focus-within > div {
  border-color: #4F46E5 !important;
  box-shadow: 0 0 0 2px #EEF2FF !important;
}

/* ══ EXPANDERS ═══════════════════════════════════════════ */
.stExpander {
  background: #ffffff !important;
  border: 1px solid #E2E8F0 !important;
  border-radius: 10px !important;
  margin-bottom: 8px !important;
  box-shadow: none !important;
}
.stExpander summary {
  font-size: 0.85rem !important; font-weight: 500 !important;
  color: #1E293B !important; padding: 10px 14px !important;
}
.stExpander [data-testid="stExpanderToggleIcon"] { color: #4F46E5 !important; }

/* ══ DATAFRAME / TABLE ═══════════════════════════════════ */
.stDataFrame {
  border: 1px solid #E2E8F0 !important;
  border-radius: 10px !important; overflow: hidden !important;
}
.stDataFrame thead th {
  background: #F8FAFC !important; color: #64748B !important;
  font-size: 0.76rem !important; font-weight: 600 !important;
  text-transform: uppercase !important; letter-spacing: 0.04em !important;
  padding: 8px 12px !important; border-bottom: 1px solid #E2E8F0 !important;
}
.stDataFrame tbody td {
  font-size: 0.82rem !important; padding: 7px 12px !important;
  color: #334155 !important; border-bottom: 1px solid #F1F5F9 !important;
}
.stDataFrame tbody tr:hover td { background: #FAFBFC !important; }

/* ══ PROGRESS BAR ════════════════════════════════════════ */
.stProgress > div > div {
  background: #4F46E5 !important; border-radius: 4px !important;
}
.stProgress > div {
  background: #E2E8F0 !important; border-radius: 4px !important;
  height: 6px !important;
}

/* ══ DIVIDER ═════════════════════════════════════════════ */
hr { border-color: #E2E8F0 !important; margin: 14px 0 !important; }

/* ══ ALERTS ══════════════════════════════════════════════ */
div[data-testid="stAlert"] {
  border-radius: 10px !important; font-size: 0.83rem !important;
  border: 1px solid #E2E8F0 !important;
}

/* ══ SPINNER ═════════════════════════════════════════════ */
.stSpinner > div { border-color: #4F46E5 !important; }

/* ══ DIALOG ══════════════════════════════════════════════ */
div[data-testid="stModal"] > div {
  border-radius: 12px !important; border: 1px solid #E2E8F0 !important;
}

/* ══ COLUMN GAP ══════════════════════════════════════════ */
div[data-testid="column"] { gap: 14px !important; }

/* ══════════════════════════════════════════════════════════
   CUSTOM COMPONENT CLASSES — mirrors HTML showcase exactly
   ══════════════════════════════════════════════════════════ */

/* RAG badges */
.rag-green { background:#DCFCE7; color:#166534; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.78rem; display:inline-block; white-space:nowrap; }
.rag-amber { background:#FEF9C3; color:#854D0E; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.78rem; display:inline-block; white-space:nowrap; }
.rag-red   { background:#FEE2E2; color:#991B1B; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.78rem; display:inline-block; white-space:nowrap; }
.rag-zero  { background:#F1F5F9; color:#64748B; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.78rem; display:inline-block; white-space:nowrap; }

/* Signal badges */
.sig-pos { background:#DCFCE7; color:#166534; padding:2px 9px; border-radius:10px; font-size:0.73rem; font-weight:600; display:inline-block; margin:2px; }
.sig-neu { background:#FEF9C3; color:#854D0E; padding:2px 9px; border-radius:10px; font-size:0.73rem; font-weight:600; display:inline-block; margin:2px; }
.sig-neg { background:#FEE2E2; color:#991B1B; padding:2px 9px; border-radius:10px; font-size:0.73rem; font-weight:600; display:inline-block; margin:2px; }
.sig-src { background:#EDE9FE; color:#5B21B6; padding:2px 8px; border-radius:8px; font-size:0.7rem; display:inline-block; margin:2px; }

/* Info card — matches HTML .card */
.info-card {
  background: #fff; border: 1px solid #E2E8F0; border-radius: 10px;
  padding: 16px 18px; margin-bottom: 14px;
}
.card-title { font-size: 0.83rem; font-weight: 600; color: #1E293B; margin-bottom: 10px; }

/* Section label */
.section-label {
  font-size: 0.68rem; font-weight: 700; color: #94A3B8;
  text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 5px;
}

/* Section heading with left border — matches HTML .section-head */
.section-head {
  font-size: 0.85rem; font-weight: 600; color: #1E293B;
  margin-bottom: 10px; margin-top: 16px;
  display: flex; align-items: center; gap: 8px;
  border-left: 3px solid #4F46E5; padding-left: 10px;
}

/* Venture card header */
.venture-header {
  background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
  color: #fff; border-radius: 10px; padding: 14px 18px; margin-bottom: 14px;
}
.venture-header h2 { color: #fff !important; font-size: 1rem !important; }
.venture-header p  { color: #C7D2FE !important; font-size: 0.78rem !important; margin-top: 3px !important; }

/* Session cards */
.session-card {
  background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px;
  padding: 14px 16px; margin-bottom: 10px;
}

/* Signal rows */
.signal-row { padding: 8px 0; border-bottom: 1px solid #F1F5F9; }
.signal-row:last-child { border-bottom: none; }
.signal-evidence {
  font-size: 0.78rem; color: #64748B; font-style: italic;
  margin-top: 4px; padding-left: 4px; line-height: 1.5;
}
.signal-source { font-size: 0.7rem; color: #94A3B8; margin-left: 6px; }

/* Bar chart rows — matches HTML */
.bar-row  { display:flex; align-items:center; gap:10px; margin-bottom:8px; }
.bar-label{ font-size:0.82rem; color:#334155; min-width:160px; }
.bar-wrap { flex:1; background:#E2E8F0; border-radius:4px; height:6px; overflow:hidden; }
.bar-fill { height:100%; border-radius:4px; background:#4F46E5; }
.bar-count{ font-size:0.76rem; color:#64748B; min-width:80px; text-align:right; }

/* Summary box */
.summary-box {
  background: #F8FAFC; border-left: 3px solid #4F46E5;
  border-radius: 0 8px 8px 0; padding: 12px 14px;
  font-size: 0.82rem; color: #334155; line-height: 1.7; margin-bottom: 12px;
}

/* Direction list */
.dir-item { display:flex; gap:10px; align-items:flex-start; margin-bottom:6px; }
.dir-num  {
  background:#4F46E5; color:#fff; border-radius:50%;
  width:20px; height:20px; display:flex; align-items:center;
  justify-content:center; font-size:10px; font-weight:600;
  flex-shrink:0; margin-top:2px;
}
.dir-text { font-size:0.83rem; color:#334155; line-height:1.5; }

/* Problem badge */
.prob-item { display:flex; align-items:flex-start; gap:10px; padding:8px 0; border-bottom:1px solid #F8FAFC; }
.prob-item:last-child { border-bottom:none; }
.prob-count { background:#DCFCE7; color:#166534; padding:3px 10px; border-radius:20px; font-size:0.73rem; font-weight:600; white-space:nowrap; flex-shrink:0; }
.prob-text strong { font-size:0.83rem; color:#1E293B; display:block; margin-bottom:2px; }
.prob-text span   { font-size:0.75rem; color:#64748B; }

/* NPS display */
.nps-big   { font-size:3rem; font-weight:700; color:#166534; line-height:1; text-align:center; }
.nps-label { font-size:0.76rem; color:#64748B; text-align:center; margin-top:4px; }

/* Hub badge */
.hub-badge {
  background:#EEF2FF; color:#3730A3;
  padding:2px 8px; border-radius:8px;
  font-size:0.73rem; font-weight:500;
}

/* Mentor card */
.mentor-card {
  background:#fff; border:1px solid #E2E8F0; border-radius:10px;
  padding:14px 16px; margin-bottom:10px;
}
.mentor-name { font-size:0.9rem; font-weight:600; color:#1E293B; }
.mentor-meta { font-size:0.73rem; color:#64748B; margin-top:2px; }

/* Table wrapper */
.tbl-wrap {
  overflow-x:auto; border:1px solid #E2E8F0;
  border-radius:10px; margin-bottom:14px;
}
.tbl-wrap table { width:100%; border-collapse:collapse; font-size:0.78rem; }
.tbl-wrap th {
  padding:8px 12px; text-align:left; background:#F8FAFC;
  color:#64748B; font-weight:600; border-bottom:1px solid #E2E8F0;
  white-space:nowrap; font-size:0.73rem; text-transform:uppercase; letter-spacing:0.04em;
}
.tbl-wrap td {
  padding:8px 12px; color:#334155;
  border-bottom:1px solid #F1F5F9; vertical-align:top;
}
.tbl-wrap tr:last-child td { border-bottom:none; }
.tbl-wrap tr:hover td { background:#FAFBFC; }

/* Score colours */
.score-green { font-weight:700; color:#166534; }
.score-amber { font-weight:700; color:#D97706; }
.score-red   { font-weight:700; color:#DC2626; }

</style>
""", unsafe_allow_html=True)

# ── password gate ──────────────────────────────────────
def check_password():
    if st.session_state.get("authenticated"): return True
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center;padding:40px 30px;background:#ffffff;
        border-radius:16px;border:1px solid #e2e8f0;box-shadow:0 4px 24px rgba(0,0,0,0.06)'>
            <div style='font-size:2.5rem;margin-bottom:8px'>🚀</div>
            <div style='font-size:1.3rem;font-weight:700;color:#1e293b;margin-bottom:4px'>
                Portfolio Success Intelligence</div>
            <div style='font-size:0.85rem;color:#64748b;margin-bottom:24px'>
                NEN Accelerate · Resources Network</div>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        pwd = st.text_input("Password", type="password",
                            placeholder="Enter access password...",
                            label_visibility="collapsed")
        _, cb, _ = st.columns([1,2,1])
        with cb:
            if st.button("🔐  Login", use_container_width=True):
                if pwd == ENV_PASSWORD:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
        st.caption("Contact: meenakshi.singh@wadhwanifoundation.org")
    return False

if not check_password(): st.stop()

# ── constants ──────────────────────────────────────────
RAG_ORDER  = {"Green": 0, "Amber": 1, "Red": 2, "ZERO": 3, "Unknown": 4}
RAG_EMOJI  = {"Green": "🟢", "Amber": "🟡", "Red": "🔴", "ZERO": "⚪", "Unknown": "⚪"}
RAG_COLOR  = {"Green": "rag-green", "Amber": "rag-amber",
              "Red": "rag-red", "ZERO": "rag-zero", "Unknown": "rag-zero"}

def rag_badge(rag):
    css   = RAG_COLOR.get(rag, "rag-zero")
    emoji = RAG_EMOJI.get(rag, "⚪")
    return f'<span class="{css}">{emoji} {rag}</span>'

def render_call_detail_fields(call, vname=None, journey_repo=None,
                               signals_repo=None):
    """Render the 7 structured call intelligence fields for one call record."""
    objective = call.get("call_objective", "Not Available")
    topics    = call.get("discussion_topics", [])
    updates   = call.get("company_updates", "Not Available")
    guidance  = call.get("venture_partner_guidance", "Not Available")
    decisions = call.get("key_decisions", [])
    actions   = call.get("action_items", [])
    risks     = call.get("risks_challenges", [])
    sprint_gaps_context = call.get("sprint_gaps_context", "")

    def _na(v): return "—" if not v or str(v) in ["Not Available","nan","None",""] else str(v)
    def _list(items, color="#334155"):
        if not items: return "<span style='color:#94a3b8'>—</span>"
        if isinstance(items, str): items = [items]
        return "".join(
            f"<li style='margin-bottom:3px;color:{color}'>{i}</li>"
            for i in items if i
        )

    cell_l = "padding:10px 13px;border-right:.5px solid #e2e8f0;border-bottom:.5px solid #e2e8f0"
    cell_r = "padding:10px 13px;border-bottom:.5px solid #e2e8f0"
    lbl    = "font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8;margin-bottom:4px"
    val    = "font-size:0.82rem;color:#334155"

    st.markdown(
        "<div style='background:#f8fafc;border:.5px solid #e2e8f0;"
        "border-radius:8px;overflow:hidden;margin-top:10px'>",
        unsafe_allow_html=True
    )
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        st.markdown(
            f"<div style='{cell_l}'><div style='{lbl}'>🎯 Call objective</div>"
            f"<div style='{val}'>{_na(objective)}</div></div>",
            unsafe_allow_html=True
        )
    with r1c2:
        st.markdown(
            f"<div style='{cell_r}'><div style='{lbl}'>💬 Major discussion points</div>"
            f"<ul style='margin:0;padding-left:16px;font-size:0.82rem'>{_list(topics)}</ul></div>",
            unsafe_allow_html=True
        )
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        st.markdown(
            f"<div style='{cell_l}'><div style='{lbl}'>🏢 Company updates</div>"
            f"<div style='{val}'>{_na(updates)}</div></div>",
            unsafe_allow_html=True
        )
    with r2c2:
        st.markdown(
            f"<div style='{cell_r}'><div style='{lbl}'>👤 Venture partner guidance</div>"
            f"<div style='{val}'>{_na(guidance)}</div></div>",
            unsafe_allow_html=True
        )
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        st.markdown(
            f"<div style='{cell_l}'><div style='{lbl}'>✅ Key decisions</div>"
            f"<ul style='margin:0;padding-left:16px;font-size:0.82rem'>{_list(decisions)}</ul></div>",
            unsafe_allow_html=True
        )
    with r3c2:
        st.markdown(
            f"<div style='{cell_r}'><div style='{lbl}'>📋 Action items</div>"
            f"<ul style='margin:0;padding-left:16px;font-size:0.82rem'>{_list(actions)}</ul></div>",
            unsafe_allow_html=True
        )
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Gap classification section ─────────────────────────────
    # Determine sprint status and journey data for this venture
    is_in_sprint = False
    journey_data = {}
    if vname and signals_repo:
        v = signals_repo.get("venture_summary", {}).get(vname, {})
        sp_status = v.get("overall_sprint_status", "")
        is_in_sprint = sp_status in ["Active", "Slow"]
    if vname and journey_repo:
        journey_data = journey_repo.get("ventures", {}).get(vname, {})

    render_gaps_with_classification(
        risks, sprint_gaps_context, journey_data, is_in_sprint
    )


def classify_gap(risk, sprint_gaps_context, journey_data, is_in_sprint):
    """
    Classify a gap/risk as:
      'In sprint'  — already being addressed in current sprint
      'In journey' — flagged in the Growth Journey Report
      'New gap'    — not in sprint or journey (newly identified)
      'NA'         — venture not in active sprint

    sprint_gaps_context: string from call record (e.g. "Sprint gaps: A; B | Objectives: C")
    journey_data: dict from journey_repo for this venture
    is_in_sprint: bool — True if venture has Active/Slow sprint status
    """
    if not is_in_sprint:
        return "NA"

    risk_lower = risk.lower().strip()
    risk_words = [w for w in risk_lower.split() if len(w) > 4]

    def has_overlap(text_lower):
        """Return True if 2+ significant words from risk appear in text."""
        matching = [w for w in risk_words if w in text_lower]
        return len(matching) >= 2 or (len(risk_words) == 1 and risk_lower[:25] in text_lower)

    # Check sprint gaps context
    if sprint_gaps_context and sprint_gaps_context not in ["Not Available", "—", ""]:
        sg_lower = sprint_gaps_context.lower()
        if has_overlap(sg_lower):
            return "In sprint"

    # Check journey document — multiple possible field names
    if journey_data:
        journey_text = " ".join([
            " ".join(journey_data.get("challenges", [])),
            " ".join(journey_data.get("gaps", [])),
            " ".join(journey_data.get("risks", [])),
            " ".join(journey_data.get("growth_areas", [])),
            journey_data.get("summary", ""),
            journey_data.get("growth_journey_summary", ""),
        ]).lower()
        if journey_text.strip() and has_overlap(journey_text):
            return "In journey"

    return "New gap"


def render_gap_badge(status):
    """Return HTML badge for a gap status."""
    styles = {
        "In sprint":  "background:#fef3c7;color:#92400e;border:1px solid #fde68a",
        "In journey": "background:#dbeafe;color:#1e40af;border:1px solid #bfdbfe",
        "New gap":    "background:#fee2e2;color:#991b1b;border:1px solid #fecaca",
        "NA":         "background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0",
    }
    icons = {"In sprint":"🏃","In journey":"📄","New gap":"🆕","NA":"—"}
    style = styles.get(status, styles["NA"])
    icon  = icons.get(status, "—")
    return (
        f"<span style='{style};padding:1px 8px;border-radius:20px;"
        f"font-size:10px;font-weight:600;white-space:nowrap'>"
        f"{icon} {status}</span>"
    )


def render_gaps_with_classification(risks, sprint_gaps_context,
                                     journey_data, is_in_sprint):
    """
    Render the risks/challenges section with gap classification badges.
    Shows: In sprint 🏃 / In journey 📄 / New gap 🆕 / NA —
    """
    if not risks:
        st.markdown(
            "<div style='padding:10px 13px;background:#fff5f5;border-radius:8px;"
            "font-size:0.82rem;color:#94a3b8'>⚠️ Risks / challenges — none identified</div>",
            unsafe_allow_html=True
        )
        return

    # Counts per category
    categorised = [(r, classify_gap(r, sprint_gaps_context, journey_data, is_in_sprint))
                   for r in risks]
    in_sprint_n  = sum(1 for _, s in categorised if s == "In sprint")
    in_journey_n = sum(1 for _, s in categorised if s == "In journey")
    new_gap_n    = sum(1 for _, s in categorised if s == "New gap")
    na_n         = sum(1 for _, s in categorised if s == "NA")

    # Summary badge row
    summary_parts = []
    if in_sprint_n:
        summary_parts.append(
            f"<span style='background:#fef3c7;color:#92400e;padding:2px 8px;"
            f"border-radius:12px;font-size:10px;font-weight:600'>"
            f"🏃 {in_sprint_n} in sprint</span>"
        )
    if in_journey_n:
        summary_parts.append(
            f"<span style='background:#dbeafe;color:#1e40af;padding:2px 8px;"
            f"border-radius:12px;font-size:10px;font-weight:600'>"
            f"📄 {in_journey_n} in journey</span>"
        )
    if new_gap_n:
        summary_parts.append(
            f"<span style='background:#fee2e2;color:#991b1b;padding:2px 8px;"
            f"border-radius:12px;font-size:10px;font-weight:600'>"
            f"🆕 {new_gap_n} new</span>"
        )
    if na_n and not (in_sprint_n or in_journey_n or new_gap_n):
        summary_parts.append(
            f"<span style='background:#f1f5f9;color:#64748b;padding:2px 8px;"
            f"border-radius:12px;font-size:10px;font-weight:600'>"
            f"— {na_n} (venture not in sprint)</span>"
        )

    # Header with summary badges
    st.markdown(
        f"<div style='padding:10px 13px 6px 13px;background:#fff5f5;"
        f"border-radius:8px 8px 0 0;border:.5px solid #fecaca'>"
        f"<div style='font-size:10px;text-transform:uppercase;letter-spacing:.06em;"
        f"color:#94a3b8;margin-bottom:6px'>⚠️ Risks / challenges ({len(risks)})</div>"
        f"<div style='display:flex;gap:6px;flex-wrap:wrap'>"
        f"{''.join(summary_parts)}</div></div>",
        unsafe_allow_html=True
    )

    # Per-risk rows
    for i, (risk, status) in enumerate(categorised):
        border_b = "border-radius:0 0 8px 8px" if i == len(risks) - 1 else ""
        bg = {
            "In sprint":  "#fffbeb",
            "In journey": "#eff6ff",
            "New gap":    "#fef2f2",
            "NA":         "#fafafa",
        }.get(status, "#fafafa")
        st.markdown(
            f"<div style='padding:7px 13px;background:{bg};"
            f"border:.5px solid #fecaca;border-top:none;{border_b};"
            f"display:flex;align-items:flex-start;gap:8px'>"
            f"{render_gap_badge(status)}"
            f"<span style='font-size:0.82rem;color:#334155;flex:1'>{risk}</span>"
            f"</div>",
            unsafe_allow_html=True
        )

    # Legend
    if is_in_sprint:
        st.markdown(
            "<div style='padding:4px 13px 6px;background:#fafafa;"
            "border:.5px solid #fecaca;border-top:none;border-radius:0 0 8px 8px;"
            "font-size:10px;color:#94a3b8'>"
            "🏃 In sprint — already being addressed &nbsp;|&nbsp; "
            "📄 In journey — flagged in Growth Journey Report &nbsp;|&nbsp; "
            "🆕 New gap — newly identified in this call"
            "</div>",
            unsafe_allow_html=True
        )


def safe_copy(src):
    dst = os.path.join(tempfile.gettempdir(), "nen_fe_" + os.path.basename(src))
    shutil.copy2(src, dst); return dst

def find_col(df, patterns):
    for c in df.columns:
        for p in patterns:
            if p.lower() in str(c).lower(): return c
    return None

def get_stage_bucket(pct):
    try:
        if pct is None or str(pct).strip() in ["","nan","None","-"]: return "Unknown"
        p = float(str(pct).replace("%","").replace(",","").strip())
        if p <= 1.0: p *= 100
        if p >= 100: return "100%"
        if p >= 76:  return "76-99%"
        if p >= 51:  return "51-75%"
        if p >= 26:  return "26-50%"
        return "0-25%"
    except: return "Unknown"

# ── Source file reference chip ────────────────────────────────────────────
# Instead of downloading files, show a reference chip telling the user
# which file to look at in SharePoint. No download, no blob, no SP API call.

CALL_TYPE_FILE_HINT = {
    "Sprint call":          "Transcript file in venture folder",
    "Panel call":           "Panel call transcript / notes in venture folder",
    "VP call":              "VP call transcript / notes in venture folder",
    "Mid-sprint VP review": "Mid-sprint VP report in venture folder",
    "WhatsApp / Email":     "WhatsApp export or email thread in venture folder",
    "Other":                "Session document in venture folder",
}

SP_FOLDER_DISPLAY = "SharePoint › 04. Advisors › 2026 › Portfolio Success Dashboard"

def source_file_ref(source_document=None, call_type=None,
                    vname=None, extra_note=None,
                    sp_path=None, folder=None):
    """
    Render a compact file-reference chip showing the user exactly which
    file to look at in SharePoint. No download — just a clear reference.

    source_document: filename (e.g. "Agro_Transcript_Mar.docx")
    sp_path:         full SharePoint path stored in call record by backend
    folder:          folder path stored in call record by backend
    call_type:       used to derive a descriptive hint if filename unavailable
    vname:           venture name — used as fallback folder context
    extra_note:      optional extra line below the chip
    """
    icon  = "📄"
    fname = source_document if (
        source_document and source_document not in
        ["Not Available", "—", "", "No session documents found"]
    ) else None

    hint = CALL_TYPE_FILE_HINT.get(call_type, "Session document in venture folder")

    # Prefer the SP path stored by backend; fall back to display hint
    if sp_path and sp_path not in ["Not Available", "—", ""]:
        folder_display = sp_path          # full path e.g. "04. Advisors/2026/.../vname/file.docx"
    elif folder and folder not in ["Not Available", "—", ""]:
        folder_display = folder           # folder without filename
    else:
        folder_display = (
            f"{SP_FOLDER_DISPLAY} › {vname}" if vname else SP_FOLDER_DISPLAY
        )

    label   = fname if fname else hint

    extra_html = (
        f"<div style='font-size:10px;color:#94a3b8;margin-top:2px'>{extra_note}</div>"
        if extra_note else ""
    )

    return (
        f"<div style='display:inline-flex;flex-direction:column;"
        f"margin-top:6px;max-width:100%'>"
        f"<div style='display:inline-flex;align-items:center;gap:5px;"
        f"background:#f1f5f9;border:.5px solid #e2e8f0;border-radius:6px;"
        f"padding:4px 10px;font-size:11px;color:#475569' "
        f"title='SharePoint: {folder_display}'>"
        f"<span>{icon}</span>"
        f"<span style='font-weight:500'>{label}</span>"
        f"</div>"
        f"<div style='font-size:10px;color:#94a3b8;margin-top:2px;"
        f"padding-left:2px'>📁 {folder_display}</div>"
        f"{extra_html}"
        f"</div>"
    )

def source_file_ref_for_call(call, vname=None):
    """
    Convenience wrapper — extract all file reference fields from a
    call intelligence record and render the file reference chip.

    Reads source_document, source_sp_path, source_folder from the
    call record — these are stored by the backend during Step 5.
    """
    return source_file_ref(
        source_document = call.get("source_document"),
        call_type       = call.get("call_type"),
        vname           = vname or call.get("venture_name"),
        sp_path         = call.get("source_sp_path"),
        folder          = call.get("source_folder"),
    )

# Keep sp_download_link and sp_call_document_link as no-ops for
# any remaining references that haven't been updated yet
def sp_download_link(file_path, label="Document", sp=None, style=""):
    return source_file_ref(source_document=file_path.split("/")[-1] if file_path else None)

def sp_call_document_link(vname, call_type, call_intel_repo=None, sp=None, label=None):
    return source_file_ref(call_type=call_type, vname=vname)


    try:
        if pct is None or str(pct).strip() in ["","nan","None","-"]: return "Unknown"
        p = float(str(pct).replace("%","").replace(",","").strip())
        if p <= 1.0: p *= 100
        if p >= 100: return "100%"
        if p >= 76:  return "76–99%"
        if p >= 51:  return "51–75%"
        if p >= 26:  return "26–50%"
        return "0–25%"
    except: return "Unknown"

# ── settings controls — compact single row ─────────────
tb1, tb2, tb3 = st.columns([1.5, 1.5, 5])
with tb1:
    use_sp = st.toggle("☁️ SharePoint", value=bool(ENV_CLIENT_ID), key="use_sp_top")
with tb2:
    api_key = ENV_API_KEY or st.text_input(
        "API Key", type="password", key="api_key_top",
        label_visibility="collapsed", placeholder="Anthropic API key...")
with tb3:
    pass  # spacer

# ── SP connection ──────────────────────────────────────
sp_reader = None
if use_sp and ENV_CLIENT_ID:
    try:
        from sharepoint_reader import SharePointReader
        if "sp_reader_fe" not in st.session_state:
            with st.spinner("Connecting to SharePoint..."):
                st.session_state["sp_reader_fe"] = SharePointReader(
                    ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        sp_reader = st.session_state["sp_reader_fe"]
    except Exception as e:
        st.error(f"SharePoint connection failed: {e}"); st.stop()

client = Anthropic(api_key=api_key) if api_key else None
sp_id  = id(sp_reader) if sp_reader else 0

# ── load dashboard Excel ───────────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_dashboard(_sp_id, use_sp):
    try:
        if use_sp and ENV_CLIENT_ID:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            content = sp.download_file(f"{SP_FOLDER}/{DASHBOARD_FILE}")
            tmp = os.path.join(tempfile.gettempdir(), "nen_fe_dashboard.xlsx")
            with open(tmp,"wb") as f: f.write(content)
            fpath = tmp
        else:
            return None, "SharePoint not configured"
        company_df = pd.read_excel(fpath, sheet_name="Company", header=2)
        return company_df, None
    except Exception as e: return None, str(e)

with st.spinner("Loading portfolio data..."):
    company_df, err = load_dashboard(sp_id, use_sp)

if err: st.error(f"❌ {err}"); st.stop()
if company_df is None: st.error("Could not load data."); st.stop()

company_df.columns = [str(c).strip() for c in company_df.columns]
col_name   = find_col(company_df, ["company name"])
col_hub    = find_col(company_df, ["hub / state","hub","state"])
col_pct    = find_col(company_df, ["% sprint completion","sprint completion"])
col_notes  = find_col(company_df, ["notes / comments","notes","remarks"])
col_rev    = find_col(company_df, ["revenue ly"])
col_tgt    = find_col(company_df, ["3-year target","3 year target"])
col_vp     = find_col(company_df, ["venture partner","venture manager","vp"])
col_sprint = None
for c in company_df.columns:
    if "sprint type" in str(c).lower(): col_sprint = c; break
if not col_sprint:
    for c in company_df.columns:
        cl = str(c).lower()
        if "sprint" in cl and not any(x in cl for x in ["commit","complet","score","%","task"]):
            col_sprint = c; break
if col_pct is None and len(company_df.columns) > 37:
    col_pct = company_df.columns[37]

skip = ["nan","none","","company name","name","venture name","-","—"]
all_rows = company_df[company_df[col_name].notna()].copy() if col_name else company_df.copy()
all_rows = all_rows[~all_rows[col_name].astype(str).str.strip().str.lower().isin(skip)]
seen_set = set(); ventures_list = []
for v in all_rows[col_name].astype(str).str.strip().tolist():
    if v not in seen_set: seen_set.add(v); ventures_list.append(v)

def get_row(name):
    m = all_rows[all_rows[col_name].astype(str).str.strip() == name]
    return m.iloc[0] if not m.empty else None

def cv(row, col, default="—"):
    if row is None or col is None: return default
    v = str(row[col]).strip()
    if v in ["nan","None","NaT",""]: return default
    try:
        f = float(v)
        if f == int(f): return str(int(f))
        return v
    except: return v

# ── load Company Basics sheet ─────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_company_basics(_sp_id, use_sp):
    """Load Company Basics sheet from the dashboard Excel file."""
    try:
        if use_sp and ENV_CLIENT_ID:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            content = sp.download_file(f"{SP_FOLDER}/{DASHBOARD_FILE}")
            tmp = os.path.join(tempfile.gettempdir(), "nen_fe_dashboard.xlsx")
            with open(tmp,"wb") as f: f.write(content)
            fpath = tmp
        else:
            return None, "SharePoint not configured"

        df_raw = pd.read_excel(fpath, sheet_name="Company Basics", header=None)

        # Col mapping: row 2 = headers, row 3+ = data
        clean_cols = {
            1:"venture_name", 2:"size_revenue", 3:"existing_product",
            4:"existing_market_segments", 5:"existing_geographies",
            6:"incremental_rev_3yr", 7:"incremental_jobs_3yr",
            8:"new_product", 9:"new_market_segments", 10:"new_geographies",
            11:"stream_status_gtm", 12:"stream_status_product",
            13:"stream_status_operations", 14:"stream_status_supply_chain",
            15:"stream_status_hr", 16:"stream_status_finance",
            17:"stream_support_gtm", 18:"stream_support_product",
            19:"stream_support_operations", 20:"stream_support_supply_chain",
            21:"stream_support_hr", 22:"stream_support_finance",
            23:"goal_gtm", 24:"goal_product", 25:"goal_operations",
            26:"goal_supply_chain", 27:"goal_people", 28:"goal_finance",
            29:"month_year", 30:"startup_smb", 31:"sprint_status",
            32:"sprint_type", 33:"hub", 34:"program_category",
            35:"revenue_fy2425_26",
            37:"unlock_gtm", 38:"unlock_product", 39:"unlock_operations",
            40:"unlock_supply_chain", 41:"unlock_people", 42:"unlock_finance",
        }

        df = df_raw.iloc[3:].copy()
        df.columns = range(len(df.columns))
        df = df.rename(columns=clean_cols)
        available = [c for c in clean_cols.values() if c in df.columns]
        df = df[available]

        skip_vals = ["nan","None","NaT","","0","x","X"]
        df = df[df["venture_name"].notna()]
        df = df[~df["venture_name"].astype(str).str.strip().isin(skip_vals)]
        df = df.reset_index(drop=True)

        # Build lookup dict keyed by venture name
        basics = {}
        for _, row in df.iterrows():
            vn = str(row["venture_name"]).strip()
            if not vn or vn in skip_vals: continue
            basics[vn] = {c: (None if str(row.get(c,"")).strip() in skip_vals
                              else str(row.get(c,"")).strip())
                          for c in available}
        return basics, None
    except Exception as e:
        return None, str(e)

with st.spinner("Loading Company Basics..."):
    company_basics, cb_err = load_company_basics(sp_id, use_sp)

if cb_err:
    company_basics = {}

# ── load knowledge repositories from SharePoint ────────
@st.cache_data(show_spinner=False, ttl=300)
def load_signals_repo(_sp_id, use_sp):
    if not use_sp or not ENV_CLIENT_ID: return None, "SharePoint not configured"
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        content = sp.download_file(SIGNALS_REPO_PATH)
        data = json.loads(content.decode("utf-8"))
        return data, None
    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False, ttl=300)
def load_feedback_repo(_sp_id, use_sp):
    if not use_sp or not ENV_CLIENT_ID: return None, "SharePoint not configured"
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        content = sp.download_file(FEEDBACK_REPO_PATH)
        data = json.loads(content.decode("utf-8"))
        return data, None
    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False, ttl=300)
def load_journey_repo(_sp_id, use_sp):
    if not use_sp or not ENV_CLIENT_ID: return None, "SharePoint not configured"
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        content = sp.download_file(JOURNEY_REPO_PATH)
        return json.loads(content.decode("utf-8")), None
    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False, ttl=300)
def load_people_repo(_sp_id, use_sp):
    if not use_sp or not ENV_CLIENT_ID: return None, "SharePoint not configured"
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        content = sp.download_file(PEOPLE_REPO_PATH)
        return json.loads(content.decode("utf-8")), None
    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False, ttl=300)
def load_call_intel_repo(_sp_id, use_sp):
    if not use_sp or not ENV_CLIENT_ID: return None, "SharePoint not configured"
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        content = sp.download_file(CALL_INTEL_REPO_PATH)
        return json.loads(content.decode("utf-8")), None
    except Exception as e:
        return None, str(e)

with st.spinner("Loading Knowledge Repositories from SharePoint..."):
    signals_repo, sig_err  = load_signals_repo(sp_id, use_sp)
    feedback_repo, fb_err  = load_feedback_repo(sp_id, use_sp)
    journey_repo,  j_err   = load_journey_repo(sp_id, use_sp)
    people_repo,   p_err   = load_people_repo(sp_id, use_sp)
    call_intel_repo, ci_err = load_call_intel_repo(sp_id, use_sp)

# ── RAG formula (mirrors processor.py — single source of truth) ───
def _nps_from_signals(signals_list):
    """
    NPS from GREEN/AMBER/RED signals.
    GREEN=Promoters, AMBER=Passives, RED=Detractors
    NPS = %Green - %Red
    NPS>=20 → Green, 0-19 → Amber, <0 → Red, 0 signals → ZERO
    """
    if not signals_list:
        return "ZERO", 0, 0, 0, 0, 0
    green = sum(1 for s in signals_list if s.get("category","GREEN") == "GREEN")
    amber = sum(1 for s in signals_list if s.get("category","GREEN") == "AMBER")
    red   = sum(1 for s in signals_list if s.get("category","GREEN") == "RED")
    total = green + amber + red
    nps   = round(green/total*100) - round(red/total*100)
    if   nps >= 20: rag = "Green"
    elif nps >= 0:  rag = "Amber"
    else:           rag = "Red"
    return rag, nps, green, amber, red, total

def compute_rag_from_signals(signals):
    """Compute full RAG dict live from signals — NPS based. Single source of truth."""
    m_sigs = signals.get("momentum",  [])
    i_sigs = signals.get("investment", [])
    m_rag, m_nps, m_g, m_a, m_r, m_tot = _nps_from_signals(m_sigs)
    i_rag, i_nps, i_g, i_a, i_r, i_tot = _nps_from_signals(i_sigs)

    order   = {"Red": 0, "Amber": 1, "Green": 2, "ZERO": 3}
    present = [r for r in [m_rag, i_rag] if r != "ZERO"]
    overall = min(present, key=lambda x: order.get(x, 3)) if present else "ZERO"

    score_matrix = {
        ("Green","Green"):10, ("Green","Amber"):8, ("Green","Red"):5, ("Green","ZERO"):5,
        ("Amber","Green"):8,  ("Amber","Amber"):7, ("Amber","Red"):3, ("Amber","ZERO"):3,
        ("Red",  "Green"):5,  ("Red",  "Amber"):3, ("Red",  "Red"):1, ("Red",  "ZERO"):0,
        ("ZERO", "Green"):5,  ("ZERO", "Amber"):3, ("ZERO", "Red"):0, ("ZERO", "ZERO"):0,
    }
    score = score_matrix.get((m_rag, i_rag), 0)

    def _reason(rag, nps, g, a, r, tot, cat):
        if rag == "ZERO": return f"No {cat} signals found."
        return f"Signal NPS {nps:+d} — {g} Green, {a} Amber, {r} Red of {tot} → {rag}."

    # Portfolio Signal NPS = Signal NPS across ALL signals combined
    all_sigs  = m_sigs + i_sigs
    _, p_nps, p_g, p_a, p_r, p_tot = _nps_from_signals(all_sigs)

    # Overall note — formula-based, zero API cost
    all_g   = m_g + i_g
    all_a   = m_a + i_a
    all_r   = m_r + i_r
    all_tot = m_tot + i_tot
    overall_nps = (
        round(all_g/all_tot*100) - round(all_r/all_tot*100)
    ) if all_tot else 0

    if overall == "Green":
        overall_note = (
            f"Signal NPS {overall_nps:+d} from {all_tot} signals "
            f"({all_g} Green, {all_a} Amber, {all_r} Red). "
            f"Founder actively progressing on sprint with concrete actions and investment committed."
        )
    elif overall == "Amber":
        overall_note = (
            f"Signal NPS {overall_nps:+d} from {all_tot} signals "
            f"({all_g} Green, {all_a} Amber, {all_r} Red). "
            f"Some positive actions taken but sprint progress is partial or delayed. "
            f"Follow-up recommended to unblock momentum."
        )
    elif overall == "Red":
        overall_note = (
            f"Signal NPS {overall_nps:+d} from {all_tot} signals "
            f"({all_g} Green, {all_a} Amber, {all_r} Red). "
            f"Founder disengaged or not investing in sprint. "
            f"Immediate intervention recommended."
        )
    else:
        overall_note = (
            "No signals extracted from available documents. "
            "Prioritise document collection or venture re-engagement."
        )

    return {
        "overall_rag":       overall,
        "momentum_rag":      m_rag,
        "investment_rag":    i_rag,
        "momentum_reason":   _reason(m_rag, m_nps, m_g, m_a, m_r, m_tot, "momentum"),
        "investment_reason": _reason(i_rag, i_nps, i_g, i_a, i_r, i_tot, "investment"),
        "overall_note":      overall_note,
        "momentum_score":    score,
        "investment_score":  score,
        "momentum_nps":      m_nps,  "investment_nps":   i_nps,
        "momentum_green":    m_g,    "investment_green":  i_g,
        "momentum_amber":    m_a,    "investment_amber":  i_a,
        "momentum_red":      m_r,    "investment_red":    i_r,
        "momentum_total":    m_tot,  "investment_total":  i_tot,
        "portfolio_nps":     p_nps,
        "portfolio_green":   p_g,    "portfolio_amber":   p_a,
        "portfolio_red":     p_r,    "portfolio_total":   p_tot,
    }

# ── parse repositories into lookups ───────────────────
venture_rag      = {}  # {vname: rag dict computed live from signals}
venture_signals  = {}  # {vname: {momentum: [...], investment: [...]}}
venture_feedback = {}  # {vname: [session_dicts]}

if signals_repo:
    vsummary = signals_repo.get("venture_summary", {})
    for vn, vdata in vsummary.items():
        signals = vdata.get("signals", {"momentum":[], "investment":[]})
        venture_signals[vn] = signals
        # Always compute RAG live from signals — keeps RAG Score tab and Signals tab in sync
        live_rag = compute_rag_from_signals(signals)
        # Carry over non-RAG fields (hub, vp, sprint, sources, processed_at etc.)
        venture_rag[vn] = {**vdata, **live_rag}

if feedback_repo:
    for vn, vdata in feedback_repo.get("ventures", {}).items():
        venture_feedback[vn] = vdata.get("sessions", [])

mentor_insights = {}  # {mentor_name: {mentor_name, total_sessions, avg_rating, sessions:[...]}}
if feedback_repo:
    mentor_insights = feedback_repo.get("mentor_insights", {})

# ── Journey Document data lookup ──────────────────────
# {vname: {existing_product, new_product, goal_gtm, ...}}
journey_data = {}
if journey_repo:
    for vn, vdata in journey_repo.get("ventures", {}).items():
        journey_data[vn] = vdata
# Also check signals_repository venture_summary for merged journey_data
if signals_repo:
    for vn, vdata in signals_repo.get("venture_summary", {}).items():
        if vdata.get("journey_data") and vn not in journey_data:
            journey_data[vn] = vdata["journey_data"]

# ── Build unified session list per venture from ALL sources ──
# Source 1+2: Session Tracker + Feedback Quality Tracker (via mentor_insights)
# Source 3+4+5: Transcript/feedback files (via venture_feedback Claude extraction)
# Both tabs use this same master lookup — Mentor Insights just re-groups by mentor

venture_tracker_sessions = {}   # {vname: [tracker sessions normalised]}
all_sessions_by_venture  = {}   # {vname: [ALL sessions from all sources]}

# Step A — normalise tracker sessions per venture
for mn, mdata in mentor_insights.items():
    for s in mdata.get("sessions", []):
        vn = s.get("venture_name","")
        if not vn: continue
        ff  = s.get("founder_feedback") or {}
        mfb = s.get("mentor_feedback")  or {}
        normalised = {
            "source":             "Session Tracker",
            "mentor_name":        mn,
            "session_date":       s.get("meeting_date","Not Available"),
            "session_type":       s.get("session_type","Not Available"),
            "topics_discussed":   s.get("ask","Not Available"),
            "key_outputs":        s.get("next_steps","Not Available"),
            "session_summary":    s.get("meeting_meeting","Not Available") or s.get("meeting_summary","Not Available"),
            "founder_feedback":   ff.get("verbatim") or s.get("tracker_feedback","Not Available"),
            "founder_rating":     ff.get("overall_rating") or s.get("tracker_rating"),
            "founder_usefulness": ff.get("usefulness","Not Available"),
            "followup_required":  s.get("followup_required","Not Available"),
            "mentor_engagement":  mfb.get("mentee_engaged","Not Available"),
            "mentor_prepared":    mfb.get("mentee_prepared","Not Available"),
        }
        venture_tracker_sessions.setdefault(vn, []).append(normalised)
        all_sessions_by_venture.setdefault(vn, []).append(normalised)

# Step B — add Claude-extracted transcript sessions per venture
for vn, sess_list in venture_feedback.items():
    for s in sess_list:
        normalised = {
            "source":             "Transcript",
            "mentor_name":        s.get("mentor_name","Not Available"),
            "session_date":       s.get("session_date","Not Available"),
            "session_type":       "Not Available",
            "topics_discussed":   s.get("topics_discussed","Not Available"),
            "key_outputs":        s.get("key_outputs","Not Available"),
            "session_summary":    s.get("session_summary","Not Available"),
            "founder_feedback":   s.get("founder_feedback","Not Available"),
            "founder_rating":     None,
            "founder_usefulness": "Not Available",
            "followup_required":  "Not Available",
            "mentor_engagement":  "Not Available",
            "mentor_prepared":    "Not Available",
        }
        all_sessions_by_venture.setdefault(vn, []).append(normalised)

# Step C — sort all sessions per venture by date descending
for vn in all_sessions_by_venture:
    all_sessions_by_venture[vn].sort(
        key=lambda s: s.get("session_date",""), reverse=True)

# Step D — rebuild mentor_insights_unified grouping ALL sessions by mentor
# This ensures Mentor Insights tab shows tracker + transcript sessions together
mentor_insights_unified = {}   # {mentor_name: {sessions: [...], ventures: [...], ...}}

for vn, sess_list in all_sessions_by_venture.items():
    for s in sess_list:
        mn = s.get("mentor_name","Not Available")
        if mn in ["Not Available","—",""]: mn = "Unknown Mentor"
        if mn not in mentor_insights_unified:
            mentor_insights_unified[mn] = {
                "mentor_name":     mn,
                "total_sessions":  0,
                "ventures_worked": [],
                "ratings":         [],
                "sessions":        [],
            }
        # Attach venture name to session for Mentor Insights display
        s_with_venture = {**s, "venture_name": vn}
        mentor_insights_unified[mn]["sessions"].append(s_with_venture)
        mentor_insights_unified[mn]["total_sessions"] += 1
        if vn not in mentor_insights_unified[mn]["ventures_worked"]:
            mentor_insights_unified[mn]["ventures_worked"].append(vn)
        if s.get("founder_rating"):
            mentor_insights_unified[mn]["ratings"].append(s["founder_rating"])

# Compute avg rating and sort sessions by date
for mn, mdata in mentor_insights_unified.items():
    ratings = mdata.pop("ratings", [])
    mdata["avg_rating"] = round(sum(ratings)/len(ratings), 1) if ratings else None
    mdata["sessions"].sort(key=lambda s: s.get("session_date",""), reverse=True)

repo_loaded = signals_repo is not None or feedback_repo is not None

# ── compact header banner ───────────────────────────────
sig_status  = f"✅ Signals · {signals_repo.get('generated_at','—')[:10]}" if signals_repo else "⚠️ Signals missing"
fb_status   = "✅ Feedback loaded" if feedback_repo else "ℹ️ Feedback missing"
jrn_status  = "✅ Journey data" if journey_data else "⬜ No journey data"
sig_color   = "#86EFAC" if signals_repo else "#FDE68A"
fb_color    = "#86EFAC" if feedback_repo else "#C7D2FE"

st.markdown(
    f"<div style='background:#4F46E5;border-radius:10px;padding:14px 20px;"
    f"margin-bottom:16px;display:flex;align-items:center;"
    f"justify-content:space-between;flex-wrap:wrap;gap:10px'>"
    f"<div style='display:flex;align-items:center;gap:10px'>"
    f"<span style='font-size:1.4rem'>🚀</span>"
    f"<div>"
    f"<div style='color:#fff;font-size:0.95rem;font-weight:700;line-height:1.2'>"
    f"Portfolio Success Intelligence</div>"
    f"<div style='color:#C7D2FE;font-size:0.72rem;margin-top:1px'>"
    f"NEN Accelerate · Resources Network</div>"
    f"</div></div>"
    f"<div style='display:flex;gap:8px;flex-wrap:wrap;align-items:center'>"
    f"<span style='background:rgba(255,255,255,0.12);color:{sig_color};"
    f"padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:500'>"
    f"{sig_status}</span>"
    f"<span style='background:rgba(255,255,255,0.12);color:{fb_color};"
    f"padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:500'>"
    f"{fb_status}</span>"
    f"<span style='background:rgba(255,255,255,0.12);color:#C7D2FE;"
    f"padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:500'>"
    f"{jrn_status}</span>"
    f"</div></div>",
    unsafe_allow_html=True
)

if not repo_loaded:
    st.error("❌ No Knowledge Repository found on SharePoint. Run the Backend app to generate and upload the repository files first.")
    st.info(f"Expected paths:\n- `{SIGNALS_REPO_PATH}`\n- `{FEEDBACK_REPO_PATH}`")
    st.stop()

# ── main tabs ──────────────────────────────────────────
tab_definitions, tab_overview, tab_companies, tab_ventures, tab_intelligence = st.tabs([
    "📖  How It Works",
    "📊  Portfolio Overview",
    "🏢  Companies",
    "🏗  Venture Cards",
    "🔬  Intelligence",
])

# ══════════════════════════════════════════════════════
#  TAB 1: HOW IT WORKS — DEFINITIONS
# ══════════════════════════════════════════════════════
with tab_definitions:
    st.title("📖 How It Works")
    st.caption("Definitions, scoring rules and methodology for the Portfolio Success Intelligence Dashboard")
    st.divider()

    def def_card(title, content_html):
        return f"""<div style='background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;
        padding:24px 28px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,0.04)'>
        <div style='font-size:1.05rem;font-weight:700;color:#1e293b;margin-bottom:14px'>
        {title}</div>{content_html}</div>"""

    def rule_row(rag, emoji, definition):
        colors = {"Green":("#dcfce7","#166534"),"Amber":("#fef9c3","#854d0e"),
                  "Red":("#fee2e2","#991b1b"),"ZERO":("#f1f5f9","#64748b")}
        bg, fg = colors.get(rag, ("#f1f5f9","#475569"))
        return (f"<div style='display:flex;align-items:flex-start;gap:14px;"
                f"padding:10px 0;border-bottom:1px solid #f1f5f9'>"
                f"<span style='background:{bg};color:{fg};padding:3px 14px;"
                f"border-radius:20px;font-weight:700;font-size:0.83rem;"
                f"white-space:nowrap;min-width:70px;text-align:center'>"
                f"{emoji} {rag}</span>"
                f"<div style='font-size:0.87rem;color:#334155;line-height:1.6'>"
                f"{definition}</div></div>")

    # Row 1: What is a Signal + Signal Categories
    r1c1, r1c2 = st.columns(2)

    with r1c1:
        signal_def = (
            rule_row("GREEN","🟢","Founder has <strong>taken action</strong> related to the sprint — completed or meaningfully in progress. E.g. hired an export manager, won an export order, completed a sprint task.") +
            rule_row("AMBER","🟡","Founder has <strong>stated a plan</strong> — intends to act but has not started yet. E.g. plans to attend a trade show, intends to hire by next quarter.") +
            rule_row("RED","🔴","Founder has <strong>not acted and has no plan</strong>, or is disengaged from the sprint. E.g. not responding, dropped out, no sprint progress.") +
            "<div style='font-size:0.78rem;color:#94a3b8;margin-top:10px;padding-top:8px;border-top:1px solid #f1f5f9'>"
            "Note: Background context, market observations, and current-state descriptions are NOT signals. "
            "Only founder actions count.</div>"
        )
        st.markdown(def_card("✦ What is a Signal?", signal_def), unsafe_allow_html=True)

    with r1c2:
        signal_nps_def = (
            "<div style='font-size:0.87rem;color:#334155;line-height:1.8;margin-bottom:14px'>"
            "<strong>Signal NPS</strong> measures the quality of founder actions using a Net Promoter Score approach:</div>"
            "<div style='background:#f8fafc;border-radius:8px;padding:14px 16px;font-family:monospace;"
            "font-size:0.85rem;color:#1e293b;margin-bottom:14px'>"
            "Signal NPS = % Green signals − % Red signals<br>"
            "Range: −100 to +100"
            "</div>"
            "<div style='font-size:0.87rem;color:#334155;margin-bottom:10px'>"
            "<strong>🟢 Green = Promoters</strong> — actions taken<br>"
            "<strong>🟡 Amber = Passives</strong> — plans stated (counted in total, not in NPS)<br>"
            "<strong>🔴 Red = Detractors</strong> — no action, no plan</div>"
            + rule_row("Green","🟢","Signal NPS ≥ 20")
            + rule_row("Amber","🟡","Signal NPS 0 to 19")
            + rule_row("Red","🔴","Signal NPS < 0")
            + rule_row("ZERO","⚪","No signals found — insufficient data")
        )
        st.markdown(def_card("📐 Signal NPS — How RAG is Calculated", signal_nps_def), unsafe_allow_html=True)

    # Row 2: Momentum RAG + Investment RAG
    r2c1, r2c2 = st.columns(2)

    with r2c1:
        mom_def = (
            "<div style='font-size:0.87rem;color:#334155;line-height:1.6;margin-bottom:12px'>"
            "Measures whether the founder is <strong>engaged and progressing</strong> on their sprint. "
            "Calculated from Sprint Momentum signals only.</div>"
            + rule_row("Green","🟢","Founder actively engaged — attending sessions, completing tasks, winning orders, sprint on track.")
            + rule_row("Amber","🟡","Partial engagement — some actions taken but sprint delayed or mixed progress.")
            + rule_row("Red","🔴","Founder disengaged — missing sessions, no sprint progress, likely to not complete.")
            + rule_row("ZERO","⚪","No momentum signals found in any document.")
        )
        st.markdown(def_card("🏃 Sprint Momentum RAG", mom_def), unsafe_allow_html=True)

    with r2c2:
        inv_def = (
            "<div style='font-size:0.87rem;color:#334155;line-height:1.6;margin-bottom:12px'>"
            "Measures whether the founder is <strong>committing resources</strong> (money, staff, time) "
            "specifically toward the sprint topic. Only sprint-relevant investment counts.</div>"
            + rule_row("Green","🟢","Founder has already invested — hired relevant staff, purchased tools, committed capital to sprint goals.")
            + rule_row("Amber","🟡","Founder has stated intent to invest — plans to hire or spend but not yet committed.")
            + rule_row("Red","🔴","No investment intent — not ready, unsure, or withdrawing resources from sprint.")
            + rule_row("ZERO","⚪","No investment signals found in any document.")
        )
        st.markdown(def_card("💰 Self Investment RAG", inv_def), unsafe_allow_html=True)

    # Row 3: Overall Venture RAG + Portfolio RAG
    r3c1, r3c2 = st.columns(2)

    with r3c1:
        overall_def = (
            "<div style='font-size:0.87rem;color:#334155;line-height:1.6;margin-bottom:12px'>"
            "The <strong>worst</strong> of Sprint Momentum RAG and Self Investment RAG. "
            "ZERO is treated as no-data and does not pull the score down.</div>"
            "<div style='background:#f8fafc;border-radius:8px;padding:12px 16px;"
            "font-size:0.83rem;font-family:monospace;color:#1e293b;margin-bottom:14px'>"
            "Overall RAG = min(Momentum RAG, Investment RAG)<br>"
            "If one is ZERO → Overall = the other one<br>"
            "If both ZERO → Overall = ZERO"
            "</div>"
            "<div style='font-size:0.85rem;font-weight:600;color:#475569;margin-bottom:8px'>"
            "10-Point Numeric Score Matrix:</div>"
            "<table style='width:100%;font-size:0.78rem;border-collapse:collapse'>"
            "<tr style='background:#f1f5f9'>"
            "<th style='padding:4px 8px;text-align:left'></th>"
            "<th style='padding:4px 8px;text-align:center'>🟢 Inv</th>"
            "<th style='padding:4px 8px;text-align:center'>🟡 Inv</th>"
            "<th style='padding:4px 8px;text-align:center'>🔴 Inv</th>"
            "<th style='padding:4px 8px;text-align:center'>⚪ Inv</th>"
            "</tr>"
            "<tr><td style='padding:4px 8px;font-weight:600'>🟢 Mom</td>"
            "<td style='padding:4px 8px;text-align:center;color:#16a34a;font-weight:700'>10</td>"
            "<td style='padding:4px 8px;text-align:center;color:#16a34a'>8</td>"
            "<td style='padding:4px 8px;text-align:center;color:#d97706'>5</td>"
            "<td style='padding:4px 8px;text-align:center;color:#d97706'>5</td></tr>"
            "<tr style='background:#f8fafc'><td style='padding:4px 8px;font-weight:600'>🟡 Mom</td>"
            "<td style='padding:4px 8px;text-align:center;color:#16a34a'>8</td>"
            "<td style='padding:4px 8px;text-align:center;color:#d97706;font-weight:700'>7</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>3</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>3</td></tr>"
            "<tr><td style='padding:4px 8px;font-weight:600'>🔴 Mom</td>"
            "<td style='padding:4px 8px;text-align:center;color:#d97706'>5</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>3</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626;font-weight:700'>1</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>0</td></tr>"
            "<tr style='background:#f8fafc'><td style='padding:4px 8px;font-weight:600'>⚪ Mom</td>"
            "<td style='padding:4px 8px;text-align:center;color:#d97706'>5</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>3</td>"
            "<td style='padding:4px 8px;text-align:center;color:#dc2626'>0</td>"
            "<td style='padding:4px 8px;text-align:center;color:#94a3b8;font-weight:700'>0</td></tr>"
            "</table>"
        )
        st.markdown(def_card("🎯 Overall Venture RAG", overall_def), unsafe_allow_html=True)

    with r3c2:
        port_def = (
            "<div style='font-size:0.87rem;color:#334155;line-height:1.6;margin-bottom:12px'>"
            "Reflects the health of the <strong>entire portfolio</strong> using Signal NPS "
            "aggregated across all ventures and all signals.</div>"
            "<div style='background:#f8fafc;border-radius:8px;padding:12px 16px;"
            "font-size:0.83rem;font-family:monospace;color:#1e293b;margin-bottom:14px'>"
            "Portfolio Signal NPS =<br>"
            "  % Green signals (all ventures) − % Red signals (all ventures)"
            "</div>"
            + rule_row("Green","🟢","Portfolio Signal NPS ≥ 20 — majority of ventures taking strong action on sprints.")
            + rule_row("Amber","🟡","Portfolio Signal NPS 0–19 — mixed portfolio, many ventures with plans but not yet acting.")
            + rule_row("Red","🔴","Portfolio Signal NPS < 0 — more ventures disengaged than engaged across the portfolio.")
            + "<div style='font-size:0.78rem;color:#94a3b8;margin-top:12px;padding-top:8px;"
            "border-top:1px solid #f1f5f9'>"
            "Filters (Hub, Venture Partner, RAG) apply to Portfolio Signal NPS — "
            "the score updates dynamically as you filter.</div>"
        )
        st.markdown(def_card("🌐 Overall Portfolio RAG", port_def), unsafe_allow_html=True)

    # Row 3.5: Stream Support definitions
    stream_def = (
        "<div style='font-size:0.87rem;color:#334155;line-height:1.6;margin-bottom:12px'>"
        "Shown in <strong>Company Basics</strong> for each of the 6 streams "
        "(GTM, Product, Operations, Supply Chain, People/HR, Finance) — "
        "indicates the level of support the venture needs in that stream.</div>"
        + rule_row("Green","🟢","On track — minimal or no support needed in this stream.")
        + rule_row("Amber","🟡","Partial / moderate need — some support required.")
        + "<div style='display:flex;align-items:flex-start;gap:14px;padding:10px 0;"
          "border-bottom:1px solid #f1f5f9'>"
          "<span style='background:#e0e7ff;color:#3730a3;padding:3px 14px;"
          "border-radius:20px;font-weight:700;font-size:0.83rem;white-space:nowrap;"
          "min-width:70px;text-align:center'>🔵 DEEP</span>"
          "<div style='font-size:0.87rem;color:#334155;line-height:1.6'>"
          "<strong>Deep Support</strong> — venture needs intensive, hands-on support "
          "in this stream, beyond standard mentoring. Typically means more frequent "
          "sessions and closer RN team involvement.</div></div>"
        + rule_row("Red","🔴","Critical gap — significant support needed in this stream.")
        + "<div style='display:flex;align-items:flex-start;gap:14px;padding:10px 0'>"
          "<span style='background:#fee2e2;color:#991b1b;padding:3px 14px;"
          "border-radius:20px;font-weight:700;font-size:0.83rem;white-space:nowrap;"
          "min-width:70px;text-align:center'>🔴 RED+D</span>"
          "<div style='font-size:0.87rem;color:#334155;line-height:1.6'>"
          "<strong>Red (Deep Support)</strong> — critical gap AND needs intensive, "
          "hands-on intervention. Highest priority for RN team attention.</div></div>"
    )
    st.markdown(def_card("🔧 Stream Support Levels", stream_def), unsafe_allow_html=True)

    # Row 4: Data sources
    st.markdown(
        def_card("📂 Data Sources Used Per Venture",
            "<div style='display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;font-size:0.83rem'>"
            "<div><div style='font-weight:600;color:#1e293b;margin-bottom:6px'>📁 Venture Folder</div>"
            "<div style='color:#475569;line-height:1.8'>"
            "Feedback file<br>Session transcript<br>Sprint plan<br>Growth Journey report<br>Other venture documents</div></div>"
            "<div><div style='font-weight:600;color:#1e293b;margin-bottom:6px'>📂 Common Documents</div>"
            "<div style='color:#475569;line-height:1.8'>"
            "Attendance tracker<br>Growth Journey reports<br>Export signal documents<br>"
            "Session Transcripts folder<br>Any other shared files</div></div>"
            "<div><div style='font-weight:600;color:#1e293b;margin-bottom:6px'>📊 Tracker Files</div>"
            "<div style='color:#475569;line-height:1.8'>"
            "05_Session_Management_Tracker<br>06_Feedback_Quality_Tracker<br>"
            "Portfolio Dashboard Excel<br>(Company sheet, row 3 headers)</div></div>"
            "</div>"
        ),
        unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════
#  TAB: COMPANIES  (sub-tabs: Company Basics | Company & Score)
# ══════════════════════════════════════════════════════
with tab_companies:
    tab_company, tab_scores = st.tabs(["🏢 Company Basics", "🎯 Company & Score"])

with tab_company:
    st.title("🏢 Company Basics")
    st.divider()

    if not company_basics:
        st.warning("Company Basics data not loaded.")
    else:
        # ── Helpers ───────────────────────────────────
        def safe(v, default="—"):
            if not v or str(v).strip() in ["None","nan","NaT","","0","null"]:
                return default
            return str(v).strip()

        def jd(vname, field, default="—"):
            """Get field from journey_data for a venture."""
            return safe(journey_data.get(vname, {}).get(field), default)

        def stream_badge_html(val):
            v = safe(val)
            if v == "—": return "<span style='color:#94a3b8'>—</span>"
            color_map = {
                "RED":                ("#fee2e2","#991b1b"),
                "RED (DEEP SUPPORT)": ("#fee2e2","#991b1b"),
                "AMBER":              ("#fef9c3","#854d0e"),
                "DEEP SUPPORT":       ("#e0e7ff","#3730a3"),
                "GREEN":              ("#dcfce7","#166534"),
            }
            bg, fg = color_map.get(v.upper(), ("#f1f5f9","#475569"))
            short = {"RED":"RED","AMBER":"AMB","DEEP SUPPORT":"DEEP",
                     "RED (DEEP SUPPORT)":"RED+D","GREEN":"GRN"}.get(v.upper(), v[:6])
            return (f"<span style='background:{bg};color:{fg};padding:2px 8px;"
                    f"border-radius:10px;font-size:0.73rem;font-weight:700'>{short}</span>")

        def sprint_status_badge(val):
            v = safe(val)
            if v == "—": return "—"
            colors = {"In Progress":("#dcfce7","#166534"),
                      "Completed":  ("#e0e7ff","#3730a3"),
                      "Not Started":("#fee2e2","#991b1b")}
            bg, fg = colors.get(v, ("#f1f5f9","#475569"))
            return (f"<span style='background:{bg};color:{fg};padding:2px 8px;"
                    f"border-radius:10px;font-size:0.73rem;font-weight:600'>{v}</span>")

        STREAMS = [
            ("GTM",          "stream_support_gtm",          "goal_gtm",          "unlock_gtm"),
            ("Product",      "stream_support_product",       "goal_product",      "unlock_product"),
            ("Operations",   "stream_support_operations",    "goal_operations",   "unlock_operations"),
            ("Supply Chain", "stream_support_supply_chain",  "goal_supply_chain", "unlock_supply_chain"),
            ("People / HR",  "stream_support_hr",            "goal_people",       "unlock_people"),
            ("Finance",      "stream_support_finance",       "goal_finance",      "unlock_finance"),
        ]

        TH = ("style='padding:8px 12px;text-align:left;color:#475569;"
              "font-weight:600;font-size:0.78rem;background:#f1f5f9;"
              "white-space:nowrap;border-bottom:2px solid #e2e8f0'")
        TD = ("style='padding:8px 12px;font-size:0.8rem;color:#334155;"
              "vertical-align:middle;border-bottom:1px solid #f1f5f9'")

        # ── Filters ───────────────────────────────────
        fc1, fc2, fc3, fc4 = st.columns(4)
        cb_search  = fc1.text_input("🔍 Search Venture", key="cb_search")
        hub_vals   = sorted(set(v.get("hub","—") or "—" for v in company_basics.values()
                                if (v.get("hub","—") or "—") != "—"))
        cb_hub     = fc2.selectbox("Hub", ["All"] + hub_vals, key="cb_hub")
        spr_vals   = sorted(set(v.get("sprint_type","—") or "—" for v in company_basics.values()
                                if (v.get("sprint_type","—") or "—") not in ["—","None"]))
        cb_sprint  = fc3.selectbox("Sprint Type", ["All"] + spr_vals, key="cb_sprint")
        type_vals  = sorted(set(v.get("startup_smb","—") or "—" for v in company_basics.values()
                                if (v.get("startup_smb","—") or "—") not in ["—","None","0"]))
        cb_type    = fc4.selectbox("Startup/SMB", ["All"] + type_vals, key="cb_type")

        filtered_basics = {}
        for vn, vdata in company_basics.items():
            if cb_search and cb_search.lower() not in vn.lower(): continue
            if cb_hub    != "All" and (vdata.get("hub") or "—") != cb_hub: continue
            if cb_sprint != "All" and (vdata.get("sprint_type") or "—") != cb_sprint: continue
            if cb_type   != "All" and (vdata.get("startup_smb") or "—") != cb_type: continue
            filtered_basics[vn] = vdata

        st.caption(f"{len(filtered_basics)} ventures · "
                   f"Journey data available for {sum(1 for v in filtered_basics if v in journey_data)} ventures")
        st.markdown("<br>", unsafe_allow_html=True)

        # ════════════════════════════════════════════
        # MAIN TABLE — Company Overview (from Excel)
        # ════════════════════════════════════════════
        # MAIN TABLE — Company Overview (st.dataframe with row selection)
        # ════════════════════════════════════════════
        st.markdown("#### 📋 Company Overview")
        st.caption("Source: Journey_Accelerate_Portfolio_Dashboard.xlsx · Click a row to view venture detail below")

        # Build dataframe for display
        rows_df = []
        venture_order = list(filtered_basics.keys())

        for vn in venture_order:
            vdata = filtered_basics[vn]
            rev_size = safe(vdata.get("size_revenue"))
            rev_fy   = safe(vdata.get("revenue_fy2425_26"))
            incr_rev = safe(vdata.get("incremental_rev_3yr"))
            incr_job = safe(vdata.get("incremental_jobs_3yr"))
            try: incr_rev_fmt = f"${float(incr_rev):.1f}M" if incr_rev != "—" else "—"
            except: incr_rev_fmt = incr_rev
            try: incr_job_fmt = f"{int(float(incr_job)):,}" if incr_job != "—" else "—"
            except: incr_job_fmt = incr_job

            rows_df.append({
                "Venture":          vn,
                "Hub":              safe(vdata.get("hub")),
                "Type":             safe(vdata.get("startup_smb")),
                "Sprint Type":      safe(vdata.get("sprint_type")),
                "Sprint Status":    safe(vdata.get("sprint_status")),
                "Category":         safe(vdata.get("program_category")),
                "Month-Year":       safe(vdata.get("month_year")),
                "Revenue (Size)":   f"${rev_size}M" if rev_size != "—" else "—",
                "Rev FY25-26":      f"${rev_fy}M"   if rev_fy   != "—" else "—",
                "Incr Rev (3Yr)":   incr_rev_fmt,
                "Incr Jobs (3Yr)":  incr_job_fmt,
                "Journey Doc":      "✅" if vn in journey_data else "⬜",
            })

        df_overview = pd.DataFrame(rows_df)

        # Row selection event
        selection = st.dataframe(
            df_overview,
            use_container_width=True,
            height=320,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="cb_table_selection",
        )

        # Determine selected venture
        selected_rows = selection.selection.get("rows", [])
        selected_venture = venture_order[selected_rows[0]] if selected_rows else None

        # ════════════════════════════════════════════
        # VENTURE DETAIL — shown when a row is selected
        # From Journey Documents
        # ════════════════════════════════════════════
        if selected_venture:
            vdata = filtered_basics.get(selected_venture, {})
            jdata = journey_data.get(selected_venture, {})

            st.markdown(
                f"<div style='background:linear-gradient(135deg,#6366f1,#8b5cf6);"
                f"border-radius:10px;padding:14px 20px;color:white;margin:16px 0 12px'>"
                f"<div style='font-size:1.05rem;font-weight:700'>{selected_venture}</div>"
                f"<div style='font-size:0.82rem;opacity:0.85;margin-top:3px'>"
                f"📍 {safe(vdata.get('hub'))} &nbsp;·&nbsp; "
                f"🏃 {safe(vdata.get('sprint_type'))} &nbsp;·&nbsp; "
                f"{safe(vdata.get('startup_smb'))} &nbsp;·&nbsp; "
                f"{safe(vdata.get('program_category'))}"
                f"</div></div>",
                unsafe_allow_html=True
            )

            if not jdata:
                st.info(
                    "No Journey Document data found for this venture. "
                    "Run Backend → Step 3 to extract from Sign off Journey Documents."
                )
            else:
                source_note = "<div style='font-size:0.72rem;color:#94a3b8;margin-bottom:12px'>Source: Sign off Journey Documents</div>"
                st.markdown(source_note, unsafe_allow_html=True)

                # ── Section 1: Existing → New ──────────────
                st.markdown("**📦 Existing → New Venture**")
                s1, arr, s2 = st.columns([5, 0.3, 5])

                with s1:
                    sc1, sc2, sc3 = st.columns(3)
                    with sc1:
                        st.markdown(
                            f"<div class='section-label'>Existing Product</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'existing_product')}</div>",
                            unsafe_allow_html=True)
                    with sc2:
                        st.markdown(
                            f"<div class='section-label'>Existing Markets</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'existing_market_segments')}</div>",
                            unsafe_allow_html=True)
                    with sc3:
                        st.markdown(
                            f"<div class='section-label'>Existing Geographies</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'existing_geographies')}</div>",
                            unsafe_allow_html=True)

                with arr:
                    st.markdown(
                        "<div style='text-align:center;font-size:1.5rem;"
                        "margin-top:20px;color:#6366f1'>→</div>",
                        unsafe_allow_html=True)

                with s2:
                    sc4, sc5, sc6 = st.columns(3)
                    with sc4:
                        st.markdown(
                            f"<div class='section-label'>New Product</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'new_product')}</div>",
                            unsafe_allow_html=True)
                    with sc5:
                        st.markdown(
                            f"<div class='section-label'>New Markets</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'new_market_segments')}</div>",
                            unsafe_allow_html=True)
                    with sc6:
                        st.markdown(
                            f"<div class='section-label'>New Geographies</div>"
                            f"<div style='font-size:0.83rem;color:#334155;white-space:pre-line'>"
                            f"{jd(selected_venture,'new_geographies')}</div>",
                            unsafe_allow_html=True)

                # Incremental metrics
                incr_rev = jd(selected_venture, "incremental_rev_3yr")
                incr_job = jd(selected_venture, "incremental_jobs_3yr")
                if incr_rev != "—" or incr_job != "—":
                    st.markdown("<br>", unsafe_allow_html=True)
                    im1, im2, _ = st.columns([2, 2, 6])
                    if incr_rev != "—":
                        try: im1.metric("Incremental Revenue (3 Yr)", f"${float(incr_rev):.1f}M")
                        except: im1.metric("Incremental Revenue (3 Yr)", incr_rev)
                    if incr_job != "—":
                        try: im2.metric("Incremental Jobs (3 Yr)", f"{int(float(incr_job)):,}")
                        except: im2.metric("Incremental Jobs (3 Yr)", incr_job)

                st.divider()

                # ── Section 2: Stream Support ───────────────
                st.markdown("**🔧 Stream Support Requirement**")
                sc_cols = st.columns(6)
                for idx, (label, sup_key, _, _) in enumerate(STREAMS):
                    with sc_cols[idx]:
                        badge = stream_badge_html(jdata.get(sup_key))
                        st.markdown(
                            f"<div style='text-align:center'>"
                            f"<div class='section-label' style='text-align:center'>{label}</div>"
                            f"<div style='margin-top:6px'>{badge}</div></div>",
                            unsafe_allow_html=True)

                st.divider()

                # ── Section 3: Goals ────────────────────────
                st.markdown("**🎯 Stream Goals (36 Months)**")
                gc = st.columns(3)
                for idx, (label, sup_key, goal_key, _) in enumerate(STREAMS):
                    goal  = jd(selected_venture, goal_key)
                    badge = stream_badge_html(jdata.get(sup_key))
                    with gc[idx % 3]:
                        st.markdown(
                            f"<div style='background:#ffffff;border:1px solid #e2e8f0;"
                            f"border-radius:8px;padding:12px 14px;margin-bottom:10px;min-height:80px'>"
                            f"<div style='display:flex;justify-content:space-between;"
                            f"align-items:center;margin-bottom:6px'>"
                            f"<span style='font-weight:600;font-size:0.83rem;color:#1e293b'>"
                            f"{label}</span>{badge}</div>"
                            f"<div style='font-size:0.8rem;color:#475569;line-height:1.5'>"
                            f"{goal}</div></div>",
                            unsafe_allow_html=True)

                # ── Section 4: Stream Unlocks ───────────────
                has_unlocks = any(jd(selected_venture, uk) != "—"
                                  for _, _, _, uk in STREAMS)
                if has_unlocks:
                    st.divider()
                    st.markdown("**🔓 Stream Unlocks**")
                    uc = st.columns(3)
                    for idx, (label, _, _, unlock_key) in enumerate(STREAMS):
                        unlock = jd(selected_venture, unlock_key)
                        if unlock == "—": continue
                        with uc[idx % 3]:
                            st.markdown(
                                f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                                f"border-radius:8px;padding:12px 14px;margin-bottom:10px'>"
                                f"<div style='font-weight:600;font-size:0.83rem;color:#1e293b;"
                                f"margin-bottom:6px'>{label}</div>"
                                f"<div style='font-size:0.79rem;color:#475569;"
                                f"line-height:1.6;white-space:pre-line'>{unlock}</div></div>",
                                unsafe_allow_html=True)
        else:
            st.info("👆 Click any row in the table above to view venture detail.")


# ══════════════════════════════════════════════════════
#  TAB: COMPANY & SCORE
# ══════════════════════════════════════════════════════
with tab_scores:
    st.title("🎯 Company & Score")
    st.divider()

    if not signals_repo:
        st.warning(
            "Signals repository not found. "
            "Run Backend → Step 1 and upload signals_repository.json to SharePoint."
        )
    else:
        # Use venture_rag — computed LIVE from signals via NPS formula
        # This is always in sync with actual signals, never stale
        # venture_rag is built at load time in the repo parsing section above

        # ── Filters ───────────────────────────────────
        f1, f2, f3, f4 = st.columns(4)
        cs_search  = f1.text_input("🔍 Search", key="cs_search")
        rag_opts   = ["All", "🟢 Green", "🟡 Amber", "🔴 Red", "⚪ ZERO"]
        cs_overall = f2.selectbox("Overall RAG", rag_opts, key="cs_overall")
        cs_mom     = f3.selectbox("Momentum RAG", rag_opts, key="cs_mom")
        cs_inv     = f4.selectbox("Investment RAG", rag_opts, key="cs_inv")

        RAG_EMOJI  = {"Green":"🟢","Amber":"🟡","Red":"🔴","ZERO":"⚪","Unknown":"⚪"}
        RAG_COLOR  = {
            "Green":  ("#dcfce7","#166534"),
            "Amber":  ("#fef9c3","#854d0e"),
            "Red":    ("#fee2e2","#991b1b"),
            "ZERO":   ("#f1f5f9","#64748b"),
            "Unknown":("#f1f5f9","#64748b"),
        }

        def rag_cell(rag):
            emoji  = RAG_EMOJI.get(rag,"⚪")
            bg, fg = RAG_COLOR.get(rag,("#f1f5f9","#64748b"))
            return (f"<span style='background:{bg};color:{fg};padding:3px 12px;"
                    f"border-radius:20px;font-weight:700;font-size:0.8rem;"
                    f"white-space:nowrap'>{emoji} {rag}</span>")

        def score_cell(score):
            try:
                s = int(score)
                color = ("#16a34a" if s >= 8 else
                         "#d97706" if s >= 5 else
                         "#dc2626" if s >= 1 else "#94a3b8")
                return (f"<span style='font-weight:800;font-size:1rem;"
                        f"color:{color}'>{s}/10</span>")
            except: return "—"

        rows = []
        rag_order = {"Red":0,"Amber":1,"Green":2,"ZERO":3,"Unknown":4}

        for vn, vdata in venture_rag.items():
            overall  = vdata.get("overall_rag","ZERO")
            m_rag    = vdata.get("momentum_rag","ZERO")
            i_rag    = vdata.get("investment_rag","ZERO")
            m_reason = vdata.get("momentum_reason","—")
            i_reason = vdata.get("investment_reason","—")
            note     = vdata.get("overall_note","—")
            score    = vdata.get("momentum_score",0)
            hub      = vdata.get("hub","—")
            vp       = vdata.get("venture_partner","—")

            if cs_search and cs_search.lower() not in vn.lower(): continue
            if cs_overall != "All" and overall != cs_overall.split(" ",1)[1]: continue
            if cs_mom     != "All" and m_rag   != cs_mom.split(" ",1)[1]:     continue
            if cs_inv     != "All" and i_rag   != cs_inv.split(" ",1)[1]:     continue

            rows.append({
                "vn":vn,"hub":hub,"vp":vp,
                "overall":overall,"m_rag":m_rag,"i_rag":i_rag,
                "m_reason":m_reason,"i_reason":i_reason,
                "score":score,"note":note,
            })

        rows.sort(key=lambda r: rag_order.get(r["overall"],4))
        st.caption(f"{len(rows)} ventures · sorted by RAG (Red first)")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Dialog for signals ────────────────────────
        @st.dialog("Signals Detail", width="large")
        def show_signals_dialog(vname, signal_type):
            sigs = venture_signals.get(vname, {})
            sig_list = sigs.get(signal_type, [])
            st.markdown(
                f"**{vname}** — "
                f"{'🏃 Sprint Momentum' if signal_type=='momentum' else '💰 Self Investment'} Signals"
            )
            if not sig_list:
                st.info("No signals found for this venture.")
                return
            # Group by category
            for cat, emoji, bg, fg in [
                ("GREEN","🟢","#dcfce7","#166534"),
                ("AMBER","🟡","#fef9c3","#854d0e"),
                ("RED",  "🔴","#fee2e2","#991b1b"),
            ]:
                cat_sigs = [s for s in sig_list if s.get("category","GREEN") == cat]
                if not cat_sigs: continue
                st.markdown(
                    f"<div style='background:{bg};color:{fg};padding:4px 12px;"
                    f"border-radius:8px;font-weight:700;font-size:0.82rem;"
                    f"margin:10px 0 6px'>{emoji} {cat} — {len(cat_sigs)} signal(s)</div>",
                    unsafe_allow_html=True
                )
                for s in cat_sigs:
                    st.markdown(
                        f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                        f"border-radius:8px;padding:10px 14px;margin-bottom:6px'>"
                        f"<div style='font-weight:600;font-size:0.83rem;color:#1e293b'>"
                        f"{s.get('type','')}</div>"
                        f"<div style='font-size:0.8rem;color:#475569;margin-top:4px'>"
                        f"{s.get('evidence','')}</div>"
                        f"<div style='font-size:0.72rem;color:#94a3b8;margin-top:4px'>"
                        f"📄 {s.get('source','')}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

        # Track which dialog to open
        if "cs_dialog_venture" not in st.session_state:
            st.session_state["cs_dialog_venture"] = None
            st.session_state["cs_dialog_type"]    = None

        if st.session_state["cs_dialog_venture"]:
            show_signals_dialog(
                st.session_state["cs_dialog_venture"],
                st.session_state["cs_dialog_type"]
            )
            st.session_state["cs_dialog_venture"] = None
            st.session_state["cs_dialog_type"]    = None

        # ── Table: one row per venture ────────────────
        TH = ("style='padding:9px 14px;text-align:left;color:#475569;"
              "font-weight:600;font-size:0.78rem;background:#f1f5f9;"
              "white-space:nowrap;border-bottom:2px solid #e2e8f0'")
        TD = ("style='padding:9px 14px;font-size:0.8rem;color:#334155;"
              "vertical-align:top;border-bottom:1px solid #f1f5f9'")
        TDW = ("style='padding:9px 14px;font-size:0.8rem;color:#475569;"
               "vertical-align:top;border-bottom:1px solid #f1f5f9;max-width:200px'")

        # Header
        h0,h1,h2,h3,h4,h5,h6,h7 = st.columns([2.5,1,2,1,2,0.8,1,2.5])
        h0.markdown("**Company**")
        h1.markdown("**Momentum RAG**")
        h2.markdown("**Momentum Reason**")
        h3.markdown("**Investment RAG**")
        h4.markdown("**Investment Reason**")
        h5.markdown("**Score**")
        h6.markdown("**Overall RAG**")
        h7.markdown("**Note on Overall Score**")
        st.divider()

        for ri, r in enumerate(rows):
            vn       = r["vn"]
            m_sigs_v = venture_signals.get(vn,{}).get("momentum",[])
            i_sigs_v = venture_signals.get(vn,{}).get("investment",[])

            c0,c1,c2,c3,c4,c5,c6,c7 = st.columns([2.5,1,2,1,2,0.8,1,2.5])

            with c0:
                st.markdown(
                    f"<div style='font-weight:600;font-size:0.83rem'>{vn}</div>"
                    f"<div style='font-size:0.72rem;color:#94a3b8'>"
                    f"{r['hub']} · {r['vp']}</div>",
                    unsafe_allow_html=True
                )
            with c1:
                st.markdown(rag_cell(r["m_rag"]), unsafe_allow_html=True)
            with c2:
                st.markdown(
                    f"<div style='font-size:0.79rem;color:#475569'>{r['m_reason']}</div>",
                    unsafe_allow_html=True
                )
                if m_sigs_v:
                    if st.button("🔍", key=f"m_sig_{ri}",
                                 help=f"View {len(m_sigs_v)} momentum signals"):
                        st.session_state["cs_dialog_venture"] = vn
                        st.session_state["cs_dialog_type"]    = "momentum"
                        st.rerun()
            with c3:
                st.markdown(rag_cell(r["i_rag"]), unsafe_allow_html=True)
            with c4:
                st.markdown(
                    f"<div style='font-size:0.79rem;color:#475569'>{r['i_reason']}</div>",
                    unsafe_allow_html=True
                )
                if i_sigs_v:
                    if st.button("🔍", key=f"i_sig_{ri}",
                                 help=f"View {len(i_sigs_v)} investment signals"):
                        st.session_state["cs_dialog_venture"] = vn
                        st.session_state["cs_dialog_type"]    = "investment"
                        st.rerun()
            with c5:
                st.markdown(score_cell(r["score"]), unsafe_allow_html=True)
            with c6:
                st.markdown(rag_cell(r["overall"]), unsafe_allow_html=True)
            with c7:
                st.markdown(
                    f"<div style='font-size:0.79rem;color:#475569'>{r['note']}</div>",
                    unsafe_allow_html=True
                )

            st.markdown(
                "<hr style='margin:4px 0;border:none;border-top:1px solid #f1f5f9'>",
                unsafe_allow_html=True
            )

        # ── Download CSV ──────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        def esc(v): return f'"{str(v).replace(chr(34), chr(39))}"'
        csv_lines = ["Company,Hub,Venture Partner,Momentum RAG,Momentum Reason,"
                     "Investment RAG,Investment Reason,Score,Overall RAG,Note"]
        for r in rows:
            csv_lines.append(
                f"{esc(r['vn'])},{esc(r['hub'])},{esc(r['vp'])},"
                f"{esc(r['m_rag'])},{esc(r['m_reason'])},"
                f"{esc(r['i_rag'])},{esc(r['i_reason'])},"
                f"{esc(r['score'])},{esc(r['overall'])},{esc(r['note'])}"
            )
        st.download_button(
            "⬇️ Download as CSV",
            data="\n".join(csv_lines),
            file_name="company_scores.csv",
            mime="text/csv",
            key="cs_download"
        )

with tab_overview:
    st.title("📊 Portfolio Overview")

    # ── filters row ───────────────────────────────────
    f1, f2, f3, f4 = st.columns(4)
    hub_opts    = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_list if cv(get_row(v),col_hub) != "—"))
    vp_opts     = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_list if col_vp and cv(get_row(v),col_vp) != "—"))

    rag_opts    = ["All","🟢 Green","🟡 Amber","🔴 Red","⚪ ZERO"]

    hub_f   = f1.selectbox("Hub",              hub_opts,   key="ov_hub")
    vp_f    = f2.selectbox("Venture Partner",  vp_opts,    key="ov_vp")
    stage_f = "All"  # Sprint Stage filter removed
    rag_f   = f4.selectbox("Overall RAG",      rag_opts,   key="ov_rag")

    # Build filtered venture data
    venture_data = []
    for vname in ventures_list:
        row    = get_row(vname)
        hub    = cv(row, col_hub)
        vp     = cv(row, col_vp) if col_vp else "—"
        sprint = cv(row, col_sprint)
        rev    = cv(row, col_rev)
        pct    = row[col_pct] if (row is not None and col_pct) else None
        bucket = get_stage_bucket(pct)
        notes  = cv(row, col_notes, default="")

        rag_data   = venture_rag.get(vname, {})
        overall    = rag_data.get("overall_rag","ZERO")
        m_rag      = rag_data.get("momentum_rag","ZERO")
        i_rag      = rag_data.get("investment_rag","ZERO")
        m_reason   = rag_data.get("momentum_reason","—")
        i_reason   = rag_data.get("investment_reason","—")
        m_score    = rag_data.get("momentum_score",0)
        i_score    = rag_data.get("investment_score",0)

        venture_data.append({
            "name": vname, "hub": hub, "vp": vp, "sprint": sprint,
            "rev": rev, "bucket": bucket, "pct_raw": pct, "notes": notes,
            "overall_rag": overall, "momentum_rag": m_rag, "investment_rag": i_rag,
            "momentum_reason": m_reason, "investment_reason": i_reason,
            "momentum_score": m_score, "investment_score": i_score,
        })

    # Apply filters
    filtered = venture_data
    if hub_f   != "All": filtered = [v for v in filtered if v["hub"]    == hub_f]
    if vp_f    != "All": filtered = [v for v in filtered if v["vp"]     == vp_f]

    if rag_f   != "All":
        rag_val = rag_f.split(" ",1)[1]
        filtered = [v for v in filtered if v["overall_rag"] == rag_val]

    total  = len(filtered)
    greens = sum(1 for v in filtered if v["overall_rag"] == "Green")
    ambers = sum(1 for v in filtered if v["overall_rag"] == "Amber")
    reds   = sum(1 for v in filtered if v["overall_rag"] == "Red")
    zeros  = sum(1 for v in filtered if v["overall_rag"] == "ZERO")

    if reds   > 0: portfolio_rag = "Red"
    elif ambers>0: portfolio_rag = "Amber"
    elif greens>0: portfolio_rag = "Green"
    else:          portfolio_rag = "ZERO"

    st.markdown(f"### Overall Portfolio RAG: {rag_badge(portfolio_rag)}", unsafe_allow_html=True)
    st.caption("Based on Knowledge Repository · Updates with filters")
    st.markdown("<br>", unsafe_allow_html=True)

    # ── summary metrics — 5 columns ───────────────────
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Ventures", total)
    c2.metric("🟢 Green", greens, f"{round(greens/total*100) if total else 0}%")
    c3.metric("🟡 Amber", ambers, f"{round(ambers/total*100) if total else 0}%")
    c4.metric("🔴 Red",   reds,   f"{round(reds/total*100)   if total else 0}%")
    c5.metric("⚪ No Data",zeros)

    st.divider()

    # ── 3-column layout: RAG + Stage + Hub ────────────
    col_a, col_b, col_c = st.columns([3, 3, 4])

    with col_a:
        st.subheader("RAG Distribution")
        st.markdown("<br>", unsafe_allow_html=True)
        for label, count, color in [
            ("🟢 Green", greens, "#16a34a"),
            ("🟡 Amber", ambers, "#d97706"),
            ("🔴 Red",   reds,   "#dc2626"),
            ("⚪ No Data", zeros, "#94a3b8"),
        ]:
            pct = round(count/total*100) if total else 0
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;margin-bottom:4px'>"
                f"<span style='font-weight:600;color:{color}'>{label}</span>"
                f"<span style='color:#64748b'>{count} &nbsp;({pct}%)</span></div>",
                unsafe_allow_html=True)
            st.progress(pct/100)
            st.markdown("<div style='margin-bottom:12px'></div>", unsafe_allow_html=True)

    with col_b:
        st.subheader("Portfolio Signal NPS")
        st.markdown("<br>", unsafe_allow_html=True)

        # Aggregate Signal NPS across all filtered ventures
        all_signals_flat = []
        for v in filtered:
            sigs = venture_signals.get(v["name"], {"momentum":[],"investment":[]})
            all_signals_flat.extend(sigs.get("momentum",[]))
            all_signals_flat.extend(sigs.get("investment",[]))

        if all_signals_flat:
            total_sigs = len(all_signals_flat)
            port_g = sum(1 for s in all_signals_flat if s.get("category","GREEN") == "GREEN")
            port_a = sum(1 for s in all_signals_flat if s.get("category","GREEN") == "AMBER")
            port_r = sum(1 for s in all_signals_flat if s.get("category","GREEN") == "RED")
            port_nps = round(port_g/total_sigs*100) - round(port_r/total_sigs*100)
            nps_rag  = "Green" if port_nps >= 20 else ("Amber" if port_nps >= 0 else "Red")
            nps_color = {"Green":"#16a34a","Amber":"#d97706","Red":"#dc2626"}.get(nps_rag,"#64748b")

            st.markdown(
                f"<div style='text-align:center;margin-bottom:16px'>"
                f"<div style='font-size:2.8rem;font-weight:800;color:{nps_color}'>{port_nps:+d}</div>"
                f"<div style='font-size:0.82rem;color:#64748b'>Portfolio Signal NPS</div>"
                f"<div style='margin-top:6px'>{rag_badge(nps_rag)}</div>"
                f"</div>",
                unsafe_allow_html=True
            )
            for label, cnt, color in [
                ("🟢 Green (Promoters)",  port_g, "#16a34a"),
                ("🟡 Amber (Passives)",   port_a, "#d97706"),
                ("🔴 Red (Detractors)",   port_r, "#dc2626"),
            ]:
                pct_s = round(cnt/total_sigs*100) if total_sigs else 0
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;margin-bottom:4px'>"
                    f"<span style='font-weight:500;color:{color}'>{label}</span>"
                    f"<span style='color:#64748b'>{cnt} ({pct_s}%)</span></div>",
                    unsafe_allow_html=True)
                st.progress(pct_s/100)
                st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)
            st.caption(f"Based on {total_sigs} total signals across {len(filtered)} ventures")
        else:
            st.caption("No signals loaded yet — run backend to generate repository.")

    with col_c:
        st.subheader("Hub-wise RAG Count")
        hub_rag = {}
        for v in filtered:
            h = v["hub"] if v["hub"] not in ["—","Other"] else "Other"
            if h not in hub_rag:
                hub_rag[h] = {"Green":0,"Amber":0,"Red":0,"ZERO":0,"Total":0}
            hub_rag[h][v["overall_rag"]] = hub_rag[h].get(v["overall_rag"],0)+1
            hub_rag[h]["Total"] += 1

        h0,h1,h2,h3,h4,h5 = st.columns([2.2,0.7,0.7,0.7,0.7,0.7])
        for lbl,col in zip(["**Hub**","**Tot**","**🟢**","**🟡**","**🔴**","**⚪**"],
                            [h0,h1,h2,h3,h4,h5]):
            col.markdown(lbl)
        st.divider()
        for hub, counts in sorted(hub_rag.items(), key=lambda x: (x[0]=="Other",-x[1]["Total"])):
            r0,r1,r2,r3,r4,r5 = st.columns([2.2,0.7,0.7,0.7,0.7,0.7])
            r0.markdown(f"**{hub}**")
            r1.markdown(f"**{counts['Total']}**")
            r2.markdown(f"<span style='color:#16a34a;font-weight:700'>{counts.get('Green',0)}</span>", unsafe_allow_html=True)
            r3.markdown(f"<span style='color:#d97706;font-weight:700'>{counts.get('Amber',0)}</span>", unsafe_allow_html=True)
            r4.markdown(f"<span style='color:#dc2626;font-weight:700'>{counts.get('Red',0)}</span>",   unsafe_allow_html=True)
            r5.markdown(f"<span style='color:#94a3b8;font-weight:700'>{counts.get('ZERO',0)}</span>",  unsafe_allow_html=True)
        st.divider()
        t0,t1,t2,t3,t4,t5 = st.columns([2.2,0.7,0.7,0.7,0.7,0.7])
        t0.markdown("**Total**"); t1.markdown(f"**{total}**")
        t2.markdown(f"**{greens}**"); t3.markdown(f"**{ambers}**")
        t4.markdown(f"**{reds}**");   t5.markdown(f"**{zeros}**")

    # ── venture partner breakdown ──────────────────────
    st.divider()
    st.subheader("Venture Partner — Portfolio Breakdown")
    vp_data = {}
    for v in filtered:
        vp = v["vp"] if v["vp"] not in ["—",""] else "Unassigned"
        if vp not in vp_data:
            vp_data[vp] = {"Green":0,"Amber":0,"Red":0,"ZERO":0,"Total":0}
        vp_data[vp][v["overall_rag"]] = vp_data[vp].get(v["overall_rag"],0)+1
        vp_data[vp]["Total"] += 1

    vp_cols = st.columns(min(len(vp_data), 4))
    for idx, (vp_name, counts) in enumerate(sorted(vp_data.items(), key=lambda x: -x[1]["Total"])):
        col = vp_cols[idx % len(vp_cols)]
        with col:
            st.markdown(
                f"<div class='info-card' style='text-align:center'>"
                f"<div style='font-weight:700;color:#1e293b;margin-bottom:8px'>{vp_name}</div>"
                f"<div style='font-size:1.6rem;font-weight:700;color:#6366f1'>{counts['Total']}</div>"
                f"<div style='font-size:0.75rem;color:#64748b'>ventures</div>"
                f"<div style='margin-top:10px;display:flex;justify-content:center;gap:8px;flex-wrap:wrap'>"
                f"<span style='color:#16a34a;font-weight:600'>🟢 {counts.get('Green',0)}</span>"
                f"<span style='color:#d97706;font-weight:600'>🟡 {counts.get('Amber',0)}</span>"
                f"<span style='color:#dc2626;font-weight:600'>🔴 {counts.get('Red',0)}</span>"
                f"</div></div>",
                unsafe_allow_html=True
            )

    # ── hub pivot score table ──────────────────────────
    st.divider()
    st.subheader("Hub-wise Venture Score (Sprint Momentum × Self Investment)")
    st.caption("Based on 10-point scoring matrix")

    hub_pivot = {}
    for v in filtered:
        h = v["hub"] if v["hub"] not in ["—","Other"] else "Other"
        m = v["momentum_rag"]
        if h not in hub_pivot: hub_pivot[h] = {}
        if m not in hub_pivot[h]: hub_pivot[h][m] = {"count":0,"scores":[]}
        hub_pivot[h][m]["count"]  += 1
        hub_pivot[h][m]["scores"].append(v.get("momentum_score",0))

    ph0,ph1,ph2,ph3,ph4,ph5,ph6 = st.columns([2.5,1.2,0.8,0.8,0.8,0.8,0.8])
    for lbl,col in zip(["**Hub**","**Momentum**","**Count**","**Avg**","**🟢**","**🟡**","**🔴**"],
                        [ph0,ph1,ph2,ph3,ph4,ph5,ph6]):
        col.markdown(lbl)
    st.divider()

    grand_total = 0; grand_scores = []
    for hub in sorted(hub_pivot.keys(), key=lambda x: (x=="Other",x)):
        hub_total=0; hub_scores_all=[]
        hub_g = hub_pivot[hub].get("Green",{}).get("count",0)
        hub_a = hub_pivot[hub].get("Amber",{}).get("count",0)
        hub_r = hub_pivot[hub].get("Red",{}).get("count",0)
        first = True
        for m_rag in ["Green","Amber","Red","ZERO"]:
            if m_rag not in hub_pivot[hub]: continue
            data = hub_pivot[hub][m_rag]
            avg  = round(sum(data["scores"])/len(data["scores"]),1) if data["scores"] else 0.0
            hub_total     += data["count"]
            hub_scores_all.extend(data["scores"])
            grand_scores.extend(data["scores"])
            c0,c1,c2,c3,c4,c5,c6 = st.columns([2.5,1.2,0.8,0.8,0.8,0.8,0.8])
            c0.markdown(f"**{hub}**" if first else "")
            c1.markdown(f"{RAG_EMOJI.get(m_rag,'⚪')} {m_rag}")
            c2.markdown(str(data["count"])); c3.markdown(f"**{avg}**")
            if first:
                c4.markdown(f"<span style='color:#16a34a;font-weight:700'>{hub_g}</span>", unsafe_allow_html=True)
                c5.markdown(f"<span style='color:#d97706;font-weight:700'>{hub_a}</span>", unsafe_allow_html=True)
                c6.markdown(f"<span style='color:#dc2626;font-weight:700'>{hub_r}</span>",   unsafe_allow_html=True)
            first = False
        hub_avg = round(sum(hub_scores_all)/len(hub_scores_all),1) if hub_scores_all else 0.0
        grand_total += hub_total
        t0,t1,t2,t3,_,_,_ = st.columns([2.5,1.2,0.8,0.8,0.8,0.8,0.8])
        t0.markdown(f"**{hub} Total**"); t2.markdown(f"**{hub_total}**"); t3.markdown(f"**{hub_avg}**")
        st.divider()

    grand_avg = round(sum(grand_scores)/len(grand_scores),1) if grand_scores else 0.0
    g0,_,g2,g3,_,_,_ = st.columns([2.5,1.2,0.8,0.8,0.8,0.8,0.8])
    g0.markdown("**Grand Total**"); g2.markdown(f"**{grand_total}**"); g3.markdown(f"**{grand_avg}**")


# ══════════════════════════════════════════════════════
#  TAB 2: VENTURE CARDS
# ══════════════════════════════════════════════════════
with tab_ventures:
    st.title("🏢 Venture Cards")
    st.divider()

    # ── filters ───────────────────────────────────────
    fc1,fc2,fc3,fc4 = st.columns(4)
    search   = fc1.text_input("🔍 Search", key="vc_search")
    hub_list = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_list if cv(get_row(v),col_hub) != "—"))
    vp_list  = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_list if col_vp and cv(get_row(v),col_vp) != "—"))
    hub_vc   = fc2.selectbox("Hub",              hub_list, key="vc_hub")
    vp_vc    = fc3.selectbox("Venture Partner",  vp_list,  key="vc_vp")
    rag_vc   = fc4.selectbox("RAG Filter",
                ["All","🟢 Green","🟡 Amber","🔴 Red","⚪ ZERO"], key="vc_rag")

    filtered_v = ventures_list
    if search:    filtered_v = [v for v in filtered_v if search.lower() in v.lower()]
    if hub_vc  != "All": filtered_v = [v for v in filtered_v if cv(get_row(v),col_hub) == hub_vc]
    if vp_vc   != "All" and col_vp:
        filtered_v = [v for v in filtered_v if cv(get_row(v),col_vp) == vp_vc]
    if rag_vc  != "All":
        rag_val = rag_vc.split(" ",1)[1]
        filtered_v = [v for v in filtered_v if venture_rag.get(v,{}).get("overall_rag","ZERO") == rag_val]

    st.caption(f"{len(filtered_v)} ventures")
    st.markdown("<br>", unsafe_allow_html=True)

    for vname in filtered_v:
        row    = get_row(vname)
        hub    = cv(row, col_hub)
        sprint = cv(row, col_sprint)
        vp     = cv(row, col_vp) if col_vp else "—"
        rev    = cv(row, col_rev)
        tgt    = cv(row, col_tgt)
        notes  = cv(row, col_notes, default="")
        pct_raw= row[col_pct] if (row is not None and col_pct) else None
        bucket = get_stage_bucket(pct_raw)
        try:
            pct_num = float(str(pct_raw).replace("%","").strip())
            if pct_num <= 1: pct_num *= 100
        except: pct_num = 0

        rag_data  = venture_rag.get(vname,{})
        overall   = rag_data.get("overall_rag","ZERO")
        m_rag     = rag_data.get("momentum_rag","ZERO")
        i_rag     = rag_data.get("investment_rag","ZERO")
        m_reason  = rag_data.get("momentum_reason","—")
        i_reason  = rag_data.get("investment_reason","—")
        m_score   = rag_data.get("momentum_score",0)
        i_score   = rag_data.get("investment_score",0)
        signals   = venture_signals.get(vname, {"momentum":[],"investment":[]})
        sessions  = venture_feedback.get(vname, [])

        rag_emoji = RAG_EMOJI.get(overall,"⚪")

        with st.expander(
            f"{rag_emoji} **{vname}**  ·  {hub}  ·  RAG: {overall}"
        ):
            # ── Venture header banner ──────────────────
            st.markdown(
                f"<div class='venture-header'>"
                f"<div style='display:flex;align-items:flex-start;flex-wrap:wrap;gap:12px'>"
                f"<div>"
                f"<div style='font-size:1.2rem;font-weight:700'>{vname}</div>"
                f"<div style='font-size:0.83rem;opacity:0.85;margin-top:4px'>"
                f"📍 {hub} &nbsp;·&nbsp; 👤 {vp} &nbsp;·&nbsp; 🏃 {sprint}</div>"
                f"</div>"
                f"</div></div>",
                unsafe_allow_html=True
            )

            # ── Sub-tabs ───────────────────────────────
            tab_basic, tab_rag, tab_signals, tab_sessions, tab_ai = st.tabs([
                "📋 Basic Details",
                "🎯 RAG Scores",
                "✦ Signals",
                "🎙 Sessions & Feedback",
                "🤖 AI Insights"
            ])

            # ── TAB: BASIC DETAILS ─────────────────────
            with tab_basic:
                st.markdown("<br>", unsafe_allow_html=True)
                bd1, bd2, bd3 = st.columns([3, 1.5, 1.5])
                with bd1:
                    st.markdown(f"<div class='section-label'>Program Remarks</div>", unsafe_allow_html=True)
                    if notes:
                        st.info(notes)
                    else:
                        st.caption("No remarks in dashboard")
                with bd2:
                    st.markdown(f"<div class='section-label'>Financials</div>", unsafe_allow_html=True)
                    if rev != "—": st.metric("Revenue LY (Cr)", rev)
                    if tgt != "—": st.metric("3-Yr Target (Cr)", tgt)
                with bd3:
                    st.markdown(f"<div class='section-label'>Repository</div>", unsafe_allow_html=True)
                    proc_at = rag_data.get("processed_at","—")
                    st.caption(f"📅 Processed: {proc_at}")
                    nsig_m = len(signals.get("momentum",[]))
                    nsig_i = len(signals.get("investment",[]))
                    st.caption(f"📊 {nsig_m} momentum signals")
                    st.caption(f"💰 {nsig_i} investment signals")
                    nsess = len(sessions)
                    st.caption(f"🎙 {nsess} session(s)")

            # ── TAB: RAG SCORES ────────────────────────
            with tab_rag:
                st.markdown("<br>", unsafe_allow_html=True)
                rag1, rag2, rag3 = st.columns(3)
                with rag1:
                    st.markdown(
                        f"<div class='info-card' style='text-align:center'>"
                        f"<div class='section-label'>Overall RAG</div>"
                        f"<div style='font-size:2.5rem;margin:8px 0'>{RAG_EMOJI.get(overall,'⚪')}</div>"
                        f"<div>{rag_badge(overall)}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                def nps_card(label, rag, reason, nps, g, a, r, tot, score):
                    nps_color = {"Green":"#16a34a","Amber":"#d97706","Red":"#dc2626"}.get(rag,"#64748b")
                    bar_g = round(g/tot*100) if tot else 0
                    bar_a = round(a/tot*100) if tot else 0
                    bar_r = round(r/tot*100) if tot else 0
                    return (
                        f"<div class='info-card'>"
                        f"<div class='section-label'>{label}</div>"
                        f"<div style='display:flex;align-items:center;gap:12px;margin-bottom:8px'>"
                        f"{rag_badge(rag)}"
                        f"<span style='font-size:1.4rem;font-weight:800;color:{nps_color}'>Signal NPS {nps:+d}</span>"
                        f"</div>"
                        f"<div style='font-size:0.82rem;color:#475569;margin-bottom:10px'>{reason}</div>"
                        f"<div style='display:flex;gap:10px;font-size:0.78rem;margin-bottom:6px'>"
                        f"<span style='color:#16a34a;font-weight:600'>🟢 {g}</span>"
                        f"<span style='color:#d97706;font-weight:600'>🟡 {a}</span>"
                        f"<span style='color:#dc2626;font-weight:600'>🔴 {r}</span>"
                        f"<span style='color:#94a3b8'>/ {tot} signals</span>"
                        f"</div>"
                        f"<div style='display:flex;height:8px;border-radius:6px;overflow:hidden;background:#f1f5f9'>"
                        f"<div style='background:#16a34a;width:{bar_g}%'></div>"
                        f"<div style='background:#fbbf24;width:{bar_a}%'></div>"
                        f"<div style='background:#dc2626;width:{bar_r}%'></div>"
                        f"</div>"
                        f"<div style='margin-top:8px;font-size:0.78rem;color:#94a3b8'>Score: {score}/10</div>"
                        f"</div>"
                    )

                with rag2:
                    m_nps = rag_data.get("momentum_nps", 0)
                    m_g   = rag_data.get("momentum_green", 0)
                    m_a   = rag_data.get("momentum_amber", 0)
                    m_r   = rag_data.get("momentum_red",   0)
                    m_tot = rag_data.get("momentum_total", 0)
                    st.markdown(nps_card("Sprint Momentum", m_rag, m_reason,
                                         m_nps, m_g, m_a, m_r, m_tot, m_score),
                                unsafe_allow_html=True)
                with rag3:
                    i_nps = rag_data.get("investment_nps", 0)
                    i_g   = rag_data.get("investment_green", 0)
                    i_a   = rag_data.get("investment_amber", 0)
                    i_r   = rag_data.get("investment_red",   0)
                    i_tot = rag_data.get("investment_total", 0)
                    st.markdown(nps_card("Self Investment", i_rag, i_reason,
                                         i_nps, i_g, i_a, i_r, i_tot, i_score),
                                unsafe_allow_html=True)

                # Progress bars
                st.markdown("<br>", unsafe_allow_html=True)
                sc1, sc2 = st.columns(2)
                with sc1:
                    st.markdown("**Sprint Completion**")
                    st.progress(max(0.0, min(pct_num/100, 1.0)))
                    st.caption(f"{pct_num:.0f}% — {bucket}")

            # ── TAB: SIGNALS ───────────────────────────
            with tab_signals:
                st.markdown("<br>", unsafe_allow_html=True)

                m_sigs = signals.get("momentum",[])
                i_sigs = signals.get("investment",[])

                sig_col1, sig_col2 = st.columns(2)

                with sig_col1:
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:12px'>"
                        f"<div style='font-weight:700;font-size:1rem'>🏃 Sprint Momentum</div>"
                        f"<span style='background:#dcfce7;color:#166534;padding:2px 10px;"
                        f"border-radius:12px;font-size:0.78rem;font-weight:600'>"
                        f"{len(m_sigs)} signals</span></div>",
                        unsafe_allow_html=True
                    )
                    if m_sigs:
                        for sig in m_sigs:
                            sig_type = sig.get("type","")
                            evidence = sig.get("evidence","")
                            source   = sig.get("source","")
                            cat      = sig.get("category","GREEN")
                            cat_style = {
                                "GREEN": "background:#dcfce7;color:#166534",
                                "AMBER": "background:#fef9c3;color:#854d0e",
                                "RED":   "background:#fee2e2;color:#991b1b",
                            }.get(cat, "background:#f1f5f9;color:#475569")
                            cat_emoji = {"GREEN":"🟢","AMBER":"🟡","RED":"🔴"}.get(cat,"⚪")
                            st.markdown(
                                f"<div class='signal-row'>"
                                f"<span style='{cat_style};padding:2px 10px;border-radius:12px;"
                                f"font-size:0.77rem;font-weight:600;display:inline-block;margin:2px'>"
                                f"{cat_emoji} {sig_type}</span>"
                                f"<span class='sig-src'>📄 {source}</span>"
                                f"<div style='font-size:0.81rem;color:#475569;margin-top:4px'>{evidence}</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("No momentum signals found")

                with sig_col2:
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:12px'>"
                        f"<div style='font-weight:700;font-size:1rem'>💰 Self Investment</div>"
                        f"<span style='background:#e0e7ff;color:#3730a3;padding:2px 10px;"
                        f"border-radius:12px;font-size:0.78rem;font-weight:600'>"
                        f"{len(i_sigs)} signals</span></div>",
                        unsafe_allow_html=True
                    )
                    if i_sigs:
                        for sig in i_sigs:
                            sig_type = sig.get("type","")
                            evidence = sig.get("evidence","")
                            source   = sig.get("source","")
                            cat      = sig.get("category","GREEN")
                            cat_style = {
                                "GREEN": "background:#dcfce7;color:#166534",
                                "AMBER": "background:#fef9c3;color:#854d0e",
                                "RED":   "background:#fee2e2;color:#991b1b",
                            }.get(cat, "background:#f1f5f9;color:#475569")
                            cat_emoji = {"GREEN":"🟢","AMBER":"🟡","RED":"🔴"}.get(cat,"⚪")
                            st.markdown(
                                f"<div class='signal-row'>"
                                f"<span style='{cat_style};padding:2px 10px;border-radius:12px;"
                                f"font-size:0.77rem;font-weight:600;display:inline-block;margin:2px'>"
                                f"{cat_emoji} {sig_type}</span>"
                                f"<span class='sig-src'>📄 {source}</span>"
                                f"<div style='font-size:0.81rem;color:#475569;margin-top:4px'>{evidence}</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                    else:
                        st.caption("No investment signals found")

                # Sources used
                srcs = rag_data.get("sources_used",[])
                if srcs:
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown(
                        f"<div class='section-label'>Sources Read</div>"
                        f"<div>{''.join(f'<span class=sig-src>{s}</span>' for s in srcs)}</div>",
                        unsafe_allow_html=True
                    )

            # ── TAB: SESSIONS & FEEDBACK ───────────────
            with tab_sessions:
                st.markdown("<br>", unsafe_allow_html=True)

                # Call intelligence records for this venture (7-field structured data)
                ci_calls = []
                if call_intel_repo:
                    ci_v = call_intel_repo.get("ventures", {}).get(vname, {})
                    ci_calls = ci_v.get("calls", [])

                # All sessions from all sources — unified lookup
                tracker_sessions     = venture_tracker_sessions.get(vname, [])
                transcript_sessions  = sessions  # Claude-extracted from transcripts
                all_venture_sessions = all_sessions_by_venture.get(vname, [])

                total_sess = len(all_venture_sessions)

                # ── Call Intelligence section (7-field structured records) ──
                if ci_calls:
                    st.markdown(
                        f"<div style='font-weight:700;font-size:0.95rem;margin-bottom:4px'>"
                        f"📞 Call Intelligence — {len(ci_calls)} call record(s)</div>",
                        unsafe_allow_html=True
                    )
                    st.caption("Structured call data: objective, discussion topics, company updates, VP guidance, key decisions, action items, risks")
                    st.divider()

                    # ── Venture-level gap summary (all calls combined) ──
                    all_risks_raw = []
                    for c in ci_calls:
                        for r in c.get("risks_challenges", []):
                            if r and r not in all_risks_raw:
                                all_risks_raw.append(r)

                    if all_risks_raw:
                        v_sig  = signals_repo.get("venture_summary", {}).get(vname, {}) if signals_repo else {}
                        v_sp   = v_sig.get("overall_sprint_status", "")
                        v_in_sprint = v_sp in ["Active", "Slow"]
                        v_journey   = journey_repo.get("ventures", {}).get(vname, {}) if journey_repo else {}
                        # Use the sprint_gaps_context from the most recent call
                        all_sprint_ctx = " | ".join(
                            c.get("sprint_gaps_context","") for c in ci_calls
                            if c.get("sprint_gaps_context","") not in ["Not Available","","—"]
                        )
                        with st.expander(
                            f"⚠️ All gaps identified across {len(ci_calls)} call(s) — {len(all_risks_raw)} total",
                            expanded=True
                        ):
                            render_gaps_with_classification(
                                all_risks_raw, all_sprint_ctx, v_journey, v_in_sprint
                            )
                        st.markdown("<br>", unsafe_allow_html=True)

                    CALL_TYPE_COLORS = {
                        "Sprint call":         "#5b21b6",
                        "Panel call":          "#854f0b",
                        "VP call":             "#854f0b",
                        "Mid-sprint VP review":"#854f0b",
                    }

                    for ci, call in enumerate(ci_calls):
                        expert     = call.get("expert_name", "Not Available")
                        expert_role= call.get("expert_role", "Expert")
                        date       = call.get("session_date", "Not Available")
                        call_type  = call.get("call_type", "Sprint call")
                        fb_rating  = call.get("venture_feedback_rating")
                        fb_text    = call.get("venture_feedback_text", "Not Available")
                        type_color = CALL_TYPE_COLORS.get(call_type, "#334155")

                        rating_str = f"⭐ {fb_rating}" if fb_rating else ""

                        st.markdown(
                            f"<div style='background:#fff;border:.5px solid #e2e8f0;"
                            f"border-radius:10px;padding:12px 16px;margin-bottom:4px'>"
                            f"<div style='display:flex;justify-content:space-between;"
                            f"align-items:center;flex-wrap:wrap;gap:6px'>"
                            f"<div>"
                            f"<span style='font-weight:700;font-size:0.92rem'>Call {ci+1}: {expert}</span>"
                            f"<span style='background:#ede9fe;color:#5b21b6;padding:2px 8px;"
                            f"border-radius:8px;font-size:0.72rem;margin-left:8px'>{expert_role}</span>"
                            f"</div>"
                            f"<div style='display:flex;gap:6px;align-items:center;flex-wrap:wrap'>"
                            f"<span style='background:#f1f5f9;color:#475569;padding:2px 8px;"
                            f"border-radius:8px;font-size:0.75rem'>📅 {date}</span>"
                            f"<span style='background:{type_color}22;color:{type_color};"
                            f"padding:2px 8px;border-radius:8px;font-size:0.72rem;font-weight:600'>"
                            f"{call_type}</span>"
                            + (f"<span style='font-weight:700;color:#16a34a'>{rating_str}</span>"
                               if rating_str else "")
                            + f"</div></div></div>",
                            unsafe_allow_html=True
                        )

                        # 7-field detail panel
                        render_call_detail_fields(
                            call,
                            vname=vname,
                            journey_repo=journey_repo,
                            signals_repo=signals_repo
                        )

                        # Venture feedback
                        if fb_text and fb_text not in ["Not Available", "—"]:
                            st.markdown(
                                f"<div style='background:#f0fdf4;border:.5px solid #bbf7d0;"
                                f"border-radius:8px;padding:8px 12px;margin-top:6px;"
                                f"font-size:0.82rem;color:#166534'>"
                                f"<span style='font-weight:600'>Venture feedback:</span> {fb_text}</div>",
                                unsafe_allow_html=True
                            )

                        # Source file reference
                        st.markdown(
                            source_file_ref_for_call(call, vname=vname),
                            unsafe_allow_html=True
                        )

                        if ci < len(ci_calls) - 1:
                            st.markdown("<br>", unsafe_allow_html=True)

                    if tracker_sessions or transcript_sessions:
                        st.divider()
                        st.markdown(
                            "<div style='font-size:0.78rem;color:#94a3b8;margin-bottom:12px'>"
                            "Additional session data from tracker and transcript files below</div>",
                            unsafe_allow_html=True
                        )

                elif total_sess == 0:
                    st.info("No session data found for this venture.")
                    st.caption(
                        "Run Backend → Step 2 Section A (upload tracker files) and "
                        "Section B (transcript extraction) to populate this tab. "
                        "For 7-field call intelligence, also run Backend → Step 5."
                    )

                if total_sess > 0:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Sessions from Tracker",     len(tracker_sessions))
                    m2.metric("Sessions from Transcripts", len(transcript_sessions))
                    m3.metric("Total Sessions",            total_sess)
                    st.divider()

                    def session_card(si, session, source_label, source_color):
                        mentor  = session.get("mentor_name","Not Available")
                        date    = session.get("session_date","Not Available")
                        stype   = session.get("session_type","")
                        topics  = session.get("topics_discussed","Not Available")
                        outputs = session.get("key_outputs","Not Available")
                        summary = session.get("session_summary","Not Available")
                        fb_text = session.get("founder_feedback","Not Available")
                        rating  = session.get("founder_rating")
                        useful  = session.get("founder_usefulness","Not Available")
                        followup= session.get("followup_required","Not Available")
                        m_eng   = session.get("mentor_engagement","Not Available")
                        m_prep  = session.get("mentor_prepared","Not Available")

                        rating_str = f"⭐ {rating}" if rating else ""
                        followup_badge = (
                            "<span style='background:#fef9c3;color:#854d0e;padding:2px 8px;"
                            "border-radius:8px;font-size:0.75rem'>🔄 Follow-up</span>"
                            if str(followup).lower() == "yes" else ""
                        )
                        stype_badge = (
                            f"<span style='background:#e0e7ff;color:#3730a3;padding:2px 8px;"
                            f"border-radius:8px;font-size:0.75rem'>{stype}</span>"
                            if stype and stype != "Not Available" else ""
                        )

                        st.markdown(
                            f"<div class='session-card'>"
                            f"<div style='display:flex;justify-content:space-between;"
                            f"align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:12px'>"
                            f"<div>"
                            f"<span style='font-weight:700;font-size:1rem'>"
                            f"Session {si+1}: {mentor}</span>"
                            f"<span style='background:{source_color};color:white;padding:2px 8px;"
                            f"border-radius:8px;font-size:0.72rem;margin-left:8px'>"
                            f"{source_label}</span>"
                            f"</div>"
                            f"<div style='display:flex;gap:6px;align-items:center;flex-wrap:wrap'>"
                            f"<span style='background:#f1f5f9;color:#475569;padding:2px 8px;"
                            f"border-radius:8px;font-size:0.75rem'>📅 {date}</span>"
                            f"{stype_badge}{followup_badge}"
                            + (f"<span style='font-weight:700;color:#16a34a'>{rating_str}</span>"
                               if rating_str else "")
                            + f"</div></div>",
                            unsafe_allow_html=True
                        )

                        sc1, sc2, sc3 = st.columns(3)
                        with sc1:
                            st.markdown(
                                f"<div class='section-label'>Ask / Topics Discussed</div>"
                                f"<div style='font-size:0.83rem;color:#334155'>{topics}</div>",
                                unsafe_allow_html=True
                            )
                        with sc2:
                            st.markdown(
                                f"<div class='section-label'>Key Outputs / Next Steps</div>"
                                f"<div style='font-size:0.83rem;color:#334155'>{outputs}</div>",
                                unsafe_allow_html=True
                            )
                        with sc3:
                            fb_color = "#166534" if (fb_text and fb_text != "Not Available") else "#94a3b8"
                            st.markdown(
                                f"<div class='section-label'>Founder Feedback on Mentor</div>"
                                f"<div style='font-size:0.83rem;color:{fb_color}'>{fb_text}</div>",
                                unsafe_allow_html=True
                            )
                            if useful and useful != "Not Available":
                                st.markdown(
                                    f"<div style='font-size:0.75rem;color:#64748b;margin-top:4px'>"
                                    f"Usefulness: {useful}</div>",
                                    unsafe_allow_html=True
                                )

                        if summary and summary != "Not Available":
                            st.markdown(
                                f"<div class='section-label' style='margin-top:10px'>"
                                f"Meeting Summary</div>"
                                f"<div style='font-size:0.84rem;color:#475569;background:#f8fafc;"
                                f"padding:10px 14px;border-radius:8px;"
                                f"border-left:3px solid #6366f1'>{summary}</div>",
                                unsafe_allow_html=True
                            )

                        if m_eng and m_eng != "Not Available":
                            st.markdown(
                                f"<div style='font-size:0.78rem;color:#64748b;margin-top:8px'>"
                                f"Mentor view — Engagement: {m_eng} · Prepared: {m_prep}</div>",
                                unsafe_allow_html=True
                            )

                        st.markdown("</div>", unsafe_allow_html=True)

                    # ── Section 1: Tracker Sessions ──────────
                    if tracker_sessions:
                        st.markdown(
                            f"<div style='font-weight:700;font-size:0.95rem;margin-bottom:10px'>"
                            f"📊 From Session Tracker ({len(tracker_sessions)} sessions)</div>",
                            unsafe_allow_html=True
                        )
                        for si, sess in enumerate(tracker_sessions):
                            session_card(si, sess,
                                         source_label="Session Tracker",
                                         source_color="#6366f1")
                            if si < len(tracker_sessions)-1:
                                st.markdown("<br>", unsafe_allow_html=True)

                    # ── Section 2: Transcript Sessions ───────
                    if transcript_sessions:
                        if tracker_sessions:
                            st.divider()
                        st.markdown(
                            f"<div style='font-weight:700;font-size:0.95rem;margin-bottom:10px'>"
                            f"🎙 From Transcripts ({len(transcript_sessions)} sessions)</div>",
                            unsafe_allow_html=True
                        )
                        for si, sess in enumerate(transcript_sessions):
                            session_card(si, sess,
                                         source_label="Transcript",
                                         source_color="#0891b2")
                            if si < len(transcript_sessions)-1:
                                st.markdown("<br>", unsafe_allow_html=True)

            # ── TAB: AI INSIGHTS ───────────────────────
            with tab_ai:
                st.markdown("<br>", unsafe_allow_html=True)
                ai_key = f"ai_insight_{vname}"

                ai_c1, ai_c2 = st.columns([2,1])
                ai_question = ai_c1.text_input(
                    "Ask about this venture",
                    placeholder="e.g. What are the key risks? What should the VP focus on next?",
                    key=f"ai_q_{vname}"
                )
                ai_c2.markdown("<br>", unsafe_allow_html=True)
                run_ai = ai_c2.button("🤖 Ask Claude", key=f"ai_btn_{vname}", use_container_width=True)

                if run_ai and ai_question and client:
                    # Build context from both repos
                    m_list = "\n".join(f"  - {s['type']}: {s['evidence']}" for s in signals.get("momentum",[]))
                    i_list = "\n".join(f"  - {s['type']}: {s['evidence']}" for s in signals.get("investment",[]))
                    sess_summary = "\n".join(
                        f"  Session {idx+1}: Mentor={s.get('mentor_name','?')} | "
                        f"Date={s.get('session_date','?')} | "
                        f"Summary={s.get('session_summary','?')[:200]}"
                        for idx,s in enumerate(sessions)
                    )
                    context = f"""Venture: {vname} | Hub: {hub} | Sprint: {sprint} | VP: {vp}
Sprint Completion: {pct_num:.0f}% ({bucket})
Revenue LY: {rev} | 3-Yr Target: {tgt}
Notes: {notes}

OVERALL RAG: {overall}
Sprint Momentum RAG: {m_rag} — {m_reason}
Self Investment RAG: {i_rag} — {i_reason}

MOMENTUM SIGNALS:
{m_list or "None found"}

INVESTMENT SIGNALS:
{i_list or "None found"}

SESSIONS ({len(sessions)} total):
{sess_summary or "No session data"}

QUESTION: {ai_question}"""
                    with st.spinner("🤖 Analysing..."):
                        try:
                            resp = client.messages.create(
                                model="claude-sonnet-4-5", max_tokens=800,
                                messages=[{"role":"user","content":context}])
                            st.session_state[ai_key] = resp.content[0].text
                        except Exception as e:
                            st.error(f"AI error: {e}")

                if ai_key in st.session_state:
                    st.markdown(
                        f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                        f"border-radius:10px;padding:16px 20px;margin-top:12px;"
                        f"font-size:0.87rem;color:#1e293b;line-height:1.7'>"
                        f"{st.session_state[ai_key].replace(chr(10),'<br>')}</div>",
                        unsafe_allow_html=True
                    )
                elif not client:
                    st.warning("Add ANTHROPIC_API_KEY to use AI Insights.")


# ══════════════════════════════════════════════════════
#  TAB: INTELLIGENCE  (sub-tabs: Mentor Insights | Value Delivered | Call Intelligence | People Hired)
# ══════════════════════════════════════════════════════
with tab_intelligence:
    tab_mentors, tab_value, tab_call_intel, tab_people = st.tabs([
        "👥 Mentor Insights",
        "🌱 Value Delivered",
        "📞 Call Intelligence",
        "👷 People Hired",
    ])

with tab_mentors:
    st.title("👥 Mentor Insights")
    st.divider()

    # Use unified mentor insights — combines tracker + transcript sessions
    active_mentor_insights = mentor_insights_unified if mentor_insights_unified else mentor_insights

    if not active_mentor_insights:
        st.info(
            "No mentor data found. Run Backend → Step 2 Section A (tracker files) "
            "and Section B (transcripts) then upload the updated feedback_repository.json."
        )
    else:
        # ── Summary metrics ───────────────────────────
        all_sessions_flat = [
            s for m in active_mentor_insights.values() for s in m.get("sessions", [])
        ]
        total_mentors  = len(active_mentor_insights)
        total_sessions = len(all_sessions_flat)

        ratings_all = [
            s["founder_rating"]
            for s in all_sessions_flat
            if s.get("founder_rating") is not None
        ]
        avg_rating = round(sum(ratings_all)/len(ratings_all), 1) if ratings_all else None

        followup_count = sum(
            1 for s in all_sessions_flat
            if str(s.get("followup_required","")).lower() == "yes"
        )
        followup_rate = round(followup_count/total_sessions*100) if total_sessions else 0

        m1,m2,m3,m4 = st.columns(4)
        m1.metric("Total Mentors",   total_mentors)
        m2.metric("Total Sessions",  total_sessions)
        m3.metric("Avg Rating",      f"{avg_rating} ⭐" if avg_rating else "N/A")
        m4.metric("Follow-up Rate",  f"{followup_rate}%")

        st.divider()

        # ── Filters — Row 1 ──────────────────────────────
        f1, f2, f3, f4 = st.columns(4)
        mentor_search = f1.text_input("🔍 Search Mentor", key="mi_search")

        all_hubs = sorted(set(
            s.get("hub","—")
            for m in active_mentor_insights.values()
            for s in m.get("sessions",[])
            if s.get("hub","—") not in ["—","Not Available"]
        ))
        hub_mi = f2.selectbox("Hub", ["All"] + all_hubs, key="mi_hub")

        all_types = sorted(set(
            s.get("session_type","—")
            for m in active_mentor_insights.values()
            for s in m.get("sessions",[])
            if s.get("session_type","—") not in ["—","Not Available"]
        ))
        stype_mi = f3.selectbox("Session Type", ["All"] + all_types, key="mi_stype")

        rating_mi = f4.selectbox("Rating Filter",
            ["All","⭐⭐⭐⭐⭐ (4.5+)","⭐⭐⭐⭐ (4+)","⚠️ Flagged (≤3)"],
            key="mi_rating")

        # ── Filters — Row 2: Mentor Name + Overall Rating ─
        f5, f6, _, _ = st.columns(4)
        all_mentor_names = sorted(active_mentor_insights.keys())
        mentor_name_filter = f5.selectbox(
            "Mentor Name", ["All"] + all_mentor_names, key="mi_mentor_name"
        )

        # Compute avg rating per mentor for filter
        def mentor_avg_rating(mdata):
            ratings = [s.get("founder_rating") for s in mdata.get("sessions",[])
                       if s.get("founder_rating")]
            return round(sum(ratings)/len(ratings),1) if ratings else None

        overall_rating_filter = f6.selectbox(
            "Overall Mentor Rating",
            ["All","⭐ 4.5+","⭐ 4.0+","⭐ 3.5+","⚠️ Below 3.5"],
            key="mi_overall_rating"
        )

        # ── Build filtered mentor list ─────────────────
        def session_passes_filters(s):
            if hub_mi   != "All" and s.get("hub","—") != hub_mi: return False
            if stype_mi != "All" and s.get("session_type","—") != stype_mi: return False
            if rating_mi != "All":
                r  = s.get("founder_rating")
                if rating_mi == "⭐⭐⭐⭐⭐ (4.5+)" and (not r or r < 4.5): return False
                if rating_mi == "⭐⭐⭐⭐ (4+)"   and (not r or r < 4.0): return False
                if rating_mi == "⚠️ Flagged (≤3)" and (not r or r > 3.0): return False
            return True

        filtered_mentors = {}
        for mn, mdata in active_mentor_insights.items():
            # Mentor name filter
            if mentor_search and mentor_search.lower() not in mn.lower(): continue
            if mentor_name_filter != "All" and mn != mentor_name_filter: continue
            # Overall rating filter
            if overall_rating_filter != "All":
                avg_r = mentor_avg_rating(mdata)
                if overall_rating_filter == "⭐ 4.5+"       and (not avg_r or avg_r < 4.5): continue
                if overall_rating_filter == "⭐ 4.0+"       and (not avg_r or avg_r < 4.0): continue
                if overall_rating_filter == "⭐ 3.5+"       and (not avg_r or avg_r < 3.5): continue
                if overall_rating_filter == "⚠️ Below 3.5" and (not avg_r or avg_r >= 3.5): continue
            # Session-level filters
            filtered_sessions = [s for s in mdata.get("sessions",[]) if session_passes_filters(s)]
            if filtered_sessions:
                filtered_mentors[mn] = {**mdata, "sessions": filtered_sessions}

        st.caption(f"{len(filtered_mentors)} mentors · "
                   f"{sum(len(m['sessions']) for m in filtered_mentors.values())} sessions")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Per-mentor expanders ───────────────────────
        for mn in sorted(filtered_mentors.keys()):
            mdata    = filtered_mentors[mn]
            sessions = mdata["sessions"]
            avg_r    = mdata.get("avg_rating")
            ventures = mdata.get("ventures_worked", [])
            nsess    = len(sessions)

            rating_str = f"⭐ {avg_r}" if avg_r else "No rating"
            ventures_str = ", ".join(ventures[:4]) + ("..." if len(ventures)>4 else "")

            with st.expander(
                f"👤 **{mn}**  ·  {nsess} session(s)  ·  {rating_str}  ·  {ventures_str}"
            ):
                # Mentor summary row
                ms1, ms2, ms3, ms4 = st.columns(4)
                ms1.metric("Sessions",  nsess)
                ms2.metric("Avg Rating", f"{avg_r} ⭐" if avg_r else "N/A")
                ms3.metric("Ventures",   len(ventures))
                followups = sum(1 for s in sessions
                                if str(s.get("followup_required","")).lower() == "yes")
                ms4.metric("Follow-ups Required", followups)

                st.markdown(
                    f"<div style='font-size:0.8rem;color:#64748b;margin-bottom:16px'>"
                    f"Ventures: {', '.join(ventures)}</div>",
                    unsafe_allow_html=True
                )
                st.divider()

                # ── Sessions table ─────────────────────────
                def na(v): return "—" if not v or str(v) in ["Not Available","nan","None",""] else str(v)

                rows_html = ""
                for session in sessions:
                    venture    = na(session.get("venture_name"))
                    date       = na(session.get("meeting_date"))
                    ask        = na(session.get("ask"))
                    stype      = na(session.get("session_type"))
                    summary    = na(session.get("meeting_summary"))
                    next_steps = na(session.get("next_steps"))
                    followup   = na(session.get("followup_required"))
                    t_rating   = session.get("founder_rating")
                    t_feedback = na(session.get("founder_feedback"))
                    eff_rating = t_rating
                    verbatim   = t_feedback
                    usefulness = na(session.get("founder_usefulness"))
                    mfb_eng    = na(session.get("mentor_engagement"))
                    engagement = mfb_eng

                    try:
                        rc = ("#16a34a" if float(eff_rating)>=4.0 else
                              "#d97706" if float(eff_rating)>=3.0 else "#dc2626") if eff_rating else "#94a3b8"
                    except: rc = "#94a3b8"

                    rating_str  = (f"<span style='font-weight:700;color:{rc}'>&#11088; {eff_rating}</span>"
                                   if eff_rating else "<span style='color:#94a3b8'>—</span>")
                    followup_str= ("<span style='background:#fef9c3;color:#854d0e;padding:1px 6px;"
                                   "border-radius:6px;font-size:0.72rem'>&#128260; Yes</span>"
                                   if str(followup).lower()=="yes"
                                   else "<span style='color:#94a3b8'>—</span>")
                    stype_str   = (f"<span style='background:#e0e7ff;color:#3730a3;padding:1px 7px;"
                                   f"border-radius:6px;font-size:0.72rem'>{stype}</span>"
                                   if stype!="—" else "—")

                    fb_cell = ""
                    if verbatim != "—":
                        fb_cell += (f"<div style='color:#166534;font-style:italic;margin-bottom:3px'>"
                                    f"&ldquo;{verbatim[:120]}{'...' if len(verbatim)>120 else ''}&rdquo;</div>")
                    if usefulness != "—":
                        fb_cell += f"<div style='font-size:0.72rem;color:#64748b'>Useful: {usefulness}</div>"
                    if not fb_cell: fb_cell = "<span style='color:#94a3b8'>—</span>"

                    mentor_cell = (f"<div style='font-size:0.78rem'>{engagement[:80]}</div>"
                                   if engagement != "—" else "<span style='color:#94a3b8'>—</span>")

                    rows_html += f"""<tr style='border-bottom:1px solid #f1f5f9'>
                        <td style='padding:8px 12px;font-weight:600;white-space:nowrap;color:#1e293b'>{venture}</td>
                        <td style='padding:8px 12px;white-space:nowrap;color:#64748b'>{date}</td>
                        <td style='padding:8px 12px'>{stype_str}</td>
                        <td style='padding:8px 12px;max-width:180px'>{ask}</td>
                        <td style='padding:8px 12px;max-width:220px;color:#475569'>{summary}</td>
                        <td style='padding:8px 12px;max-width:160px;color:#475569'>{next_steps}</td>
                        <td style='padding:8px 12px;max-width:220px'>{fb_cell}</td>
                        <td style='padding:8px 12px;text-align:center'>{rating_str}</td>
                        <td style='padding:8px 12px;text-align:center'>{followup_str}</td>
                        <td style='padding:8px 12px;max-width:140px'>{mentor_cell}</td>
                        <td style='padding:8px 12px' title='Find this file in SharePoint'>{
                            source_file_ref(call_type=stype, vname=venture)
                        }</td>
                    </tr>"""

                st.markdown(f"""
                <div style='overflow-x:auto;margin-top:8px'>
                <table style='width:100%;border-collapse:collapse;font-size:0.81rem;font-family:Inter,sans-serif'>
                    <thead><tr style='background:#f1f5f9;border-bottom:2px solid #e2e8f0'>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600;white-space:nowrap'>Venture</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600;white-space:nowrap'>Date</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Type</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Ask / Topic</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Meeting Summary</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Next Steps</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Founder Feedback</th>
                        <th style='padding:8px 12px;text-align:center;color:#475569;font-weight:600'>Rating</th>
                        <th style='padding:8px 12px;text-align:center;color:#475569;font-weight:600'>Follow-up</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Mentor View</th>
                        <th style='padding:8px 12px;text-align:left;color:#475569;font-weight:600'>Source file</th>
                    </tr></thead>
                    <tbody>{rows_html}</tbody>
                </table></div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════
#  INTELLIGENCE SUB-TAB: VALUE DELIVERED
# ══════════════════════════════════════════════════════
with tab_value:
    st.title("🌱 Value Delivered to Beneficiaries")
    st.caption("Portfolio-level synthesis · Powered by Claude AI · Based on signals, sessions and journey documents")
    st.divider()

    # ── Check data availability ───────────────────────
    has_signals  = bool(signals_repo and signals_repo.get("venture_summary"))
    has_feedback = bool(feedback_repo and feedback_repo.get("mentor_insights"))

    if not has_signals and not has_feedback:
        st.warning(
            "No repository data found. Generate and upload both "
            "`signals_repository.json` and `feedback_repository.json` first."
        )
    else:
        # ── Static metrics (no API needed) ────────────
        total_ventures  = len(signals_repo.get("venture_summary",{})) if signals_repo else 0
        total_mentors   = len(feedback_repo.get("mentor_insights",{})) if feedback_repo else 0
        total_sessions  = sum(
            m.get("total_sessions",0)
            for m in (feedback_repo.get("mentor_insights",{}) or {}).values()
        ) if feedback_repo else 0

        # Jobs target from company basics
        total_jobs_target = 0
        if company_basics:
            for vdata in company_basics.values():
                try:
                    j = float(str(vdata.get("incremental_jobs_3yr","0")).replace(",",""))
                    total_jobs_target += j
                except: pass

        # Avg rating
        all_ratings = [
            s.get("founder_rating")
            for m in (feedback_repo.get("mentor_insights",{}) or {}).values()
            for s in m.get("sessions",[])
            if s.get("founder_rating")
        ] if feedback_repo else []
        avg_rating = round(sum(all_ratings)/len(all_ratings),1) if all_ratings else None

        # ── Metric row ────────────────────────────────
        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        mc1.metric("Ventures Supported",    total_ventures)
        mc2.metric("Mentors Engaged",       total_mentors)
        mc3.metric("Sessions Delivered",    total_sessions)
        mc4.metric("Avg Session Rating",    f"{avg_rating} ⭐" if avg_rating else "N/A")
        mc5.metric("Jobs Target (3 Yr)",
                   f"{int(total_jobs_target):,}" if total_jobs_target else "N/A")

        st.divider()

        # ── Generate insights button ──────────────────
        SYNTHESIS_KEY = "value_synthesis_result"

        if not client:
            st.warning("Add ANTHROPIC_API_KEY to generate AI insights.")
        else:
            gen_col, _ = st.columns([2, 5])
            with gen_col:
                if st.button("✨ Generate Value Insights", key="gen_value",
                             use_container_width=True):
                    with st.spinner("Claude is synthesising portfolio value insights..."):
                        from processor import synthesise_value_delivered
                        result, err = synthesise_value_delivered(
                            client, signals_repo, feedback_repo, company_basics or {}
                        )
                        if err:
                            st.error(f"Error: {err}")
                        else:
                            st.session_state[SYNTHESIS_KEY] = result
                            st.rerun()

            if SYNTHESIS_KEY in st.session_state:
                if st.button("🔄 Regenerate", key="regen_value"):
                    del st.session_state[SYNTHESIS_KEY]
                    st.rerun()

        # ── Display synthesis ─────────────────────────
        if SYNTHESIS_KEY in st.session_state:
            result = st.session_state[SYNTHESIS_KEY]

            # ── Section 1: Jobs Generated 2026 ────────
            st.markdown("### 👷 Jobs Generated in 2026")
            jobs_data = result.get("jobs_2026", {})
            jobs_total    = jobs_data.get("total_estimated", 0)
            jobs_conf     = jobs_data.get("confidence", "—")
            jobs_ev_count = jobs_data.get("evidence_count", 0)
            jobs_examples = jobs_data.get("top_examples", [])

            conf_color = {"High":"#16a34a","Medium":"#d97706","Low":"#dc2626"}.get(
                jobs_conf, "#64748b")

            jc1, jc2, jc3 = st.columns([1.5, 1.5, 5])
            jc1.metric("Estimated Jobs Created", f"{jobs_total:,}")
            jc2.metric("Ventures with Hiring Evidence", jobs_ev_count)
            with jc3:
                st.markdown(
                    f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                    f"border-radius:10px;padding:14px 18px;margin-top:4px'>"
                    f"<div style='font-size:0.75rem;font-weight:600;color:#94a3b8;"
                    f"text-transform:uppercase;margin-bottom:8px'>Confidence</div>"
                    f"<span style='background:{conf_color}22;color:{conf_color};"
                    f"padding:3px 12px;border-radius:20px;font-weight:700;"
                    f"font-size:0.83rem'>{jobs_conf}</span>"
                    f"<div style='font-size:0.78rem;color:#64748b;margin-top:8px'>"
                    f"Based on {jobs_ev_count} ventures with explicit hiring mentions "
                    f"in signals and session notes</div></div>",
                    unsafe_allow_html=True
                )

            if jobs_examples:
                with st.expander(f"📋 Hiring evidence ({len(jobs_examples)} ventures)"):
                    for ex in jobs_examples:
                        st.markdown(
                            f"<div style='padding:6px 0;border-bottom:1px solid #f1f5f9'>"
                            f"<span style='font-weight:600;color:#1e293b'>"
                            f"{ex.get('venture','')}</span>"
                            f"<div style='font-size:0.82rem;color:#475569;margin-top:2px'>"
                            f"{ex.get('jobs_detail','')}</div></div>",
                            unsafe_allow_html=True
                        )

            st.divider()

            # ── Section 2: Gap Categories Resolved ────
            st.markdown("### 🔧 Gap Categories Resolved")
            st.caption("Based on session summaries across all mentoring sessions")

            gap_categories = result.get("gap_categories", [])
            if gap_categories:
                # Sort by session count descending
                gap_categories = sorted(gap_categories,
                                        key=lambda x: x.get("session_count",0),
                                        reverse=True)
                max_count = max(g.get("session_count",0) for g in gap_categories) or 1

                gc_cols = st.columns(2)
                for idx, gap in enumerate(gap_categories):
                    cat   = gap.get("category","—")
                    count = gap.get("session_count", 0)
                    desc  = gap.get("description","")
                    examples = gap.get("example_actions", [])
                    pct   = round(count / max_count * 100)

                    with gc_cols[idx % 2]:
                        ex_html = "".join(
                            f"<div style='font-size:0.76rem;color:#475569;"
                            f"margin-top:3px'>• {e}</div>"
                            for e in examples[:3]
                        )
                        st.markdown(
                            f"<div style='background:#ffffff;border:1px solid #e2e8f0;"
                            f"border-radius:10px;padding:14px 16px;margin-bottom:12px'>"
                            f"<div style='display:flex;justify-content:space-between;"
                            f"align-items:center;margin-bottom:6px'>"
                            f"<span style='font-weight:700;font-size:0.9rem;color:#1e293b'>"
                            f"{cat}</span>"
                            f"<span style='background:#e0e7ff;color:#3730a3;padding:2px 10px;"
                            f"border-radius:12px;font-size:0.78rem;font-weight:700'>"
                            f"{count} sessions</span></div>"
                            f"<div style='background:#f1f5f9;border-radius:4px;height:6px;"
                            f"margin-bottom:8px;overflow:hidden'>"
                            f"<div style='background:#6366f1;height:100%;width:{pct}%'>"
                            f"</div></div>"
                            f"<div style='font-size:0.82rem;color:#475569;margin-bottom:6px'>"
                            f"{desc}</div>"
                            f"{ex_html}</div>",
                            unsafe_allow_html=True
                        )

            st.divider()

            # ── Section 3: Actionable Business Direction ──
            st.markdown("### 🧭 Actionable Business Direction")
            st.caption("Synthesised from session notes and next steps across all ventures")

            directions_data = result.get("actionable_directions", {})
            summary     = directions_data.get("summary","")
            top_dirs    = directions_data.get("top_directions", [])

            if summary:
                st.markdown(
                    f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                    f"border-left:4px solid #6366f1;border-radius:8px;"
                    f"padding:16px 20px;font-size:0.87rem;color:#334155;"
                    f"line-height:1.8;margin-bottom:16px'>{summary}</div>",
                    unsafe_allow_html=True
                )

            if top_dirs:
                st.markdown("**Top Actionable Directions:**")
                for i, direction in enumerate(top_dirs, 1):
                    st.markdown(
                        f"<div style='display:flex;gap:12px;align-items:flex-start;"
                        f"padding:10px 0;border-bottom:1px solid #f1f5f9'>"
                        f"<span style='background:#6366f1;color:white;border-radius:50%;"
                        f"width:24px;height:24px;display:flex;align-items:center;"
                        f"justify-content:center;font-size:0.75rem;font-weight:700;"
                        f"flex-shrink:0'>{i}</span>"
                        f"<div style='font-size:0.85rem;color:#334155;padding-top:3px'>"
                        f"{direction}</div></div>",
                        unsafe_allow_html=True
                    )

            st.divider()

            # ── Section 4: Problems Solved ─────────────
            st.markdown("### ✅ Problems Solved (Category Level)")
            st.caption("Categories of business challenges addressed across the portfolio")

            problems = result.get("problems_solved", [])
            if problems:
                problems = sorted(problems, key=lambda x: x.get("count",0), reverse=True)

                # Badge-style display
                badges_html = ""
                for p in problems:
                    cat   = p.get("category","")
                    count = p.get("count", 0)
                    desc  = p.get("description","")
                    badges_html += (
                        f"<div style='background:#ffffff;border:1px solid #e2e8f0;"
                        f"border-radius:10px;padding:12px 16px;margin-bottom:10px;"
                        f"display:flex;align-items:flex-start;gap:12px'>"
                        f"<div style='background:#dcfce7;color:#166534;padding:4px 12px;"
                        f"border-radius:20px;font-weight:700;font-size:0.82rem;"
                        f"white-space:nowrap;flex-shrink:0'>{count} ventures</div>"
                        f"<div><div style='font-weight:600;color:#1e293b;font-size:0.87rem'>"
                        f"{cat}</div>"
                        f"<div style='font-size:0.8rem;color:#64748b;margin-top:2px'>"
                        f"{desc}</div></div></div>"
                    )
                pc1, pc2 = st.columns(2)
                half = len(problems) // 2 + len(problems) % 2
                with pc1:
                    for p in problems[:half]:
                        cat   = p.get("category","")
                        count = p.get("count", 0)
                        desc  = p.get("description","")
                        st.markdown(
                            f"<div style='background:#ffffff;border:1px solid #e2e8f0;"
                            f"border-radius:10px;padding:12px 16px;margin-bottom:10px;"
                            f"display:flex;align-items:flex-start;gap:12px'>"
                            f"<div style='background:#dcfce7;color:#166534;padding:4px 12px;"
                            f"border-radius:20px;font-weight:700;font-size:0.82rem;"
                            f"white-space:nowrap;flex-shrink:0'>{count} ventures</div>"
                            f"<div><div style='font-weight:600;color:#1e293b;font-size:0.87rem'>"
                            f"{cat}</div>"
                            f"<div style='font-size:0.8rem;color:#64748b;margin-top:2px'>"
                            f"{desc}</div></div></div>",
                            unsafe_allow_html=True
                        )
                with pc2:
                    for p in problems[half:]:
                        cat   = p.get("category","")
                        count = p.get("count", 0)
                        desc  = p.get("description","")
                        st.markdown(
                            f"<div style='background:#ffffff;border:1px solid #e2e8f0;"
                            f"border-radius:10px;padding:12px 16px;margin-bottom:10px;"
                            f"display:flex;align-items:flex-start;gap:12px'>"
                            f"<div style='background:#dcfce7;color:#166534;padding:4px 12px;"
                            f"border-radius:20px;font-weight:700;font-size:0.82rem;"
                            f"white-space:nowrap;flex-shrink:0'>{count} ventures</div>"
                            f"<div><div style='font-weight:600;color:#1e293b;font-size:0.87rem'>"
                            f"{cat}</div>"
                            f"<div style='font-size:0.8rem;color:#64748b;margin-top:2px'>"
                            f"{desc}</div></div></div>",
                            unsafe_allow_html=True
                        )

        else:
            st.info(
                "Click **✨ Generate Value Insights** above to synthesise portfolio value "
                "from signals, session notes and journey documents."
            )


# ══════════════════════════════════════════════════════
#  INTELLIGENCE SUB-TAB: CALL INTELLIGENCE
# ══════════════════════════════════════════════════════
with tab_call_intel:
    st.title("📞 Call Intelligence")
    st.caption(
        "Per-call structured records — objective, discussion topics, company updates, "
        "VP guidance, key decisions, action items, risks/challenges · "
        "generated in Backend Step 5 · read from call_intelligence_repository.json"
    )
    st.divider()

    if not call_intel_repo:
        st.warning(
            "Call Intelligence repository not found. "
            "Run **Backend → Step 5 — Call Intelligence** to generate it, "
            "then upload `call_intelligence_repository.json` to SharePoint."
        )
        st.info(f"Expected path: `{CALL_INTEL_REPO_PATH}`")
    else:
        ci_ventures = call_intel_repo.get("ventures", {})
        gen_at      = call_intel_repo.get("generated_at", "—")
        total_calls = call_intel_repo.get("total_calls", 0)

        st.caption(
            f"Generated: {gen_at} · {len(ci_ventures)} ventures · {total_calls} call records"
        )

        # ── Filters ──────────────────────────────────
        cif1, cif2, cif3, cif4 = st.columns([2, 1.5, 1.5, 1.5])
        with cif1:
            ci_search = st.text_input("🔍 Search venture", key="ci_search_v2")
        with cif2:
            ci_hub_f = st.selectbox("Hub", ["All"] + sorted(set(
                v.get("hub", "—") for v in ci_ventures.values() if v.get("hub","—") != "—"
            )), key="ci_hub_f")
        with cif3:
            ci_call_type_f = st.selectbox("Call type", [
                "All", "Sprint call", "Panel call", "VP call", "Mid-sprint VP review"
            ], key="ci_call_type_f")
        with cif4:
            ci_sprint_f = st.selectbox("Sprint", ["All"] + sorted(set(
                v.get("sprint", "—") for v in ci_ventures.values() if v.get("sprint","—") not in ["—",""]
            )), key="ci_sprint_f")

        # ── Portfolio KPIs ────────────────────────────
        all_ci_calls = []
        for vn, vd in ci_ventures.items():
            for c in vd.get("calls", []):
                c["_venture_name"] = vn
                c["_hub"]          = vd.get("hub","—")
                c["_sprint"]       = vd.get("sprint","—")
                all_ci_calls.append(c)

        # Apply filters
        def ci_call_passes(c):
            vn = c.get("_venture_name","")
            if ci_search and ci_search.lower() not in vn.lower(): return False
            if ci_hub_f != "All" and c.get("_hub","") != ci_hub_f: return False
            if ci_call_type_f != "All" and c.get("call_type","") != ci_call_type_f: return False
            if ci_sprint_f != "All" and c.get("_sprint","") != ci_sprint_f: return False
            return True

        filtered_calls = [c for c in all_ci_calls if ci_call_passes(c)]

        # Gap analysis from risks_challenges across all calls
        in_sprint_count = len([c for c in filtered_calls
                                if c.get("sprint_gaps_context","Not Available") != "Not Available"])
        new_gap_count   = sum(
            len(c.get("risks_challenges",[])) for c in filtered_calls
        )
        ventures_with_risks = len(set(
            c["_venture_name"] for c in filtered_calls
            if c.get("risks_challenges")
        ))
        rated_calls = [c for c in filtered_calls
                       if c.get("venture_feedback_rating") not in [None, "None"]]
        avg_rating = (
            round(sum(float(c["venture_feedback_rating"]) for c in rated_calls) / len(rated_calls), 1)
            if rated_calls else None
        )

        pm1, pm2, pm3, pm4, pm5, pm6 = st.columns(6)
        pm1.metric("Ventures analysed",      len(ci_ventures))
        pm2.metric("Call records",           len(filtered_calls))
        pm3.metric("Risks identified",       new_gap_count)
        pm4.metric("Ventures with risks",    ventures_with_risks)
        pm5.metric("Avg founder rating",
                   f"{avg_rating}/5" if avg_rating else "—")
        pm6.metric("Calls with sprint gaps", in_sprint_count)

        st.divider()

        # ── Summary: Call types breakdown ─────────────
        type_counts = {}
        for c in filtered_calls:
            ct = c.get("call_type","Other")
            type_counts[ct] = type_counts.get(ct, 0) + 1

        # ── Risk categories from all calls ────────────
        risk_category_count = {}
        for c in filtered_calls:
            for r in c.get("risks_challenges",[]):
                r_lower = r.lower()
                if any(w in r_lower for w in ["finance","capital","credit","funding"]):
                    cat = "Finance"
                elif any(w in r_lower for w in ["hire","staff","team","hr","recruit"]):
                    cat = "HR / People"
                elif any(w in r_lower for w in ["distribution","freight","logistics","delivery"]):
                    cat = "Distribution"
                elif any(w in r_lower for w in ["compliance","legal","regulation","license"]):
                    cat = "Legal / Compliance"
                elif any(w in r_lower for w in ["tech","software","digital","platform"]):
                    cat = "Technology"
                else:
                    cat = "Operations"
                risk_category_count[cat] = risk_category_count.get(cat, 0) + 1

        if type_counts or risk_category_count:
            sum_c1, sum_c2 = st.columns(2)
            with sum_c1:
                st.markdown(
                    "<div class='card-title'>Call type breakdown</div>",
                    unsafe_allow_html=True
                )
                max_tc = max(type_counts.values()) if type_counts else 1
                TYPE_COLORS = {
                    "Sprint call": "#5b21b6", "Panel call": "#854f0b",
                    "VP call": "#854f0b", "Mid-sprint VP review": "#854f0b"
                }
                for ct, cnt in sorted(type_counts.items(), key=lambda x: -x[1]):
                    pct = round(cnt / max_tc * 100)
                    col = TYPE_COLORS.get(ct, "#334155")
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:10px;margin-bottom:8px'>"
                        f"<div style='min-width:160px;font-size:0.83rem;color:#1e293b'>{ct}</div>"
                        f"<div style='flex:1;background:#e2e8f0;border-radius:4px;height:7px;overflow:hidden'>"
                        f"<div style='background:{col};height:100%;width:{pct}%'></div></div>"
                        f"<div style='min-width:28px;font-size:0.8rem;font-weight:700;color:{col}'>{cnt}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            with sum_c2:
                if risk_category_count:
                    st.markdown(
                        "<div class='card-title'>Risks by category (across all calls)</div>",
                        unsafe_allow_html=True
                    )
                    max_rc = max(risk_category_count.values())
                    for cat, cnt in sorted(risk_category_count.items(), key=lambda x: -x[1]):
                        pct = round(cnt / max_rc * 100)
                        st.markdown(
                            f"<div style='display:flex;align-items:center;gap:10px;margin-bottom:8px'>"
                            f"<div style='min-width:160px;font-size:0.83rem;color:#1e293b'>{cat}</div>"
                            f"<div style='flex:1;background:#e2e8f0;border-radius:4px;height:7px;overflow:hidden'>"
                            f"<div style='background:#dc2626;height:100%;width:{pct}%'></div></div>"
                            f"<div style='min-width:28px;font-size:0.8rem;font-weight:700;color:#dc2626'>{cnt}</div>"
                            f"</div>",
                            unsafe_allow_html=True
                        )

        st.divider()

        # ── Per-venture call records ───────────────────
        st.markdown(
            f"<div style='font-weight:700;font-size:0.95rem;margin-bottom:10px'>"
            f"🏢 Call records by venture — {len(filtered_calls)} calls shown</div>",
            unsafe_allow_html=True
        )

        # Group filtered calls by venture
        calls_by_venture = {}
        for c in filtered_calls:
            vn = c["_venture_name"]
            calls_by_venture.setdefault(vn, []).append(c)

        CALL_TYPE_COLORS = {
            "Sprint call": "#5b21b6", "Panel call": "#854f0b",
            "VP call": "#854f0b", "Mid-sprint VP review": "#854f0b"
        }

        for vn, v_calls in calls_by_venture.items():
            hub_str    = v_calls[0].get("_hub","—")
            sprint_str = v_calls[0].get("_sprint","—")
            n_risks    = sum(len(c.get("risks_challenges",[])) for c in v_calls)

            with st.expander(
                f"**{vn}** · {hub_str} · {sprint_str} · "
                f"{len(v_calls)} call(s) · {n_risks} risk(s)"
            ):
                for ci, call in enumerate(v_calls):
                    expert      = call.get("expert_name","Not Available")
                    expert_role = call.get("expert_role","Expert")
                    date        = call.get("session_date","Not Available")
                    call_type   = call.get("call_type","Sprint call")
                    fb_rating   = call.get("venture_feedback_rating")
                    fb_text     = call.get("venture_feedback_text","Not Available")
                    type_color  = CALL_TYPE_COLORS.get(call_type, "#334155")
                    rating_str  = f"⭐ {fb_rating}" if fb_rating else ""

                    # Call header
                    st.markdown(
                        f"<div style='background:#fff;border:.5px solid #e2e8f0;"
                        f"border-radius:10px;padding:10px 14px;margin-bottom:4px;"
                        f"margin-top:{12 if ci>0 else 0}px'>"
                        f"<div style='display:flex;justify-content:space-between;"
                        f"align-items:center;flex-wrap:wrap;gap:6px'>"
                        f"<div>"
                        f"<span style='font-weight:700;font-size:0.92rem'>"
                        f"Call {ci+1}: {expert}</span>"
                        f"<span style='background:#ede9fe;color:#5b21b6;padding:2px 8px;"
                        f"border-radius:8px;font-size:0.72rem;margin-left:8px'>{expert_role}</span>"
                        f"</div>"
                        f"<div style='display:flex;gap:6px;flex-wrap:wrap;align-items:center'>"
                        f"<span style='background:#f1f5f9;color:#475569;padding:2px 8px;"
                        f"border-radius:8px;font-size:0.75rem'>📅 {date}</span>"
                        f"<span style='background:{type_color}22;color:{type_color};"
                        f"padding:2px 8px;border-radius:8px;font-size:0.72rem;font-weight:600'>"
                        f"{call_type}</span>"
                        + (f"<span style='font-weight:700;color:#16a34a'>{rating_str}</span>"
                           if rating_str else "")
                        + "</div></div></div>",
                        unsafe_allow_html=True
                    )

                    # 7-field detail
                    render_call_detail_fields(
                        call,
                        vname=vn,
                        journey_repo=journey_repo,
                        signals_repo=signals_repo
                    )

                    # Venture feedback if available
                    if fb_text and fb_text not in ["Not Available","—"]:
                        st.markdown(
                            f"<div style='background:#f0fdf4;border:.5px solid #bbf7d0;"
                            f"border-radius:8px;padding:8px 12px;margin-top:6px;"
                            f"font-size:0.82rem;color:#166534'>"
                            f"<span style='font-weight:600'>Venture feedback:</span> {fb_text}</div>",
                            unsafe_allow_html=True
                        )

                    # Source file reference
                    st.markdown(
                        source_file_ref_for_call(call, vname=vn),
                        unsafe_allow_html=True
                    )

                    # Sprint gaps context
                    gaps_ctx = call.get("sprint_gaps_context","")
                    if gaps_ctx and gaps_ctx not in ["Not Available","—",""]:
                        with st.expander("🎯 Sprint gaps context (from Sprint Plan + Journey Doc)"):
                            st.caption(gaps_ctx)


# ══════════════════════════════════════════════════════
#  INTELLIGENCE SUB-TAB: PEOPLE HIRED
# ══════════════════════════════════════════════════════
with tab_people:
    st.title("👷 People Hired")
    st.caption("Resources hired, 6-month hiring plans, and role breakdown by function — read from all documents")
    st.divider()

    if not people_repo or not people_repo.get("ventures"):
        st.warning(
            "People Hired data not found. "
            "Run Backend → Step 4 and upload people_repository.json to SharePoint."
        )
    else:
        people_ventures = people_repo.get("ventures", {})

        # ── Summary metrics ────────────────────────────
        total_hired   = sum(v.get("resources_hired_count",0) for v in people_ventures.values())
        total_planned = sum(v.get("hiring_plan_6mo_count",0) for v in people_ventures.values())
        ventures_with_hires = sum(1 for v in people_ventures.values()
                                  if v.get("resources_hired_count",0) > 0)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Resources Hired",     total_hired)
        m2.metric("Total 6-Mo Hiring Plan",    total_planned)
        m3.metric("Ventures with Hires",       ventures_with_hires)
        m4.metric("Ventures Tracked",          len(people_ventures))

        st.divider()

        # ── Filters ────────────────────────────────────
        f1, f2 = st.columns(2)
        ph_search = f1.text_input("🔍 Search Venture", key="ph_search")
        ph_filter = f2.selectbox(
            "Filter",
            ["All", "Has Hires", "Has 6-Mo Plan", "No Data"],
            key="ph_filter"
        )

        def na(v):
            if not v or str(v).strip() in ["None","nan","NaT","","Data Not Available"]:
                return "Data Not Available"
            return str(v).strip()

        FUNCTIONS = [("gtm","GTM"),("product","Product"),("operations","Operations"),
                     ("supply_chain","Supply Chain"),("hr","HR"),("finance","Finance")]

        rows = []
        for vn, vdata in people_ventures.items():
            if ph_search and ph_search.lower() not in vn.lower(): continue
            hired_count   = vdata.get("resources_hired_count", 0)
            planned_count = (vdata.get("planned_2026_count",0) +
                             vdata.get("planned_2027_count",0) +
                             vdata.get("hiring_plan_6mo_count",0))
            if ph_filter == "Has Hires"     and hired_count   == 0: continue
            if ph_filter == "Has 6-Mo Plan" and planned_count == 0: continue
            if ph_filter == "No Data"       and (hired_count>0 or planned_count>0): continue
            rows.append({
                "vn":           vn,
                "current_emp":  na(vdata.get("current_employee_count")),
                "hired_count":  hired_count,
                "hired_desc":   na(vdata.get("resources_hired_description")),
                "h26_count":    vdata.get("hired_2026_count",0),
                "h26_desc":     na(vdata.get("hired_2026_description")),
                "h27_count":    vdata.get("hired_2027_count",0),
                "h27_desc":     na(vdata.get("hired_2027_description")),
                "p26_count":    vdata.get("planned_2026_count",0),
                "p26_desc":     na(vdata.get("planned_2026_description")),
                "p27_count":    vdata.get("planned_2027_count",0),
                "p27_desc":     na(vdata.get("planned_2027_description")),
                "role_breakdown": vdata.get("role_breakdown",{}),
            })

        rows.sort(key=lambda r: r["hired_count"]+r["h26_count"]+r["h27_count"], reverse=True)
        st.caption(f"{len(rows)} ventures")
        st.markdown("<br>", unsafe_allow_html=True)

        TH = ("style='padding:8px 12px;text-align:left;color:#475569;font-weight:600;"
              "font-size:0.76rem;background:#f1f5f9;white-space:nowrap;"
              "border-bottom:2px solid #e2e8f0;border-right:1px solid #e2e8f0'")
        THG= ("style='padding:8px 12px;text-align:center;color:#166534;font-weight:700;"
              "font-size:0.76rem;background:#dcfce7;white-space:nowrap;"
              "border-bottom:2px solid #bbf7d0;border-right:1px solid #e2e8f0'")
        THA= ("style='padding:8px 12px;text-align:center;color:#854d0e;font-weight:700;"
              "font-size:0.76rem;background:#fef9c3;white-space:nowrap;"
              "border-bottom:2px solid #fde68a;border-right:1px solid #e2e8f0'")
        TD = ("style='padding:8px 12px;font-size:0.79rem;color:#334155;"
              "vertical-align:top;border-bottom:1px solid #f1f5f9;"
              "border-right:1px solid #f1f5f9'")
        TDW= ("style='padding:8px 12px;font-size:0.78rem;color:#475569;"
              "vertical-align:top;border-bottom:1px solid #f1f5f9;"
              "border-right:1px solid #f1f5f9;max-width:180px'")

        def count_badge(n, color):
            return (f"<span style='font-weight:800;font-size:1rem;color:{color}'>{n}</span>"
                    if n > 0 else "<span style='color:#94a3b8'>0</span>")

        rows_html = ""
        for r in rows:
            rb = r["role_breakdown"]
            rb_html = "<div style='font-size:0.73rem;line-height:1.7'>"
            for key, label in FUNCTIONS:
                val = na(rb.get(key))
                color = "#94a3b8" if val == "Data Not Available" else "#334155"
                rb_html += (f"<div><strong style='color:#1e293b'>{label}:</strong> "
                            f"<span style='color:{color}'>{val}</span></div>")
            rb_html += "</div>"

            rows_html += (
                f"<tr>"
                f"<td {TD}><strong>{r['vn']}</strong></td>"
                f"<td {TD} style='text-align:center'><strong>{r['current_emp']}</strong></td>"
                f"<td {TD} style='text-align:center'>{count_badge(r['hired_count'],'#16a34a')}</td>"
                f"<td {TDW}>{r['hired_desc']}</td>"
                f"<td {TD} style='text-align:center'>{count_badge(r['h26_count'],'#16a34a')}</td>"
                f"<td {TDW}>{r['h26_desc']}</td>"
                f"<td {TD} style='text-align:center'>{count_badge(r['h27_count'],'#16a34a')}</td>"
                f"<td {TDW}>{r['h27_desc']}</td>"
                f"<td {TD} style='text-align:center'>{count_badge(r['p26_count'],'#d97706')}</td>"
                f"<td {TDW}>{r['p26_desc']}</td>"
                f"<td {TD} style='text-align:center'>{count_badge(r['p27_count'],'#d97706')}</td>"
                f"<td {TDW}>{r['p27_desc']}</td>"
                f"<td {TDW}>{rb_html}</td>"
                f"</tr>"
            )

        st.markdown(
            f"<div style='overflow-x:auto;max-height:650px;border:1px solid #e2e8f0;"
            f"border-radius:10px;overflow-y:auto'>"
            f"<table style='width:100%;border-collapse:collapse;font-family:Inter,sans-serif'>"
            f"<thead>"
            f"<tr>"
            f"<th {TH} rowspan='2'>Company</th>"
            f"<th {TH} rowspan='2'>Current<br>Employees</th>"
            f"<th {TH} colspan='2' style='text-align:center'>Total Hired</th>"
            f"<th {THG} colspan='2'>Hired in 2026</th>"
            f"<th {THG} colspan='2'>Hired in 2027</th>"
            f"<th {THA} colspan='2'>Planned 2026</th>"
            f"<th {THA} colspan='2'>Planned 2027</th>"
            f"<th {TH} rowspan='2'>Role Breakdown<br>by Function</th>"
            f"</tr><tr>"
            f"<th {TH}>Count</th><th {TH}>Description</th>"
            f"<th {THG}>Count</th><th {THG}>Description</th>"
            f"<th {THG}>Count</th><th {THG}>Description</th>"
            f"<th {THA}>Count</th><th {THA}>Description</th>"
            f"<th {THA}>Count</th><th {THA}>Description</th>"
            f"</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            f"</table></div>",
            unsafe_allow_html=True
        )

        st.markdown("<br>", unsafe_allow_html=True)
        def esc_ph(v): return f'"{str(v).replace(chr(34), chr(39))}"' 
        csv_lines = [
            "Company,Current Employees,Total Hired,Hired Desc,"
            "Hired 2026,Hired 2026 Desc,Hired 2027,Hired 2027 Desc,"
            "Planned 2026,Planned 2026 Desc,Planned 2027,Planned 2027 Desc,"
            "GTM,Product,Operations,Supply Chain,HR,Finance"
        ]
        for r in rows:
            rb = r["role_breakdown"]
            csv_lines.append(
                f"{esc_ph(r['vn'])},{esc_ph(r['current_emp'])},"
                f"{esc_ph(r['hired_count'])},{esc_ph(r['hired_desc'])},"
                f"{esc_ph(r['h26_count'])},{esc_ph(r['h26_desc'])},"
                f"{esc_ph(r['h27_count'])},{esc_ph(r['h27_desc'])},"
                f"{esc_ph(r['p26_count'])},{esc_ph(r['p26_desc'])},"
                f"{esc_ph(r['p27_count'])},{esc_ph(r['p27_desc'])},"
                f"{esc_ph(na(rb.get('gtm')))},{esc_ph(na(rb.get('product')))},"                f"{esc_ph(na(rb.get('operations')))},{esc_ph(na(rb.get('supply_chain')))},"                f"{esc_ph(na(rb.get('hr')))},{esc_ph(na(rb.get('finance')))}"
            )
        st.download_button(
            "⬇️ Download as CSV",
            data="\n".join(csv_lines),
            file_name="people_hired.csv",
            mime="text/csv",
            key="ph_download"
        )

