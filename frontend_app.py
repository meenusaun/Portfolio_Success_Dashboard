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

SP_FOLDER        = "Documents/04. Advisors/2026/Portfolio Success Dashboard"
COMMON_FOLDER    = f"{SP_FOLDER}/Common Documents"
REPO_FOLDER      = f"{COMMON_FOLDER}/Knowledge Repository"
DASHBOARD_FILE   = "0. Journey_Accelerate_Portfolio Dashboard.xlsx"

SIGNALS_REPO_PATH  = f"{REPO_FOLDER}/signals_repository.json"
FEEDBACK_REPO_PATH = f"{REPO_FOLDER}/feedback_repository.json"

st.set_page_config(
    page_title="Portfolio Success Intelligence",
    page_icon="🚀",
    layout="wide"
)

# ── CSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f5f7fa; }
.main .block-container { padding: 1.5rem 2.5rem; max-width: 1600px; }
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e2e8f0; }
section[data-testid="stSidebar"] * { color: #1e293b !important; }
div[data-testid="stMetricValue"] { color: #1e293b !important; font-weight: 700; font-size: 1.6rem; }
div[data-testid="stMetricLabel"] { font-size: 0.82rem; color: #64748b; }
.stProgress > div > div { background: #6366f1 !important; }
h1,h2,h3,h4 { color: #1e293b !important; }
.stExpander { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; margin-bottom: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.04); }
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] { padding: 8px 18px; border-radius: 8px; }

/* RAG badges */
.rag-green  { background:#dcfce7; color:#166534; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.83rem; display:inline-block; }
.rag-amber  { background:#fef9c3; color:#854d0e; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.83rem; display:inline-block; }
.rag-red    { background:#fee2e2; color:#991b1b; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.83rem; display:inline-block; }
.rag-zero   { background:#f1f5f9; color:#64748b; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.83rem; display:inline-block; }

/* Signal badges */
.sig-pos { background:#dcfce7; color:#166534; padding:2px 10px; border-radius:12px; font-size:0.77rem; font-weight:600; display:inline-block; margin:2px; }
.sig-neg { background:#fee2e2; color:#991b1b; padding:2px 10px; border-radius:12px; font-size:0.77rem; font-weight:600; display:inline-block; margin:2px; }
.sig-src { background:#e0e7ff; color:#3730a3; padding:2px 8px; border-radius:10px; font-size:0.72rem; display:inline-block; margin:2px; }

/* Card sections */
.info-card { background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:20px 24px; margin-bottom:16px; }
.section-label { font-size:0.72rem; font-weight:700; color:#94a3b8; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:6px; }
.venture-header { background: linear-gradient(135deg,#6366f1 0%,#8b5cf6 100%); color:white; border-radius:12px; padding:20px 24px; margin-bottom:20px; }
.session-card { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:16px 20px; margin-bottom:12px; }
.signal-row { padding:8px 0; border-bottom:1px solid #f1f5f9; }
.signal-row:last-child { border-bottom: none; }
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
        if p >= 76:  return "76–99%"
        if p >= 51:  return "51–75%"
        if p >= 26:  return "26–50%"
        return "0–25%"
    except: return "Unknown"

# ── sidebar ────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    use_sp = st.toggle("☁️ Use SharePoint", value=bool(ENV_CLIENT_ID))
    if use_sp and ENV_CLIENT_ID:
        st.success("✅ SharePoint configured")
    elif use_sp:
        st.warning("Add Azure credentials to secrets")
    api_key = ENV_API_KEY or st.text_input("Anthropic API Key", type="password")
    st.markdown("---")
    st.caption("NEN Accelerate · Portfolio Intelligence\nResources Network Team")

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

with st.spinner("Loading Knowledge Repositories from SharePoint..."):
    signals_repo, sig_err  = load_signals_repo(sp_id, use_sp)
    feedback_repo, fb_err  = load_feedback_repo(sp_id, use_sp)

# ── parse repositories into lookups ───────────────────
venture_rag      = {}  # {vname: {overall_rag, momentum_rag, investment_rag, ...}}
venture_signals  = {}  # {vname: {momentum: [...], investment: [...]}}
venture_feedback = {}  # {vname: [session_dicts]}

if signals_repo:
    vsummary = signals_repo.get("venture_summary", {})
    for vn, vdata in vsummary.items():
        venture_rag[vn]     = vdata
        venture_signals[vn] = vdata.get("signals", {"momentum":[], "investment":[]})

if feedback_repo:
    for vn, vdata in feedback_repo.get("ventures", {}).items():
        venture_feedback[vn] = vdata.get("sessions", [])

repo_loaded = signals_repo is not None or feedback_repo is not None

# ── header + repo status ───────────────────────────────
hc1, hc2 = st.columns([3,1])
with hc1:
    st.markdown("""
    <div style='display:flex;align-items:center;gap:12px;margin-bottom:4px'>
        <span style='font-size:2rem'>🚀</span>
        <div>
            <div style='font-size:1.4rem;font-weight:700;color:#1e293b'>
                Portfolio Success Intelligence</div>
            <div style='font-size:0.82rem;color:#64748b'>
                NEN Accelerate · Resources Network · Knowledge Repository</div>
        </div>
    </div>""", unsafe_allow_html=True)
with hc2:
    if signals_repo:
        gen_at = signals_repo.get("generated_at","—")
        st.success(f"✅ Signals repo loaded\n{gen_at}")
    else:
        st.warning(f"⚠️ Signals repo not found\n{sig_err or ''}")
    if feedback_repo:
        st.success(f"✅ Feedback repo loaded")
    else:
        st.info(f"ℹ️ Feedback repo not found\n{fb_err or ''}")

st.divider()

if not repo_loaded:
    st.error("❌ No Knowledge Repository found on SharePoint. Run the Backend app to generate and upload the repository files first.")
    st.info(f"Expected paths:\n- `{SIGNALS_REPO_PATH}`\n- `{FEEDBACK_REPO_PATH}`")
    st.stop()

# ── main tabs ──────────────────────────────────────────
tab_overview, tab_ventures = st.tabs([
    "📊  Portfolio Overview",
    "🏢  Venture Cards"
])

# ══════════════════════════════════════════════════════
#  TAB 1: PORTFOLIO OVERVIEW
# ══════════════════════════════════════════════════════
with tab_overview:
    st.title("📊 Portfolio Overview")

    # ── filters row ───────────────────────────────────
    f1, f2, f3, f4 = st.columns(4)
    hub_opts    = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_list if cv(get_row(v),col_hub) != "—"))
    vp_opts     = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_list if col_vp and cv(get_row(v),col_vp) != "—"))
    stage_opts  = ["All","0–25%","26–50%","51–75%","76–99%","100%","Unknown"]
    rag_opts    = ["All","🟢 Green","🟡 Amber","🔴 Red","⚪ ZERO"]

    hub_f   = f1.selectbox("Hub",              hub_opts,   key="ov_hub")
    vp_f    = f2.selectbox("Venture Partner",  vp_opts,    key="ov_vp")
    stage_f = f3.selectbox("Sprint Stage",     stage_opts, key="ov_stage")
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
    if stage_f != "All": filtered = [v for v in filtered if v["bucket"] == stage_f]
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
        st.subheader("Sprint Stage Distribution")
        st.markdown("<br>", unsafe_allow_html=True)
        stage_counts = {}
        for v in filtered:
            stage_counts[v["bucket"]] = stage_counts.get(v["bucket"],0) + 1
        for stage in ["100%","76–99%","51–75%","26–50%","0–25%","Unknown"]:
            cnt = stage_counts.get(stage,0)
            pct = round(cnt/total*100) if total else 0
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;margin-bottom:4px'>"
                f"<span style='font-weight:500'>{stage}</span>"
                f"<span style='color:#64748b'>{cnt} &nbsp;({pct}%)</span></div>",
                unsafe_allow_html=True)
            st.progress(pct/100)
            st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)

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
            f"{rag_emoji} **{vname}**  ·  {hub}  ·  {bucket}  ·  RAG: {overall}"
        ):
            # ── Venture header banner ──────────────────
            st.markdown(
                f"<div class='venture-header'>"
                f"<div style='display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px'>"
                f"<div>"
                f"<div style='font-size:1.2rem;font-weight:700'>{vname}</div>"
                f"<div style='font-size:0.83rem;opacity:0.85;margin-top:4px'>"
                f"📍 {hub} &nbsp;·&nbsp; 👤 {vp} &nbsp;·&nbsp; 🏃 {sprint}</div>"
                f"</div>"
                f"<div style='text-align:right'>"
                f"<div style='font-size:0.72rem;opacity:0.75;text-transform:uppercase;letter-spacing:0.06em'>Sprint Completion</div>"
                f"<div style='font-size:1.8rem;font-weight:700'>{pct_num:.0f}%</div>"
                f"<div style='font-size:0.78rem;opacity:0.75'>{bucket}</div>"
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
                with rag2:
                    st.markdown(
                        f"<div class='info-card'>"
                        f"<div class='section-label'>Sprint Momentum</div>"
                        f"<div style='margin-bottom:8px'>{rag_badge(m_rag)}</div>"
                        f"<div style='font-size:0.82rem;color:#475569'>{m_reason}</div>"
                        f"<div style='margin-top:10px;font-size:0.78rem;color:#94a3b8'>Score: {m_score}/10</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with rag3:
                    st.markdown(
                        f"<div class='info-card'>"
                        f"<div class='section-label'>Self Investment</div>"
                        f"<div style='margin-bottom:8px'>{rag_badge(i_rag)}</div>"
                        f"<div style='font-size:0.82rem;color:#475569'>{i_reason}</div>"
                        f"<div style='margin-top:10px;font-size:0.78rem;color:#94a3b8'>Score: {i_score}/10</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

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
                            # Simple positive/negative heuristic
                            neg_words = ["disengaged","drop out","no progress","not responding","red","delay"]
                            is_neg = any(w in (sig_type+evidence).lower() for w in neg_words)
                            badge_cls = "sig-neg" if is_neg else "sig-pos"
                            st.markdown(
                                f"<div class='signal-row'>"
                                f"<span class='{badge_cls}'>{sig_type}</span>"
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
                            neg_words = ["not ready","unsure","no investment","not committed"]
                            is_neg = any(w in (sig_type+evidence).lower() for w in neg_words)
                            badge_cls = "sig-neg" if is_neg else "sig-pos"
                            st.markdown(
                                f"<div class='signal-row'>"
                                f"<span class='{badge_cls}'>{sig_type}</span>"
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

                if not sessions:
                    st.info("No session or feedback data found for this venture in the Knowledge Repository.")
                    st.caption("Run the Backend app → Step 2 to generate the Feedback Repository.")
                else:
                    st.markdown(
                        f"<div style='margin-bottom:16px'>"
                        f"<span style='font-weight:700'>{len(sessions)} session(s) found</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                    for si, session in enumerate(sessions):
                        mentor   = session.get("mentor_name","Not Available")
                        date     = session.get("session_date","Not Available")
                        topics   = session.get("topics_discussed","Not Available")
                        outputs  = session.get("key_outputs","Not Available")
                        fb_text  = session.get("founder_feedback","Not Available")
                        summary  = session.get("session_summary","Not Available")
                        srcs     = session.get("sources_used",[])

                        st.markdown(
                            f"<div class='session-card'>"
                            f"<div style='display:flex;justify-content:space-between;align-items:center;"
                            f"margin-bottom:14px;flex-wrap:wrap;gap:8px'>"
                            f"<div style='font-weight:700;font-size:1rem'>Session {si+1}: {mentor}</div>"
                            f"<div style='display:flex;gap:8px;flex-wrap:wrap'>"
                            f"<span style='background:#f1f5f9;color:#475569;padding:3px 10px;"
                            f"border-radius:10px;font-size:0.78rem'>📅 {date}</span>"
                            + (f"<span style='background:#e0e7ff;color:#3730a3;padding:3px 10px;"
                               f"border-radius:10px;font-size:0.78rem'>{'  '.join(srcs)}</span>"
                               if srcs else "")
                            + f"</div></div>",
                            unsafe_allow_html=True
                        )

                        # 3-column layout inside session card
                        sc1, sc2, sc3 = st.columns(3)
                        with sc1:
                            st.markdown(f"<div class='section-label'>Topics Discussed</div>", unsafe_allow_html=True)
                            st.markdown(
                                f"<div style='font-size:0.83rem;color:#334155'>{topics}</div>",
                                unsafe_allow_html=True
                            )
                        with sc2:
                            st.markdown(f"<div class='section-label'>Key Outputs / Actions</div>", unsafe_allow_html=True)
                            st.markdown(
                                f"<div style='font-size:0.83rem;color:#334155'>{outputs}</div>",
                                unsafe_allow_html=True
                            )
                        with sc3:
                            st.markdown(f"<div class='section-label'>Founder Feedback on Mentor</div>", unsafe_allow_html=True)
                            fb_color = "#166534" if fb_text != "Not Available" else "#94a3b8"
                            st.markdown(
                                f"<div style='font-size:0.83rem;color:{fb_color}'>{fb_text}</div>",
                                unsafe_allow_html=True
                            )

                        st.markdown(f"<div class='section-label' style='margin-top:12px'>Session Summary</div>", unsafe_allow_html=True)
                        st.markdown(
                            f"<div style='font-size:0.84rem;color:#475569;background:#f8fafc;"
                            f"padding:10px 14px;border-radius:8px;border-left:3px solid #6366f1'>"
                            f"{summary}</div>",
                            unsafe_allow_html=True
                        )
                        st.markdown("</div>", unsafe_allow_html=True)
                        if si < len(sessions)-1:
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
