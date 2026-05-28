import streamlit as st
import pandas as pd
import os, json, re, tempfile, shutil, io
from pathlib import Path
from anthropic import Anthropic

# ── load .env ─────────────────────────────────────────
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
ENV_ROOT_PATH     = os.environ.get("PORTFOLIO_ROOT_PATH", "")
ENV_CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "")
ENV_TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "")
ENV_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# SharePoint constants
SP_USER    = "meenakshi.singh@wadhwanifoundation.org"
SP_FOLDER  = "Documents/04. Advisors/2026/Portfolio Success Dashboard"
DASHBOARD_FILE = "0. Journey_Accelerate_Portfolio Dashboard.xlsx"

st.set_page_config(page_title="Portfolio Success Intelligence", page_icon="🚀", layout="wide")

# ── password protection ───────────────────────────────
ENV_PASSWORD = os.environ.get("APP_PASSWORD", "nen2026")

def check_password():
    """Simple password gate."""
    if st.session_state.get("authenticated"):
        return True

    # Center the login form
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
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        password = st.text_input("Enter access password", 
                                  type="password",
                                  placeholder="Enter password...",
                                  label_visibility="collapsed")
        
        col_a, col_b, col_c = st.columns([1,2,1])
        with col_b:
            login_btn = st.button("🔐  Login", use_container_width=True)

        if login_btn:
            if password == ENV_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption("For access contact: meenakshi.singh@wadhwanifoundation.org")
    
    return False

if not check_password():
    st.stop()

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f5f7fa; }
.main .block-container { padding: 1.5rem 2rem; max-width: 1400px; }
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e2e8f0; }
section[data-testid="stSidebar"] * { color: #1e293b !important; }
div[data-testid="stMetricValue"] { color: #1e293b !important; font-weight: 700; }
.stProgress > div > div { background: #6366f1 !important; }
h1,h2,h3 { color: #1e293b !important; }
.stExpander { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

# ── constants ─────────────────────────────────────────
STAGE_COLORS  = {"0–25%":"#ef4444","26–50%":"#f59e0b","51–75%":"#3b82f6","76–99%":"#8b5cf6","100%":"#10b981"}
ACCENT_COLORS = ["#6366f1","#10b981","#f59e0b","#8b5cf6","#ef4444","#3b82f6","#ec4899","#14b8a6"]
SIGNAL_KEYWORDS = {
    "Hired":               ["hired","new hire","appointed","joined","onboard","recruit","new employee","new manager","export manager","dedicated"],
    "Investment / Spend":  ["invested","purchased","bought","commissioned","capex","new machine","new facility","production facility","spent"],
    "New Stream / Market": ["new market","new segment","new product","new geography","entered","launching","private label","supermarket","new category"],
    "Self-Funded Sprint":  ["own money","paid for","self-funded","sprint 2","sprint 3","another sprint","continuing sprint"],
    "Revenue / Order Win": ["revenue","order","contract","deal","new client","purchase order","po received","first export"],
}

def detect_signals(text):
    if not text: return []
    tl = text.lower()
    out = []
    for stype, kws in SIGNAL_KEYWORDS.items():
        for kw in kws:
            if kw in tl:
                out.append({"type": stype, "keyword": kw})
                break
    return out

def get_stage_bucket(pct):
    try:
        if pct is None or str(pct).strip() in ["","nan","None","-"]: return "Unknown"
        p = float(str(pct).replace("%","").replace(",","").strip())
        if p <= 1.0: p = p * 100
        if p >= 100: return "100%"
        if p >= 76:  return "76–99%"
        if p >= 51:  return "51–75%"
        if p >= 26:  return "26–50%"
        return "0–25%"
    except: return "Unknown"

def safe_copy(src):
    dst = os.path.join(tempfile.gettempdir(), "nen_" + os.path.basename(src))
    shutil.copy2(src, dst)
    return dst

def find_col(df, patterns):
    for c in df.columns:
        for p in patterns:
            if p.lower() in str(c).lower():
                return c
    return None

def extract_text_local(fpath):
    ext = Path(fpath).suffix.lower()
    try:
        fp = safe_copy(fpath)
        if ext in [".xlsx",".xls"]:
            xl = pd.ExcelFile(fp)
            parts = [pd.read_excel(fp, sheet_name=s, header=None).fillna("").astype(str).to_string() for s in xl.sheet_names]
            return "\n".join(parts)[:6000]
        elif ext == ".docx":
            from docx import Document
            return "\n".join(p.text for p in Document(fp).paragraphs)[:6000]
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(fp) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)[:6000]
    except Exception as e:
        return f"[Error: {e}]"
    return ""

def extract_text_bytes(content_bytes, filename):
    """Extract text from file bytes (for SharePoint downloads)."""
    ext = Path(filename).suffix.lower()
    try:
        tmp = os.path.join(tempfile.gettempdir(), f"nen_sp_{filename}")
        with open(tmp, "wb") as f:
            f.write(content_bytes)
        return extract_text_local(tmp)
    except Exception as e:
        return f"[Error: {e}]"

# ── SharePoint connection ─────────────────────────────
@st.cache_resource(show_spinner=False)
def get_sp_reader(client_id, tenant_id, client_secret):
    try:
        from sharepoint_reader import SharePointReader
        return SharePointReader(client_id, tenant_id, client_secret), None
    except Exception as e:
        return None, str(e)

# ── sidebar ───────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")

    # Data source toggle
    use_sharepoint = st.toggle("☁️ Use SharePoint/OneDrive", 
                                value=bool(ENV_CLIENT_ID),
                                help="Connect to SharePoint using Azure AD credentials")

    if use_sharepoint:
        if ENV_CLIENT_ID:
            st.success("✅ Azure AD credentials configured")
        else:
            st.warning("Add AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET to .env")
        root_path = None
    else:
        if ENV_ROOT_PATH:
            root_path = ENV_ROOT_PATH
            st.success("✅ Local folder configured")
            st.text_input("Portfolio Root Folder", value=ENV_ROOT_PATH, disabled=True)
        else:
            root_path = st.text_input("Portfolio Root Folder", 
                                       placeholder=r"C:\Users\...\Portfolio")

    if ENV_API_KEY:
        api_key = ENV_API_KEY
        st.success("✅ API key configured")
    else:
        api_key = st.text_input("Anthropic API Key", type="password")

    st.markdown("---")
    view = st.radio("📂 View", 
                     ["📊 Portfolio Overview", "🏢 Venture Cards", "💡 Success Signals"],
                     label_visibility="collapsed")
    view = view.split(" ", 1)[1]

    st.markdown("---")
    st.caption("NEN Accelerate · Portfolio Intelligence\nResources Network Team")

# ── load data ─────────────────────────────────────────
if not use_sharepoint and not root_path:
    st.title("🚀 Portfolio Success Intelligence")
    st.info("👈 Configure your data source in the sidebar to get started.")
    st.stop()

@st.cache_data(show_spinner=False, ttl=300)
def load_data_sharepoint(client_id, tenant_id, client_secret):
    """Load portfolio data from SharePoint."""
    from sharepoint_reader import SharePointReader
    sp = SharePointReader(client_id, tenant_id, client_secret)
    file_path = f"{SP_FOLDER}/{DASHBOARD_FILE}"
    try:
        content = sp.download_file(file_path)
        tmp = os.path.join(tempfile.gettempdir(), "nen_dashboard.xlsx")
        with open(tmp, "wb") as f:
            f.write(content)
        company_df = pd.read_excel(tmp, sheet_name="Company",       header=2)
        try:
            basics_df  = pd.read_excel(tmp, sheet_name="Company Basics", header=2)
        except:
            basics_df = None
        return company_df, basics_df, None
    except Exception as e:
        return None, None, str(e)

@st.cache_data(show_spinner=False, ttl=300)
def load_data_local(root_path):
    """Load portfolio data from local drive."""
    dash = None
    hc = r"C:\Users\MeenakshiSingh\OneDrive - National Entrepreneurship Network\04. Advisors\2026\Portfolio Success Dashboard\0. Journey_Accelerate_Portfolio Dashboard.xlsx"
    if os.path.exists(hc):
        dash = hc
    else:
        for f in os.listdir(root_path):
            if "Journey_Accelerate_Portfolio" in f and f.endswith(".xlsx") and "~$" not in f:
                dash = os.path.join(root_path, f); break
    if not dash:
        return None, None, "Dashboard file not found."
    try:
        fp = safe_copy(dash)
        company_df = pd.read_excel(fp, sheet_name="Company",       header=2)
        try:
            basics_df  = pd.read_excel(fp, sheet_name="Company Basics", header=2)
        except:
            basics_df = None
        return company_df, basics_df, None
    except Exception as e:
        return None, None, str(e)

# Load data
with st.spinner("Loading portfolio data..."):
    if use_sharepoint and ENV_CLIENT_ID:
        try:
            from sharepoint_reader import SharePointReader
            sp_test = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            
            # Debug: browse all possible paths
            with st.expander("🔍 Debug: Browse SharePoint folders", expanded=False):
                paths_to_try = [
                    "Documents",
                    "Documents/04. Advisors/2026/Portfolio Success Dashboard",
                ]
                for try_path in paths_to_try:
                    try:
                        items = sp_test.list_folder(try_path)
                        label = try_path if try_path else "ROOT"
                        st.write(f"**Contents of {label}/:**")
                        for item in items:
                            icon = "📁" if "folder" in item else "📄"
                            st.write(f"{icon} {item['name']}")
                        st.divider()
                    except Exception as e:
                        st.write(f"❌ `{try_path}`: {e}")
                        st.divider()

            company_df, basics_df, err = load_data_sharepoint(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        except Exception as auth_err:
            st.error(f"❌ {str(auth_err)}")
            st.stop()
    else:
        company_df, basics_df, err = load_data_local(root_path or "")

if err:
    st.error(f"❌ {err}")
    st.stop()
if company_df is None:
    st.error("Could not load data.")
    st.stop()

# ── detect columns ────────────────────────────────────
company_df.columns = [str(c).strip() for c in company_df.columns]
col_name  = find_col(company_df, ["company name"])
col_hub   = find_col(company_df, ["hub / state","hub","state"])
col_pct   = find_col(company_df, ["% sprint completion","sprint completion","% completion"])
col_notes = find_col(company_df, ["notes / comments","notes","remarks","comment"])
col_rev   = find_col(company_df, ["revenue ly","revenue ly ("])
col_tgt   = find_col(company_df, ["3-year target","3 year target","target ("])
col_sprint = None
for c in company_df.columns:
    cl = str(c).lower()
    if "sprint type" in cl:
        col_sprint = c; break
if not col_sprint:
    for c in company_df.columns:
        cl = str(c).lower()
        if "sprint" in cl and not any(x in cl for x in ["commit","complet","score","%","task"]):
            col_sprint = c; break
if col_pct is None and len(company_df.columns) > 37:
    col_pct = company_df.columns[37]

# ── build ventures list ───────────────────────────────
skip = ["nan","none","","company name","name","venture name","-","—"]
if col_name:
    all_rows = company_df[company_df[col_name].notna()].copy()
    all_rows = all_rows[~all_rows[col_name].astype(str).str.strip().str.lower().isin(skip)]
    ventures_raw = all_rows[col_name].astype(str).str.strip().tolist()
else:
    st.error("Could not find 'Company Name' column.")
    st.stop()

# Deduplicate
seen = set(); ventures_deduped = []
for v in ventures_raw:
    if v not in seen: seen.add(v); ventures_deduped.append(v)
ventures_raw = ventures_deduped

# ── sidebar debug ─────────────────────────────────────
with st.sidebar:
    with st.expander("🔍 Column Mapping (debug)"):
        st.write(f"**Source:** {'SharePoint ☁️' if use_sharepoint else 'Local 💻'}")
        st.write(f"**Name:** {col_name}")
        st.write(f"**Hub:** {col_hub}")
        st.write(f"**Sprint:** {col_sprint}")
        st.write(f"**% Completion:** {col_pct}")
        st.write(f"**Notes:** {col_notes}")
        st.write(f"**Total rows:** {len(all_rows)}")
        st.write(f"**Ventures loaded:** {len(ventures_raw)}")
        if col_pct:
            sample = all_rows[col_pct].dropna().head(5).tolist()
            st.write(f"**Sample %:** {sample}")

# ── helpers ───────────────────────────────────────────
def get_row(name):
    matches = all_rows[all_rows[col_name].astype(str).str.strip() == name]
    return matches.iloc[0] if not matches.empty else None

def cv(row, col, default="—"):
    if row is None or col is None: return default
    v = str(row[col]).strip()
    if v in ["nan","None","NaT",""]: return default
    try:
        f = float(v)
        if f == int(f): return str(int(f))
        return v
    except: return v

def load_v_files_local(vname):
    vpath = os.path.join(root_path or "", vname)
    if not os.path.isdir(vpath): return {}
    files = {}
    for f in os.listdir(vpath):
        fl = f.lower(); fp = os.path.join(vpath, f)
        if not os.path.isfile(fp): continue
        if "transcript" in fl:   files["transcript"] = fp
        elif "feedback"  in fl:  files["feedback"]   = fp
        elif "sprint plan" in fl or "growth sprint" in fl: files["sprint"] = fp
    return files

def load_v_files_sharepoint(vname, sp):
    """Load venture files from SharePoint."""
    folder = f"{SP_FOLDER}/{vname}"
    files  = {}
    try:
        items = sp.list_files(folder)
        for fname in items:
            fl = fname.lower()
            fpath = f"{folder}/{fname}"
            if "transcript" in fl:   files["transcript"] = fpath
            elif "feedback"  in fl:  files["feedback"]   = fpath
            elif "sprint plan" in fl or "growth sprint" in fl: files["sprint"] = fpath
    except: pass
    return files

def get_file_text(fpath, sp=None):
    """Get text from file — local or SharePoint."""
    if sp and not os.path.exists(str(fpath)):
        try:
            content = sp.download_file(fpath)
            return extract_text_bytes(content, Path(fpath).name)
        except: return ""
    return extract_text_local(fpath)

client = Anthropic(api_key=api_key) if api_key else None
sp_reader = None
if use_sharepoint and ENV_CLIENT_ID:
    with st.spinner("Connecting to SharePoint..."):
        try:
            from sharepoint_reader import SharePointReader
            sp_reader = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        except Exception as e:
            st.error(f"SharePoint connection failed: {e}")

# ══════════════════════════════════════════════════════
#  VIEW 1: PORTFOLIO OVERVIEW
# ══════════════════════════════════════════════════════
if view == "Portfolio Overview":
    st.title("📊 Portfolio Overview")
    st.caption("NEN Accelerate · All ventures at a glance")
    st.divider()

    total = len(ventures_raw)
    stage_counts = {k:0 for k in STAGE_COLORS}; stage_counts["Unknown"] = 0
    hub_data, sprint_types = {}, {}

    for v in ventures_raw:
        row    = get_row(v)
        pct    = row[col_pct] if (row is not None and col_pct) else None
        bucket = get_stage_bucket(pct)
        stage_counts[bucket] = stage_counts.get(bucket,0) + 1
        hub = cv(row, col_hub, "Other")
        hub_data[hub] = hub_data.get(hub, 0) + 1
        sp = cv(row, col_sprint, "")
        if sp and sp != "—":
            sprint_types[sp] = sprint_types.get(sp, 0) + 1

    active    = stage_counts.get("51–75%",0) + stage_counts.get("76–99%",0)
    completed = stage_counts.get("100%",0)
    early     = stage_counts.get("0–25%",0)  + stage_counts.get("26–50%",0)

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Ventures",         total)
    c2.metric("Sprint Active (51–99%)", active,    f"{round(active/total*100) if total else 0}%")
    c3.metric("Completed (100%)",       completed, f"{round(completed/total*100) if total else 0}%")
    c4.metric("Early Stage (0–50%)",    early,     f"{round(early/total*100) if total else 0}%")
    c5.metric("Sprint Types",           len(sprint_types))

    st.divider()
    col_l, col_r = st.columns(2)

    with col_l:
        st.subheader("Sprint Completion Distribution")
        for bucket, color in STAGE_COLORS.items():
            cnt = stage_counts.get(bucket,0)
            pct = round(cnt/total*100) if total else 0
            st.markdown(f"**{bucket}** — {cnt} ventures ({pct}%)")
            st.progress(pct/100)

    with col_r:
        st.subheader("Hub-wise Breakdown")
        for hub, cnt in sorted(hub_data.items(), key=lambda x:-x[1]):
            pct = round(cnt/total*100) if total else 0
            st.markdown(f"**{hub}** — {cnt} ventures ({pct}%)")
            st.progress(pct/100)

    if sprint_types:
        st.divider()
        st.subheader("Sprint Types in Portfolio")
        cols = st.columns(min(len(sprint_types), 4))
        for i,(sp,cnt) in enumerate(sorted(sprint_types.items(), key=lambda x:-x[1])):
            cols[i % len(cols)].metric(sp, f"{cnt} ventures")

# ══════════════════════════════════════════════════════
#  VIEW 2: VENTURE CARDS
# ══════════════════════════════════════════════════════
elif view == "Venture Cards":
    st.title("🏢 Venture Cards")
    st.caption("Individual venture status, sessions, and insights")
    st.divider()

    fc1,fc2,fc3 = st.columns([2,1,1])
    search       = fc1.text_input("🔍 Search ventures")
    hub_opts     = ["All Hubs"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_raw if cv(get_row(v),col_hub) != "—"))
    hub_filter   = fc2.selectbox("Hub", hub_opts)
    stage_filter = fc3.selectbox("Stage", ["All Stages"]+list(STAGE_COLORS.keys()))
    use_ai       = st.checkbox("✨ Generate AI insights", help="Uses Claude API")

    filtered = ventures_raw
    if search:       filtered = [v for v in filtered if search.lower() in v.lower()]
    if hub_filter != "All Hubs":
        filtered = [v for v in filtered if cv(get_row(v), col_hub) == hub_filter]
    if stage_filter != "All Stages":
        filtered = [v for v in filtered if get_stage_bucket(
            get_row(v)[col_pct] if (get_row(v) is not None and col_pct) else None) == stage_filter]

    st.caption(f"{len(filtered)} ventures shown")

    for i, vname in enumerate(filtered):
        row    = get_row(vname)
        hub    = cv(row, col_hub)
        sprint = cv(row, col_sprint)
        rev    = cv(row, col_rev)
        tgt    = cv(row, col_tgt)
        notes  = cv(row, col_notes, default="")
        pct_raw= row[col_pct] if (row is not None and col_pct) else None
        bucket = get_stage_bucket(pct_raw)

        try:
            pct_num = float(str(pct_raw).replace("%","").strip())
            if pct_num <= 1: pct_num *= 100
        except: pct_num = 0

        # Load venture files
        if use_sharepoint and sp_reader:
            vfiles = load_v_files_sharepoint(vname, sp_reader)
        else:
            vfiles = load_v_files_local(vname)

        fb_text = get_file_text(vfiles["feedback"], sp_reader) if "feedback"   in vfiles else ""
        tr_text = get_file_text(vfiles["transcript"], sp_reader) if "transcript" in vfiles else ""
        signals = detect_signals((notes or "")+" "+(fb_text or ""))

        with st.expander(f"**{vname}**  ·  {hub}  ·  {bucket}  ·  Sprint: {sprint}"):
            ca, cb = st.columns([3,1])
            with ca:
                st.markdown(f"### {vname}")
                st.caption(f"📍 {hub}")
                if sprint and sprint != "—": st.caption(f"🏃 Sprint: {sprint}")
            with cb:
                if rev != "—": st.metric("Revenue LY (Cr)", rev)
                if tgt != "—": st.metric("3-Yr Target (Cr)", tgt)

            st.markdown(f"**Sprint Completion: {pct_num:.0f}%** &nbsp; `{bucket}`")
            st.progress(min(pct_num/100, 1.0))

            if vfiles:
                found = []
                if "feedback"   in vfiles: found.append("📋 Feedback")
                if "transcript" in vfiles: found.append("🎙 Transcript")
                if "sprint"     in vfiles: found.append("📌 Sprint Plan")
                st.caption("Files: " + "  ".join(found))

            if signals:
                st.markdown("**✦ Success Signals:**")
                sig_cols = st.columns(min(len(signals), 3))
                for j, s in enumerate(signals[:3]):
                    sig_cols[j].success(s["type"])

            if notes:
                st.markdown("**📝 Remarks:**")
                st.info(notes[:600] + ("..." if len(notes)>600 else ""))

            if fb_text or tr_text:
                st.markdown("**🎙 Session Feedback:**")
                st.info((fb_text or tr_text)[:500] + ("..." if len(fb_text or tr_text)>500 else ""))

            if use_ai and client:
                with st.spinner("Generating AI insight..."):
                    try:
                        resp = client.messages.create(
                            model="claude-sonnet-4-20250514", max_tokens=400,
                            messages=[{"role":"user","content":
                                f"Venture: {vname}\nNotes: {notes}\nFeedback: {fb_text[:500]}\n\n"
                                "Write 3-4 sentences: current momentum, risks to watch, recommended next action."}])
                        st.markdown("**✦ AI Insight:**")
                        st.success(resp.content[0].text.strip())
                    except Exception as e:
                        st.warning(f"AI insight unavailable: {e}")

# ══════════════════════════════════════════════════════
#  VIEW 3: SUCCESS SIGNALS
# ══════════════════════════════════════════════════════
elif view == "Success Signals":
    st.title("💡 Success Signals")
    st.caption("Growth actions taken by ventures — hiring, spending, new streams")
    st.divider()

    use_ai_signals = st.checkbox("✨ Use Claude AI for deeper signal detection", 
                                  help="Reads documents with AI understanding — much richer than keywords")

    all_signals = []
    prog = st.progress(0, text="Scanning ventures...")

    for idx, vname in enumerate(ventures_raw):
        row   = get_row(vname)
        notes = cv(row, col_notes, default="")

        if use_sharepoint and sp_reader:
            vfiles = load_v_files_sharepoint(vname, sp_reader)
        else:
            vfiles = load_v_files_local(vname)

        fb   = get_file_text(vfiles["feedback"], sp_reader) if "feedback"   in vfiles else ""
        tr   = get_file_text(vfiles["transcript"], sp_reader) if "transcript" in vfiles else ""
        combined = (notes or "") + " " + (fb or "") + " " + (tr or "")

        if use_ai_signals and client:
            try:
                resp = client.messages.create(
                    model="claude-sonnet-4-20250514", max_tokens=600,
                    messages=[{"role":"user","content":
                        f"""Analyze this text from venture '{vname}' and identify success signals.
Success signals = concrete evidence venture spent money or effort for growth:
- Hired someone new
- Invested in equipment/facility
- Entered new market/geography
- Paid for another sprint themselves
- Won new client or order

Text: {combined[:3000]}

Return ONLY a JSON array. Each item: {{"type": "signal type", "evidence": "exact quote from text"}}
If none found return []. No other text."""}])
                raw = resp.content[0].text.strip()
                raw = re.sub(r"```json|```","",raw).strip()
                sigs = json.loads(raw)
            except: sigs = detect_signals(combined)
        else:
            sigs = detect_signals(combined)

        for s in sigs:
            all_signals.append({
                "venture": vname, "hub": cv(row,col_hub),
                "sprint":  cv(row,col_sprint),
                "type":    s.get("type","Signal"),
                "evidence":s.get("evidence", s.get("keyword","—")),
                "notes":   notes[:300]
            })
        prog.progress((idx+1)/max(len(ventures_raw),1), text=f"Scanning {vname}...")
    prog.empty()

    # Summary
    v_with = len(set(s["venture"] for s in all_signals))
    type_counts = {}
    for s in all_signals: type_counts[s["type"]] = type_counts.get(s["type"],0)+1

    c1,c2,c3 = st.columns(3)
    c1.metric("Total Signals",         len(all_signals))
    c2.metric("Ventures with Signals", v_with, f"of {len(ventures_raw)}")
    top = sorted(type_counts.items(), key=lambda x:-x[1])
    if top: c3.metric(f"Top Signal: {top[0][0]}", top[0][1])

    # Signal type breakdown
    if type_counts:
        st.divider()
        st.subheader("Signals by Type")
        cols = st.columns(min(len(type_counts),5))
        for i,(t,c) in enumerate(sorted(type_counts.items(),key=lambda x:-x[1])):
            cols[i%len(cols)].metric(t, c)

    st.divider()
    if not all_signals:
        st.info("No signals detected yet. Make sure Notes column has data, or enable AI detection.")
    else:
        f1,f2 = st.columns([1,2])
        type_filter = f1.selectbox("Signal Type", ["All"]+sorted(type_counts.keys()))
        name_filter = f2.text_input("Filter by venture")

        filtered_s = [s for s in all_signals
                      if (type_filter=="All" or s["type"]==type_filter)
                      and (not name_filter or name_filter.lower() in s["venture"].lower())]

        st.caption(f"{len(filtered_s)} signals shown")
        for s in filtered_s:
            with st.expander(f"**{s['venture']}** · `{s['type']}`"):
                st.markdown(f"**Hub:** {s['hub']}  ·  **Sprint:** {s['sprint']}")
                st.markdown(f"**Signal:** `{s['type']}`")
                st.markdown(f"**Evidence:** {s['evidence']}")
                if s["notes"]: st.info(s["notes"])
