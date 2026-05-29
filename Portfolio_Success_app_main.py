import streamlit as st
import pandas as pd
import os, json, re, tempfile, shutil
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
ENV_PASSWORD      = os.environ.get("APP_PASSWORD", "nen2026")

SP_FOLDER       = "04. Advisors/2026/Portfolio Success Dashboard"
COMMON_FOLDER   = f"{SP_FOLDER}/Common Documents"
DASHBOARD_FILE  = "0. Journey_Accelerate_Portfolio Dashboard.xlsx"

st.set_page_config(page_title="Portfolio Success Intelligence", page_icon="🚀", layout="wide")

# ── CSS ───────────────────────────────────────────────
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
.rag-green  { background:#dcfce7; color:#166534; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.85rem; }
.rag-amber  { background:#fef9c3; color:#854d0e; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.85rem; }
.rag-red    { background:#fee2e2; color:#991b1b; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.85rem; }
.rag-zero   { background:#f1f5f9; color:#64748b; padding:3px 12px; border-radius:20px; font-weight:700; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

# ── password gate ─────────────────────────────────────
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
        pwd = st.text_input("Password", type="password", placeholder="Enter access password...", label_visibility="collapsed")
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

# ── constants ─────────────────────────────────────────
RAG_ORDER  = {"Green": 0, "Amber": 1, "Red": 2, "ZERO": 3, "Unknown": 4}
RAG_EMOJI  = {"Green": "🟢", "Amber": "🟡", "Red": "🔴", "ZERO": "⚪", "Unknown": "⚪"}
RAG_COLOR  = {"Green": "rag-green", "Amber": "rag-amber", "Red": "rag-red", "ZERO": "rag-zero", "Unknown": "rag-zero"}

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
                out.append({"type": stype, "keyword": kw}); break
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
    shutil.copy2(src, dst); return dst

def find_col(df, patterns):
    for c in df.columns:
        for p in patterns:
            if p.lower() in str(c).lower(): return c
    return None

def extract_text_local(fpath):
    ext = Path(fpath).suffix.lower()
    try:
        fp = safe_copy(fpath)
        if ext in [".xlsx",".xls"]:
            xl = pd.ExcelFile(fp)
            return "\n".join(pd.read_excel(fp,sheet_name=s,header=None).fillna("").astype(str).to_string() for s in xl.sheet_names)[:6000]
        elif ext == ".docx":
            from docx import Document
            return "\n".join(p.text for p in Document(fp).paragraphs)[:6000]
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(fp) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)[:6000]
    except Exception as e: return f"[Error: {e}]"
    return ""

def extract_text_bytes(content_bytes, filename):
    ext = Path(filename).suffix.lower()
    try:
        tmp = os.path.join(tempfile.gettempdir(), f"nen_sp_{filename}")
        with open(tmp,"wb") as f: f.write(content_bytes)
        return extract_text_local(tmp)
    except Exception as e: return f"[Error: {e}]"

def combine_rag(momentum, investment):
    """Compute overall RAG from two scores."""
    scores = [s for s in [momentum, investment] if s not in ["ZERO","Unknown",None,"","—"]]
    if not scores: return "ZERO"
    order = {"Red":0,"Amber":1,"Green":2}
    worst = min(scores, key=lambda x: order.get(x,1))
    return worst

# ── sidebar ───────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    use_sp = st.toggle("☁️ Use SharePoint", value=bool(ENV_CLIENT_ID))
    if use_sp:
        root_path = None
        if ENV_CLIENT_ID: st.success("✅ SharePoint configured")
        else: st.warning("Add Azure credentials to secrets")
    else:
        root_path = ENV_ROOT_PATH or st.text_input("Local Folder Path")

    api_key = ENV_API_KEY or st.text_input("Anthropic API Key", type="password")

    st.markdown("---")
    st.caption("NEN Accelerate · Portfolio Intelligence\nResources Network Team")

# ── SP reader ─────────────────────────────────────────
sp_reader = None
if use_sp and ENV_CLIENT_ID:
    try:
        from sharepoint_reader import SharePointReader
        if "sp_reader" not in st.session_state:
            with st.spinner("Connecting to SharePoint..."):
                st.session_state["sp_reader"] = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        sp_reader = st.session_state["sp_reader"]
    except Exception as e:
        st.error(f"SharePoint connection failed: {e}"); st.stop()

def sp_download(fpath):
    if sp_reader:
        content = sp_reader.download_file(fpath)
        return extract_text_bytes(content, Path(fpath).name)
    return ""

def sp_list(folder):
    if sp_reader:
        try: return sp_reader.list_folder(folder)
        except: return []
    return []

# ── load dashboard ────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_dashboard(_sp_id, root_path, use_sp):
    try:
        if use_sp and ENV_CLIENT_ID:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            content = sp.download_file(f"{SP_FOLDER}/{DASHBOARD_FILE}")
            tmp = os.path.join(tempfile.gettempdir(), "nen_dashboard.xlsx")
            with open(tmp,"wb") as f: f.write(content)
            fpath = tmp
        else:
            hc = r"C:\Users\MeenakshiSingh\OneDrive - National Entrepreneurship Network\04. Advisors\2026\Portfolio Success Dashboard\0. Journey_Accelerate_Portfolio Dashboard.xlsx"
            fpath = hc if os.path.exists(hc) else None
            if not fpath:
                for f in os.listdir(root_path or ""):
                    if "Journey_Accelerate_Portfolio" in f and f.endswith(".xlsx") and "~$" not in f:
                        fpath = os.path.join(root_path, f); break
            if fpath: fpath = safe_copy(fpath)
        if not fpath: return None, None, "Dashboard file not found."
        company_df = pd.read_excel(fpath, sheet_name="Company", header=2)
        try: basics_df = pd.read_excel(fpath, sheet_name="Company Basics", header=2)
        except: basics_df = None
        return company_df, basics_df, None
    except Exception as e: return None, None, str(e)

with st.spinner("Loading portfolio data..."):
    sp_id = id(sp_reader) if sp_reader else 0
    company_df, basics_df, err = load_dashboard(sp_id, root_path, use_sp)

if err: st.error(f"❌ {err}"); st.stop()
if company_df is None: st.error("Could not load data."); st.stop()

# ── detect columns ────────────────────────────────────
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
    cl = str(c).lower()
    if "sprint type" in cl: col_sprint = c; break
if not col_sprint:
    for c in company_df.columns:
        cl = str(c).lower()
        if "sprint" in cl and not any(x in cl for x in ["commit","complet","score","%","task"]): col_sprint = c; break
if col_pct is None and len(company_df.columns) > 37: col_pct = company_df.columns[37]

# ── build venture list ────────────────────────────────
skip = ["nan","none","","company name","name","venture name","-","—"]
all_rows = company_df[company_df[col_name].notna()].copy() if col_name else company_df.copy()
all_rows = all_rows[~all_rows[col_name].astype(str).str.strip().str.lower().isin(skip)]
ventures_raw = all_rows[col_name].astype(str).str.strip().tolist()
seen = set(); ventures_deduped = []
for v in ventures_raw:
    if v not in seen: seen.add(v); ventures_deduped.append(v)
ventures_raw = ventures_deduped

client = Anthropic(api_key=api_key) if api_key else None

# ── helpers ───────────────────────────────────────────
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

def load_v_files(vname):
    files = {}
    if use_sp and sp_reader:
        folder = f"{SP_FOLDER}/{vname}"
        try:
            items = sp_reader.list_files(folder)
            for fname in items:
                fl = fname.lower(); fp = f"{folder}/{fname}"
                if "transcript" in fl: files["transcript"] = fp
                elif "feedback" in fl: files["feedback"]   = fp
                elif "sprint plan" in fl or "growth sprint" in fl: files["sprint"] = fp
        except: pass
    else:
        vpath = os.path.join(root_path or "", vname)
        if os.path.isdir(vpath):
            for f in os.listdir(vpath):
                fl = f.lower(); fp = os.path.join(vpath, f)
                if not os.path.isfile(fp): continue
                if "transcript" in fl: files["transcript"] = fp
                elif "feedback"  in fl: files["feedback"]   = fp
                elif "sprint plan" in fl or "growth sprint" in fl: files["sprint"] = fp
    return files

def get_text(fpath):
    if use_sp and sp_reader and not os.path.exists(str(fpath)):
        try:
            content = sp_reader.download_file(fpath)
            return extract_text_bytes(content, Path(fpath).name)
        except: return ""
    return extract_text_local(fpath)

@st.cache_data(show_spinner=False, ttl=600)
def load_common_docs_cached(_sp_id, use_sp, root_path):
    """Load all Common Documents and return full combined text."""
    texts = []
    if use_sp and ENV_CLIENT_ID:
        try:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            items = sp.list_files(COMMON_FOLDER)
            for fname in items:
                ext = Path(fname).suffix.lower()
                if ext not in [".xlsx",".xls",".docx",".pdf",".pptx",".ppt"]: continue
                fp = f"{COMMON_FOLDER}/{fname}"
                try:
                    content = sp.download_file(fp)
                    text = extract_text_bytes(content, fname)
                    if text and len(text) > 50:
                        texts.append(f"=== FILE: {fname} ===\n{text[:3000]}")
                except: pass
        except: pass
    elif root_path:
        cpath = os.path.join(root_path, "Common Documents")
        if os.path.isdir(cpath):
            for f in os.listdir(cpath):
                ext = Path(f).suffix.lower()
                if ext not in [".xlsx",".xls",".docx",".pdf",".pptx"]: continue
                fp = os.path.join(cpath, f)
                if os.path.isfile(fp):
                    text = extract_text_local(fp)
                    if text: texts.append(f"=== FILE: {f} ===\n{text[:3000]}")
    return "\n\n".join(texts)

def extract_venture_from_common(venture_name, common_text):
    """Extract only sections mentioning this venture from common docs."""
    if not common_text or not venture_name: return ""
    # Split by file sections
    relevant = []
    sections = common_text.split("=== FILE:")
    for section in sections:
        # Check if venture name appears in this section (case-insensitive)
        if venture_name.lower() in section.lower():
            # Extract paragraphs/lines mentioning the venture
            lines = section.split("\n")
            venture_lines = []
            for i, line in enumerate(lines):
                if venture_name.lower() in line.lower():
                    # Include surrounding context (2 lines before and after)
                    start = max(0, i-2)
                    end   = min(len(lines), i+5)
                    venture_lines.extend(lines[start:end])
            if venture_lines:
                fname = section.split("===")[0].strip() if "===" in section else "Common Doc"
                relevant.append(f"[{fname}]\n" + "\n".join(venture_lines))
    return "\n\n".join(relevant)[:2000]

def load_common_docs():
    """Wrapper to load common docs using cache."""
    sp_id = id(sp_reader) if sp_reader else 0
    return load_common_docs_cached(sp_id, use_sp, root_path)

@st.cache_data(show_spinner=False, ttl=600)
def compute_rag_ai(vname, notes, fb_text, tr_text, common_text, pct_raw, _client_key):
    """Use Claude to compute RAG scores for a venture."""
    if not client: return "Unknown", "Unknown", "No AI key", "No AI key"
    
    try:
        pct = float(str(pct_raw).replace("%","").strip())
        if pct <= 1: pct *= 100
    except: pct = 0

    # Extract only sections mentioning this venture from common docs
    venture_common = extract_venture_from_common(vname, common_text) if common_text else ""

    combined = f"""
Venture: {vname}
Sprint Completion: {pct:.0f}%
Program Notes/Remarks: {notes[:800]}
Session Feedback: {fb_text[:600]}
Transcript: {tr_text[:600]}
Common Documents (venture-specific excerpts): {venture_common[:800]}
"""
    prompt = f"""Analyze this venture data and return ONLY a JSON object with 4 keys:
- "momentum_rag": one of "Green", "Amber", "Red", "ZERO"
- "momentum_reason": one sentence explanation
- "investment_rag": one of "Green", "Amber", "Red", "ZERO"  
- "investment_reason": one sentence explanation

Scoring guide:
Sprint Momentum Score:
- Green: Active, completing tasks, engaged founder, positive feedback
- Amber: Some progress but inconsistent, needs nudging, partial completion
- Red: Not engaged, founder dissatisfied, no progress, wants to drop out
- ZERO: No data available

Sprint Self Investment Signal Score:
- Green: Hired staff, invested in equipment/facility, self-funded activities
- Amber: Planning to invest, exploring options, partial commitment
- Red: No investment intent, focused elsewhere, withdrawing
- ZERO: No data available

Venture data:
{combined}

Return ONLY the JSON, no other text."""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5", max_tokens=300,
            messages=[{"role":"user","content":prompt}])
        raw = re.sub(r"```json|```","",resp.content[0].text.strip()).strip()
        data = json.loads(raw)
        return (data.get("momentum_rag","Unknown"), data.get("investment_rag","Unknown"),
                data.get("momentum_reason","—"), data.get("investment_reason","—"))
    except Exception as e:
        return "Unknown","Unknown",f"Error: {e}","—"

def rag_badge(rag):
    css = RAG_COLOR.get(rag, "rag-zero")
    emoji = RAG_EMOJI.get(rag, "⚪")
    return f'<span class="{css}">{emoji} {rag}</span>'

# ── top navigation tabs ──────────────────────────────
tab_overview, tab_ventures = st.tabs(["📊  Portfolio Overview", "🏢  Venture Cards"])

# ══════════════════════════════════════════════════════
#  VIEW 1: PORTFOLIO OVERVIEW
# ══════════════════════════════════════════════════════
with tab_overview:
    st.title("📊 Portfolio Overview")
    st.caption("NEN Accelerate · All ventures at a glance — RAG scores powered by AI")
    st.divider()

    # ── filters ──────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns(4)
    hub_list = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_raw if cv(get_row(v),col_hub) != "—"))
    vp_list  = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_raw if col_vp and cv(get_row(v),col_vp)  != "—"))
    hub_f    = fc1.selectbox("Hub", hub_list, key="ov_hub")
    vp_f     = fc2.selectbox("Venture Partner", vp_list, key="ov_vp")
    rag_f    = fc3.selectbox("RAG Filter", ["All","🟢 Green","🟡 Amber","🔴 Red","⚪ ZERO"], key="ov_rag")
    stage_f  = fc4.selectbox("Sprint Stage", ["All","0–25%","26–50%","51–75%","76–99%","100%"], key="ov_stage")

    # ── compute scores (AI by default) ───────────────
    venture_data = []
    cache_key    = "rag_scores_ai"

    col_ref, col_info = st.columns([1, 4])
    if col_ref.button("🔄 Refresh Scores", help="Recompute all RAG scores"):
        st.session_state.pop(cache_key, None)
        st.session_state.pop("common_docs_cache", None)

    if cache_key in st.session_state:
        venture_data = st.session_state[cache_key]
        col_info.caption("✅ Using cached AI scores — click 'Refresh Scores' to recompute")
    else:
        if not client:
            col_info.warning("⚠️ No Anthropic API key found. Add it to Streamlit secrets as ANTHROPIC_API_KEY.")

        # Step 1: Build basic venture list instantly
        for vname in ventures_raw:
            row = get_row(vname)
            venture_data.append({
                "name":    vname,
                "hub":     cv(row, col_hub),
                "vp":      cv(row, col_vp) if col_vp else "—",
                "sprint":  cv(row, col_sprint),
                "rev":     cv(row, col_rev),
                "bucket":  get_stage_bucket(row[col_pct] if (row is not None and col_pct) else None),
                "pct_raw": row[col_pct] if (row is not None and col_pct) else None,
                "notes":   cv(row, col_notes, default=""),
                "momentum_rag":    "ZERO",
                "investment_rag":  "ZERO",
                "overall_rag":     "ZERO",
                "momentum_reason": "Pending",
                "investment_reason": "Pending",
            })

        # Step 2: Load common documents once
        with st.spinner("📂 Loading Common Documents..."):
            common_text = load_common_docs()
        col_info.caption(f"📂 Common Documents loaded: {len(common_text)} chars")

        # Step 3: AI scoring in parallel — ALL ventures scored
        # Common Documents is primary source, Notes + files are supplementary
        if client:
            import concurrent.futures

            prog = st.progress(0, text=f"🤖 AI scoring all {len(venture_data)} ventures using Common Documents + Notes + Files...")
            completed = [0]

            def score_one(v):
                # Load venture-specific files
                vfiles  = load_v_files(v["name"])
                fb_text = get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
                tr_text = get_text(vfiles["transcript"]) if "transcript" in vfiles else ""

                # Extract venture-specific section from common documents
                venture_common = extract_venture_from_common(v["name"], common_text)

                # Only skip if absolutely no data from ANY source
                has_data = any([
                    v["notes"] and v["notes"] not in ["—","Pending",""],
                    fb_text,
                    tr_text,
                    venture_common
                ])

                if not has_data:
                    return v["name"], "ZERO", "ZERO", "No data found in any source", "No data found in any source"

                m, i, mr, ir = compute_rag_ai(
                    v["name"], v["notes"], fb_text, tr_text,
                    common_text, v["pct_raw"], api_key)
                return v["name"], m, i, mr, ir

            results = {}
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
                futures = {ex.submit(score_one, v): v for v in venture_data}
                for fut in concurrent.futures.as_completed(futures):
                    try:
                        vn, m, i, mr, ir = fut.result()
                        results[vn] = (m, i, mr, ir)
                    except Exception as e:
                        pass
                    completed[0] += 1
                    prog.progress(
                        completed[0] / len(venture_data),
                        text=f"🤖 Scored {completed[0]}/{len(venture_data)} ventures..."
                    )
            prog.empty()

            for v in venture_data:
                if v["name"] in results:
                    m, i, mr, ir = results[v["name"]]
                    v["momentum_rag"]      = m
                    v["investment_rag"]    = i
                    v["overall_rag"]       = combine_rag(m, i)
                    v["momentum_reason"]   = mr
                    v["investment_reason"] = ir
        else:
            st.warning("⚠️ No API key — add ANTHROPIC_API_KEY to Streamlit secrets for AI scoring.")

        st.session_state[cache_key] = venture_data
        st.success(f"✅ AI scores computed for {len([v for v in venture_data if v['overall_rag'] != 'ZERO'])} ventures")

    # ── apply filters ─────────────────────────────────
    filtered = venture_data
    if hub_f   != "All": filtered = [v for v in filtered if v["hub"]    == hub_f]
    if vp_f    != "All": filtered = [v for v in filtered if v["vp"]     == vp_f]
    if stage_f != "All": filtered = [v for v in filtered if v["bucket"] == stage_f]
    if rag_f   != "All":
        rag_val = rag_f.split(" ",1)[1]
        filtered = [v for v in filtered if v["overall_rag"] == rag_val]

    # ── summary metrics ───────────────────────────────
    total   = len(filtered)
    greens  = sum(1 for v in filtered if v["overall_rag"] == "Green")
    ambers  = sum(1 for v in filtered if v["overall_rag"] == "Amber")
    reds    = sum(1 for v in filtered if v["overall_rag"] == "Red")
    zeros   = sum(1 for v in filtered if v["overall_rag"] == "ZERO")

    # Overall portfolio RAG — worst of the lot
    if reds   > 0: portfolio_rag = "Red"
    elif ambers > 0: portfolio_rag = "Amber"
    elif greens > 0: portfolio_rag = "Green"
    else:            portfolio_rag = "ZERO"

    # Top row — overall RAG + counts
    st.markdown(f"### Overall Portfolio RAG: {rag_badge(portfolio_rag)}", unsafe_allow_html=True)
    st.caption("Updates automatically based on filters applied above")
    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Ventures",  total)
    c2.metric("🟢 Green",  greens, f"{round(greens/total*100) if total else 0}%")
    c3.metric("🟡 Amber",  ambers, f"{round(ambers/total*100) if total else 0}%")
    c4.metric("🔴 Red",    reds,   f"{round(reds/total*100)   if total else 0}%")
    c5.metric("⚪ No Data", zeros)

    # RAG progress bars
    st.markdown("<br>", unsafe_allow_html=True)
    for label, count, color in [
        ("🟢 Green", greens, "#16a34a"),
        ("🟡 Amber", ambers, "#d97706"),
        ("🔴 Red",   reds,   "#dc2626"),
    ]:
        pct = round(count/total*100) if total else 0
        st.markdown(f"**{label}** — {count} ventures ({pct}%)")
        st.progress(pct/100)

    # Sprint stage distribution
    st.divider()
    st.subheader("Sprint Stage Distribution")
    stage_counts = {}
    for v in filtered:
        stage_counts[v["bucket"]] = stage_counts.get(v["bucket"],0) + 1
    sc1, sc2 = st.columns(2)
    for i, (stage, cnt) in enumerate(sorted(stage_counts.items())):
        pct = round(cnt/total*100) if total else 0
        col = sc1 if i % 2 == 0 else sc2
        col.markdown(f"**{stage}** — {cnt} ventures ({pct}%)")
        col.progress(pct/100)

    # ── hub-wise RAG table ────────────────────────────
    st.divider()
    st.subheader("Hub-wise Venture Count")

    # Build hub data
    hub_rag = {}
    for v in filtered:
        h = v["hub"] if v["hub"] != "—" else "Other"
        if h not in hub_rag:
            hub_rag[h] = {"Green":0,"Amber":0,"Red":0,"ZERO":0,"Total":0}
        hub_rag[h][v["overall_rag"]] = hub_rag[h].get(v["overall_rag"],0) + 1
        hub_rag[h]["Total"] += 1

    # Table header
    h0,h1,h2,h3,h4,h5 = st.columns([2,1,1,1,1,1])
    h0.markdown("**Hub**")
    h1.markdown("**Total**")
    h2.markdown("**🟢 Green**")
    h3.markdown("**🟡 Amber**")
    h4.markdown("**🔴 Red**")
    h5.markdown("**⚪ No Data**")
    st.divider()

    for hub, counts in sorted(hub_rag.items(), key=lambda x: -x[1]["Total"]):
        c0,c1,c2,c3,c4,c5 = st.columns([2,1,1,1,1,1])
        total_h = counts["Total"]
        c0.markdown(f"**{hub}**")
        c1.markdown(f"**{total_h}**")
        c2.markdown(f"<span style='color:#16a34a;font-weight:700'>{counts.get('Green',0)}</span>", unsafe_allow_html=True)
        c3.markdown(f"<span style='color:#d97706;font-weight:700'>{counts.get('Amber',0)}</span>", unsafe_allow_html=True)
        c4.markdown(f"<span style='color:#dc2626;font-weight:700'>{counts.get('Red',0)}</span>", unsafe_allow_html=True)
        c5.markdown(f"<span style='color:#94a3b8;font-weight:700'>{counts.get('ZERO',0)}</span>", unsafe_allow_html=True)
    
    st.divider()
    # Totals row
    t0,t1,t2,t3,t4,t5 = st.columns([2,1,1,1,1,1])
    t0.markdown("**Total**")
    t1.markdown(f"**{len(filtered)}**")
    t2.markdown(f"**{sum(c.get('Green',0) for c in hub_rag.values())}**")
    t3.markdown(f"**{sum(c.get('Amber',0) for c in hub_rag.values())}**")
    t4.markdown(f"**{sum(c.get('Red',0) for c in hub_rag.values())}**")
    t5.markdown(f"**{sum(c.get('ZERO',0) for c in hub_rag.values())}**")

# ══════════════════════════════════════════════════════
#  VIEW 2: VENTURE CARDS
# ══════════════════════════════════════════════════════
with tab_ventures:
    st.title("🏢 Venture Cards")
    st.caption("Individual venture deep-dive")
    st.divider()

    # ── filters ──────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns(4)
    search   = fc1.text_input("🔍 Search", key="vc_search")
    hub_list = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_raw if cv(get_row(v),col_hub) != "—"))
    vp_list  = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_raw if col_vp and cv(get_row(v),col_vp)  != "—"))
    hub_f    = fc2.selectbox("Hub", hub_list, key="vc_hub")
    vp_f     = fc3.selectbox("Venture Partner", vp_list, key="vc_vp")
    rag_f    = fc4.selectbox("RAG", ["All","🟢 Green","🟡 Amber","🔴 Red"], key="vc_rag")

    filtered = ventures_raw
    if search: filtered = [v for v in filtered if search.lower() in v.lower()]
    if hub_f != "All": filtered = [v for v in filtered if cv(get_row(v),col_hub) == hub_f]
    if vp_f  != "All" and col_vp: filtered = [v for v in filtered if cv(get_row(v),col_vp) == vp_f]

    # Handle jump from portfolio view
    selected = st.session_state.get("selected_venture")
    if selected and selected in ventures_raw:
        filtered = [selected] + [v for v in filtered if v != selected]
        st.session_state.pop("selected_venture", None)
        st.session_state.pop("jump_to_venture", None)

    st.caption(f"{len(filtered)} ventures")

    for vname in filtered:
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

        vfiles  = load_v_files(vname)
        fb_text = get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
        tr_text = get_text(vfiles["transcript"]) if "transcript" in vfiles else ""
        sp_text = get_text(vfiles["sprint"])     if "sprint"     in vfiles else ""
        signals = detect_signals((notes or "")+" "+(fb_text or "")+" "+(tr_text or ""))

        with st.expander(f"**{vname}**  ·  {hub}  ·  {bucket}"):

            # ── sub-tabs ─────────────────────────────
            tab1, tab2, tab3, tab4 = st.tabs(["📋 Basic Details", "🎙 Sessions", "✦ Success Signals", "🤖 AI Insights"])

            # TAB 1: Basic Details
            with tab1:
                ca, cb = st.columns([3,1])
                with ca:
                    st.markdown(f"### {vname}")
                    st.caption(f"📍 {hub}  ·  👤 VP: {vp}  ·  🏃 Sprint: {sprint}")
                with cb:
                    if rev != "—": st.metric("Revenue LY (Cr)", rev)
                    if tgt != "—": st.metric("3-Yr Target (Cr)", tgt)

                st.markdown(f"**Sprint Completion: {pct_num:.0f}%** — `{bucket}`")
                safe_pct = max(0.0, min(pct_num/100, 1.0)) if pct_num else 0.0
                st.progress(safe_pct)

                # ── RAG scores ───────────────────────
                st.markdown("#### RAG Scores")

                # Get cached scores if available
                cached = st.session_state.get("rag_scores_ai") or st.session_state.get("rag_scores_kw") or []
                v_score = next((s for s in cached if s["name"] == vname), None)

                if v_score:
                    rs1, rs2, rs3 = st.columns(3)
                    rs1.markdown(f"**Overall RAG**<br>{rag_badge(v_score['overall_rag'])}<br><small>{''}</small>", unsafe_allow_html=True)
                    rs2.markdown(f"**Sprint Momentum**<br>{rag_badge(v_score['momentum_rag'])}<br><small>{v_score.get('momentum_reason','')[:80]}</small>", unsafe_allow_html=True)
                    rs3.markdown(f"**Self Investment**<br>{rag_badge(v_score['investment_rag'])}<br><small>{v_score.get('investment_reason','')[:80]}</small>", unsafe_allow_html=True)
                else:
                    st.caption("RAG scores not computed yet. Go to Portfolio Overview and enable scoring first.")

                if notes:
                    st.markdown("**📝 Remarks:**")
                    st.info(notes[:800])

            # TAB 2: Sessions
            with tab2:
                col_a, col_b = st.columns(2)
                # Count sessions from session tracker if available
                col_a.metric("Files Found", len(vfiles))
                col_b.metric("Signal Count", len(signals))

                if fb_text:
                    st.markdown("**📋 Session Feedback:**")
                    st.info(fb_text[:800] + ("..." if len(fb_text)>800 else ""))
                if tr_text:
                    st.markdown("**🎙 Transcript:**")
                    st.info(tr_text[:800] + ("..." if len(tr_text)>800 else ""))
                if not fb_text and not tr_text:
                    st.info("No session files found in venture folder.")

            # TAB 3: Success Signals
            with tab3:
                if signals:
                    for s in signals:
                        st.success(f"✦ **{s['type']}** — keyword: `{s['keyword']}`")
                else:
                    st.info("No success signals detected from available text.")

                # Common docs signals
                if use_sp and sp_reader:
                    st.caption("Note: Common Documents signals will appear here when AI mode is enabled.")

            # TAB 4: AI Insights
            with tab4:
                if not client:
                    st.warning("Add Anthropic API key to enable AI insights.")
                else:
                    if st.button("Generate AI Insight", key=f"ai_{vname}"):
                        with st.spinner("Analyzing..."):
                            try:
                                resp = client.messages.create(
                                    model="claude-sonnet-4-5", max_tokens=600,
                                    messages=[{"role":"user","content":
                                        f"""Venture: {vname}
Hub: {hub} | Sprint: {sprint} | Completion: {pct_num:.0f}%
Notes: {notes[:600]}
Feedback: {fb_text[:400]}
Transcript: {tr_text[:400]}

Provide a concise analysis in plain text (no markdown headers, no # symbols).
Use these section labels followed by a colon:

MOMENTUM: 2-3 sentences on current activity and progress.
RISKS: 2-3 sentences on key concerns to watch.
SIGNALS: List 2-3 success signals observed (one per line, start each with •).
NEXT ACTION: 1-2 sentences on recommended follow-up.
MOMENTUM RAG: Green/Amber/Red — one sentence reason.
INVESTMENT RAG: Green/Amber/Red — one sentence reason."""}])
                                
                                insight_text = resp.content[0].text.strip()
                                # Display in styled container
                                sections = insight_text.splitlines()
                                for line in sections:
                                    line = line.strip()
                                    if not line: continue
                                    if line.startswith("MOMENTUM RAG:") or line.startswith("INVESTMENT RAG:"):
                                        st.caption(f"🎯 {line}")
                                    elif any(line.startswith(x) for x in ["MOMENTUM:","RISKS:","SIGNALS:","NEXT ACTION:"]):
                                        label, _, rest = line.partition(":")
                                        st.markdown(f"**{label.title()}**")
                                        if rest.strip(): st.write(rest.strip())
                                    elif line.startswith("•"):
                                        st.write(line)
                                    else:
                                        st.write(line)
                            except Exception as e:
                                st.error(f"AI error: {e}")
