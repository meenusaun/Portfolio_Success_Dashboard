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
    view = st.radio("📂 View", ["📊 Portfolio Overview", "🏢 Venture Cards"], label_visibility="collapsed")
    view = view.split(" ",1)[1]
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

def load_common_docs():
    """Load all text from Common Documents folder."""
    texts = []
    if use_sp and sp_reader:
        try:
            items = sp_reader.list_files(COMMON_FOLDER)
            for fname in items[:10]:  # limit to avoid timeout
                fp = f"{COMMON_FOLDER}/{fname}"
                try:
                    content = sp_reader.download_file(fp)
                    text = extract_text_bytes(content, fname)
                    if text and len(text) > 50:
                        texts.append(f"[{fname}]\n{text[:2000]}")
                except: pass
        except: pass
    elif root_path:
        cpath = os.path.join(root_path, "Common Documents")
        if os.path.isdir(cpath):
            for f in os.listdir(cpath)[:10]:
                fp = os.path.join(cpath, f)
                if os.path.isfile(fp):
                    text = extract_text_local(fp)
                    if text: texts.append(f"[{f}]\n{text[:2000]}")
    return "\n\n".join(texts)

@st.cache_data(show_spinner=False, ttl=600)
def compute_rag_ai(vname, notes, fb_text, tr_text, common_text, pct_raw, _client_key):
    """Use Claude to compute RAG scores for a venture."""
    if not client: return "Unknown", "Unknown", "No AI key", "No AI key"
    
    try:
        pct = float(str(pct_raw).replace("%","").strip())
        if pct <= 1: pct *= 100
    except: pct = 0

    combined = f"""
Venture: {vname}
Sprint Completion: {pct:.0f}%
Program Notes/Remarks: {notes[:800]}
Session Feedback: {fb_text[:600]}
Transcript: {tr_text[:600]}
Common Documents (relevant excerpts): {common_text[:800]}
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

# ══════════════════════════════════════════════════════
#  VIEW 1: PORTFOLIO OVERVIEW
# ══════════════════════════════════════════════════════
if view == "Portfolio Overview":
    st.title("📊 Portfolio Overview")
    st.caption("NEN Accelerate · All ventures at a glance")

    use_ai_rag = st.toggle("🤖 Compute RAG scores using AI", value=False,
                            help="Uses Claude to score each venture. Takes 1-2 mins for full portfolio.")

    st.divider()

    # ── filters ──────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns(4)
    hub_list = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_raw if cv(get_row(v),col_hub) != "—"))
    vp_list  = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_raw if col_vp and cv(get_row(v),col_vp)  != "—"))
    hub_f    = fc1.selectbox("Hub",             hub_list)
    vp_f     = fc2.selectbox("Venture Partner", vp_list)
    rag_f    = fc3.selectbox("RAG Filter",      ["All","🟢 Green","🟡 Amber","🔴 Red","⚪ ZERO"])
    stage_f  = fc4.selectbox("Sprint Stage",    ["All","0–25%","26–50%","51–75%","76–99%","100%"])

    # ── compute scores ────────────────────────────────
    venture_data = []
    common_text  = ""
    if use_ai_rag:
        with st.spinner("Loading common documents..."):
            common_text = load_common_docs()

    prog = st.progress(0, text="Computing venture scores...")
    for idx, vname in enumerate(ventures_raw):
        row    = get_row(vname)
        hub    = cv(row, col_hub)
        vp     = cv(row, col_vp) if col_vp else "—"
        sprint = cv(row, col_sprint)
        rev    = cv(row, col_rev)
        notes  = cv(row, col_notes, default="")
        pct_raw= row[col_pct] if (row is not None and col_pct) else None
        bucket = get_stage_bucket(pct_raw)

        if use_ai_rag and client:
            vfiles  = load_v_files(vname)
            fb_text = get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
            tr_text = get_text(vfiles["transcript"]) if "transcript" in vfiles else ""
            m_rag, i_rag, m_reason, i_reason = compute_rag_ai(
                vname, notes, fb_text, tr_text, common_text,
                pct_raw, api_key)
        else:
            # Keyword-based fallback
            all_text = notes
            signals  = detect_signals(all_text)
            sig_types= [s["type"] for s in signals]
            i_rag    = "Green" if any(t in sig_types for t in ["Hired","Investment / Spend","Self-Funded Sprint"]) else "ZERO"
            try:
                p = float(str(pct_raw).replace("%","").strip())
                if p <= 1: p *= 100
                m_rag = "Green" if p >= 60 else ("Amber" if p >= 30 else ("Red" if p > 0 else "ZERO"))
            except: m_rag = "ZERO"
            m_reason = f"Sprint completion: {bucket}"
            i_reason = ", ".join(sig_types) if sig_types else "No investment signals found"

        overall = combine_rag(m_rag, i_rag)
        venture_data.append({
            "name": vname, "hub": hub, "vp": vp, "sprint": sprint,
            "rev": rev, "bucket": bucket, "pct_raw": pct_raw,
            "momentum_rag": m_rag, "investment_rag": i_rag,
            "overall_rag": overall,
            "momentum_reason": m_reason, "investment_reason": i_reason,
            "notes": notes
        })
        prog.progress((idx+1)/max(len(ventures_raw),1))
    prog.empty()

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

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Ventures", total)
    c2.metric("🟢 Green", greens, f"{round(greens/total*100) if total else 0}%")
    c3.metric("🟡 Amber", ambers, f"{round(ambers/total*100) if total else 0}%")
    c4.metric("🔴 Red",   reds,   f"{round(reds/total*100)   if total else 0}%")
    c5.metric("⚪ No Data", zeros)

    # ── hub breakdown ─────────────────────────────────
    st.divider()
    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("Hub-wise RAG")
        hubs = {}
        for v in filtered:
            h = v["hub"]
            if h not in hubs: hubs[h] = {"Green":0,"Amber":0,"Red":0,"ZERO":0}
            hubs[h][v["overall_rag"]] = hubs[h].get(v["overall_rag"],0) + 1
        for hub, counts in sorted(hubs.items()):
            total_h = sum(counts.values())
            st.markdown(f"**{hub}** — {total_h} ventures")
            hc1,hc2,hc3 = st.columns(3)
            hc1.metric("🟢", counts.get("Green",0))
            hc2.metric("🟡", counts.get("Amber",0))
            hc3.metric("🔴", counts.get("Red",0))

    with col_r:
        st.subheader("Sprint Stage Distribution")
        stage_counts = {}
        for v in filtered:
            stage_counts[v["bucket"]] = stage_counts.get(v["bucket"],0) + 1
        for stage, cnt in sorted(stage_counts.items()):
            pct = round(cnt/total*100) if total else 0
            st.markdown(f"**{stage}** — {cnt} ventures ({pct}%)")
            st.progress(pct/100)

    # ── venture table ─────────────────────────────────
    st.divider()
    st.subheader(f"Venture List ({len(filtered)} ventures)")

    # Sort by RAG severity
    filtered_sorted = sorted(filtered, key=lambda x: RAG_ORDER.get(x["overall_rag"],4))

    for v in filtered_sorted:
        m_badge = rag_badge(v["momentum_rag"])
        i_badge = rag_badge(v["investment_rag"])
        o_badge = rag_badge(v["overall_rag"])

        with st.expander(f"{RAG_EMOJI.get(v['overall_rag'],'⚪')} {v['name']}  ·  {v['hub']}  ·  {v['bucket']}"):
            r1, r2, r3, r4 = st.columns(4)
            r1.markdown(f"**Overall RAG**<br>{o_badge}", unsafe_allow_html=True)
            r2.markdown(f"**Momentum**<br>{m_badge}", unsafe_allow_html=True)
            r3.markdown(f"**Self Investment**<br>{i_badge}", unsafe_allow_html=True)
            r4.metric("Revenue LY (Cr)", v["rev"])

            if v["momentum_reason"] != "—":
                st.caption(f"📈 Momentum: {v['momentum_reason']}")
            if v["investment_reason"] != "—":
                st.caption(f"💰 Investment: {v['investment_reason']}")
            if v["notes"]:
                st.info(v["notes"][:400])

            # Link to venture card
            if st.button(f"Open Venture Card →", key=f"open_{v['name']}"):
                st.session_state["selected_venture"] = v["name"]
                st.session_state["jump_to_venture"]  = True
                st.rerun()

# ══════════════════════════════════════════════════════
#  VIEW 2: VENTURE CARDS
# ══════════════════════════════════════════════════════
elif view == "Venture Cards":
    st.title("🏢 Venture Cards")
    st.caption("Individual venture deep-dive")
    st.divider()

    # ── filters ──────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns(4)
    search   = fc1.text_input("🔍 Search")
    hub_list = ["All"] + sorted(set(cv(get_row(v),col_hub) for v in ventures_raw if cv(get_row(v),col_hub) != "—"))
    vp_list  = ["All"] + sorted(set(cv(get_row(v),col_vp)  for v in ventures_raw if col_vp and cv(get_row(v),col_vp)  != "—"))
    hub_f    = fc2.selectbox("Hub",             hub_list)
    vp_f     = fc3.selectbox("Venture Partner", vp_list)
    rag_f    = fc4.selectbox("RAG",             ["All","🟢 Green","🟡 Amber","🔴 Red"])

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

Provide a structured analysis:
1. Current Momentum (2-3 sentences)
2. Key Risks to Watch (2-3 sentences)  
3. Success Signals Observed (bullet points)
4. Recommended Next Action (1-2 sentences)

Sprint Momentum RAG: score and reason
Self Investment Signal RAG: score and reason"""}])
                                st.markdown(resp.content[0].text)
                            except Exception as e:
                                st.error(f"AI error: {e}")
