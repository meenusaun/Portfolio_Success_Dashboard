"""
NEN Accelerate — Knowledge Repository Generator (Backend / Admin)
Part 1 of 2: Generates signals_repository.json and feedback_repository.json
This page is for admin use only — not shown to end users.
"""
import streamlit as st
import pandas as pd
import os, json, re, tempfile, shutil
from pathlib import Path
from anthropic import Anthropic
from datetime import datetime

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
ADMIN_PASSWORD    = os.environ.get("ADMIN_PASSWORD", "nenadmin2026")

SP_FOLDER        = "04. Advisors/2026/Portfolio Success Dashboard"
COMMON_FOLDER    = f"{SP_FOLDER}/Common Documents"
TRANSCRIPT_FOLDER= f"{COMMON_FOLDER}/Session Transcripts"
REPO_FOLDER      = f"{COMMON_FOLDER}/Knowledge Repository"
DASHBOARD_FILE   = "0. Journey_Accelerate_Portfolio Dashboard.xlsx"

SIGNALS_REPO_PATH  = f"{REPO_FOLDER}/signals_repository.json"
FEEDBACK_REPO_PATH = f"{REPO_FOLDER}/feedback_repository.json"

st.set_page_config(
    page_title="NEN — Knowledge Repository Generator",
    page_icon="⚙️",
    layout="wide"
)

# ── CSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f5f7fa; }
.main .block-container { padding: 1.5rem 2rem; max-width: 1400px; }
h1,h2,h3,h4 { color: #1e293b !important; }
.stProgress > div > div { background: #6366f1 !important; }
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e2e8f0; }
section[data-testid="stSidebar"] * { color: #1e293b !important; }
div[data-testid="stMetricValue"] { color: #1e293b !important; font-weight: 700; }
div[data-testid="stMetricLabel"] { font-size: 0.82rem; color: #64748b; }
.stExpander { background: #ffffff; border: 1px solid #e2e8f0 !important; border-radius: 10px; margin-bottom: 8px; }
.stButton > button { background: #6366f1 !important; color: white !important; border: none !important; border-radius: 6px !important; font-weight: 600; }
.stButton > button:hover { background: #4f46e5 !important; }
.admin-badge { background: #7c3aed; color: white; padding: 3px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; letter-spacing: 0.05em; }
.rag-green { background:#dcfce7; color:#166534; padding:3px 10px; border-radius:20px; font-weight:700; font-size:0.82rem; }
.rag-amber { background:#fef9c3; color:#854d0e; padding:3px 10px; border-radius:20px; font-weight:700; font-size:0.82rem; }
.rag-red   { background:#fee2e2; color:#991b1b; padding:3px 10px; border-radius:20px; font-weight:700; font-size:0.82rem; }
.rag-zero  { background:#f1f5f9; color:#64748b; padding:3px 10px; border-radius:20px; font-weight:700; font-size:0.82rem; }
.status-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 16px; margin: 8px 0; font-size: 0.82rem; }
</style>
""", unsafe_allow_html=True)

# ── admin password gate ────────────────────────────────
def check_admin():
    if st.session_state.get("admin_authenticated"): return True
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center;padding:40px 30px;background:#ffffff;
        border-radius:16px;border:1px solid #e2e8f0;box-shadow:0 4px 24px rgba(0,0,0,0.06)'>
            <div style='font-size:2rem;margin-bottom:8px'>⚙️</div>
            <div style='font-size:1.2rem;font-weight:700;color:#1e293b;margin-bottom:4px'>
                Knowledge Repository Generator</div>
            <div style='font-size:0.82rem;color:#64748b;margin-bottom:12px'>
                NEN Accelerate · Admin Access Only</div>
            <span class='admin-badge'>BACKEND</span>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        pwd = st.text_input("Admin Password", type="password",
                            placeholder="Enter admin password...",
                            label_visibility="collapsed")
        _, cb, _ = st.columns([1,2,1])
        with cb:
            if st.button("🔐  Login", use_container_width=True):
                if pwd in [ADMIN_PASSWORD, ENV_PASSWORD]:
                    st.session_state["admin_authenticated"] = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
    return False

if not check_admin(): st.stop()

# ── header ─────────────────────────────────────────────
st.markdown("""
<div style='display:flex;align-items:center;gap:14px;margin-bottom:8px'>
    <div style='font-size:1.8rem'>⚙️</div>
    <div>
        <div style='font-size:1.3rem;font-weight:700;color:#1e293b'>Knowledge Repository Generator</div>
        <div style='font-size:0.8rem;color:#64748b'>
            NEN Accelerate · Admin · Generates signals_repository.json + feedback_repository.json
        </div>
    </div>
    <span class='admin-badge' style='margin-left:auto'>BACKEND ONLY</span>
</div>
""", unsafe_allow_html=True)
st.divider()

# ── sidebar ────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    use_sp = st.toggle("☁️ Use SharePoint", value=bool(ENV_CLIENT_ID))
    if use_sp and ENV_CLIENT_ID:
        st.success("✅ SharePoint configured")
    elif use_sp:
        st.warning("Add Azure credentials to secrets")
    api_key = ENV_API_KEY or st.text_input("Anthropic API Key", type="password")
    st.markdown("---")
    st.markdown("**📁 Repository paths:**")
    st.code(f"Signals:\n{SIGNALS_REPO_PATH}\n\nFeedback:\n{FEEDBACK_REPO_PATH}", language="text")
    st.markdown("---")
    st.caption("NEN Accelerate · Backend\nResources Network Team")

# ── SharePoint connection ──────────────────────────────
sp_reader = None
if use_sp and ENV_CLIENT_ID:
    try:
        from sharepoint_reader import SharePointReader
        if "sp_reader_backend" not in st.session_state:
            with st.spinner("Connecting to SharePoint..."):
                st.session_state["sp_reader_backend"] = SharePointReader(
                    ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        sp_reader = st.session_state["sp_reader_backend"]
    except Exception as e:
        st.error(f"SharePoint connection failed: {e}"); st.stop()

client = Anthropic(api_key=api_key) if api_key else None

# ── helpers ────────────────────────────────────────────
def safe_copy(src):
    dst = os.path.join(tempfile.gettempdir(), "nen_bk_" + os.path.basename(src))
    shutil.copy2(src, dst); return dst

def find_col(df, patterns):
    for c in df.columns:
        for p in patterns:
            if p.lower() in str(c).lower(): return c
    return None

def extract_text_bytes(content_bytes, filename):
    ext = Path(filename).suffix.lower()
    try:
        tmp = os.path.join(tempfile.gettempdir(), f"nen_bk_{filename}")
        with open(tmp,"wb") as f: f.write(content_bytes)
        if ext in [".xlsx",".xls"]:
            xl = pd.ExcelFile(tmp)
            return "\n".join(pd.read_excel(tmp,sheet_name=s,header=None).fillna("").astype(str).to_string()
                             for s in xl.sheet_names)
        elif ext == ".docx":
            from docx import Document
            return "\n".join(p.text for p in Document(tmp).paragraphs)
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(tmp) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e: return f"[Error: {e}]"
    return ""

def sp_get_text(fpath):
    if sp_reader:
        try:
            content = sp_reader.download_file(fpath)
            return extract_text_bytes(content, Path(fpath).name)
        except: return ""
    return ""

def sp_list(folder):
    if sp_reader:
        try: return sp_reader.list_folder(folder)
        except: return []
    return []

# ── load dashboard ─────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_dashboard(_sp_id, use_sp):
    try:
        if use_sp and ENV_CLIENT_ID:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            content = sp.download_file(f"{SP_FOLDER}/{DASHBOARD_FILE}")
            tmp = os.path.join(tempfile.gettempdir(), "nen_bk_dashboard.xlsx")
            with open(tmp,"wb") as f: f.write(content)
            fpath = tmp
        else:
            return None, "SharePoint not configured"
        company_df = pd.read_excel(fpath, sheet_name="Company", header=2)
        return company_df, None
    except Exception as e: return None, str(e)

with st.spinner("Loading portfolio data..."):
    sp_id = id(sp_reader) if sp_reader else 0
    company_df, err = load_dashboard(sp_id, use_sp)

if err: st.error(f"❌ Dashboard load failed: {err}"); st.stop()
if company_df is None: st.error("Could not load dashboard."); st.stop()

company_df.columns = [str(c).strip() for c in company_df.columns]
col_name   = find_col(company_df, ["company name"])
col_hub    = find_col(company_df, ["hub / state","hub","state"])
col_notes  = find_col(company_df, ["notes / comments","notes","remarks"])
col_rev    = find_col(company_df, ["revenue ly"])
col_vp     = find_col(company_df, ["venture partner","venture manager","vp"])
col_sprint = None
for c in company_df.columns:
    if "sprint type" in str(c).lower(): col_sprint = c; break
if not col_sprint:
    for c in company_df.columns:
        cl = str(c).lower()
        if "sprint" in cl and not any(x in cl for x in ["commit","complet","score","%","task"]):
            col_sprint = c; break
col_pct = find_col(company_df, ["% sprint completion","sprint completion"])
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

st.success(f"✅ Dashboard loaded — {len(ventures_list)} ventures found")

# ── load venture files ─────────────────────────────────
def load_v_files(vname):
    files = {}
    VALID_EXT = [".xlsx",".xls",".docx",".pdf",".pptx",".ppt",".txt"]
    if use_sp and sp_reader:
        folder = f"{SP_FOLDER}/{vname}"
        try:
            items = sp_reader.list_files(folder)
            for fname in items:
                fl  = fname.lower()
                fp  = f"{folder}/{fname}"
                ext = Path(fname).suffix.lower()
                if ext not in VALID_EXT: continue
                if "transcript"   in fl: files["transcript"] = fp
                elif "feedback"   in fl: files["feedback"]   = fp
                elif "growth sprint" in fl or "sprint plan" in fl: files["sprint"] = fp
                elif "growth journey" in fl or "journey report" in fl: files["journey"] = fp
                else:
                    key = f"other_{len([k for k in files if k.startswith('other')])}"
                    files[key] = fp
        except: pass
    return files

# ── load common docs ───────────────────────────────────
@st.cache_data(show_spinner=False, ttl=600)
def load_common_docs_cached(_sp_id, use_sp):
    texts = []
    def process_file(fname, content_bytes=None):
        ext = Path(fname).suffix.lower()
        if ext not in [".xlsx",".xls",".docx",".pdf",".pptx",".ppt"]: return
        if "journey_accelerate_portfolio" in fname.lower(): return
        try:
            text = extract_text_bytes(content_bytes, fname)
            if text and len(text) >= 50:
                texts.append(f"=== FILE: {fname} ===\n{text}")
        except: pass

    if use_sp and ENV_CLIENT_ID:
        try:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
            def scan_sp_folder(folder_path):
                try:
                    items = sp.list_folder(folder_path)
                    for item in items:
                        iname = item.get("name","")
                        ipath = f"{folder_path}/{iname}"
                        if "folder" in item: scan_sp_folder(ipath)
                        elif "file" in item:
                            try:
                                content = sp.download_file(ipath)
                                process_file(iname, content_bytes=content)
                            except: pass
                except: pass
            scan_sp_folder(COMMON_FOLDER)
        except: pass
    return "\n\n".join(texts)

def extract_venture_from_common(venture_name, common_text):
    if not venture_name: return ""
    relevant = []
    try:
        if col_name and company_df is not None:
            match = all_rows[all_rows[col_name].astype(str).str.strip() == venture_name]
            if not match.empty:
                row = match.iloc[0]
                row_parts = []
                for col in company_df.columns:
                    val = str(row[col]).strip()
                    if val and val not in ["nan","None","NaT",""]:
                        row_parts.append(f"{col}: {val}")
                if row_parts:
                    relevant.append(f"[Portfolio Dashboard — {venture_name}]\n" + "\n".join(row_parts))
    except: pass
    if common_text:
        sections = common_text.split("=== FILE:")
        for section in sections:
            if not section.strip(): continue
            if venture_name.lower() not in section.lower(): continue
            lines = section.split("\n")
            venture_lines = []
            i = 0
            while i < len(lines):
                if venture_name.lower() in lines[i].lower():
                    start = max(0, i-5)
                    end   = min(len(lines), i+15)
                    venture_lines.extend(lines[start:end])
                    venture_lines.append("---")
                    i = end
                else: i += 1
            if venture_lines:
                fname_part = section.split("===")[0].strip() if "===" in section else ""
                relevant.append(f"[{fname_part}]\n" + "\n".join(venture_lines))
    return "\n\n".join(relevant)

# ── load session transcripts from Session Transcripts folder ──
@st.cache_data(show_spinner=False, ttl=600)
def load_session_transcripts_cached(_sp_id, use_sp):
    """Load all files from Common Documents/Session Transcripts/ folder."""
    transcripts = {}  # {venture_name_lower: [{"filename": ..., "text": ...}]}
    if not use_sp or not ENV_CLIENT_ID:
        return transcripts
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        def scan_folder(folder_path):
            try:
                items = sp.list_folder(folder_path)
                for item in items:
                    iname = item.get("name","")
                    ipath = f"{folder_path}/{iname}"
                    if "folder" in item:
                        scan_folder(ipath)
                    elif "file" in item:
                        ext = Path(iname).suffix.lower()
                        if ext not in [".docx",".pdf",".xlsx",".xls",".txt"]: continue
                        try:
                            content = sp.download_file(ipath)
                            text = extract_text_bytes(content, iname)
                            if text and len(text) > 50:
                                # Match transcript to venture by filename
                                fname_lower = iname.lower()
                                transcripts.setdefault("_all", []).append({
                                    "filename": iname,
                                    "path":     ipath,
                                    "text":     text,
                                })
                        except: pass
            except: pass
        scan_folder(TRANSCRIPT_FOLDER)
    except: pass
    return transcripts

def get_transcript_for_venture(vname, all_transcripts):
    """Find transcript files matching a venture name."""
    matched = []
    vname_lower = vname.lower()
    vwords = [w for w in vname_lower.split() if len(w) > 3]
    for entry in all_transcripts.get("_all", []):
        fname_lower = entry["filename"].lower()
        text_lower  = entry["text"].lower()
        # Match by filename or text content
        name_match = vname_lower in fname_lower or any(w in fname_lower for w in vwords)
        text_match = vname_lower in text_lower[:2000]  # Check first 2000 chars
        if name_match or text_match:
            matched.append(entry["text"])
    return "\n\n".join(matched) if matched else ""

# ── attendance ─────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=600)
def load_attendance_cached(_sp_id, use_sp):
    attendance = {}
    ATTENDANCE_KEYWORDS = ["attendance", "d&v attendance"]
    if not use_sp or not ENV_CLIENT_ID: return attendance
    try:
        from sharepoint_reader import SharePointReader
        sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
        def find_in_sp(folder_path):
            try:
                items = sp.list_folder(folder_path)
                for item in items:
                    iname = item.get("name","")
                    ipath = f"{folder_path}/{iname}"
                    if "folder" in item:
                        r = find_in_sp(ipath)
                        if r[0]: return r
                    elif any(kw in iname.lower() for kw in ATTENDANCE_KEYWORDS):
                        return sp.download_file(ipath), iname
            except: pass
            return None, None
        content_bytes, fname_found = find_in_sp(COMMON_FOLDER)
        if not content_bytes: return attendance
        ext = Path(fname_found).suffix.lower()
        tmp = os.path.join(tempfile.gettempdir(), f"nen_bk_att{ext}")
        with open(tmp,"wb") as f: f.write(content_bytes)
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(tmp) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            date_pattern = re.compile(r'\d{2}-\d{2}')
            for line in text.split("\n"):
                line = line.strip()
                if not line: continue
                dates = date_pattern.findall(line)
                if dates:
                    first_date_idx = line.find(dates[0])
                    vname_raw = re.sub(r'\s*\([^)]*\)\s*$','',line[:first_date_idx].strip()).strip()
                    if vname_raw and len(vname_raw) > 3:
                        attendance[vname_raw] = {"sessions":len(dates),"dates":dates}
        elif ext in [".xlsx",".xls"]:
            xl = pd.ExcelFile(tmp)
            for sheet in xl.sheet_names:
                df = pd.read_excel(tmp,sheet_name=sheet,header=None)
                date_pattern = re.compile(r'\d{2}-\d{2}')
                for _, row in df.iterrows():
                    row_vals = [str(v).strip() for v in row if str(v).strip() not in ["nan","None",""]]
                    if len(row_vals) < 2: continue
                    dates = [v for v in row_vals[1:] if date_pattern.match(v)]
                    if dates:
                        vname_clean = re.sub(r'\s*\([^)]*\)\s*$','',row_vals[0]).strip()
                        if vname_clean and len(vname_clean) > 3:
                            attendance[vname_clean] = {"sessions":len(dates),"dates":dates}
    except: pass
    return attendance

def get_attendance_for_venture(vname, attendance_data):
    if not attendance_data: return None
    if vname in attendance_data: return attendance_data[vname]
    vname_lower = vname.lower()
    for att_name, data in attendance_data.items():
        att_lower = att_name.lower()
        v_words = [w for w in vname_lower.split() if len(w) > 3]
        a_words = [w for w in att_lower.split() if len(w) > 3]
        matches = sum(1 for w in v_words if any(w in aw or aw in w for aw in a_words))
        if matches >= 2 or (len(v_words) == 1 and matches == 1):
            return data
    return None

# ══════════════════════════════════════════════════════
#  MAIN UI — Two generation steps
# ══════════════════════════════════════════════════════

step1_tab, step2_tab, status_tab = st.tabs([
    "📊 Step 1 — Signals Repository",
    "🎙 Step 2 — Feedback Repository",
    "📁 Status & Downloads"
])

# ══════════════════════════════════════════════════════
#  STEP 1: SIGNALS REPOSITORY
# ══════════════════════════════════════════════════════
with step1_tab:
    st.markdown("### 📊 Generate Signals Repository")
    st.caption(
        "Reads all venture folders + Common Documents → Extracts signals → "
        "Scores RAG → Saves to `signals_repository.json`"
    )

    if not client:
        st.error("❌ Anthropic API key required. Add ANTHROPIC_API_KEY to secrets."); st.stop()

    # Controls row
    ctrl1, ctrl2, ctrl3 = st.columns([2,1,1])
    with ctrl1:
        venture_filter = st.multiselect(
            "Ventures to process (leave empty = all)",
            options=ventures_list, key="sig_filter"
        )
    with ctrl2:
        batch_size = st.selectbox("Batch size", [5, 10, 15, 20], index=1, key="sig_batch_size")
    with ctrl3:
        st.markdown("<br>", unsafe_allow_html=True)
        run_all_sig = st.button("▶ Generate All Signals", key="run_all_signals",
                                use_container_width=True)

    target_ventures = venture_filter if venture_filter else ventures_list
    batches = [target_ventures[i:i+batch_size] for i in range(0, len(target_ventures), batch_size)]

    st.caption(f"{len(target_ventures)} ventures · {len(batches)} batches of {batch_size}")
    st.divider()

    SIG_RESULTS_KEY = "sig_repo_results"
    if SIG_RESULTS_KEY not in st.session_state:
        st.session_state[SIG_RESULTS_KEY] = {}

    sig_results = st.session_state[SIG_RESULTS_KEY]
    done_sig = sum(1 for v in sig_results.values() if v.get("status") == "done")

    # ── Step 0: Explicit pre-load (never triggers on page render) ──
    preload_done = ("common_text_sig" in st.session_state and
                    "att_data_sig"    in st.session_state)

    if not preload_done:
        st.info("**Step 0 — Pre-load Common Documents first** before running any batch.")
        if st.button("📂 Pre-load Common Documents & Attendance", key="preload_btn"):
            status_box = st.empty()
            prog_pre   = st.progress(0, text="Scanning Common Documents folder...")
            file_counter   = [0]
            texts_collected = []

            def process_file_progress(fname, content_bytes=None):
                ext = Path(fname).suffix.lower()
                if ext not in [".xlsx",".xls",".docx",".pdf",".pptx",".ppt"]: return
                if "journey_accelerate_portfolio" in fname.lower(): return
                try:
                    text = extract_text_bytes(content_bytes, fname)
                    if text and len(text) >= 50:
                        texts_collected.append(f"=== FILE: {fname} ===\n{text}")
                        file_counter[0] += 1
                        status_box.caption(f"📄 Loaded {file_counter[0]} files... latest: {fname[:60]}")
                except: pass

            if use_sp and ENV_CLIENT_ID:
                try:
                    from sharepoint_reader import SharePointReader
                    sp_pre = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)
                    all_files = []
                    def collect_files(folder_path):
                        try:
                            items = sp_pre.list_folder(folder_path)
                            for item in items:
                                iname = item.get("name","")
                                ipath = f"{folder_path}/{iname}"
                                if "folder" in item: collect_files(ipath)
                                elif "file" in item:
                                    ext = Path(iname).suffix.lower()
                                    if ext in [".xlsx",".xls",".docx",".pdf",".pptx",".ppt"]:
                                        if "journey_accelerate_portfolio" not in iname.lower():
                                            all_files.append((iname, ipath))
                        except: pass
                    collect_files(COMMON_FOLDER)
                    total_files = len(all_files)
                    prog_pre.progress(0, text=f"Found {total_files} files — downloading...")
                    for fi, (fname, fpath) in enumerate(all_files):
                        prog_pre.progress((fi+1)/max(total_files,1),
                                          text=f"Reading {fi+1}/{total_files}: {fname[:50]}")
                        try:
                            content = sp_pre.download_file(fpath)
                            process_file_progress(fname, content_bytes=content)
                        except: pass
                except Exception as e:
                    st.error(f"Common docs load error: {e}")

            st.session_state["common_text_sig"] = "\n\n".join(texts_collected)
            status_box.caption(f"📅 Loading attendance data...")
            att = load_attendance_cached(sp_id, use_sp)
            st.session_state["att_data_sig"] = att
            prog_pre.progress(1.0, text=f"✅ Done — {file_counter[0]} files loaded")
            st.success(f"✅ Pre-load complete — {file_counter[0]} Common Document files · {len(att)} attendance records")
            st.rerun()
        st.stop()

    # Already pre-loaded
    common_text_sig = st.session_state["common_text_sig"]
    att_data_sig    = st.session_state["att_data_sig"]

    cdocs_files = common_text_sig.count("=== FILE:")
    col_pre1, col_pre2 = st.columns([4,1])
    col_pre1.success(f"✅ Common Documents loaded — {cdocs_files} files · {len(att_data_sig)} attendance records")
    if col_pre2.button("🔄 Reload Docs", key="reload_preload"):
        del st.session_state["common_text_sig"]
        del st.session_state["att_data_sig"]
        st.rerun()

    for bi, batch in enumerate(batches):
        batch_done = sum(1 for v in batch if sig_results.get(v,{}).get("status") == "done")
        icon = "✅" if batch_done == len(batch) else ("🔄" if batch_done > 0 else "⬜")
        with st.expander(f"{icon} Batch {bi+1}  —  {batch_done}/{len(batch)} done  —  {', '.join(batch[:3])}{'...' if len(batch)>3 else ''}"):
            for vn in batch:
                vr = sig_results.get(vn,{})
                vs = vr.get("status","pending")
                ic = {"done":"✅","error":"❌","processing":"🔄"}.get(vs,"⬜")
                if vs == "done":
                    rag = vr.get("rag",{})
                    nsig = len(vr.get("signals",{}).get("momentum",[])) + \
                           len(vr.get("signals",{}).get("investment",[]))
                    st.caption(f"{ic} {vn}  ·  Overall: {rag.get('overall_rag','')}  ·  "
                               f"{nsig} signals  ·  {vr.get('total_chars',0):,} chars")
                elif vs == "error":
                    st.caption(f"{ic} {vn}  ·  Error: {vr.get('error','')[:80]}")
                else:
                    st.caption(f"{ic} {vn}  ·  Not processed")

            st.markdown("")
            run_batch_btn = st.button(f"▶ Run Batch {bi+1}", key=f"run_sig_batch_{bi}")
            should_run = run_batch_btn or run_all_sig

            if should_run:

                prog = st.progress(0, text=f"Starting batch {bi+1}...")
                for vi, vname in enumerate(batch):
                    if sig_results.get(vname,{}).get("status") == "done":
                        prog.progress((vi+1)/len(batch), text=f"Skipping {vname} (done)")
                        continue
                    prog.progress(vi/len(batch), text=f"Processing {vname} ({vi+1}/{len(batch)})...")

                    row    = get_row(vname)
                    notes  = cv(row, col_notes, default="")
                    sprint = cv(row, col_sprint)
                    pct    = row[col_pct] if (row is not None and col_pct) else None
                    hub    = cv(row, col_hub)
                    vp     = cv(row, col_vp) if col_vp else "—"

                    # Load venture files
                    vfiles  = load_v_files(vname)
                    fb_text = sp_get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
                    tr_text = sp_get_text(vfiles["transcript"]) if "transcript" in vfiles else ""
                    sp_text = sp_get_text(vfiles["sprint"])     if "sprint"     in vfiles else ""
                    jr_text = sp_get_text(vfiles["journey"])    if "journey"    in vfiles else ""
                    others  = [sp_get_text(p) for k,p in vfiles.items() if k.startswith("other_")]
                    others  = [t for t in others if t]

                    venture_common = extract_venture_from_common(vname, common_text_sig)

                    sources = {
                        "Notes": notes or "", "Feedback": fb_text,
                        "Transcript": tr_text, "Sprint Plan": sp_text,
                        "Growth Journey": jr_text, "Common Docs": venture_common,
                    }
                    for idx, ot in enumerate(others):
                        sources[f"Venture File {idx+1}"] = ot

                    full_text   = "\n\n".join(f"=== {k} ===\n{v}" for k,v in sources.items() if v)
                    total_chars = len(full_text)
                    srcs_used   = [k for k,v in sources.items() if v]

                    att  = get_attendance_for_venture(vname, att_data_sig)
                    att_sessions = att["sessions"] if att else 0
                    att_dates    = att["dates"]    if att else []
                    att_summary  = (f"{att_sessions} sessions ({', '.join(att_dates)})"
                                    if att_sessions else "No attendance data")

                    from processor import extract_signals_from_text, score_rag_from_signals
                    signals, num_chunks = extract_signals_from_text(client, vname, sprint, full_text)
                    rag = score_rag_from_signals(client, vname, sprint, notes,
                                                  att_summary, signals, pct)

                    sig_results[vname] = {
                        "status":        "done",
                        "venture_name":  vname,
                        "hub":           hub,
                        "venture_partner": vp,
                        "sprint":        sprint,
                        "rag":           rag,
                        "signals":       signals,
                        "total_chars":   total_chars,
                        "num_chunks":    num_chunks,
                        "sources_used":  srcs_used,
                        "att_sessions":  att_sessions,
                        "att_dates":     att_dates,
                        "processed_at":  datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                    }
                    st.session_state[SIG_RESULTS_KEY] = sig_results

                prog.progress(1.0, text=f"✅ Batch {bi+1} complete!")
                st.rerun()

    # ── download / upload signals repo ─────────────────
    st.divider()
    done_count_s = sum(1 for v in sig_results.values() if v.get("status") == "done")
    st.markdown(f"**{done_count_s}/{len(ventures_list)} ventures processed**")

    if done_count_s > 0:
        dl_col, up_col = st.columns(2)
        with dl_col:
            st.markdown("**💾 Download signals_repository.json**")
            # Build flat records for repository
            repo_records = []
            for vn, vr in sig_results.items():
                if vr.get("status") != "done": continue
                rag = vr.get("rag",{})
                for sig in vr.get("signals",{}).get("momentum",[]):
                    repo_records.append({
                        "venture_name":    vn,
                        "hub":             vr.get("hub","—"),
                        "venture_partner": vr.get("venture_partner","—"),
                        "sprint":          vr.get("sprint","—"),
                        "signal_type":     "Momentum",
                        "rag":             rag.get("momentum_rag","ZERO"),
                        "rag_description": rag.get("momentum_reason","—"),
                        "signal_category": sig.get("type",""),
                        "evidence":        sig.get("evidence",""),
                        "source":          sig.get("source",""),
                        "overall_rag":     rag.get("overall_rag","ZERO"),
                        "momentum_rag":    rag.get("momentum_rag","ZERO"),
                        "investment_rag":  rag.get("investment_rag","ZERO"),
                        "processed_at":    vr.get("processed_at",""),
                    })
                for sig in vr.get("signals",{}).get("investment",[]):
                    repo_records.append({
                        "venture_name":    vn,
                        "hub":             vr.get("hub","—"),
                        "venture_partner": vr.get("venture_partner","—"),
                        "sprint":          vr.get("sprint","—"),
                        "signal_type":     "Self Investment",
                        "rag":             rag.get("investment_rag","ZERO"),
                        "rag_description": rag.get("investment_reason","—"),
                        "signal_category": sig.get("type",""),
                        "evidence":        sig.get("evidence",""),
                        "source":          sig.get("source",""),
                        "overall_rag":     rag.get("overall_rag","ZERO"),
                        "momentum_rag":    rag.get("momentum_rag","ZERO"),
                        "investment_rag":  rag.get("investment_rag","ZERO"),
                        "processed_at":    vr.get("processed_at",""),
                    })
                # Add summary row per venture (no individual signal)
                if not vr.get("signals",{}).get("momentum") and \
                   not vr.get("signals",{}).get("investment"):
                    repo_records.append({
                        "venture_name":    vn,
                        "hub":             vr.get("hub","—"),
                        "venture_partner": vr.get("venture_partner","—"),
                        "sprint":          vr.get("sprint","—"),
                        "signal_type":     "No Signals",
                        "rag":             rag.get("overall_rag","ZERO"),
                        "rag_description": "No signals extracted",
                        "signal_category": "",
                        "evidence":        "",
                        "source":          "",
                        "overall_rag":     rag.get("overall_rag","ZERO"),
                        "momentum_rag":    rag.get("momentum_rag","ZERO"),
                        "investment_rag":  rag.get("investment_rag","ZERO"),
                        "processed_at":    vr.get("processed_at",""),
                    })

            save_payload = {
                "generated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                "venture_count": done_count_s,
                "records":       repo_records,
                "venture_summary": {
                    vn: {
                        "venture_name":    vn,
                        "hub":             vr.get("hub","—"),
                        "venture_partner": vr.get("venture_partner","—"),
                        "sprint":          vr.get("sprint","—"),
                        "overall_rag":     vr.get("rag",{}).get("overall_rag","ZERO"),
                        "momentum_rag":    vr.get("rag",{}).get("momentum_rag","ZERO"),
                        "investment_rag":  vr.get("rag",{}).get("investment_rag","ZERO"),
                        "momentum_reason": vr.get("rag",{}).get("momentum_reason","—"),
                        "investment_reason":vr.get("rag",{}).get("investment_reason","—"),
                        "momentum_score":  vr.get("rag",{}).get("momentum_score",0),
                        "investment_score":vr.get("rag",{}).get("investment_score",0),
                        "signals":         vr.get("signals",{}),
                        "att_sessions":    vr.get("att_sessions",0),
                        "att_dates":       vr.get("att_dates",[]),
                        "sources_used":    vr.get("sources_used",[]),
                        "processed_at":    vr.get("processed_at",""),
                    }
                    for vn, vr in sig_results.items() if vr.get("status") == "done"
                }
            }
            st.download_button(
                "⬇️ Download signals_repository.json",
                data=json.dumps(save_payload, indent=2, default=str),
                file_name="signals_repository.json",
                mime="application/json",
            )
            st.caption(f"Upload this file to SharePoint at:\n`{SIGNALS_REPO_PATH}`")

# ══════════════════════════════════════════════════════
#  STEP 2: FEEDBACK REPOSITORY
# ══════════════════════════════════════════════════════
with step2_tab:
    st.markdown("### 🎙 Generate Feedback Repository")
    st.caption(
        "Reads venture feedback files + Session Transcripts folder → "
        "Extracts session data → Saves to `feedback_repository.json`"
    )

    if not client:
        st.error("❌ Anthropic API key required."); st.stop()

    fb_ctrl1, fb_ctrl2, fb_ctrl3 = st.columns([2,1,1])
    with fb_ctrl1:
        fb_venture_filter = st.multiselect(
            "Ventures to process (leave empty = all)",
            options=ventures_list, key="fb_filter"
        )
    with fb_ctrl2:
        fb_batch_size = st.selectbox("Batch size", [5, 10, 15, 20], index=1, key="fb_batch_size")
    with fb_ctrl3:
        st.markdown("<br>", unsafe_allow_html=True)
        run_all_fb = st.button("▶ Generate All Feedback", key="run_all_feedback",
                               use_container_width=True)

    fb_target = fb_venture_filter if fb_venture_filter else ventures_list
    fb_batches = [fb_target[i:i+fb_batch_size] for i in range(0, len(fb_target), fb_batch_size)]

    st.caption(f"{len(fb_target)} ventures · {len(fb_batches)} batches of {fb_batch_size}")
    st.divider()

    FB_RESULTS_KEY = "fb_repo_results"
    if FB_RESULTS_KEY not in st.session_state:
        st.session_state[FB_RESULTS_KEY] = {}

    fb_results = st.session_state[FB_RESULTS_KEY]

    if run_all_fb or any(
        st.session_state.get(f"run_fb_batch_{b}", False) for b in range(len(fb_batches))
    ):
        with st.spinner("📂 Loading Session Transcripts from SharePoint..."):
            all_transcripts = load_session_transcripts_cached(sp_id, use_sp)
        st.info(f"📁 Found {len(all_transcripts.get('_all',[]))} transcript files in Session Transcripts folder")
    else:
        all_transcripts = None

    for bi, batch in enumerate(fb_batches):
        batch_done = sum(1 for v in batch if fb_results.get(v,{}).get("status") == "done")
        icon = "✅" if batch_done == len(batch) else ("🔄" if batch_done > 0 else "⬜")
        with st.expander(f"{icon} Batch {bi+1}  —  {batch_done}/{len(batch)} done  —  {', '.join(batch[:3])}{'...' if len(batch)>3 else ''}"):
            for vn in batch:
                vr = fb_results.get(vn,{})
                vs = vr.get("status","pending")
                ic = {"done":"✅","error":"❌"}.get(vs,"⬜")
                if vs == "done":
                    nsess = len(vr.get("sessions",[]))
                    st.caption(f"{ic} {vn}  ·  {nsess} session(s) extracted")
                else:
                    st.caption(f"{ic} {vn}  ·  Not processed")

            st.markdown("")
            run_fb_btn = st.button(f"▶ Run Batch {bi+1}", key=f"run_fb_batch_{bi}")
            should_run_fb = run_fb_btn or run_all_fb

            if should_run_fb:
                if all_transcripts is None:
                    all_transcripts = load_session_transcripts_cached(sp_id, use_sp)

                prog_fb = st.progress(0, text=f"Starting batch {bi+1}...")
                for vi, vname in enumerate(batch):
                    if fb_results.get(vname,{}).get("status") == "done":
                        prog_fb.progress((vi+1)/len(batch), text=f"Skipping {vname}")
                        continue
                    prog_fb.progress(vi/len(batch), text=f"Processing {vname} ({vi+1}/{len(batch)})...")

                    row = get_row(vname)
                    hub = cv(row, col_hub)
                    vp  = cv(row, col_vp) if col_vp else "—"

                    # Load venture feedback file
                    vfiles  = load_v_files(vname)
                    fb_text = sp_get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
                    tr_text = sp_get_text(vfiles["transcript"]) if "transcript" in vfiles else ""

                    # Also check Session Transcripts folder
                    session_tr_text = get_transcript_for_venture(vname, all_transcripts)

                    # Combine all transcript sources
                    combined_transcript = "\n\n".join(t for t in [tr_text, session_tr_text] if t)

                    from processor import extract_session_feedback
                    sessions = extract_session_feedback(
                        client, vname, combined_transcript, fb_text)

                    fb_results[vname] = {
                        "status":          "done",
                        "venture_name":    vname,
                        "hub":             hub,
                        "venture_partner": vp,
                        "sessions":        sessions,
                        "processed_at":    datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                    }
                    st.session_state[FB_RESULTS_KEY] = fb_results

                prog_fb.progress(1.0, text=f"✅ Batch {bi+1} complete!")
                st.rerun()

    # ── download feedback repo ──────────────────────────
    st.divider()
    done_count_f = sum(1 for v in fb_results.values() if v.get("status") == "done")
    st.markdown(f"**{done_count_f}/{len(ventures_list)} ventures processed**")

    if done_count_f > 0:
        fb_dl_col, _ = st.columns(2)
        with fb_dl_col:
            st.markdown("**💾 Download feedback_repository.json**")
            fb_payload = {
                "generated_at":   datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                "venture_count":  done_count_f,
                "ventures":       {
                    vn: vr for vn, vr in fb_results.items()
                    if vr.get("status") == "done"
                }
            }
            st.download_button(
                "⬇️ Download feedback_repository.json",
                data=json.dumps(fb_payload, indent=2, default=str),
                file_name="feedback_repository.json",
                mime="application/json",
            )
            st.caption(f"Upload this file to SharePoint at:\n`{FEEDBACK_REPO_PATH}`")

# ══════════════════════════════════════════════════════
#  STATUS TAB
# ══════════════════════════════════════════════════════
with status_tab:
    st.markdown("### 📁 Repository Status & Download")

    sig_results  = st.session_state.get(SIG_RESULTS_KEY, {})
    fb_results   = st.session_state.get(FB_RESULTS_KEY, {})
    done_s = sum(1 for v in sig_results.values() if v.get("status") == "done")
    done_f = sum(1 for v in fb_results.values()  if v.get("status") == "done")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Ventures",       len(ventures_list))
    m2.metric("Signals Processed",    done_s)
    m3.metric("Feedback Processed",   done_f)
    m4.metric("Both Complete",
              sum(1 for v in ventures_list
                  if sig_results.get(v,{}).get("status") == "done"
                  and fb_results.get(v,{}).get("status")  == "done"))

    st.divider()
    st.markdown("#### Per-Venture Status")

    RAG_EMOJI = {"Green":"🟢","Amber":"🟡","Red":"🔴","ZERO":"⚪","Unknown":"⚪"}

    h0,h1,h2,h3,h4,h5 = st.columns([2.5, 1.0, 0.8, 0.8, 0.8, 1.5])
    h0.markdown("**Venture**"); h1.markdown("**Hub**")
    h2.markdown("**Signals**"); h3.markdown("**RAG**")
    h4.markdown("**Feedback**"); h5.markdown("**Processed At**")
    st.divider()

    for vname in ventures_list:
        row    = get_row(vname)
        hub    = cv(row, col_hub)
        sr     = sig_results.get(vname,{})
        fr     = fb_results.get(vname,{})
        s_done = sr.get("status") == "done"
        f_done = fr.get("status") == "done"
        rag    = sr.get("rag",{}).get("overall_rag","—") if s_done else "—"
        nsig   = (len(sr.get("signals",{}).get("momentum",[])) +
                  len(sr.get("signals",{}).get("investment",[]))) if s_done else "—"
        nsess  = len(fr.get("sessions",[])) if f_done else "—"
        proc   = sr.get("processed_at","—")[:16] if s_done else "—"

        c0,c1,c2,c3,c4,c5 = st.columns([2.5,1.0,0.8,0.8,0.8,1.5])
        c0.markdown(vname)
        c1.markdown(hub)
        c2.markdown(f"{'✅' if s_done else '⬜'} {nsig}")
        rag_emoji = RAG_EMOJI.get(rag, "⚪")
        c3.markdown(f"{rag_emoji} {rag}")
        c4.markdown(f"{'✅' if f_done else '⬜'} {nsess}")
        c5.markdown(f"<small>{proc}</small>", unsafe_allow_html=True)

    st.divider()
    st.markdown("#### 📤 Upload Instructions")
    st.info(
        "After downloading both JSON files, upload them manually to SharePoint at:\n\n"
        f"**Signals:** `{SIGNALS_REPO_PATH}`\n\n"
        f"**Feedback:** `{FEEDBACK_REPO_PATH}`\n\n"
        "The frontend app will automatically read them from there."
    )
