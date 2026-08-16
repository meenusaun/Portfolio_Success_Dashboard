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
TRANSCRIPT_FOLDER= f"{COMMON_FOLDER}/Transcripts (Accelerate Mentor Meetings)"
REPO_FOLDER      = f"{COMMON_FOLDER}/Knowledge Repository"

EXTRA_SCAN_FOLDERS = [
    f"{SP_FOLDER}/Transcripts",
    f"{SP_FOLDER}/Session Transcripts",
]

TRANSCRIPT_FOLDER_VARIANTS = [
    "transcripts (accelerate mentor meetings)",
    "accelerate mentor meetings",
    "session transcripts",
    "mentor meetings",
    "transcripts",
]
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

# ── extract session transcripts from already-loaded common docs ──
# No separate download — Session Transcripts folder was already scanned
# during Step 0 pre-load as part of Common Documents recursive scan.

def extract_transcripts_from_common(common_text):
    """
    Parse already-loaded common_text and return only sections
    that came from files inside the Session Transcripts folder.

    Now uses full SharePoint path stored in FILE header (since Step 0 fix)
    so any file inside Session Transcripts/ is included regardless of filename.
    Returns list of {"filename": str, "text": str}.
    """
    TRANSCRIPT_FOLDER_LOWER = "session transcripts"
    results = []
    sections = common_text.split("=== FILE:")
    for section in sections:
        if not section.strip(): continue
        header_end = section.find("===")
        if header_end == -1: continue
        fpath = section[:header_end].strip()
        text  = section[header_end+3:].strip()
        # Match by folder path (full path now stored) OR filename keyword fallback
        in_transcript_folder = TRANSCRIPT_FOLDER_LOWER in fpath.lower()
        fname_match = ("transcript" in fpath.lower().split("/")[-1] or
                       "session"    in fpath.lower().split("/")[-1])
        if (in_transcript_folder or fname_match) and text and len(text) > 50:
            results.append({
                "filename": fpath.split("/")[-1],  # just the filename for display
                "path":     fpath,
                "text":     text
            })
    return results

def get_transcript_for_venture(vname, all_transcripts):
    """
    Find transcript entries matching a venture name.
    all_transcripts is the list returned by extract_transcripts_from_common().
    Matches by filename or first 2000 chars of content.
    """
    matched = []
    vname_lower = vname.lower()
    vwords = [w for w in vname_lower.split() if len(w) > 3]
    for entry in all_transcripts:
        fname_lower = entry["filename"].lower()
        text_lower  = entry["text"].lower()
        name_match  = vname_lower in fname_lower or any(w in fname_lower for w in vwords)
        text_match  = vname_lower in text_lower[:2000]
        if name_match or text_match:
            matched.append(entry["text"])
    return "\n\n".join(matched) if matched else ""

# ── File index storage ────────────────────────────────────────────────────
# Pre-load stores only a dict of SP paths — no file content.
# Tiny enough to survive in session state without memory issues.

def _store_file_index(index: dict, att: dict):
    st.session_state["_file_index"] = index
    st.session_state["_att_cache"]  = att

def _retrieve_file_index():
    return (
        st.session_state.get("_file_index", {}),
        st.session_state.get("_att_cache",  {})
    )

def _read_venture_files_from_index(vname, sp):
    """
    Download only Common Documents files relevant to this venture.

    Strategy (fast-to-slow):
    1. Always include files where venture name appears in the FILE PATH
       (covers venture-specific sub-folders, named files)
    2. For transcript-folder files where name isn't in path:
       check FILENAME only — download only if filename contains a venture
       word match. Skip the rest — transcript files not named after a venture
       will be caught when the venture folder's own transcript files are read.

    This avoids downloading 80 transcript files per venture.
    """
    index, _ = _retrieve_file_index()
    if not index or not sp:
        return ""
    vname_lower = vname.lower()
    skip   = {"and","the","pvt","ltd","llp","inc","for","with","from","that"}
    vwords = [w for w in vname_lower.split() if len(w) > 3 and w not in skip]

    parts = []
    for fpath, fmeta in index.items():
        # fmeta is now a dict {type, modified, size} — handle both old and new format
        ftype       = fmeta["type"] if isinstance(fmeta, dict) else fmeta
        fpath_lower = fpath.lower()
        fname       = fpath.split("/")[-1]
        fname_lower = fname.lower()

        is_transcript_folder = any(v in fpath_lower for v in TRANSCRIPT_FOLDER_VARIANTS)

        # Check if venture name or key words appear in the path or filename
        name_in_path  = (vname_lower in fpath_lower or
                         any(w in fpath_lower for w in vwords))
        name_in_fname = (vname_lower in fname_lower or
                         any(w in fname_lower for w in vwords))

        if is_transcript_folder:
            # For transcript folders: only download if venture name is in
            # the filename — avoids downloading every transcript for every venture
            if not name_in_fname:
                continue
        else:
            # For other common doc files: require name in path
            if not name_in_path:
                continue

        try:
            cb   = sp.download_file(fpath)
            text = extract_text_bytes(cb, fname)
            if text and len(text) >= 50:
                parts.append(f"=== SOURCE: {fpath} ===\n{text}")
        except Exception:
            pass

    return "\n\n".join(parts)

def _get_venture_file_signatures(vname):
    """
    Return {sp_path: lastModifiedDateTime} for all files belonging to
    this venture in the file index. Used to detect changes since last run.
    """
    index, _ = _retrieve_file_index()
    vname_lower = vname.lower()
    skip   = {"and","the","pvt","ltd","llp","inc","for","with","from","that"}
    vwords = [w for w in vname_lower.split() if len(w) > 3 and w not in skip]
    sigs   = {}
    for fpath, fmeta in index.items():
        fpath_lower = fpath.lower()
        fname_lower = fpath.split("/")[-1].lower()
        modified    = fmeta["modified"] if isinstance(fmeta, dict) else ""
        in_tr_folder = any(v in fpath_lower for v in TRANSCRIPT_FOLDER_VARIANTS)
        if in_tr_folder:
            # Only include transcript files named after this venture
            if vname_lower in fname_lower or any(w in fname_lower for w in vwords):
                sigs[fpath] = modified
        else:
            if vname_lower in fpath_lower or any(w in fpath_lower for w in vwords):
                sigs[fpath] = modified
    return sigs

def _venture_files_changed(vname, stored_sigs):
    """
    Return True if any source file for this venture has changed since
    stored_sigs was captured, or if new files were added.
    stored_sigs: {sp_path: lastModifiedDateTime} saved in the repository JSON.
    Returns True (must reprocess) if stored_sigs is empty or None.
    """
    if not stored_sigs:
        return True
    current = _get_venture_file_signatures(vname)
    if set(current.keys()) != set(stored_sigs.keys()):
        return True  # files added or removed
    return any(stored_sigs.get(p) != m for p, m in current.items())

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

step1_tab, step2_tab, step3_tab, step4_tab, status_tab = st.tabs([
    "📊 Step 1 — Signals Repository",
    "🎙 Step 2 — Feedback Repository",
    "📄 Step 3 — Journey Documents",
    "👷 Step 4 — People Hired",
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
    # ── Step 0: Build File Index ──────────────────────────────────────────────
    # Scans SP folder names only — no file content downloaded.
    # Stores {sp_path: file_type} dict in session state (tiny, safe).
    # File content is read on-demand per venture during each batch run.

    _idx, _att = _retrieve_file_index()
    preload_done = bool(_idx)

    with st.expander(
        "✅ File index ready — click to rebuild" if preload_done
        else "📂 Step 0 — Build File Index first (required)",
        expanded=not preload_done
    ):
        if preload_done:
            n_tr = sum(1 for v in _idx.values() if (v["type"] if isinstance(v,dict) else v) == "transcript")
            st.success(
                f"{len(_idx)} files indexed · {n_tr} transcript files "
                f"· {len(_att)} attendance records  "
                f"· File content read on-demand per batch (no memory issues)"
            )
            if st.button("🔄 Rebuild Index", key="reload_preload"):
                st.session_state.pop("_file_index", None)
                st.session_state.pop("_att_cache",  None)
                st.rerun()
        else:
            st.info(
                "Scans SharePoint for file names only — no content downloaded yet.  "
                "Takes 10-30 seconds. File content is downloaded per venture during batch runs."
            )
            if st.button("📂 Build File Index", key="preload_btn"):
                status_box = st.empty()
                prog_pre   = st.progress(0, text="Scanning folders...")
                file_index  = {}
                folder_list = []
                fi          = [0]
                VALID_EXT   = {".xlsx",".xls",".docx",".pdf",".pptx",".ppt"}

                if use_sp and ENV_CLIENT_ID:
                    try:
                        from sharepoint_reader import SharePointReader
                        sp_pre = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)

                        # Venture folder names — excluded from index (handled by load_v_files)
                        venture_folder_names = set(
                            v.strip().lower() for v in ventures_list if v.strip()
                        )

                        def index_folder(folder_path, depth=0):
                            try:
                                items = sp_pre.list_folder(folder_path)
                            except Exception:
                                return
                            sub_folders = [i for i in items if "folder" in i]
                            files_here  = [i for i in items if "file"   in i]
                            folder_name = folder_path.split("/")[-1]
                            # Skip venture-named sub-folders — handled by load_v_files()
                            if depth > 0 and folder_name.lower() in venture_folder_names:
                                return
                            folder_list.append("  " * depth + "📁 " + folder_name
                                               + " (" + str(len(files_here)) + " files)")
                            if len(folder_list) % 5 == 0:
                                status_box.caption(
                                    "Scanning: " + folder_name
                                    + " | Folders: " + str(len(folder_list))
                                    + " | Files: " + str(fi[0])
                                )
                            for item in files_here:
                                iname = item.get("name", "")
                                ipath = f"{folder_path}/{iname}"
                                if Path(iname).suffix.lower() not in VALID_EXT: continue
                                if "journey_accelerate_portfolio" in iname.lower(): continue
                                _ftype = "transcript" if any(
                                    v in folder_path.lower() for v in TRANSCRIPT_FOLDER_VARIANTS
                                ) else (
                                    "feedback"   if "feedback"   in iname.lower() else
                                    "journey"    if "journey"    in iname.lower() else
                                    "sprint"     if "sprint"     in iname.lower() else
                                    "attendance" if "attendance" in iname.lower() else
                                    "other"
                                )
                                file_index[ipath] = {
                                    "type":     _ftype,
                                    "modified": item.get("lastModifiedDateTime", ""),
                                    "size":     item.get("size", 0),
                                }
                                fi[0] += 1
                                if fi[0] % 5 == 0:
                                    prog_pre.progress(
                                        min(fi[0] / max(fi[0] + 20, 1), 0.95),
                                        text="Indexed " + str(fi[0]) + " | " + iname[:50]
                                    )
                            for item in sub_folders:
                                iname = item.get("name", "")
                                if iname.lower() in venture_folder_names: continue
                                index_folder(f"{folder_path}/{iname}", depth + 1)

                        index_folder(COMMON_FOLDER)
                        for extra in EXTRA_SCAN_FOLDERS:
                            try:
                                if sp_pre.list_folder(extra):
                                    index_folder(extra)
                            except Exception:
                                pass

                        status_box.markdown(
                            "**Index complete** — "
                            + str(len(folder_list)) + " folders · "
                            + str(fi[0]) + " files indexed"
                        )
                    except Exception as e:
                        st.error(f"Index error: {e}")

                att = load_attendance_cached(sp_id, use_sp)
                _store_file_index(file_index, att)
                try:
                    prog_pre.progress(1.0, text=f"✅ Done — {fi[0]} files indexed")
                except Exception:
                    pass
                st.success(
                    f"✅ Index built — {fi[0]} files across {len(folder_list)} folders "
                    f"· {len(att)} attendance records"
                )
                st.rerun()

    if not preload_done:
        st.info("👆 Build the file index above to enable batch processing.")
        st.stop()

    # Index ready — get att from index cache; common_text_sig not used (on-demand reading)
    _, att_data_sig = _retrieve_file_index()
    common_text_sig = ""  # not used — batches call _read_venture_files_from_index instead

    # Status already shown in the expander above

    # ── OPT 3: Upload existing repo to skip already-processed ventures ──
    st.markdown("---")
    st.markdown("**⚡ Skip already-processed ventures (saves API credits)**")
    skip_col1, skip_col2 = st.columns([3,1])
    with skip_col1:
        uploaded_repo = st.file_uploader(
            "Upload existing signals_repository.json — ventures with signals already extracted will be skipped",
            type=["json"], key="existing_repo_upload",
            help="Only ventures with zero signals (failed or new) will be re-processed"
        )
    if uploaded_repo and "existing_sig_repo" not in st.session_state:
        try:
            existing = json.loads(uploaded_repo.read().decode("utf-8"))
            already_done = {
                vn for vn, vdata in existing.get("venture_summary", {}).items()
                if len(vdata.get("signals",{}).get("momentum",[])) > 0
                or len(vdata.get("signals",{}).get("investment",[])) > 0
            }
            st.session_state["existing_sig_repo"]       = existing
            st.session_state["already_done_ventures"]   = already_done
            # Store full venture data for file-signature comparison
            st.session_state["sig_existing_data"]       = existing.get("venture_summary", {})
        except Exception as e:
            st.error(f"Could not read file: {e}")

    already_done_ventures = st.session_state.get("already_done_ventures", set())
    existing_sig_repo     = st.session_state.get("existing_sig_repo", {})

    if already_done_ventures:
        sk1, sk2 = st.columns([3,1])
        sk1.info(f"⚡ {len(already_done_ventures)} ventures already have signals — will be skipped  ·  "
                 f"🆕 {len([v for v in ventures_list if v not in already_done_ventures])} ventures to process")
        if sk2.button("🗑 Clear skip list", key="clear_skip"):
            del st.session_state["existing_sig_repo"]
            del st.session_state["already_done_ventures"]
            st.rerun()

        # Pre-populate sig_results from existing repo for skipped ventures
        for vn, vdata in existing_sig_repo.get("venture_summary", {}).items():
            if vn in already_done_ventures and vn not in sig_results:
                sig_results[vn] = {
                    "status":          "done",
                    "venture_name":    vn,
                    "hub":             vdata.get("hub","—"),
                    "venture_partner": vdata.get("venture_partner","—"),
                    "sprint":          vdata.get("sprint","—"),
                    "rag":             {
                        "overall_rag":      vdata.get("overall_rag","ZERO"),
                        "momentum_rag":     vdata.get("momentum_rag","ZERO"),
                        "investment_rag":   vdata.get("investment_rag","ZERO"),
                        "momentum_reason":  vdata.get("momentum_reason","—"),
                        "investment_reason":vdata.get("investment_reason","—"),
                        "momentum_score":   vdata.get("momentum_score",0),
                        "investment_score": vdata.get("investment_score",0),
                    },
                    "signals":         vdata.get("signals", {"momentum":[],"investment":[]}),
                    "att_sessions":    vdata.get("att_sessions",0),
                    "att_dates":       vdata.get("att_dates",[]),
                    "sources_used":    vdata.get("sources_used",[]),
                    "total_chars":     vdata.get("total_chars",0),
                    "num_chunks":      vdata.get("num_chunks",0),
                    "processed_at":    vdata.get("processed_at","—"),
                }
        st.session_state[SIG_RESULTS_KEY] = sig_results
    st.markdown("---")


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
                    prog.progress(vi/len(batch), text=f"Processing {vname} ({vi+1}/{len(batch)})...")

                    # ── Skip check ───────────────────────────────────────
                    # 1. Already done in this session
                    if sig_results.get(vname,{}).get("status") == "done":
                        # Also check if source files have changed since last run
                        stored_sigs = sig_results[vname].get("file_signatures", {})
                        if not _venture_files_changed(vname, stored_sigs):
                            prog.progress((vi+1)/len(batch),
                                          text=f"⚡ {vname} — no file changes, skipping")
                            continue
                    # 2. In uploaded existing repo — check file changes
                    if vname in st.session_state.get("already_done_ventures", set()):
                        existing = st.session_state.get(
                            "sig_existing_data", {}
                        ).get(vname, {})
                        stored_sigs = existing.get("file_signatures", {})
                        if not _venture_files_changed(vname, stored_sigs):
                            prog.progress((vi+1)/len(batch),
                                          text=f"⚡ {vname} — no file changes, skipping")
                            continue

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

                    venture_common = _read_venture_files_from_index(vname, sp_reader)

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

                    # Surface credit/API errors immediately — do not silently store empty results
                    api_errors = signals.get("errors", [])
                    credit_error = any("credit balance" in str(e) or "billing" in str(e).lower()
                                       for e in api_errors)
                    if credit_error:
                        prog.empty()
                        st.error(f"❌ Anthropic API credit error on venture '{vname}'. "
                                 f"Top up credits at console.anthropic.com → Plans & Billing, then re-run.")
                        st.stop()

                    rag = score_rag_from_signals(client, vname, sprint, notes,
                                                  att_summary, signals, pct)

                    sig_results[vname] = {
                        "status":          "done",
                        "venture_name":    vname,
                        "hub":             hub,
                        "venture_partner": vp,
                        "sprint":          sprint,
                        "rag":             rag,
                        "signals":         signals,
                        "total_chars":     total_chars,
                        "num_chunks":      num_chunks,
                        "sources_used":    srcs_used,
                        "att_sessions":    att_sessions,
                        "att_dates":       att_dates,
                        "processed_at":    datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                        "file_signatures": _get_venture_file_signatures(vname),
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
        "Parses Session Tracker + Feedback Quality Tracker (no API credits needed) + "
        "reads venture transcripts via Claude → Saves to `feedback_repository.json`"
    )

    if not client:
        st.error("❌ Anthropic API key required for transcript extraction."); st.stop()

    # ── Section A: Parse Tracker Files ──────────────────
    st.markdown("#### 📊 Section A — Parse Tracker Files")
    st.caption("Upload both tracker files. These are parsed directly — no API credits used.")

    ta_col1, ta_col2 = st.columns(2)
    with ta_col1:
        sess_file = st.file_uploader(
            "05_Session_Management_Tracker.xlsx",
            type=["xlsx"], key="sess_tracker_upload"
        )
    with ta_col2:
        fb_file = st.file_uploader(
            "06_Feedback_Quality_Tracker.xlsx",
            type=["xlsx"], key="fb_tracker_upload"
        )

    MENTOR_KEY = "mentor_insights_parsed"

    if sess_file and fb_file:
        if st.button("📊 Parse Tracker Files", key="parse_trackers"):
            with st.spinner("Parsing tracker files..."):
                try:
                    from processor import parse_tracker_files
                    sess_bytes = sess_file.read()
                    fb_bytes   = fb_file.read()
                    mentor_insights = parse_tracker_files(sess_bytes, fb_bytes)
                    st.session_state[MENTOR_KEY] = mentor_insights
                    total_sess = sum(m["total_sessions"] for m in mentor_insights.values())
                    st.success(
                        f"✅ Parsed {len(mentor_insights)} mentors · "
                        f"{total_sess} sessions · 0 API credits used"
                    )
                except Exception as e:
                    st.error(f"Parse error: {e}")
    elif st.session_state.get(MENTOR_KEY):
        mi = st.session_state[MENTOR_KEY]
        total_sess = sum(m["total_sessions"] for m in mi.values())
        st.success(f"✅ Tracker data loaded — {len(mi)} mentors · {total_sess} sessions")
        if st.button("🗑 Clear tracker data", key="clear_trackers"):
            del st.session_state[MENTOR_KEY]; st.rerun()
    else:
        st.info("Upload both tracker files above to parse session and feedback data.")

    st.divider()

    # ── Section B: Venture Transcript Sessions ──────────
    st.markdown("#### 🎙 Section B — Extract Venture Sessions from Transcripts")
    st.caption("Uses Claude to extract structured sessions from transcript files. API credits used here.")

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
        run_all_fb = st.button("▶ Generate All", key="run_all_feedback",
                               use_container_width=True)

    fb_target  = fb_venture_filter if fb_venture_filter else ventures_list
    fb_batches = [fb_target[i:i+fb_batch_size] for i in range(0, len(fb_target), fb_batch_size)]
    st.caption(f"{len(fb_target)} ventures · {len(fb_batches)} batches of {fb_batch_size}")
    st.divider()

    FB_RESULTS_KEY = "fb_repo_results"
    if FB_RESULTS_KEY not in st.session_state:
        st.session_state[FB_RESULTS_KEY] = {}
    fb_results = st.session_state[FB_RESULTS_KEY]

    # Transcripts read on-demand per venture using file index
    all_transcripts = []
    _n_tr = sum(1 for v in _retrieve_file_index()[0].values() if (v["type"] if isinstance(v,dict) else v) == "transcript")
    st.info(f"📁 {_n_tr} transcript files in index — read per venture during batch")

    for bi, batch in enumerate(fb_batches):
        batch_done = sum(1 for v in batch if fb_results.get(v,{}).get("status") == "done")
        icon = "✅" if batch_done == len(batch) else ("🔄" if batch_done > 0 else "⬜")
        with st.expander(
            f"{icon} Batch {bi+1}  —  {batch_done}/{len(batch)} done  —  "
            f"{', '.join(batch[:3])}{'...' if len(batch)>3 else ''}"
        ):
            for vn in batch:
                vr = fb_results.get(vn,{})
                vs = vr.get("status","pending")
                ic = {"done":"✅","error":"❌"}.get(vs,"⬜")
                nsess = len(vr.get("sessions",[])) if vs == "done" else "—"
                st.caption(f"{ic} {vn}  ·  {nsess} session(s)")

            st.markdown("")
            run_fb_btn   = st.button(f"▶ Run Batch {bi+1}", key=f"run_fb_batch_{bi}")
            should_run_fb = run_fb_btn or run_all_fb

            if should_run_fb:
                prog_fb = st.progress(0, text=f"Starting batch {bi+1}...")
                for vi, vname in enumerate(batch):
                    if fb_results.get(vname,{}).get("status") == "done":
                        prog_fb.progress((vi+1)/len(batch), text=f"Skipping {vname}")
                        continue
                    prog_fb.progress(vi/len(batch),
                                     text=f"Processing {vname} ({vi+1}/{len(batch)})...")
                    row    = get_row(vname)
                    hub    = cv(row, col_hub)
                    vp     = cv(row, col_vp) if col_vp else "—"
                    vfiles = load_v_files(vname)
                    fb_text = sp_get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
                    tr_text = sp_get_text(vfiles["transcript"]) if "transcript" in vfiles else ""

                    # Read other venture folder files (VP reports, panel notes, etc.)
                    other_texts = [
                        f"=== SOURCE: {p.split('/')[-1]} ===\n{sp_get_text(p)}"
                        for k, p in vfiles.items()
                        if k.startswith("other_") and sp_get_text(p)
                    ]

                    # Read common transcript files from index (transcript folder)
                    common_tr_text = _read_venture_files_from_index(vname, sp_reader)

                    combined_transcript = "\n\n".join(
                        t for t in [tr_text, common_tr_text] + other_texts if t
                    )

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

    # ── Download combined feedback_repository.json ──────
    st.divider()
    done_count_f     = sum(1 for v in fb_results.values() if v.get("status") == "done")
    mentor_insights  = st.session_state.get(MENTOR_KEY, {})
    can_download     = done_count_f > 0 or bool(mentor_insights)

    st.markdown(
        f"**{done_count_f}/{len(ventures_list)} venture transcripts processed · "
        f"{len(mentor_insights)} mentors parsed**"
    )

    if can_download:
        fb_dl_col, _ = st.columns(2)
        with fb_dl_col:
            st.markdown("**💾 Download feedback_repository.json**")
            fb_payload = {
                "generated_at":  datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                "venture_count": done_count_f,
                "mentor_count":  len(mentor_insights),
                "ventures":      {
                    vn: vr for vn, vr in fb_results.items()
                    if vr.get("status") == "done"
                },
                "mentor_insights": mentor_insights,
            }
            st.download_button(
                "⬇️ Download feedback_repository.json",
                data=json.dumps(fb_payload, indent=2, default=str),
                file_name="feedback_repository.json",
                mime="application/json",
            )
            st.caption(f"Upload this file to SharePoint at:\n`{FEEDBACK_REPO_PATH}`")
            if not mentor_insights:
                st.warning("⚠️ Mentor insights not yet parsed — upload tracker files in Section A first.")
            if done_count_f == 0:
                st.info("ℹ️ No venture transcripts processed — Section B is optional.")

# ══════════════════════════════════════════════════════
#  STEP 3: JOURNEY DOCUMENTS
# ══════════════════════════════════════════════════════
with step3_tab:
    st.markdown("### 📄 Extract Journey Document Data")
    st.caption(
        "Reads Sign off Journey Documents from Common Documents → "
        "Extracts structured venture data → Stores in signals_repository.json"
    )

    if not client:
        st.error("❌ Anthropic API key required."); st.stop()

    JOURNEY_FOLDER  = f"{COMMON_FOLDER}/Sign off Journey Documents"
    JOURNEY_KEY     = "journey_repo_results"
    VALID_EXTS      = [".pdf", ".docx", ".doc"]

    if JOURNEY_KEY not in st.session_state:
        st.session_state[JOURNEY_KEY] = {}

    journey_results = st.session_state[JOURNEY_KEY]
    done_j = sum(1 for v in journey_results.values() if v.get("status") == "done")

    # ── Controls ──────────────────────────────────────
    jc1, jc2, jc3 = st.columns([2, 1, 1])
    with jc1:
        j_venture_filter = st.multiselect(
            "Ventures to process (leave empty = all)",
            options=ventures_list, key="j_filter"
        )
    with jc2:
        j_batch_size = st.selectbox("Batch size", [5, 10, 15, 20],
                                     index=1, key="j_batch_size")
    with jc3:
        st.markdown("<br>", unsafe_allow_html=True)
        run_all_j = st.button("▶ Extract All", key="run_all_journey",
                               use_container_width=True)

    j_target  = j_venture_filter if j_venture_filter else ventures_list
    j_batches = [j_target[i:i+j_batch_size]
                 for i in range(0, len(j_target), j_batch_size)]
    st.caption(f"{len(j_target)} ventures · {len(j_batches)} batches")
    st.divider()

    # Pre-load: check if common docs already loaded
    if not _retrieve_file_index()[0]:
        st.warning("⚠️ Go to Step 1 and build the file index first.")
        st.stop()

    # Build journey doc lookup from pre-loaded common docs
    # Files in Sign off Journey Documents folder are already in common_text_sig
    @st.cache_data(show_spinner=False)
    def build_journey_lookup(common_text):
        """
        Extract journey document text from pre-loaded common docs.
        Matches files in:
          - Sign off Journey Documents/ (any subfolder depth)
          - JS1/, JS2/ subfolders inside it
          - Any path containing 'sign off' or 'js1' or 'js2'
        """
        JOURNEY_MARKERS = [
            "sign off journey documents",
            "sign off journey",
            "signoff journey",
            "final journey reports",
            "journey reports ratified",
            "ratified by vp",
            "/js1/",
            "/js2/",
        ]
        lookup = {}
        sections = common_text.split("=== FILE:")
        for section in sections:
            if not section.strip(): continue
            header_end = section.find("===")
            if header_end == -1: continue
            fpath = section[:header_end].strip()
            text  = section[header_end+3:].strip()
            fpath_lower = fpath.lower()
            if any(m in fpath_lower for m in JOURNEY_MARKERS) and text and len(text) > 100:
                fname = fpath.split("/")[-1]
                lookup[fname] = {"path": fpath, "text": text}
        return lookup

    # Journey docs: find paths from index, load text on demand
    def build_journey_index_lookup():
        index, _ = _retrieve_file_index()
        JOURNEY_MARKERS = ["sign off journey","journey documents","js1","js2","growth journey","journey report"]
        lookup = {}
        for fpath, ftype in index.items():
            if any(m in fpath.lower() for m in JOURNEY_MARKERS):
                lookup[fpath.split("/")[-1]] = {"path": fpath, "text": None}
        return lookup
    journey_docs = build_journey_index_lookup()

    # ── Debug: show ALL paths in common_text_sig to diagnose missing folders ──
    with st.expander("🔍 Debug — All paths loaded in Common Documents (Step 0)"):
        all_paths = []
        for section in []:  # file index approach — no pre-loaded text available; use SP paths instead
            if not section.strip(): continue
            header_end = section.find("===")
            if header_end == -1: continue
            fpath = section[:header_end].strip()
            if fpath: all_paths.append(fpath)
        if all_paths:
            st.caption(f"Total files loaded: {len(all_paths)}")
            journey_paths = [p for p in all_paths if "journey" in p.lower() or "js1" in p.lower() or "js2" in p.lower() or "sign off" in p.lower()]
            if journey_paths:
                st.markdown("**Journey-related paths found:**")
                for p in journey_paths: st.code(p)
            else:
                st.warning("No journey-related paths found. Showing all paths:")
                for p in sorted(all_paths): st.caption(p)
        else:
            st.warning("No paths found — Step 0 may not have completed.")

    st.info(f"📁 Found {len(journey_docs)} Journey Documents matching 'Sign off Journey Documents' folder")

    if journey_docs:
        with st.expander("📋 Journey Documents found"):
            for fname in sorted(journey_docs.keys()):
                st.caption(f"📄 {fname}")

    st.divider()

    def find_journey_doc(vname, journey_docs):
        """Match venture name to journey document — loads text on demand."""
        vname_lower = vname.lower()
        vwords = [w for w in vname_lower.split() if len(w) > 3]

        def get_text(doc):
            if doc.get("text") is not None:
                return doc["text"]
            try:
                cb = sp_reader.download_file(doc["path"])
                text = extract_text_bytes(cb, doc["path"].split("/")[-1])
                doc["text"] = text or ""
            except Exception:
                doc["text"] = ""
            return doc["text"]

        for fname, doc in journey_docs.items():
            fl = fname.lower()
            if vname_lower in fl or any(w in fl for w in vwords):
                return get_text(doc)

        for fname, doc in journey_docs.items():
            text = get_text(doc)
            if text and vname_lower in text.lower()[:1000]:
                return text
        return ""

    for bi, batch in enumerate(j_batches):
        batch_done = sum(1 for v in batch
                         if journey_results.get(v,{}).get("status") == "done")
        icon = "✅" if batch_done == len(batch) else                ("🔄" if batch_done > 0 else "⬜")
        with st.expander(
            f"{icon} Batch {bi+1}  —  {batch_done}/{len(batch)} done  —  "
            f"{', '.join(batch[:3])}{'...' if len(batch)>3 else ''}"
        ):
            for vn in batch:
                jr = journey_results.get(vn, {})
                js = jr.get("status", "pending")
                ic = {"done":"✅","error":"❌","no_doc":"⚠️","pending":"⬜"}.get(js,"⬜")
                fields_found = len([k for k,v in jr.items()
                                    if k not in ["status","venture_name","processed_at","_error"]
                                    and v])
                if js == "error":
                    err_msg = jr.get("_error", "Unknown error")
                    st.caption(f"{ic} {vn}  ·  Error: {err_msg[:150]}")
                elif js == "no_doc":
                    st.caption(f"{ic} {vn}  ·  No Journey Document found for this venture")
                elif js == "done":
                    st.caption(f"{ic} {vn}  ·  {fields_found} fields extracted")
                else:
                    st.caption(f"{ic} {vn}  ·  Not processed yet")

            st.markdown("")
            run_j_btn   = st.button(f"▶ Run Batch {bi+1}", key=f"run_j_batch_{bi}")
            should_run_j = run_j_btn or run_all_j

            if should_run_j:
                prog_j = st.progress(0, text=f"Starting batch {bi+1}...")
                for vi, vname in enumerate(batch):
                    if journey_results.get(vname,{}).get("status") == "done":
                        prog_j.progress((vi+1)/len(batch),
                                        text=f"⚡ Skipping {vname}")
                        continue

                    prog_j.progress(vi/len(batch),
                                    text=f"Processing {vname} ({vi+1}/{len(batch)})...")

                    # Find journey doc
                    doc_text = find_journey_doc(vname, journey_docs)
                    if not doc_text:
                        journey_results[vname] = {
                            "status": "no_doc",
                            "venture_name": vname,
                            "processed_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                        }
                        st.session_state[JOURNEY_KEY] = journey_results
                        continue

                    from processor import extract_journey_document_data
                    extracted = extract_journey_document_data(client, vname, doc_text)

                    journey_results[vname] = {
                        "status":       "done" if not extracted.get("_error") else "error",
                        "venture_name": vname,
                        "processed_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                        **extracted
                    }
                    st.session_state[JOURNEY_KEY] = journey_results

                prog_j.progress(1.0, text=f"✅ Batch {bi+1} complete!")
                st.rerun()

    # ── Download ──────────────────────────────────────
    st.divider()
    done_j = sum(1 for v in journey_results.values() if v.get("status") == "done")
    no_doc = sum(1 for v in journey_results.values() if v.get("status") == "no_doc")
    st.markdown(f"**{done_j} extracted · {no_doc} no document found · "
                f"{len(ventures_list)-done_j-no_doc} pending**")

    if done_j > 0:
        # Merge journey data into signals repository if it exists
        sig_results = st.session_state.get("sig_repo_results", {})

        # Build download payload — standalone journey repository
        journey_payload = {
            "generated_at":  datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
            "venture_count": done_j,
            "ventures":      {
                vn: vr for vn, vr in journey_results.items()
                if vr.get("status") == "done"
            }
        }

        dl1, dl2 = st.columns(2)
        with dl1:
            st.markdown("**💾 Download journey_repository.json**")
            st.download_button(
                "⬇️ Download journey_repository.json",
                data=json.dumps(journey_payload, indent=2, default=str),
                file_name="journey_repository.json",
                mime="application/json",
            )
            journey_repo_path = f"{REPO_FOLDER}/journey_repository.json"
            st.caption(f"Upload to SharePoint: `{journey_repo_path}`")

        with dl2:
            if sig_results:
                # Merge journey data into signals repository
                merged_sig = dict(st.session_state.get("sig_repo_results", {}))
                for vn, jr in journey_results.items():
                    if vn in merged_sig and jr.get("status") == "done":
                        merged_sig[vn]["journey_data"] = {
                            k: v for k, v in jr.items()
                            if k not in ["status","venture_name","processed_at"]
                        }

                # Rebuild signals payload with journey data
                records = st.session_state.get("sig_repo_records", [])
                save_payload_merged = {
                    "generated_at":   datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                    "venture_count":  len([v for v in merged_sig.values()
                                           if v.get("status") == "done"]),
                    "records":        records,
                    "venture_summary": {
                        vn: {
                            **{k: vr.get(k) for k in [
                                "venture_name","hub","venture_partner","sprint",
                                "overall_rag","momentum_rag","investment_rag",
                                "momentum_reason","investment_reason",
                                "momentum_score","investment_score",
                                "signals","att_sessions","att_dates",
                                "sources_used","processed_at"
                            ] if vr.get(k) is not None},
                            "journey_data": vr.get("journey_data", {}),
                        }
                        for vn, vr in merged_sig.items()
                        if vr.get("status") == "done"
                    }
                }
                st.markdown("**💾 Download merged signals_repository.json**")
                st.download_button(
                    "⬇️ Download signals_repository.json (with journey data)",
                    data=json.dumps(save_payload_merged, indent=2, default=str),
                    file_name="signals_repository.json",
                    mime="application/json",
                )
                st.caption("Includes journey data merged into signals repository")
            else:
                st.info("Run Step 1 first to also merge journey data into signals repository.")

# ══════════════════════════════════════════════════════
#  STEP 4: PEOPLE HIRED
# ══════════════════════════════════════════════════════
with step4_tab:
    st.markdown("### 👷 Extract People Hired Data")
    st.caption(
        "Reads ALL documents (venture folder + Common Documents) → "
        "Extracts hiring data (resources hired, 6-month plan, role breakdown by function) → "
        "Stores in people_repository.json"
    )

    if not client:
        st.error("❌ Anthropic API key required."); st.stop()

    PEOPLE_KEY = "people_repo_results"
    if PEOPLE_KEY not in st.session_state:
        st.session_state[PEOPLE_KEY] = {}

    people_results = st.session_state[PEOPLE_KEY]
    done_p = sum(1 for v in people_results.values() if v.get("status") == "done")

    if not _retrieve_file_index()[0]:
        st.warning("⚠️ Go to Step 1 and build the file index first.")
        st.stop()

    # ── Controls ──────────────────────────────────────
    pc1, pc2, pc3 = st.columns([2, 1, 1])
    with pc1:
        p_venture_filter = st.multiselect(
            "Ventures to process (leave empty = all)",
            options=ventures_list, key="p_filter"
        )
    with pc2:
        p_batch_size = st.selectbox("Batch size", [5, 10, 15, 20],
                                     index=1, key="p_batch_size")
    with pc3:
        st.markdown("<br>", unsafe_allow_html=True)
        run_all_p = st.button("▶ Extract All", key="run_all_people",
                               use_container_width=True)

    p_target  = p_venture_filter if p_venture_filter else ventures_list
    p_batches = [p_target[i:i+p_batch_size]
                 for i in range(0, len(p_target), p_batch_size)]
    st.caption(f"{len(p_target)} ventures · {len(p_batches)} batches")
    st.divider()

    common_text_p = ""  # not used — on-demand reading via _read_venture_files_from_index

    for bi, batch in enumerate(p_batches):
        batch_done = sum(1 for v in batch
                         if people_results.get(v,{}).get("status") == "done")
        icon = "✅" if batch_done == len(batch) else                ("🔄" if batch_done > 0 else "⬜")
        with st.expander(
            f"{icon} Batch {bi+1}  —  {batch_done}/{len(batch)} done  —  "
            f"{', '.join(batch[:3])}{'...' if len(batch)>3 else ''}"
        ):
            for vn in batch:
                pr = people_results.get(vn, {})
                ps = pr.get("status", "pending")
                ic = {"done":"✅","error":"❌"}.get(ps,"⬜")
                if ps == "done":
                    hc = pr.get("resources_hired_count", 0)
                    pc = pr.get("hiring_plan_6mo_count", 0)
                    st.caption(f"{ic} {vn}  ·  {hc} hired · {pc} planned")
                else:
                    st.caption(f"{ic} {vn}  ·  Not processed")

            st.markdown("")
            run_p_btn   = st.button(f"▶ Run Batch {bi+1}", key=f"run_p_batch_{bi}")
            should_run_p = run_p_btn or run_all_p

            if should_run_p:
                prog_p = st.progress(0, text=f"Starting batch {bi+1}...")
                for vi, vname in enumerate(batch):
                    if people_results.get(vname,{}).get("status") == "done":
                        prog_p.progress((vi+1)/len(batch),
                                        text=f"⚡ Skipping {vname}")
                        continue

                    prog_p.progress(vi/len(batch),
                                    text=f"Processing {vname} ({vi+1}/{len(batch)})...")

                    # Build full_text from venture files + common docs
                    vfiles  = load_v_files(vname)
                    fb_text = sp_get_text(vfiles["feedback"])   if "feedback"   in vfiles else ""
                    tr_text = sp_get_text(vfiles["transcript"]) if "transcript" in vfiles else ""
                    sp_text = sp_get_text(vfiles["sprint"])     if "sprint"     in vfiles else ""
                    jr_text = sp_get_text(vfiles["journey"])    if "journey"    in vfiles else ""
                    others  = [sp_get_text(p) for k,p in vfiles.items() if k.startswith("other_")]
                    others  = [t for t in others if t]

                    venture_common = _read_venture_files_from_index(vname, sp_reader)

                    sources = {
                        "Feedback": fb_text, "Transcript": tr_text,
                        "Sprint Plan": sp_text, "Growth Journey": jr_text,
                        "Common Docs": venture_common,
                    }
                    for idx, ot in enumerate(others):
                        sources[f"Venture File {idx+1}"] = ot

                    full_text = "\n\n".join(
                        f"=== {k} ===\n{v}" for k,v in sources.items() if v)

                    from processor import extract_people_hired_data
                    extracted = extract_people_hired_data(client, vname, full_text)

                    people_results[vname] = {
                        "status":       "done",
                        "venture_name": vname,
                        "processed_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                        **extracted
                    }
                    st.session_state[PEOPLE_KEY] = people_results

                prog_p.progress(1.0, text=f"✅ Batch {bi+1} complete!")
                st.rerun()

    # ── Download ──────────────────────────────────────
    st.divider()
    done_p = sum(1 for v in people_results.values() if v.get("status") == "done")
    st.markdown(f"**{done_p}/{len(ventures_list)} ventures processed**")

    if done_p > 0:
        people_payload = {
            "generated_at":  datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
            "venture_count": done_p,
            "ventures":      {
                vn: vr for vn, vr in people_results.items()
                if vr.get("status") == "done"
            }
        }
        st.markdown("**💾 Download people_repository.json**")
        st.download_button(
            "⬇️ Download people_repository.json",
            data=json.dumps(people_payload, indent=2, default=str),
            file_name="people_repository.json",
            mime="application/json",
        )
        people_repo_path = f"{REPO_FOLDER}/people_repository.json"
        st.caption(f"Upload to SharePoint: `{people_repo_path}`")

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
