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
            return "\n".join(pd.read_excel(fp,sheet_name=s,header=None).fillna("").astype(str).to_string() for s in xl.sheet_names)[:20000]
        elif ext == ".docx":
            from docx import Document
            return "\n".join(p.text for p in Document(fp).paragraphs)[:20000]
        elif ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(fp) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)[:20000]
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
    """Load ALL files from venture folder — not just keyword-matched ones."""
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
                # Categorise by keyword, but also keep ALL as "other_N"
                if "transcript"   in fl:
                    files["transcript"] = fp
                elif "feedback"   in fl:
                    files["feedback"]   = fp
                elif "growth sprint" in fl or "sprint plan" in fl:
                    files["sprint"]     = fp
                elif "growth journey" in fl or "journey report" in fl:
                    files["journey"]    = fp
                else:
                    # Read all other files too
                    key = f"other_{len([k for k in files if k.startswith('other')])}"
                    files[key] = fp
        except: pass
    else:
        vpath = os.path.join(root_path or "", vname)
        if os.path.isdir(vpath):
            for f in os.listdir(vpath):
                fl  = f.lower()
                fp  = os.path.join(vpath, f)
                ext = Path(f).suffix.lower()
                if not os.path.isfile(fp): continue
                if ext not in VALID_EXT: continue
                if "transcript"   in fl:
                    files["transcript"] = fp
                elif "feedback"   in fl:
                    files["feedback"]   = fp
                elif "growth sprint" in fl or "sprint plan" in fl:
                    files["sprint"]     = fp
                elif "growth journey" in fl or "journey report" in fl:
                    files["journey"]    = fp
                else:
                    key = f"other_{len([k for k in files if k.startswith('other')])}"
                    files[key] = fp
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
    """Load all Common Documents recursively (including sub-folders)."""
    texts = []

    def process_file(fname, content_bytes=None, fpath=None):
        ext = Path(fname).suffix.lower()
        if ext not in [".xlsx",".xls",".docx",".pdf",".pptx",".ppt"]: return
        if "journey_accelerate_portfolio" in fname.lower(): return
        try:
            if content_bytes:
                text = extract_text_bytes(content_bytes, fname)
            else:
                text = extract_text_local(fpath)
            if not text or len(text) < 50: return
            # Store full content up to 25000 chars per file
            # Enough for ~15-20 pages — truncation per venture at extraction time
            texts.append(f"=== FILE: {fname} ===\n{text}")
        except: pass

    if use_sp and ENV_CLIENT_ID:
        try:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)

            def scan_sp_folder(folder_path):
                """Recursively scan SharePoint folder."""
                try:
                    items = sp.list_folder(folder_path)
                    for item in items:
                        item_name = item.get("name","")
                        item_path = f"{folder_path}/{item_name}"
                        if "folder" in item:
                            # Recurse into sub-folder
                            scan_sp_folder(item_path)
                        elif "file" in item:
                            try:
                                content = sp.download_file(item_path)
                                process_file(item_name, content_bytes=content)
                            except: pass
                except: pass

            scan_sp_folder(COMMON_FOLDER)
        except: pass

    elif root_path:
        cpath = os.path.join(root_path, "Common Documents")
        if os.path.isdir(cpath):
            # Walk all sub-folders recursively
            for dirpath, dirnames, filenames in os.walk(cpath):
                for f in filenames:
                    process_file(f, fpath=os.path.join(dirpath, f))

    return "\n\n".join(texts)

def extract_venture_from_common(venture_name, common_text, file_keyword=None):
    """Extract all sections mentioning this venture from common docs."""
    if not venture_name: return ""
    relevant = []

    # Source 1: Pull all columns from already-loaded company_df
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

    # Source 2: Common Documents — extract ALL text around venture mentions
    if common_text:
        sections = common_text.split("=== FILE:")
        for section in sections:
            if not section.strip(): continue
            fname_part = section.split("===")[0].strip() if "===" in section else ""
            if file_keyword and file_keyword.lower() not in fname_part.lower():
                continue
            if venture_name.lower() not in section.lower():
                continue

            lines = section.split("\n")
            venture_lines = []
            i = 0
            while i < len(lines):
                if venture_name.lower() in lines[i].lower():
                    # Extract a wide context window — 5 lines before, 15 lines after
                    start = max(0, i-5)
                    end   = min(len(lines), i+15)
                    chunk = lines[start:end]
                    venture_lines.extend(chunk)
                    venture_lines.append("---")
                    i = end  # skip ahead to avoid duplication
                else:
                    i += 1

            if venture_lines:
                relevant.append(f"[{fname_part}]\n" + "\n".join(venture_lines))

    return "\n\n".join(relevant)

def load_common_docs():
    """Wrapper to load common docs using cache."""
    sp_id = id(sp_reader) if sp_reader else 0
    return load_common_docs_cached(sp_id, use_sp, root_path)

@st.cache_data(show_spinner=False, ttl=600)
def load_attendance_data(_sp_id, use_sp, root_path):
    """Load attendance data from 'd&v ATTENDANCE IN LAST 2 MONTHS' file in Common Documents."""
    attendance = {}  # {venture_name: {"sessions": int, "dates": [str], "weeks_active": int}}
    
    ATTENDANCE_KEYWORDS = ["attendance", "d&v attendance", "d & v attendance"]
    
    try:
        content_bytes = None
        fname_found   = None

        if use_sp and ENV_CLIENT_ID:
            from sharepoint_reader import SharePointReader
            sp = SharePointReader(ENV_CLIENT_ID, ENV_TENANT_ID, ENV_CLIENT_SECRET)

            def find_in_sp(folder_path):
                """Recursively search for attendance file."""
                try:
                    items = sp.list_folder(folder_path)
                    for item in items:
                        iname = item.get("name","")
                        ipath = f"{folder_path}/{iname}"
                        if "folder" in item:
                            result = find_in_sp(ipath)
                            if result: return result
                        elif any(kw in iname.lower() for kw in ATTENDANCE_KEYWORDS):
                            return sp.download_file(ipath), iname
                except: pass
                return None, None

            content_bytes, fname_found = find_in_sp(COMMON_FOLDER)

        elif root_path:
            cpath = os.path.join(root_path, "Common Documents")
            if os.path.isdir(cpath):
                for dirpath, _, filenames in os.walk(cpath):
                    for f in filenames:
                        if any(kw in f.lower() for kw in ATTENDANCE_KEYWORDS):
                            with open(os.path.join(dirpath, f), "rb") as fh:
                                content_bytes = fh.read()
                            fname_found = f
                            break
                    if fname_found: break

        if not content_bytes or not fname_found:
            return attendance

        # Save to temp and parse
        ext  = Path(fname_found).suffix.lower()
        tmp  = os.path.join(tempfile.gettempdir(), f"nen_attendance{ext}")
        with open(tmp, "wb") as f: f.write(content_bytes)

        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(tmp) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            # Parse text format: "VentureName  date1  date2  date3..."
            lines = text.split("\n")
            for line in lines:
                line = line.strip()
                if not line or line.startswith("Hub") or line.startswith("Partner") or line.startswith("Status") or line.startswith("Sum"):
                    continue
                # Skip hub names (single word lines that are city names)
                parts = line.split()
                if len(parts) < 2: continue
                # Date pattern: dd-mm
                date_pattern = re.compile(r'\d{2}-\d{2}')
                dates = date_pattern.findall(line)
                if dates:
                    # Venture name is everything before the first date
                    first_date_idx = line.find(dates[0])
                    vname_raw = line[:first_date_idx].strip()
                    # Remove trailing abbreviation in parentheses if present
                    vname_clean = re.sub(r'\s*\([^)]*\)\s*$', '', vname_raw).strip()
                    if vname_clean and len(vname_clean) > 3:
                        attendance[vname_clean] = {
                            "sessions":     len(dates),
                            "dates":        dates,
                            "weeks_active": len(set(dates))
                        }

        elif ext in [".xlsx", ".xls"]:
            xl = pd.ExcelFile(tmp)
            for sheet in xl.sheet_names:
                df = pd.read_excel(tmp, sheet_name=sheet, header=None)
                for _, row in df.iterrows():
                    row_vals = [str(v).strip() for v in row if str(v).strip() not in ["nan","None",""]]
                    if len(row_vals) < 2: continue
                    date_pattern = re.compile(r'\d{2}-\d{2}')
                    dates = [v for v in row_vals[1:] if date_pattern.match(v)]
                    if dates:
                        vname_clean = re.sub(r'\s*\([^)]*\)\s*$', '', row_vals[0]).strip()
                        if vname_clean and len(vname_clean) > 3:
                            attendance[vname_clean] = {
                                "sessions":     len(dates),
                                "dates":        dates,
                                "weeks_active": len(set(dates))
                            }

    except Exception as e:
        pass

    return attendance

def get_attendance_for_venture(vname, attendance_data):
    """Find attendance record for a venture using fuzzy name matching."""
    if not attendance_data: return None
    
    # Exact match first
    if vname in attendance_data: return attendance_data[vname]
    
    # Partial match — check if venture name words appear in attendance key
    vname_lower = vname.lower()
    for att_name, data in attendance_data.items():
        att_lower = att_name.lower()
        # Check if first 2 significant words match
        v_words = [w for w in vname_lower.split() if len(w) > 3]
        a_words = [w for w in att_lower.split() if len(w) > 3]
        matches = sum(1 for w in v_words if any(w in aw or aw in w for aw in a_words))
        if matches >= 2 or (len(v_words) == 1 and matches == 1):
            return data
    return None

@st.cache_data(show_spinner=False, ttl=600)
def compute_rag_ai(vname, notes, fb_text, tr_text, common_text, pct_raw,
                   att_sessions, att_dates, att_weeks, sprint_type,
                   growth_journey_text, _client_key):
    """Use Claude to compute RAG scores for a venture."""
    if not client: return "Unknown", "Unknown", "No AI key", "No AI key"
    
    try:
        pct = float(str(pct_raw).replace("%","").strip())
        if pct <= 1: pct *= 100
    except: pct = 0

    # Extract only sections mentioning this venture from common docs
    venture_common = extract_venture_from_common(vname, common_text) if common_text else ""

    # Format attendance info
    if att_sessions and att_sessions > 0:
        att_summary = f"{att_sessions} sessions attended in last 2 months (dates: {', '.join(att_dates or [])}), active in {att_weeks} weeks"
    else:
        att_summary = "No attendance data found"

    # Sprint-specific investment signals based on sprint type
    sprint_investment_guide = get_sprint_investment_guide(sprint_type)

    combined = f"""
Venture: {vname}
Sprint Type: {sprint_type or 'Unknown'}
Sprint Completion: {pct:.0f}%
Attendance (last 2 months): {att_summary}
Program Notes/Remarks: {notes}
Session Feedback: {fb_text}
Transcript: {tr_text}
Growth Journey Report: {growth_journey_text}
Common Documents (venture-specific excerpts): {venture_common}
"""
    prompt = f"""You are scoring a venture in an accelerator program. Analyze the data and return ONLY a JSON object with 6 keys.

CRITICAL RULES:
1. Attendance absence does NOT automatically mean Red. A founder with a $68,000 export order is Green even if not in attendance tracker.
2. Business outcomes (orders won, hires made, investments) are STRONGER signals than attendance.
3. ZERO means no data available at all — not poor performance. Only use ZERO when there is genuinely NO information.
4. Score based on EVIDENCE of actions taken, not just engagement frequency.

SPRINT MOMENTUM SCORE:
- Green: Founder engaged, will complete sprint IN LINE with objectives. Evidence: active on deliverables, positive outcomes, export orders, task completion.
- Amber: Founder engaged BUT sprint running with DELAY. Founder not fully happy. Evidence: slow progress, inconsistent, needs nudging.
- Red: Founder DISENGAGED, sprint UNLIKELY to reach objective. Evidence: wants to drop out, not convinced by program, no progress despite time passing.
- ZERO: No data from any source — genuinely unknown.

SELF INVESTMENT SIGNAL (MUST be specific to sprint type '{sprint_type}'):
- Green: Founder HAS ALREADY invested in sprint-related resources. Hired staff, bought tools, spent money on sprint goals, likely to self-invest in next sprint.
- Amber: Founder has NOT YET invested but is LIKELY TO invest in next sprint. Planning, exploring, showing intent.
- Red: Founder UNSURE, not ready to invest. Waiting for value proof, no commitment signal.
- ZERO: No data available.

REAL EXAMPLES (learn from these):
- Kadillac Chemicals: Momentum=Green (purchase order $68K in hand, active), Investment=Green (subscribed to export DB, hired ops staff)
- Suman Exports: Momentum=Amber (progress but execution gaps), Investment=Green (hired 4 sales professionals, SEO investment)
- ARGE: Momentum=Red (founder found no value), Investment=ZERO (no data)
- Shubham Aquavitro: Momentum=Amber (cannot comment on value yet), Investment=ZERO (no data)
- MIPA Industries: Momentum=Green (9/10 tasks done), Investment=Green (CRM implemented, hired CFO+HR+interns)
- Hydro Meshines: Momentum=Red (wants to drop out), Investment=Red (unresponsive, wants to exit)
- Atreya Innovations: Momentum=ZERO (founder missed entire sprint, no data), Investment=ZERO
- Shree Multi Sticks: Momentum=Green (8/10 tasks, export ready), Investment=Green (₹5Cr plant, ₹10L activities, 8-person team)

NUMERIC SCORING MATRIX:
Green+Green=10, Green+Amber=8, Green+Red=5, Green+ZERO=5
Amber+Green=8, Amber+Amber=7, Amber+Red=3, Amber+ZERO=3
Red+Green=5, Red+Amber=3, Red+Red=1, Red+ZERO=0
ZERO+Green=5, ZERO+Amber=3, ZERO+ZERO=0

{sprint_investment_guide}

Venture data to score:
{combined}

Return ONLY this JSON, no other text:
{{"momentum_rag": "Green/Amber/Red/ZERO", "momentum_reason": "one sentence", "investment_rag": "Green/Amber/Red/ZERO", "investment_reason": "one sentence referencing sprint type", "momentum_score": 0, "investment_score": 0}}"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5", max_tokens=500,
            messages=[{"role":"user","content":prompt}])
        raw = re.sub(r"```json|```","",resp.content[0].text.strip()).strip()
        data = json.loads(raw)
        return (data.get("momentum_rag","Unknown"), data.get("investment_rag","Unknown"),
                data.get("momentum_reason","—"), data.get("investment_reason","—"),
                data.get("momentum_score", 0), data.get("investment_score", 0))
    except Exception as e:
        return "Unknown","Unknown",f"Error: {e}","—", 0, 0


def get_sprint_investment_guide(sprint_type):
    """Return sprint-specific investment signal criteria."""
    if not sprint_type or sprint_type in ["—","Unknown",""]:
        return """- Green: Hired staff, invested in equipment/facility, self-funded sprint activities
- Amber: Planning to invest, exploring options, partial commitment
- Red: No investment intent, focused elsewhere, withdrawing"""

    st_lower = sprint_type.lower()

    if any(x in st_lower for x in ["export","international","global"]):
        return """This venture is on an EXPORT sprint. Investment signals must be EXPORT-specific:
- Green: Hired export manager / sales rep, attended trade fairs, got export certifications, opened overseas accounts, onboarded freight forwarder, paid for market research, self-funded export activities
- Amber: Planning export hire, exploring certifications, initial buyer conversations started
- Red: No export-related investment, paused export activities, no market engagement"""

    elif any(x in st_lower for x in ["product","r&d","research","development","innovation"]):
        return """This venture is on a PRODUCT DEVELOPMENT sprint. Investment signals must be product-specific:
- Green: Invested in R&D, purchased equipment/machinery, hired technical staff, filed patents, prototype built, testing conducted
- Amber: Planning equipment purchase, technical hire in progress, design stage
- Red: No product investment, development stalled, no technical progress"""

    elif any(x in st_lower for x in ["market","segment","customer","channel"]):
        return """This venture is on a NEW MARKET/SEGMENT sprint. Investment signals must be market-specific:
- Green: Hired sales/BD staff, invested in marketing, attended industry events, opened new dealer/distributor, ran pilots in new segment
- Amber: Planning market entry, initial outreach done, evaluating channels
- Red: No market investment, no new customer acquisition effort, withdrawn from segment"""

    elif any(x in st_lower for x in ["route","distribution","retail","supply"]):
        return """This venture is on a ROUTE TO MARKET sprint. Investment signals must be distribution-specific:
- Green: Onboarded distributors, invested in logistics, hired field sales, opened new retail channels
- Amber: Distributor talks initiated, evaluating logistics partners
- Red: No distribution investment, existing channels unchanged"""

    else:
        return f"""This venture is on a '{sprint_type}' sprint. Look for investment signals SPECIFIC to this sprint type:
- Green: Hired relevant staff, invested money, took concrete paid action directly related to {sprint_type} goals
- Amber: Planning investment, initial steps taken, partial commitment
- Red: No investment related to sprint goals, focused elsewhere"""

def rag_badge(rag):
    css = RAG_COLOR.get(rag, "rag-zero")
    emoji = RAG_EMOJI.get(rag, "⚪")
    return f'<span class="{css}">{emoji} {rag}</span>'

# ── top navigation tabs ──────────────────────────────
tab_overview, tab_ventures, tab_process = st.tabs(["📊  Portfolio Overview", "🏢  Venture Cards", "⚙️  Process Batches"])

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
                "momentum_score":  0,
                "investment_score": 0,
            })

        # Step 2: Load attendance only (no file reading for portfolio)
        with st.spinner("📋 Loading attendance data..."):
            sp_id           = id(sp_reader) if sp_reader else 0
            attendance_data = load_attendance_data(sp_id, use_sp, root_path)
        col_info.caption(f"📋 Attendance records: {len(attendance_data)} ventures")

        # Step 3: AI scoring in parallel — ALL ventures scored
        # Common Documents is primary source, Notes + files are supplementary
        if client:
            import concurrent.futures

            prog = st.progress(0, text=f"🤖 AI scoring {len(venture_data)} ventures using Notes + Attendance...")
            completed = [0]

            def score_one(v):
                # PORTFOLIO SCORING: Only uses Excel data + attendance
                # No file reading — keeps portfolio load fast
                att          = get_attendance_for_venture(v["name"], attendance_data)
                att_sessions = att["sessions"]     if att else 0
                att_dates    = att["dates"]        if att else []
                att_weeks    = att["weeks_active"] if att else 0

                has_data = any([
                    v["notes"] and v["notes"] not in ["—","Pending",""],
                    att_sessions > 0
                ])

                if not has_data:
                    return v["name"], "ZERO", "ZERO", "No data in Notes or Attendance", "No data in Notes or Attendance", 0, 0

                m, i, mr, ir, ms, is_ = compute_rag_ai(
                    v["name"], v["notes"], "", "", "", v["pct_raw"],
                    att_sessions, att_dates, att_weeks,
                    v["sprint"], "", api_key)
                return v["name"], m, i, mr, ir, ms, is_

            results = {}
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
                futures = {ex.submit(score_one, v): v for v in venture_data}
                for fut in concurrent.futures.as_completed(futures):
                    try:
                        vn, m, i, mr, ir, ms, is_ = fut.result()
                        results[vn] = (m, i, mr, ir, ms, is_)
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
                    m, i, mr, ir, ms, is_ = results[v["name"]]
                    v["momentum_rag"]      = m
                    v["investment_rag"]    = i
                    v["overall_rag"]       = combine_rag(m, i)
                    v["momentum_reason"]   = mr
                    v["investment_reason"] = ir
                    v["momentum_score"]    = ms
                    v["investment_score"]  = is_
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

    st.divider()

    # ── 2-column layout ───────────────────────────────
    left_col, gap_col, right_col = st.columns([5, 0.5, 5])

    # LEFT: RAG distribution + Sprint Stage
    with left_col:
        st.subheader("RAG Distribution")
        for label, count, color in [
            ("🟢 Green", greens, "#16a34a"),
            ("🟡 Amber", ambers, "#d97706"),
            ("🔴 Red",   reds,   "#dc2626"),
        ]:
            pct = round(count/total*100) if total else 0
            st.markdown(f"**{label}** — {count} ventures ({pct}%)")
            st.progress(pct/100)

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Sprint Stage Distribution")
        stage_counts = {}
        for v in filtered:
            stage_counts[v["bucket"]] = stage_counts.get(v["bucket"],0) + 1
        for stage, cnt in sorted(stage_counts.items()):
            pct = round(cnt/total*100) if total else 0
            st.markdown(f"**{stage}** — {cnt} ventures ({pct}%)")
            st.progress(pct/100)

    # RIGHT: Hub-wise RAG count table
    with right_col:
        st.subheader("Hub-wise RAG Count")
        hub_rag = {}
        for v in filtered:
            h = v["hub"] if v["hub"] not in ["—","Other"] else "Other"
            if h not in hub_rag:
                hub_rag[h] = {"Green":0,"Amber":0,"Red":0,"ZERO":0,"Total":0}
            hub_rag[h][v["overall_rag"]] = hub_rag[h].get(v["overall_rag"],0) + 1
            hub_rag[h]["Total"] += 1

        h0,h1,h2,h3,h4,h5 = st.columns([2.2,0.8,0.8,0.8,0.8,0.8])
        h0.markdown("**Hub**"); h1.markdown("**Total**")
        h2.markdown("**🟢**");  h3.markdown("**🟡**")
        h4.markdown("**🔴**");  h5.markdown("**⚪**")
        st.divider()
        for hub, counts in sorted(hub_rag.items(), key=lambda x: (x[0] == "Other", -x[1]["Total"])):
            r0,r1,r2,r3,r4,r5 = st.columns([2.2,0.8,0.8,0.8,0.8,0.8])
            r0.markdown(f"**{hub}**")
            r1.markdown(f"**{counts['Total']}**")
            r2.markdown(f"<span style='color:#16a34a;font-weight:700'>{counts.get('Green',0)}</span>", unsafe_allow_html=True)
            r3.markdown(f"<span style='color:#d97706;font-weight:700'>{counts.get('Amber',0)}</span>", unsafe_allow_html=True)
            r4.markdown(f"<span style='color:#dc2626;font-weight:700'>{counts.get('Red',0)}</span>",   unsafe_allow_html=True)
            r5.markdown(f"<span style='color:#94a3b8;font-weight:700'>{counts.get('ZERO',0)}</span>",  unsafe_allow_html=True)
        st.divider()
        t0,t1,t2,t3,t4,t5 = st.columns([2.2,0.8,0.8,0.8,0.8,0.8])
        t0.markdown("**Total**")
        t1.markdown(f"**{len(filtered)}**")
        t2.markdown(f"**{greens}**"); t3.markdown(f"**{ambers}**")
        t4.markdown(f"**{reds}**");   t5.markdown(f"**{zeros}**")

    # ── hub pivot table with scores (full width) ──────
    st.divider()
    st.subheader("Hub-wise Venture Score (Sprint Momentum × Self Investment)")
    st.caption("Avg Score based on 10-point matrix")

    hub_pivot = {}
    for v in filtered:
        h = v["hub"] if v["hub"] not in ["—","Other"] else "Other"
        m = v["momentum_rag"]
        if h not in hub_pivot: hub_pivot[h] = {}
        if m not in hub_pivot[h]: hub_pivot[h][m] = {"count":0,"scores":[]}
        hub_pivot[h][m]["count"]  += 1
        hub_pivot[h][m]["scores"].append(v.get("momentum_score",0))

    ph0,ph1,ph2,ph3,ph4,ph5,ph6 = st.columns([2.2,1.2,0.8,0.8,0.8,0.8,0.8])
    ph0.markdown("**Hub**");           ph1.markdown("**Sprint Momentum**")
    ph2.markdown("**Count**");         ph3.markdown("**Avg Score**")
    ph4.markdown("**🟢**");            ph5.markdown("**🟡**"); ph6.markdown("**🔴**")
    st.divider()

    grand_total = 0; grand_scores = []
    for hub in sorted(hub_pivot.keys(), key=lambda x: (x == "Other", x)):
        hub_total = 0; hub_scores_all = []
        hub_greens = hub_pivot[hub].get("Green",{}).get("count",0)
        hub_ambers = hub_pivot[hub].get("Amber",{}).get("count",0)
        hub_reds   = hub_pivot[hub].get("Red",{}).get("count",0)
        first_row  = True

        for m_rag in ["Green","Amber","Red","ZERO"]:
            if m_rag not in hub_pivot[hub]: continue
            data   = hub_pivot[hub][m_rag]
            scores = data["scores"]
            avg    = round(sum(scores)/len(scores),1) if scores else 0.0
            hub_total      += data["count"]
            hub_scores_all.extend(scores)
            grand_scores.extend(scores)

            c0,c1,c2,c3,c4,c5,c6 = st.columns([2.2,1.2,0.8,0.8,0.8,0.8,0.8])
            c0.markdown(f"**{hub}**" if first_row else "")
            c1.markdown(f"{RAG_EMOJI.get(m_rag,'⚪')} {m_rag}")
            c2.markdown(f"{data['count']}")
            c3.markdown(f"**{avg}**")
            if first_row:
                c4.markdown(f"<span style='color:#16a34a;font-weight:700'>{hub_greens}</span>", unsafe_allow_html=True)
                c5.markdown(f"<span style='color:#d97706;font-weight:700'>{hub_ambers}</span>", unsafe_allow_html=True)
                c6.markdown(f"<span style='color:#dc2626;font-weight:700'>{hub_reds}</span>",   unsafe_allow_html=True)
            first_row = False

        hub_avg = round(sum(hub_scores_all)/len(hub_scores_all),1) if hub_scores_all else 0.0
        grand_total += hub_total
        t0,t1,t2,t3,_,_,_ = st.columns([2.2,1.2,0.8,0.8,0.8,0.8,0.8])
        t0.markdown(f"**{hub} Total**"); t2.markdown(f"**{hub_total}**"); t3.markdown(f"**{hub_avg}**")
        st.divider()

    grand_avg = round(sum(grand_scores)/len(grand_scores),1) if grand_scores else 0.0
    g0,_,g2,g3,_,_,_ = st.columns([2.2,1.2,0.8,0.8,0.8,0.8,0.8])
    g0.markdown("**Grand Total**"); g2.markdown(f"**{grand_total}**"); g3.markdown(f"**{grand_avg}**")

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

        # Get cached RAG for title — no file loading
        cached    = st.session_state.get("rag_scores_ai") or []
        v_score   = next((s for s in cached if s["name"] == vname), None)
        rag_label = v_score["overall_rag"] if v_score else "—"
        rag_emoji = RAG_EMOJI.get(rag_label, "⚪")

        with st.expander(f"{rag_emoji} **{vname}**  ·  {hub}  ·  {bucket}  ·  RAG: {rag_label}"):

            # All file loading deferred — nothing runs until tab is clicked
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 Basic Details", "🎙 Sessions", "✦ Success Signals", "🤖 AI Insights", "🔍 Data Sources"])

            # TAB 1: Basic Details — Excel data only, no files
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

                st.markdown("#### RAG Scores")
                if v_score:
                    rs1, rs2, rs3 = st.columns(3)
                    rs1.markdown(f"**Overall RAG**<br>{rag_badge(v_score['overall_rag'])}", unsafe_allow_html=True)
                    rs2.markdown(f"**Sprint Momentum**<br>{rag_badge(v_score['momentum_rag'])}<br><small>{v_score.get('momentum_reason','')[:100]}</small>", unsafe_allow_html=True)
                    rs3.markdown(f"**Self Investment**<br>{rag_badge(v_score['investment_rag'])}<br><small>{v_score.get('investment_reason','')[:100]}</small>", unsafe_allow_html=True)
                else:
                    st.caption("RAG scores not computed yet — go to Portfolio Overview first.")

                if notes:
                    st.markdown("**📝 Remarks:**")
                    st.info(notes)

            # Helper: load files lazily — only when needed
            def get_venture_files():
                """Load all files for this venture on demand."""
                fkey = f"vfiles_loaded_{vname}"
                if fkey not in st.session_state:
                    vf = load_v_files(vname)
                    fb   = get_text(vf["feedback"])   if "feedback"   in vf else ""
                    tr   = get_text(vf["transcript"]) if "transcript" in vf else ""
                    sp   = get_text(vf["sprint"])     if "sprint"     in vf else ""
                    jour = get_text(vf["journey"])    if "journey"    in vf else ""
                    others = [get_text(p) for k,p in vf.items() if k.startswith("other_")]
                    others = [t for t in others if t]
                    st.session_state[fkey] = {"vf":vf,"fb":fb,"tr":tr,"sp":sp,"jour":jour,"others":others}
                return st.session_state[fkey]

            def get_venture_common():
                """Load common docs venture sections on demand."""
                ckey = f"vdocs_{vname}"
                if ckey not in st.session_state:
                    with st.spinner("📂 Scanning Common Documents..."):
                        cdocs = load_common_docs()
                        st.session_state[ckey] = extract_venture_from_common(vname, cdocs)
                return st.session_state[ckey]

            # TAB 2: Sessions — loads files on click
            with tab2:
                docs = get_venture_files()
                fb_text = docs["fb"]; tr_text = docs["tr"]
                col_a, col_b = st.columns(2)
                col_a.metric("Files Found", len(docs["vf"]))
                col_b.metric("Session Docs", sum(1 for x in [fb_text, tr_text] if x))

                if fb_text:
                    st.markdown("**📋 Session Feedback:**")
                    st.info(fb_text)
                if tr_text:
                    st.markdown("**🎙 Transcript:**")
                    st.info(tr_text)
                if not fb_text and not tr_text:
                    st.info("No session files found in venture folder.")

            # TAB 3: Success Signals
            with tab3:
                st.caption("Sources: Notes · All Venture Files · Common Documents (venture-specific sections only)")

                sig_cache_key  = f"signals_{vname}"
                sig_trigger_key = f"sig_trigger_{vname}"

                col_btn, col_clear = st.columns([2,1])

                if col_btn.button("🔍 Get All Signals", key=f"get_sig_{vname}",
                                   help="Reads ALL documents — chunked, no limits"):
                    st.session_state[sig_trigger_key] = True
                    st.session_state.pop(sig_cache_key, None)

                if col_clear.button("🗑 Clear", key=f"clr_sig_{vname}"):
                    st.session_state.pop(sig_cache_key, None)
                    st.session_state.pop(sig_trigger_key, None)

                # Run extraction if triggered
                if st.session_state.get(sig_trigger_key) and sig_cache_key not in st.session_state:
                    st.session_state.pop(sig_trigger_key, None)

                    # Load all files
                    docs           = get_venture_files()
                    venture_common = get_venture_common()

                    full_sources = {
                        "Notes":          notes or "",
                        "Feedback":       docs["fb"],
                        "Transcript":     docs["tr"],
                        "Sprint Plan":    docs["sp"],
                        "Growth Journey": docs["jour"],
                        "Common Docs":    venture_common,
                    }
                    for idx_o, ot in enumerate(docs["others"]):
                        full_sources[f"Venture File {idx_o+1}"] = ot

                    found_sources = {k:v for k,v in full_sources.items() if v}
                    total_chars   = sum(len(v) for v in found_sources.values())
                    st.info(f"📄 Sources: {', '.join(found_sources.keys())}  |  📊 {total_chars:,} chars total")

                    full_combined = "\n\n".join(
                        f"=== {src} ===\n{txt}"
                        for src, txt in full_sources.items() if txt
                    )

                    CHUNK_SIZE = 120000
                    if len(full_combined) <= CHUNK_SIZE:
                        chunks = [full_combined]
                    else:
                        parts = full_combined.split("\n\n=== ")
                        current_chunk = ""
                        chunks = []
                        for part in parts:
                            section = part if part.startswith("===") else "=== " + part
                            if len(current_chunk) + len(section) > CHUNK_SIZE:
                                if current_chunk: chunks.append(current_chunk)
                                current_chunk = section
                            else:
                                current_chunk += ("\n\n" + section if current_chunk else section)
                        if current_chunk: chunks.append(current_chunk)

                    st.info(f"🔄 {len(chunks)} chunk(s) to process")

                    result     = {"momentum": [], "investment": []}
                    seen_sigs  = set()
                    PROMPT_TPL = """Venture: {vname} | Sprint: {sprint} | Chunk {n} of {total}

Extract ALL signals into TWO categories from the text below.

SPRINT MOMENTUM SIGNALS: attendance, task completion, export orders, progress, founder engagement, milestones, positive feedback.
SELF INVESTMENT SIGNALS: money spent, staff hired, tools/subscriptions bought, new markets entered, capital invested, self-funded activities — specific to sprint type '{sprint}'.

Use EXACTLY this format per signal (one per line):
MOMENTUM: [signal type] | EVIDENCE: [exact quote from text] | SOURCE: [document name]
INVESTMENT: [signal type] | EVIDENCE: [exact quote from text] | SOURCE: [document name]

Be THOROUGH. If none in a category write: MOMENTUM: None found

--- DOCUMENTS ---
{text}"""

                    if client:
                        for ci, chunk in enumerate(chunks):
                            with st.spinner(f"🤖 Chunk {ci+1}/{len(chunks)}..."):
                                try:
                                    resp = client.messages.create(
                                        model="claude-sonnet-4-5", max_tokens=2000,
                                        messages=[{"role":"user","content":
                                            PROMPT_TPL.format(vname=vname, sprint=sprint,
                                                              n=ci+1, total=len(chunks),
                                                              text=chunk)}])
                                    for line in resp.content[0].text.splitlines():
                                        line = line.strip()
                                        if line.startswith("MOMENTUM:") and "None" not in line:
                                            parts = [p.strip() for p in line.split("|")]
                                            t = parts[0].replace("MOMENTUM:","").strip()
                                            e = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                                            s = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                                            dk = f"m_{t}_{e[:40]}"
                                            if dk not in seen_sigs:
                                                seen_sigs.add(dk)
                                                result["momentum"].append({"type":t,"evidence":e,"source":s})
                                        elif line.startswith("INVESTMENT:") and "None" not in line:
                                            parts = [p.strip() for p in line.split("|")]
                                            t = parts[0].replace("INVESTMENT:","").strip()
                                            e = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                                            s = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                                            dk = f"i_{t}_{e[:40]}"
                                            if dk not in seen_sigs:
                                                seen_sigs.add(dk)
                                                result["investment"].append({"type":t,"evidence":e,"source":s})
                                except Exception as ce:
                                    st.warning(f"Chunk {ci+1} error: {ce}")

                        st.session_state[sig_cache_key] = result
                        st.success(f"✅ Done — {len(result['momentum'])} momentum + {len(result['investment'])} investment signals found")
                    else:
                        st.warning("No API key — add ANTHROPIC_API_KEY to Streamlit secrets.")

                # Display results
                if sig_cache_key in st.session_state:
                    cached_sigs = st.session_state[sig_cache_key]
                    st.divider()
                    st.markdown("### 📈 Sprint Momentum Signals")
                    m_sigs = cached_sigs.get("momentum", [])
                    if m_sigs:
                        for s in m_sigs:
                            st.markdown(f"**✦ {s['type']}**  <small style='color:#6b7280'>({s.get('source','')})</small>", unsafe_allow_html=True)
                            if s.get("evidence"): st.markdown(f"> *{s['evidence']}*")
                            st.markdown("")
                    else:
                        st.info("No momentum signals found.")
                    st.divider()
                    st.markdown("### 💰 Self Investment Signals")
                    i_sigs = cached_sigs.get("investment", [])
                    if i_sigs:
                        for s in i_sigs:
                            st.markdown(f"**✦ {s['type']}**  <small style='color:#6b7280'>({s.get('source','')})</small>", unsafe_allow_html=True)
                            if s.get("evidence"): st.markdown(f"> *{s['evidence']}*")
                            st.markdown("")
                    else:
                        st.info("No investment signals found.")
                else:
                    st.info("👆 Click 'Get All Signals' to extract signals from all documents.")

            # TAB 4: AI Insights
            with tab4:
                if not client:
                    st.warning("Add Anthropic API key to enable AI insights.")
                else:
                    if st.button("Generate AI Insight", key=f"ai_{vname}"):
                        with st.spinner("Analyzing..."):
                            try:
                                docs = get_venture_files()
                                vc   = get_venture_common()
                                resp = client.messages.create(
                                    model="claude-sonnet-4-5", max_tokens=600,
                                    messages=[{"role":"user","content":
                                        f"""Venture: {vname}
Hub: {hub} | Sprint: {sprint} | Completion: {pct_num:.0f}%
Notes: {notes}
Feedback: {docs['fb']}
Transcript: {docs['tr']}
Common Docs: {vc}

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

            # TAB 5: Data Sources
            with tab5:
                st.markdown("#### All data sources read for this venture")
                st.caption("This shows exactly what the app found and used for RAG scoring")
                st.divider()

                # Source 1: Notes column
                st.markdown("**📊 Source 1: Notes/Comments (Excel Dashboard)**")
                if notes and notes not in ["—", ""]:
                    st.info(notes)
                else:
                    st.warning("No notes found in Excel for this venture")

                # Source 2: Venture folder files — lazy loaded
                st.markdown("**📁 Source 2: Venture Folder Files**")
                docs = get_venture_files()
                if docs["vf"]:
                    for ftype, fpath in docs["vf"].items():
                        st.caption(f"Found: `{ftype}` → `{Path(str(fpath)).name}`")
                    file_map = {"feedback": docs["fb"], "transcript": docs["tr"],
                                "sprint": docs["sp"], "journey": docs["jour"]}
                    for ftype, fpath in docs["vf"].items():
                        fname_display = Path(str(fpath)).name
                        fc = file_map.get(ftype, get_text(fpath))
                        if fc:
                            with st.expander(f"📄 {fname_display}"):
                                st.text(fc[:3000])
                else:
                    st.warning("No files found in venture folder")

                # Source 3: Common Documents — lazy loaded
                st.markdown("**📂 Source 3: Common Documents (venture-specific sections only)**")
                vc = get_venture_common()
                if vc:
                    with st.expander("View Common Documents excerpt"):
                        st.text(vc[:5000])
                else:
                    st.warning(f"No sections mentioning '{vname}' found in Common Documents")

                # Source 4: Attendance
                st.markdown("**📋 Source 4: Attendance Data**")
                sp_id_att = id(sp_reader) if sp_reader else 0
                att_data_preview = load_attendance_data(sp_id_att, use_sp, root_path)
                att_preview = get_attendance_for_venture(vname, att_data_preview)
                if att_preview:
                    st.success(f"✅ Found: {att_preview['sessions']} sessions — dates: {', '.join(att_preview['dates'])}")
                else:
                    st.warning("No attendance record found for this venture")

                st.divider()
                st.caption("💡 If $68K order is not showing above in any source, add it to the Notes column in your Excel or Common Documents for accurate scoring.")

# ══════════════════════════════════════════════════════
#  PROCESS BATCHES TAB
# ══════════════════════════════════════════════════════
with tab_process:
    st.title("⚙️ Process Batches")
    st.caption("Read all documents per venture, extract signals, score RAG accurately. Run once, view dashboard instantly.")
    st.divider()

    from processor import process_venture

    BATCH_RESULTS_KEY = "batch_results"
    BATCH_SIZE        = 10

    # ── status overview ───────────────────────────────
    batch_results = st.session_state.get(BATCH_RESULTS_KEY, {})
    done_count    = sum(1 for v in batch_results.values() if v.get("status") == "done")
    error_count   = sum(1 for v in batch_results.values() if v.get("status") == "error")
    total         = len(ventures_raw)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Ventures",  total)
    c2.metric("✅ Processed",    done_count)
    c3.metric("❌ Errors",       error_count)
    c4.metric("⏳ Remaining",    total - done_count - error_count)

    # Overall progress bar
    if total > 0:
        st.progress(done_count / total, text=f"{done_count}/{total} ventures processed")

    st.divider()

    if not client:
        st.warning("⚠️ Add ANTHROPIC_API_KEY to Streamlit secrets to enable batch processing.")
        st.stop()

    # ── batch controls ────────────────────────────────
    st.subheader("Run Batches")

    # Calculate batches
    batches = []
    for i in range(0, total, BATCH_SIZE):
        batch_ventures = ventures_raw[i:i+BATCH_SIZE]
        done_in_batch  = sum(1 for v in batch_ventures if batch_results.get(v,{}).get("status") == "done")
        batches.append({
            "num":       len(batches)+1,
            "ventures":  batch_ventures,
            "done":      done_in_batch,
            "total":     len(batch_ventures),
            "complete":  done_in_batch == len(batch_ventures)
        })

    # Load attendance once
    sp_id_b        = id(sp_reader) if sp_reader else 0
    attendance_b   = load_attendance_data(sp_id_b, use_sp, root_path)

    col_run_all, col_clear_all = st.columns([2,1])
    run_all   = col_run_all.button("🚀 Run All Batches", help="Process all ventures sequentially")
    clear_all = col_clear_all.button("🗑 Clear All Results", help="Reset all batch results")

    if clear_all:
        st.session_state.pop(BATCH_RESULTS_KEY, None)
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Show each batch
    for batch in batches:
        status_icon = "✅" if batch["complete"] else ("⏳" if batch["done"] > 0 else "⬜")
        with st.expander(f"{status_icon} Batch {batch['num']}  —  {batch['done']}/{batch['total']} done  —  {', '.join(batch['ventures'][:3])}{'...' if len(batch['ventures'])>3 else ''}"):

            # Show venture status in batch
            for vname_b in batch["ventures"]:
                vr = batch_results.get(vname_b, {})
                vs = vr.get("status","pending")
                icon = {"done":"✅","error":"❌","processing":"🔄"}.get(vs,"⬜")
                if vs == "done":
                    rag  = vr.get("rag",{})
                    nsig = len(vr.get("signals",{}).get("momentum",[])) + len(vr.get("signals",{}).get("investment",[]))
                    st.caption(f"{icon} {vname_b}  ·  Overall: {RAG_EMOJI.get(rag.get('overall_rag',''),'⚪')} {rag.get('overall_rag','')}  ·  {nsig} signals  ·  {vr.get('total_chars',0):,} chars  ·  {vr.get('num_chunks',1)} chunk(s)")
                elif vs == "error":
                    st.caption(f"{icon} {vname_b}  ·  Error: {vr.get('error','')[:80]}")
                else:
                    st.caption(f"{icon} {vname_b}  ·  Not processed yet")

            st.markdown("")
            run_batch = st.button(f"▶ Run Batch {batch['num']}", key=f"run_batch_{batch['num']}")

            if run_batch or (run_all and not batch["complete"]):
                prog_b = st.progress(0, text=f"Starting batch {batch['num']}...")
                for vi, vname_b in enumerate(batch["ventures"]):
                    if batch_results.get(vname_b,{}).get("status") == "done":
                        prog_b.progress((vi+1)/len(batch["ventures"]),
                                        text=f"Skipping {vname_b} (already done)")
                        continue

                    prog_b.progress(vi/len(batch["ventures"]),
                                    text=f"Processing {vname_b} ({vi+1}/{len(batch['ventures'])})...")

                    row    = get_row(vname_b)
                    notes  = cv(row, col_notes, default="")
                    sprint = cv(row, col_sprint)
                    pct    = row[col_pct] if (row is not None and col_pct) else None

                    result = process_venture(
                        client       = client,
                        vname        = vname_b,
                        venture_data = {},
                        load_v_files_fn    = load_v_files,
                        get_text_fn        = get_text,
                        extract_common_fn  = extract_venture_from_common,
                        load_common_fn     = load_common_docs,
                        get_attendance_fn  = get_attendance_for_venture,
                        attendance_data    = attendance_b,
                        notes        = notes,
                        sprint       = sprint,
                        pct_raw      = pct,
                    )

                    if BATCH_RESULTS_KEY not in st.session_state:
                        st.session_state[BATCH_RESULTS_KEY] = {}
                    st.session_state[BATCH_RESULTS_KEY][vname_b] = result

                prog_b.progress(1.0, text=f"✅ Batch {batch['num']} complete!")
                st.rerun()

    # ── sync batch results to portfolio RAG cache ─────
    if done_count > 0:
        st.divider()
        st.subheader("📊 Sync to Dashboard")
        st.caption("Push batch results to Portfolio Overview and Venture Cards")

        if st.button("🔄 Sync Results to Dashboard"):
            rag_cache = []
            for vname_b, vr in batch_results.items():
                if vr.get("status") != "done": continue
                row    = get_row(vname_b)
                rag    = vr.get("rag",{})
                rag_cache.append({
                    "name":             vname_b,
                    "hub":              cv(row, col_hub),
                    "vp":               cv(row, col_vp) if col_vp else "—",
                    "sprint":           cv(row, col_sprint),
                    "rev":              cv(row, col_rev),
                    "bucket":           get_stage_bucket(row[col_pct] if (row is not None and col_pct) else None),
                    "pct_raw":          row[col_pct] if (row is not None and col_pct) else None,
                    "notes":            cv(row, col_notes, default=""),
                    "overall_rag":      rag.get("overall_rag","ZERO"),
                    "momentum_rag":     rag.get("momentum_rag","ZERO"),
                    "investment_rag":   rag.get("investment_rag","ZERO"),
                    "momentum_reason":  rag.get("momentum_reason","—"),
                    "investment_reason":rag.get("investment_reason","—"),
                    "momentum_score":   rag.get("momentum_score",0),
                    "investment_score": rag.get("investment_score",0),
                })
            st.session_state["rag_scores_ai"] = rag_cache

            # Also cache signals per venture
            for vname_b, vr in batch_results.items():
                if vr.get("status") == "done" and vr.get("signals"):
                    st.session_state[f"signals_{vname_b}"] = vr["signals"]

            st.success(f"✅ Synced {len(rag_cache)} ventures to dashboard! Go to 📊 Portfolio Overview to view results.")
