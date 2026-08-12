"""
Batch processor for Portfolio Success Dashboard
Processes ventures in batches, reads all documents, extracts signals, scores RAG
NO character limitations — full chunking support
"""
import json, re
from pathlib import Path

CHUNK_SIZE = 120000  # Claude API hard limit per call

def chunk_text(text, size=CHUNK_SIZE):
    """Split at file boundaries — never cuts a file in the middle."""
    if len(text) <= size:
        return [text]
    parts   = text.split("\n\n=== ")
    chunks  = []
    current = ""
    for part in parts:
        section = part if part.startswith("===") else "=== " + part
        if len(current) + len(section) > size:
            if current: chunks.append(current)
            current = section
        else:
            current += ("\n\n" + section if current else section)
    if current: chunks.append(current)
    return chunks

def extract_signals_from_text(client, vname, sprint, full_text):
    """
    Extract ALL signals from full_text using chunking.
    Each signal is labelled POSITIVE or NEGATIVE by Claude at extraction time.
    No character limit — processes documents of any size.
    """
    chunks = chunk_text(full_text)
    result = {"momentum": [], "investment": []}
    seen   = set()

    PROMPT = """Venture:{vname}|Sprint Topic:{sprint}|Chunk {n}/{total}

A SIGNAL is evidence of a FOUNDER ACTION related to the sprint topic '{sprint}'.
Extract only what the FOUNDER DID or PLANS TO DO around this sprint.

SIGNAL CATEGORIES:
GREEN  = Founder has TAKEN ACTION — completed or meaningfully in progress
AMBER  = Founder has STATED A PLAN — intends to act but has not started yet
RED    = Founder has NOT acted AND has no plan, OR is disengaged from the sprint

DO NOT extract as signals:
- Descriptions of current state ("no international exposure yet", "market is competitive")
- Background context about the company or sector
- Observations without founder action ("export documentation is complex")
- Opinions or feelings without a concrete action ("founder feels it is difficult")

ONLY extract where there is clear evidence of what the FOUNDER DID or PLANS TO DO.

SPRINT MOMENTUM — founder actions around sprint engagement and progress:
- Attended or missed sprint sessions
- Completed or started sprint tasks / milestones
- Won orders / deals / contracts related to sprint '{sprint}'
- Took any concrete step toward sprint objectives
- Explicitly disengaged or dropped out of sprint

SELF INVESTMENT — founder actions committing resources to sprint '{sprint}':
- Hired staff for sprint-relevant roles (or explicitly decided not to)
- Purchased equipment / tools / software for sprint goals
- Invested own capital into sprint-related activities
- Entered a new market or channel as part of sprint
- Stated a concrete plan to invest (AMBER) or explicitly refused (RED)

FORMAT — one signal per line:
MOMENTUM:[what the founder did/plans]|EVIDENCE:[exact quote from text]|SOURCE:[doc name]|CATEGORY:GREEN
INVESTMENT:[what the founder did/plans]|EVIDENCE:[exact quote from text]|SOURCE:[doc name]|CATEGORY:AMBER

If genuinely no founder actions found in a category: MOMENTUM:None found

---DOCUMENTS---
{text}"""

    for i, chunk in enumerate(chunks):
        try:
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001", max_tokens=1500,
                messages=[{"role":"user","content":
                    PROMPT.format(vname=vname, sprint=sprint,
                                  n=i+1, total=len(chunks), text=chunk)}])
            for line in resp.content[0].text.splitlines():
                line = line.strip()
                if line.startswith("MOMENTUM:") and "None" not in line:
                    parts = [p.strip() for p in line.split("|")]
                    t   = parts[0].replace("MOMENTUM:","").strip()
                    e   = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s   = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    cat = parts[3].replace("CATEGORY:","").strip().upper() if len(parts)>3 else "GREEN"
                    if cat not in ["GREEN","AMBER","RED"]: cat = "GREEN"
                    dk  = f"m_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["momentum"].append({"type":t,"evidence":e,"source":s,"category":cat})
                elif line.startswith("INVESTMENT:") and "None" not in line:
                    parts = [p.strip() for p in line.split("|")]
                    t   = parts[0].replace("INVESTMENT:","").strip()
                    e   = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s   = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    cat = parts[3].replace("CATEGORY:","").strip().upper() if len(parts)>3 else "GREEN"
                    if cat not in ["GREEN","AMBER","RED"]: cat = "GREEN"
                    dk  = f"i_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["investment"].append({"type":t,"evidence":e,"source":s,"category":cat})
        except Exception as e:
            result.setdefault("errors",[]).append(f"Chunk {i+1}: {e}")

    return result, len(chunks)


def _nps_from_signals(signals_list):
    """
    Calculate NPS from GREEN/AMBER/RED categorised signals.
    GREEN = Promoters, AMBER = Passives, RED = Detractors

    NPS = % Promoters - % Detractors  (range: -100 to +100)

    RAG thresholds:
      NPS >= 20  → Green
      NPS 0–19   → Amber
      NPS < 0    → Red
      0 signals  → ZERO

    Returns (rag, nps, green_count, amber_count, red_count, total)
    """
    if not signals_list:
        return "ZERO", 0, 0, 0, 0, 0

    green = sum(1 for s in signals_list if s.get("category","GREEN") == "GREEN")
    amber = sum(1 for s in signals_list if s.get("category","GREEN") == "AMBER")
    red   = sum(1 for s in signals_list if s.get("category","GREEN") == "RED")
    total = green + amber + red

    pct_promoters  = round(green / total * 100)
    pct_detractors = round(red   / total * 100)
    nps = pct_promoters - pct_detractors

    if   nps >= 20: rag = "Green"
    elif nps >= 0:  rag = "Amber"
    else:           rag = "Red"

    return rag, nps, green, amber, red, total


def score_rag_from_signals(client, vname, sprint, notes, att_summary,
                            signals, pct_raw):
    """
    Score RAG using NPS from GREEN/AMBER/RED categorised signals.
    No separate Claude API call — pure formula.

    NPS = % Green (Promoters) - % Red (Detractors)
    Amber signals = Passives (counted in total, not in NPS numerator)

    NPS >= 20  → Green RAG
    NPS 0–19   → Amber RAG
    NPS < 0    → Red RAG
    0 signals  → ZERO

    Overall RAG = worst of Momentum + Investment (ZERO = no data, not worst)
    """
    m_sigs = signals.get("momentum",  [])
    i_sigs = signals.get("investment", [])

    m_rag, m_nps, m_g, m_a, m_r, m_tot = _nps_from_signals(m_sigs)
    i_rag, i_nps, i_g, i_a, i_r, i_tot = _nps_from_signals(i_sigs)

    # Overall RAG: worst of the two, ignoring ZERO
    order   = {"Red": 0, "Amber": 1, "Green": 2, "ZERO": 3}
    present = [r for r in [m_rag, i_rag] if r != "ZERO"]
    overall = min(present, key=lambda x: order.get(x, 3)) if present else "ZERO"

    def _reason(rag, nps, g, a, r, tot, category):
        if rag == "ZERO":
            return f"No {category} signals found."
        return f"Signal NPS {nps:+d} — {g} Green, {a} Amber, {r} Red of {tot} signals → {rag}."

    m_reason = _reason(m_rag, m_nps, m_g, m_a, m_r, m_tot, "momentum")
    i_reason = _reason(i_rag, i_nps, i_g, i_a, i_r, i_tot, "investment")

    score_matrix = {
        ("Green","Green"):10, ("Green","Amber"):8, ("Green","Red"):5, ("Green","ZERO"):5,
        ("Amber","Green"):8,  ("Amber","Amber"):7, ("Amber","Red"):3, ("Amber","ZERO"):3,
        ("Red",  "Green"):5,  ("Red",  "Amber"):3, ("Red",  "Red"):1, ("Red",  "ZERO"):0,
        ("ZERO", "Green"):5,  ("ZERO", "Amber"):3, ("ZERO", "Red"):0, ("ZERO", "ZERO"):0,
    }
    numeric_score = score_matrix.get((m_rag, i_rag), 0)

    # ── Overall note — formula-based, stored in repo, zero extra API cost ──
    all_g   = m_g   + i_g
    all_a   = m_a   + i_a
    all_r   = m_r   + i_r
    all_tot = m_tot + i_tot
    overall_nps = (
        round(all_g / all_tot * 100) - round(all_r / all_tot * 100)
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
        "momentum_rag":      m_rag,
        "investment_rag":    i_rag,
        "overall_rag":       overall,
        "momentum_reason":   m_reason,
        "investment_reason": i_reason,
        "overall_note":      overall_note,
        "momentum_score":    numeric_score,
        "investment_score":  numeric_score,
        "momentum_nps":      m_nps,
        "investment_nps":    i_nps,
        "momentum_green":    m_g,
        "momentum_amber":    m_a,
        "momentum_red":      m_r,
        "momentum_total":    m_tot,
        "investment_green":  i_g,
        "investment_amber":  i_a,
        "investment_red":    i_r,
        "investment_total":  i_tot,
    }



def extract_session_feedback(client, vname, transcript_text, feedback_text):
    """
    Extract structured session data from transcript and/or feedback files.
    Returns a list of session records — one dict per session found.
    If no data, returns a single record with "Not Available" for all fields.
    """
    has_transcript = bool(transcript_text and len(transcript_text.strip()) > 50)
    has_feedback   = bool(feedback_text   and len(feedback_text.strip())   > 50)

    if not has_transcript and not has_feedback:
        return [{
            "mentor_name":     "Not Available",
            "session_date":    "Not Available",
            "topics_discussed":"Not Available",
            "key_outputs":     "Not Available",
            "founder_feedback":"Not Available",
            "session_summary": "Not Available",
            "sources_used":    [],
        }]

    sources_used = []
    combined     = ""
    if has_transcript:
        sources_used.append("Transcript")
        combined += f"\n\n=== SESSION TRANSCRIPT ===\n{transcript_text}"
    if has_feedback:
        sources_used.append("Feedback")
        combined += f"\n\n=== FEEDBACK FILE ===\n{feedback_text}"

    # Chunk if needed
    chunks = chunk_text(combined)

    PROMPT = """You are extracting structured session data for venture: {vname}

From the documents below, extract ALL sessions mentioned. For each session, extract:
1. Mentor/Advisor Name (who ran the session)
2. Session Date (any date format found)
3. Topics Discussed (what was discussed — raw extract)
4. Key Outputs / Action Items (decisions made, next steps)
5. Founder Feedback on Mentor/Session (what founder said about the session quality)
6. Overall Session Summary (2-3 sentence Claude summary of the session)

If any field is missing in the source, write exactly: Not Available

Return ONLY a JSON array. Each element is one session. Example:
[
  {{
    "mentor_name": "Rajesh Kumar",
    "session_date": "15-Mar-2026",
    "topics_discussed": "Export documentation, buyer negotiations with Germany client",
    "key_outputs": "1. Draft proforma invoice by 20 March 2. Follow up with freight forwarder",
    "founder_feedback": "Very useful session, mentor helped clarify DGFT process",
    "session_summary": "Session focused on export readiness. Mentor walked through documentation requirements. Founder confirmed next steps on buyer negotiation."
  }}
]

If multiple sessions exist, return multiple objects in the array.
If you cannot identify distinct sessions, return one object for the overall content.

--- DOCUMENTS (Chunk {n}/{total}) ---
{text}"""

    all_sessions = []
    seen_sessions = set()

    for i, chunk in enumerate(chunks):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-5", max_tokens=3000,
                messages=[{"role":"user","content":
                    PROMPT.format(vname=vname, n=i+1, total=len(chunks), text=chunk)}])
            raw = re.sub(r"```json|```","",resp.content[0].text.strip()).strip()
            sessions = json.loads(raw)
            if isinstance(sessions, dict): sessions = [sessions]
            for s in sessions:
                # Dedup by mentor+date
                dk = f"{s.get('mentor_name','')}_{s.get('session_date','')}"
                if dk not in seen_sessions:
                    seen_sessions.add(dk)
                    s["sources_used"] = sources_used
                    all_sessions.append(s)
        except Exception as e:
            pass  # If parsing fails, continue with other chunks

    if not all_sessions:
        # Extraction ran but yielded nothing parseable
        return [{
            "mentor_name":     "Not Available",
            "session_date":    "Not Available",
            "topics_discussed":"Not Available",
            "key_outputs":     "Not Available",
            "founder_feedback":"Not Available",
            "session_summary": "Could not parse session data from documents.",
            "sources_used":    sources_used,
        }]

    return all_sessions


def process_venture(client, vname, venture_data, load_v_files_fn,
                    get_text_fn, extract_common_fn, load_common_fn,
                    get_attendance_fn, attendance_data, notes, sprint, pct_raw):
    """
    Fully process one venture:
    1. Read ALL documents (no character limit)
    2. Extract ALL signals via chunking
    3. Score RAG from those signals (always in sync)
    """
    result = {"name": vname, "status": "processing"}
    try:
        # 1. Load all venture files
        vfiles = load_v_files_fn(vname)
        fb     = get_text_fn(vfiles["feedback"])   if "feedback"   in vfiles else ""
        tr     = get_text_fn(vfiles["transcript"]) if "transcript" in vfiles else ""
        sp     = get_text_fn(vfiles["sprint"])     if "sprint"     in vfiles else ""
        jour   = get_text_fn(vfiles["journey"])    if "journey"    in vfiles else ""
        others = [get_text_fn(p) for k,p in vfiles.items() if k.startswith("other_")]
        others = [t for t in others if t]

        # 2. Extract venture sections from pre-loaded common docs
        common_text    = load_common_fn()
        venture_common = extract_common_fn(vname, common_text)

        # 3. Build full text — NO limits
        sources = {
            "Notes":          notes or "",
            "Feedback":       fb,
            "Transcript":     tr,
            "Sprint Plan":    sp,
            "Growth Journey": jour,
            "Common Docs":    venture_common,
        }
        for idx, ot in enumerate(others):
            sources[f"Venture File {idx+1}"] = ot

        full_text    = "\n\n".join(f"=== {k} ===\n{v}" for k,v in sources.items() if v)
        total_chars  = len(full_text)
        sources_used = [k for k,v in sources.items() if v]

        # 4. Attendance
        att          = get_attendance_fn(vname, attendance_data)
        att_sessions = att["sessions"]     if att else 0
        att_dates    = att["dates"]        if att else []
        att_summary  = f"{att_sessions} sessions ({', '.join(att_dates)})" if att_sessions else "No attendance data"

        # 5. Extract ALL signals — chunked, no limit
        signals, num_chunks = extract_signals_from_text(client, vname, sprint, full_text)

        # 6. Score RAG from complete signals — always in sync
        rag = score_rag_from_signals(client, vname, sprint, notes,
                                      att_summary, signals, pct_raw)

        result.update({
            "status":       "done",
            "signals":      signals,
            "rag":          rag,
            "total_chars":  total_chars,
            "num_chunks":   num_chunks,
            "sources_used": sources_used,
            "att_sessions": att_sessions,
            "att_dates":    att_dates,
        })

    except Exception as e:
        result.update({
            "status": "error", "error": str(e),
            "rag": {"momentum_rag":"Unknown","investment_rag":"Unknown",
                    "overall_rag":"Unknown","momentum_reason":str(e),
                    "investment_reason":"—","momentum_score":0,"investment_score":0},
            "signals": {"momentum":[],"investment":[]}
        })
    return result


def parse_tracker_files(session_tracker_bytes, feedback_tracker_bytes):
    """
    Parse 05_Session_Management_Tracker and 06_Feedback_Quality_Tracker
    into structured mentor_insights dict keyed by mentor name.

    Returns:
        mentor_insights: {mentor_name: {mentor_name, total_sessions,
                          ventures_worked, avg_rating, sessions: [...]}}
    """
    import pandas as pd
    import io
    from difflib import SequenceMatcher

    def safe_str(val):
        if val is None: return "Not Available"
        s = str(val).strip()
        return "Not Available" if s in ["nan","None","NaT",""] else s

    def safe_float(val):
        try:
            f = float(val)
            return round(f, 1) if not (f != f) else None  # NaN check
        except: return None

    def fuzzy_match(a, b):
        a_l = str(a).lower().strip()
        b_l = str(b).lower().strip()
        if a_l == b_l: return True
        ratio = SequenceMatcher(None, a_l, b_l).ratio()
        return ratio >= 0.85

    # Load Session Tracker
    df_sess = pd.read_excel(
        io.BytesIO(session_tracker_bytes),
        sheet_name="Session Tracker"
    )

    # Load Feedback Quality Tracker — two sheets
    df_fb = pd.read_excel(
        io.BytesIO(feedback_tracker_bytes),
        sheet_name="Session Feedback"
    )
    df_mfb = pd.read_excel(
        io.BytesIO(feedback_tracker_bytes),
        sheet_name="Feedback from Mentor"
    )

    # Build founder feedback lookup: {(venture_lower, mentor_lower): row}
    fb_lookup = {}
    for _, row in df_fb.iterrows():
        vn = safe_str(row.get("Venture Name",""))
        mn = safe_str(row.get("Mentor Name",""))
        if vn != "Not Available" and mn != "Not Available":
            fb_lookup[(vn.lower(), mn.lower())] = row

    # Build mentor feedback lookup: {venture_lower: row}
    mfb_lookup = {}
    for _, row in df_mfb.iterrows():
        vn = safe_str(row.get("Venture Name",""))
        if vn != "Not Available":
            mfb_lookup[vn.lower()] = row

    def get_founder_feedback(venture, mentor):
        """Fuzzy match founder feedback by venture+mentor."""
        vl = venture.lower(); ml = mentor.lower()
        # Exact match first
        if (vl, ml) in fb_lookup:
            r = fb_lookup[(vl, ml)]
            return {
                "overall_rating":      safe_float(r.get("Overall Rating (1-5)")),
                "usefulness":          safe_str(r.get("How useful was this mentor session for your current business priorities?")),
                "actionability":       safe_str(r.get("Actionability of Advice")),
                "followup_requested":  safe_str(r.get("Follow-Up Requested?")),
                "verbatim":            safe_str(r.get("Verbatim Feedback")),
                "flagged":             safe_str(r.get("Flagged (≤3)?")),
                "feedback_date":       safe_str(r.get("Date")),
            }
        # Fuzzy match
        for (fvn, fmn), r in fb_lookup.items():
            if fuzzy_match(vl, fvn) and fuzzy_match(ml, fmn):
                return {
                    "overall_rating":     safe_float(r.get("Overall Rating (1-5)")),
                    "usefulness":         safe_str(r.get("How useful was this mentor session for your current business priorities?")),
                    "actionability":      safe_str(r.get("Actionability of Advice")),
                    "followup_requested": safe_str(r.get("Follow-Up Requested?")),
                    "verbatim":           safe_str(r.get("Verbatim Feedback")),
                    "flagged":            safe_str(r.get("Flagged (≤3)?")),
                    "feedback_date":      safe_str(r.get("Date")),
                }
        return None

    def get_mentor_feedback(venture):
        """Fuzzy match mentor feedback by venture name."""
        vl = venture.lower()
        if vl in mfb_lookup:
            r = mfb_lookup[vl]
        else:
            r = next((v for k,v in mfb_lookup.items() if fuzzy_match(vl,k)), None)
        if r is None: return None
        return {
            "mentor_name":        safe_str(r.get("Your Name")),
            "agenda_relevant":    safe_str(r.get("1. Was the session agenda relevant to the startup's current stage and specific needs?")),
            "mentee_prepared":    safe_str(r.get("5. How prepared and organized was your mentee for the session?")),
            "mentee_engaged":     safe_str(r.get("6. How engaged is your mentee during mentoring sessions, in terms of active participation and willingness to discuss challenges?")),
            "session_rating":     safe_str(r.get("8. How would you like to rate the session?")),
            "action_items_relevant": safe_str(r.get("9. How relevant were the recommended action items to the discussion during your session?")),
            "improvements":       safe_str(r.get("7. Is there anything the venture (or platform) could have done better to make this session more effective?")),
        }

    # Build mentor_insights
    mentor_insights = {}

    for _, row in df_sess.iterrows():
        venture = safe_str(row.get("Venture Name",""))
        mentor  = safe_str(row.get("Mentor Name",""))
        if venture == "Not Available" or mentor == "Not Available":
            continue

        # Get date
        raw_date = row.get("Meeting Date")
        try:
            date_str = pd.to_datetime(raw_date).strftime("%Y-%m-%d")
        except: date_str = safe_str(raw_date)

        # Get rating from tracker
        tracker_rating = safe_float(row.get("Feedback Rating (1-5)"))

        # Founder feedback from quality tracker
        founder_fb = get_founder_feedback(venture, mentor)

        # Mentor feedback
        mentor_fb  = get_mentor_feedback(venture)

        session_record = {
            "meeting_id":              safe_str(row.get("Meeting ID")),
            "venture_name":            venture,
            "hub":                     safe_str(row.get("Hub")),
            "program_tier":            safe_str(row.get("Program Tier")),
            "meeting_date":            date_str,
            "session_type":            safe_str(row.get("Session Type")),
            "ask":                     safe_str(row.get("Ask")),
            "duration_min":            safe_float(row.get("Duration (min)")),
            "meeting_summary":         safe_str(row.get("Meeting Summary")),
            "next_steps":              safe_str(row.get("Next Steps / Action Items")),
            "followup_required":       safe_str(row.get("Follow-Up Required?")),
            "followup_status":         safe_str(row.get("Follow-Up Status")),
            "tracker_rating":          tracker_rating,
            "tracker_feedback":        safe_str(row.get("Feedback Comments")),
            "rn_team_member":          safe_str(row.get("RN Team Member")),
            "session_paid":            safe_str(row.get("Session Paid or Probono")),
            # Enriched from quality tracker
            "founder_feedback":        founder_fb,
            "mentor_feedback":         mentor_fb,
        }

        # Add to mentor_insights
        if mentor not in mentor_insights:
            mentor_insights[mentor] = {
                "mentor_name":     mentor,
                "total_sessions":  0,
                "ventures_worked": [],
                "ratings":         [],
                "sessions":        [],
            }

        mentor_insights[mentor]["sessions"].append(session_record)
        mentor_insights[mentor]["total_sessions"] += 1
        if venture not in mentor_insights[mentor]["ventures_worked"]:
            mentor_insights[mentor]["ventures_worked"].append(venture)

        # Collect ratings for avg
        rating = (founder_fb.get("overall_rating") if founder_fb else None) or tracker_rating
        if rating: mentor_insights[mentor]["ratings"].append(rating)

    # Compute avg rating per mentor
    for mn, mdata in mentor_insights.items():
        ratings = mdata.pop("ratings", [])
        mdata["avg_rating"] = round(sum(ratings)/len(ratings), 1) if ratings else None
        # Sort sessions by date descending
        mdata["sessions"].sort(
            key=lambda s: s.get("meeting_date",""), reverse=True)

    return mentor_insights


def extract_journey_document_data(client, vname, journey_text):
    """
    Extract structured venture data from a Journey Document (PDF/Word).
    Returns a dict with all Company Basics fields.
    Uses claude-haiku for cost efficiency.
    """
    if not journey_text or len(journey_text.strip()) < 100:
        return {}

    prompt = f"""Extract structured data for venture: {vname}

From the Journey Document below, extract the following fields.
Return ONLY a JSON object. Use null for any field not found.

{{
  "existing_product": "current products/services offered",
  "existing_market_segments": "current customer segments served",
  "existing_geographies": "current geographies of operation",
  "new_product": "new products/services planned",
  "new_market_segments": "new market segments to be targeted",
  "new_geographies": "new geographies to enter",
  "incremental_rev_3yr": "incremental revenue target over 3 years (number only, in $ millions)",
  "incremental_jobs_3yr": "incremental jobs to be created over 3 years (number only)",
  "goal_gtm": "GTM / sales / marketing goal",
  "goal_product": "product development goal",
  "goal_operations": "operations goal",
  "goal_supply_chain": "supply chain goal",
  "goal_people": "people / HR goal",
  "goal_finance": "finance goal",
  "stream_support_gtm": "GTM support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "stream_support_product": "product support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "stream_support_operations": "operations support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "stream_support_supply_chain": "supply chain support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "stream_support_hr": "HR/people support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "stream_support_finance": "finance support needed: RED/AMBER/GREEN/DEEP SUPPORT or null",
  "unlock_gtm": "GTM stream unlock description",
  "unlock_product": "product stream unlock description",
  "unlock_operations": "operations stream unlock description",
  "unlock_supply_chain": "supply chain stream unlock description",
  "unlock_people": "people stream unlock description",
  "unlock_finance": "finance stream unlock description"
}}

--- JOURNEY DOCUMENT ---
{journey_text[:80000]}"""

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}]
        )
        raw  = re.sub(r"```json|```", "", resp.content[0].text.strip()).strip()
        data = json.loads(raw)
        # Clean nulls
        return {k: v for k, v in data.items() if v is not None and str(v).strip() not in ["","null","None"]}
    except Exception as e:
        return {"_error": str(e)}


def synthesise_value_delivered(client, signals_repo, feedback_repo, company_basics):
    """
    Synthesise portfolio-level value delivered from all available data.
    One Claude call covering:
      1. Jobs generated in 2026 (from hiring signals)
      2. Gap categories resolved (from session summaries)
      3. Actionable business directions (from next steps)
      4. Problems solved categories (from signals + sessions)

    Returns structured dict ready for frontend display.
    """
    # ── Collect all hiring signals ──────────────────
    hiring_signals = []
    all_momentum_signals = []
    all_investment_signals = []

    if signals_repo:
        for vn, vdata in signals_repo.get("venture_summary", {}).items():
            sigs = vdata.get("signals", {})
            for s in sigs.get("investment", []):
                ev = s.get("evidence","").lower()
                if any(w in ev for w in ["hire","hired","staff","employ","recruit",
                                          "team","people","headcount","joined"]):
                    hiring_signals.append({
                        "venture": vn,
                        "evidence": s.get("evidence",""),
                        "source":   s.get("source",""),
                    })
                all_investment_signals.append({
                    "venture":  vn,
                    "type":     s.get("type",""),
                    "evidence": s.get("evidence",""),
                    "category": s.get("category",""),
                })
            for s in sigs.get("momentum", []):
                all_momentum_signals.append({
                    "venture":  vn,
                    "type":     s.get("type",""),
                    "evidence": s.get("evidence",""),
                    "category": s.get("category",""),
                })

    # ── Collect all session summaries + next steps ──
    session_summaries = []
    if feedback_repo:
        for mn, mdata in feedback_repo.get("mentor_insights", {}).items():
            for s in mdata.get("sessions", []):
                summary    = s.get("meeting_summary","")
                next_steps = s.get("next_steps","")
                ask        = s.get("ask","")
                venture    = s.get("venture_name","")
                date       = s.get("meeting_date","")
                if summary or next_steps:
                    session_summaries.append({
                        "venture":    venture,
                        "mentor":     mn,
                        "date":       date,
                        "ask":        ask,
                        "summary":    summary,
                        "next_steps": next_steps,
                    })

    # ── Build context for Claude ────────────────────
    hiring_text = "\n".join(
        f"- {h['venture']}: {h['evidence']}"
        for h in hiring_signals[:80]
    ) or "No explicit hiring signals found."

    momentum_text = "\n".join(
        f"- [{s['venture']}] {s['type']}: {s['evidence']}"
        for s in all_momentum_signals[:100]
    ) or "No momentum signals."

    investment_text = "\n".join(
        f"- [{s['venture']}] {s['type']}: {s['evidence']}"
        for s in all_investment_signals[:100]
    ) or "No investment signals."

    sessions_text = "\n".join(
        f"- [{s['venture']} | {s['date']}] Ask: {s['ask']} | Summary: {s['summary'][:200]} | Next Steps: {s['next_steps'][:150]}"
        for s in session_summaries[:100]
    ) or "No session data."

    total_ventures = len(signals_repo.get("venture_summary", {})) if signals_repo else 0
    total_sessions = len(session_summaries)

    prompt = f"""You are analysing a portfolio of {total_ventures} ventures in an accelerator program.
Based on the data below, extract and synthesise portfolio-level insights.

CRITICAL: Return ONLY valid JSON. No markdown, no backticks, no explanatory text.
Use only double quotes for strings. Do not include newlines or tabs inside string values — use spaces instead.
Escape any double quotes inside string values with backslash.

Return ONLY a JSON object with this exact structure:
{{
  "jobs_2026": {{
    "total_estimated": <integer — best estimate of total jobs created in 2026 across all ventures>,
    "confidence": "High/Medium/Low",
    "evidence_count": <number of ventures with explicit hiring evidence>,
    "top_examples": [
      {{"venture": "...", "jobs_detail": "exact hiring evidence from signal"}}
    ]
  }},
  "gap_categories": [
    {{
      "category": "GTM / Market Access",
      "session_count": <int>,
      "description": "one sentence on what gaps were addressed",
      "example_actions": ["action 1", "action 2"]
    }}
  ],
  "actionable_directions": {{
    "summary": "3-5 sentence synthesis of the most common actionable directions that emerged across all sessions",
    "top_directions": [
      "Direction 1 — concise action-oriented statement",
      "Direction 2",
      "Direction 3",
      "Direction 4",
      "Direction 5"
    ]
  }},
  "problems_solved": [
    {{
      "category": "category name",
      "count": <number of ventures where this problem was addressed>,
      "description": "one line description"
    }}
  ]
}}

Use these gap categories for gap_categories: GTM / Market Access, Product / Quality & Certification,
Operations / Manufacturing, Finance & Working Capital, People / HR & Org, Supply Chain & Procurement.
Include all 6 even if count is 0.

For problems_solved, identify 5-8 distinct problem categories actually resolved (not targets).

--- HIRING SIGNALS (for jobs_2026) ---
{hiring_text}

--- MOMENTUM SIGNALS (for problems solved + directions) ---
{momentum_text}

--- INVESTMENT SIGNALS (for problems solved + directions) ---
{investment_text}

--- SESSION SUMMARIES (for gap categories + directions) ---
{sessions_text}"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = resp.content[0].text.strip()

        # Strip markdown fences
        raw = re.sub(r"```json|```", "", raw).strip()

        # Find JSON object boundaries — extract only the JSON part
        start = raw.find("{")
        end   = raw.rfind("}") + 1
        if start == -1 or end == 0:
            return None, "No JSON object found in response"
        raw = raw[start:end]

        # Fix common JSON issues: unescaped newlines inside strings
        # Replace literal newlines inside string values with \n
        def fix_json_string(s):
            result = []
            in_string = False
            escape_next = False
            for ch in s:
                if escape_next:
                    result.append(ch)
                    escape_next = False
                    continue
                if ch == "\\":
                    escape_next = True
                    result.append(ch)
                    continue
                if ch == '"' and not escape_next:
                    in_string = not in_string
                    result.append(ch)
                    continue
                if in_string and ch == "\n":
                    result.append("\\n")
                    continue
                if in_string and ch == "\t":
                    result.append("\\t")
                    continue
                result.append(ch)
            return "".join(result)

        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            raw_fixed = fix_json_string(raw)
            try:
                data = json.loads(raw_fixed)
            except json.JSONDecodeError as je:
                # Last resort: use ast.literal_eval won't work for JSON
                # Try to extract partial data
                return None, f"JSON parse error: {je}. Raw (first 300 chars): {raw[:300]}"

        data["total_ventures"]  = total_ventures
        data["total_sessions"]  = total_sessions
        data["hiring_signals"]  = len(hiring_signals)
        data["total_momentum"]  = len(all_momentum_signals)
        data["total_investment"]= len(all_investment_signals)
        return data, None
    except Exception as e:
        return None, str(e)


def extract_people_hired_data(client, vname, full_text):
    """
    Extract hiring information for a venture from all available documents.
    Returns structured dict with:
    - current_employee_count
    - hired_2026_count + hired_2026_description
    - hired_2027_count + hired_2027_description
    - planned_2026_count + planned_2026_description
    - planned_2027_count + planned_2027_description
    - resources_hired_count / description (total confirmed hires)
    - hiring_plan_6mo_count / description (near-term plan)
    - role_breakdown by function
    Uses claude-haiku for cost efficiency.
    """
    DNA = "Data Not Available"
    empty = {
        "current_employee_count":     DNA,
        "resources_hired_count":      0,
        "resources_hired_description":DNA,
        "hired_2026_count":           0,
        "hired_2026_description":     DNA,
        "hired_2027_count":           0,
        "hired_2027_description":     DNA,
        "hiring_plan_6mo_count":      0,
        "hiring_plan_6mo_description":DNA,
        "planned_2026_count":         0,
        "planned_2026_description":   DNA,
        "planned_2027_count":         0,
        "planned_2027_description":   DNA,
        "role_breakdown": {
            "gtm":DNA,"product":DNA,"operations":DNA,
            "supply_chain":DNA,"hr":DNA,"finance":DNA,
        }
    }
    if not full_text or len(full_text.strip()) < 100:
        return empty

    chunks = chunk_text(full_text)

    PROMPT = """Venture: {vname} | Chunk {n}/{total}

Extract ALL people/hiring information from the text below.

Look for:
1. CURRENT EMPLOYEES — total headcount / number of staff currently employed
2. ALREADY HIRED IN 2026 — roles confirmed hired during 2026
3. ALREADY HIRED IN 2027 — roles confirmed hired during 2027
4. PLANNED HIRE 2026 — roles the founder plans/intends to hire in 2026 (not yet hired)
5. PLANNED HIRE 2027 — roles the founder plans/intends to hire in 2027 (not yet hired)
6. 6-MONTH PLAN — any near-term hiring intent (if no year specified)

For each hire, note the role and function: GTM, Product, Operations, Supply Chain, HR, Finance.

Return ONLY valid JSON, no markdown:
{{
  "current_employee_count": <integer or null>,
  "hired_2026": [{{"role": "title", "function": "GTM/Product/Operations/Supply Chain/HR/Finance"}}],
  "hired_2027": [{{"role": "title", "function": "..."}}],
  "planned_2026": [{{"role": "title", "function": "..."}}],
  "planned_2027": [{{"role": "title", "function": "..."}}],
  "planned_6mo": [{{"role": "title", "function": "..."}}]
}}

Use empty arrays [] if nothing found. Do NOT invent — only extract what is explicitly stated.

--- DOCUMENTS (Chunk {n}/{total}) ---
{text}"""

    emp_count_raw = []
    all_hired_2026 = []; all_hired_2027 = []
    all_planned_2026 = []; all_planned_2027 = []; all_planned_6mo = []

    def dedup(lst, new_items):
        existing = {f"{x.get('role','')}_{x.get('function','')}" for x in lst}
        for item in new_items:
            dk = f"{item.get('role','')}_{item.get('function','')}"
            if dk not in existing:
                existing.add(dk)
                lst.append(item)

    for i, chunk in enumerate(chunks):
        try:
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=1500,
                messages=[{"role":"user","content":
                    PROMPT.format(vname=vname, n=i+1, total=len(chunks), text=chunk)}]
            )
            raw = resp.content[0].text.strip()
            raw = re.sub(r"```json|```","",raw).strip()
            s = raw.find("{"); e = raw.rfind("}")+1
            if s == -1 or e == 0: continue
            data = json.loads(raw[s:e])

            if data.get("current_employee_count") is not None:
                try: emp_count_raw.append(int(data["current_employee_count"]))
                except: pass

            dedup(all_hired_2026,  data.get("hired_2026",[]))
            dedup(all_hired_2027,  data.get("hired_2027",[]))
            dedup(all_planned_2026,data.get("planned_2026",[]))
            dedup(all_planned_2027,data.get("planned_2027",[]))
            dedup(all_planned_6mo, data.get("planned_6mo",[]))
        except Exception:
            continue

    # All confirmed hires = 2026 + 2027 + 6mo (already done)
    all_hired = all_hired_2026 + all_hired_2027
    all_planned = all_planned_2026 + all_planned_2027 + all_planned_6mo

    # Role breakdown by function (from all confirmed hires)
    FUNCTIONS = ["GTM","Product","Operations","Supply Chain","HR","Finance"]
    role_breakdown = {}
    for fn in FUNCTIONS:
        fn_key = fn.lower().replace(" ","_")
        fn_hired = [h for h in all_hired if h.get("function","").upper()==fn.upper()]
        role_breakdown[fn_key] = (
            f"{len(fn_hired)} ({', '.join(h.get('role','') for h in fn_hired)})"
            if fn_hired else DNA
        )

    def desc(lst, prefix=""):
        if not lst: return DNA
        return prefix + ", ".join(
            f"{h.get('role','')} ({h.get('function','')})" for h in lst)

    current_emp = max(emp_count_raw) if emp_count_raw else DNA

    if all_planned:
        planned_desc = "Planned: " + ", ".join(
            f"{p.get('role','')} ({p.get('function','')})" for p in all_planned
        )
    else:
        planned_desc = DNA

    return {
        "current_employee_count":      str(current_emp) if current_emp != DNA else DNA,
        "resources_hired_count":       len(all_hired),
        "resources_hired_description": desc(all_hired, "Hired: "),
        "hired_2026_count":            len(all_hired_2026),
        "hired_2026_description":      desc(all_hired_2026),
        "hired_2027_count":            len(all_hired_2027),
        "hired_2027_description":      desc(all_hired_2027),
        "hiring_plan_6mo_count":       len(all_planned),
        "hiring_plan_6mo_description": planned_desc,
        "planned_2026_count":          len(all_planned_2026),
        "planned_2026_description":    desc(all_planned_2026),
        "planned_2027_count":          len(all_planned_2027),
        "planned_2027_description":    desc(all_planned_2027),
        "role_breakdown":              role_breakdown,
    }


# ══════════════════════════════════════════════════════════════════
#  CALL INTELLIGENCE EXTRACTION
#  Extracts 7 structured fields per call from transcript + feedback
#  + sprint plan + journey doc
# ══════════════════════════════════════════════════════════════════

CALL_INTEL_SYSTEM = (
    "You are a venture analyst extracting structured call intelligence from NEN Accelerate "
    "program documents. Return only valid JSON arrays. Never invent data — write "
    "\"Not Available\" for any field not found in the source text."
)

CALL_INTEL_PROMPT = """Venture: {vname} | Sprint: {sprint} | Chunk {n}/{total}

Extract ALL individual call or session records from the documents below.
One document may contain multiple sessions — identify and return each separately.

CURRENT SPRINT GAPS (already being worked on — do NOT list these as new gaps):
{sprint_gaps}

For EACH call/session found, extract ALL of the following 7 fields:

1. call_objective        — What was the stated purpose or agenda of this call?
2. discussion_topics     — Major topics covered (3-6 bullet points as a list)
3. company_updates       — What did the venture/founder report as updates on their business?
4. venture_partner_guidance — What guidance or advice did the mentor/VP/expert give?
5. key_decisions         — What decisions were made or agreed during this call? (as a list)
6. action_items          — Specific tasks assigned, with owner and deadline if mentioned (as a list)
7. risks_challenges      — What risks, blockers or challenges were identified? (as a list)

Also extract:
- expert_name    — name of the mentor, expert, or VP on the call
- expert_role    — their role: "Expert" / "VP" / "Sprint Partner" / "Panel"
- session_date   — date of the call (any format found in text)
- call_type      — "Sprint call" / "Panel call" / "VP call" / "Mid-sprint VP review" / "Other"
- venture_feedback_rating — numeric rating given by founder (1-5) or null
- venture_feedback_text   — what the founder said about the quality of the session

Return ONLY a valid JSON array — one object per call:
[
  {{
    "expert_name": "Name or Not Available",
    "expert_role": "Expert",
    "session_date": "date or Not Available",
    "call_type": "Sprint call",
    "call_objective": "...",
    "discussion_topics": ["topic 1", "topic 2", "topic 3"],
    "company_updates": "...",
    "venture_partner_guidance": "...",
    "key_decisions": ["decision 1", "decision 2"],
    "action_items": ["Owner: task by date", "Owner: task by date"],
    "risks_challenges": ["risk 1", "risk 2"],
    "venture_feedback_rating": null,
    "venture_feedback_text": "Not Available"
  }}
]

Return [] if no sessions found in this chunk. No markdown fences.

--- DOCUMENTS (Chunk {n}/{total}) ---
{text}"""


def extract_call_intelligence(client, vname, sprint,
                               transcript_text, feedback_text,
                               sprint_plan_text, journey_text):
    """
    Extract structured call intelligence records from all document sources.

    Per-call record contains all 7 fields:
      1. call_objective
      2. discussion_topics
      3. company_updates
      4. venture_partner_guidance
      5. key_decisions
      6. action_items
      7. risks_challenges

    Plus metadata: expert_name, expert_role, session_date, call_type,
    venture_feedback_rating, venture_feedback_text.

    Sources:
      - transcript_text    → main source for all 7 fields
      - feedback_text      → venture_feedback_rating + venture_feedback_text
      - sprint_plan_text   → current sprint gaps (context only, not repeated as new gaps)
      - journey_text       → current sprint gaps context

    Returns list of call record dicts.
    """

    # ── Step 1: Extract current sprint gaps from Sprint Plan + Journey ──
    sprint_gaps_context = "Not Available"
    if sprint_plan_text or journey_text:
        gap_sources = []
        if sprint_plan_text:
            gap_sources.append(f"=== SPRINT PLAN ===\n{sprint_plan_text}")
        if journey_text:
            gap_sources.append(f"=== GROWTH JOURNEY DOCUMENT ===\n{journey_text}")
        gap_combined = "\n\n".join(gap_sources)

        GAP_PROMPT = f"""Venture: {vname} | Sprint: {sprint}

From the Sprint Plan and Growth Journey Document below, list:
1. The specific gaps or challenges the venture is currently working on in this sprint
2. The key sprint objectives or milestones

Return ONLY a JSON object:
{{
  "current_sprint_gaps": ["Gap 1", "Gap 2"],
  "sprint_objectives": ["Objective 1", "Objective 2"]
}}

Return empty arrays if nothing found. No markdown.

--- DOCUMENTS ---
{gap_combined[:60000]}"""

        try:
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=600,
                messages=[{"role": "user", "content": GAP_PROMPT}]
            )
            raw = re.sub(r"```json|```", "", resp.content[0].text.strip()).strip()
            s = raw.find("{"); e = raw.rfind("}") + 1
            if s != -1 and e > 0:
                parsed = json.loads(raw[s:e])
                gaps = parsed.get("current_sprint_gaps", [])
                objs = parsed.get("sprint_objectives", [])
                parts = []
                if gaps: parts.append("Sprint gaps: " + "; ".join(gaps))
                if objs: parts.append("Objectives: " + "; ".join(objs))
                sprint_gaps_context = " | ".join(parts) if parts else "Not Available"
        except Exception:
            sprint_gaps_context = "Not Available"

    # ── Step 2: Extract venture feedback from feedback file ──
    feedback_rating = None
    feedback_summary = "Not Available"
    if feedback_text and len(feedback_text.strip()) > 50:
        FB_PROMPT = f"""Venture: {vname}

From the feedback file below, extract the venture's feedback on their advisory sessions.

Return ONLY a JSON object:
{{
  "overall_rating": <number 1-5 or null>,
  "feedback_summary": "concise summary of what the venture said",
  "specific_comments": ["comment 1", "comment 2"],
  "areas_for_improvement": "any suggestions or complaints or null"
}}

No markdown. Null for any missing field.

--- FEEDBACK FILE ---
{feedback_text[:40000]}"""

        try:
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=500,
                messages=[{"role": "user", "content": FB_PROMPT}]
            )
            raw = re.sub(r"```json|```", "", resp.content[0].text.strip()).strip()
            s = raw.find("{"); e = raw.rfind("}") + 1
            if s != -1 and e > 0:
                parsed = json.loads(raw[s:e])
                feedback_rating = parsed.get("overall_rating")
                parts = []
                if parsed.get("feedback_summary") and parsed["feedback_summary"] != "null":
                    parts.append(parsed["feedback_summary"])
                if parsed.get("specific_comments"):
                    parts.append("Comments: " + "; ".join(
                        c for c in parsed["specific_comments"] if c and c != "null"
                    ))
                if parsed.get("areas_for_improvement") and \
                   parsed["areas_for_improvement"] not in [None, "null", "None"]:
                    parts.append("Improvement: " + parsed["areas_for_improvement"])
                feedback_summary = " | ".join(parts) if parts else "Not Available"
        except Exception:
            pass

    # ── Step 3: No transcript → skeleton record ──
    if not transcript_text or len(transcript_text.strip()) < 50:
        return [{
            "venture_name":             vname,
            "expert_name":              "Not Available",
            "expert_role":              "Not Available",
            "session_date":             "Not Available",
            "call_type":                "Not Available",
            "call_objective":           "Not Available",
            "discussion_topics":        [],
            "company_updates":          "Not Available",
            "venture_partner_guidance": "Not Available",
            "key_decisions":            [],
            "action_items":             [],
            "risks_challenges":         [],
            "venture_feedback_rating":  feedback_rating,
            "venture_feedback_text":    feedback_summary,
            "sprint_gaps_context":      sprint_gaps_context,
            "sources_used":             ["Feedback", "Sprint Plan", "Journey Doc"],
        }]

    # ── Step 4: Extract calls from transcript (chunked) ──
    chunks    = chunk_text(transcript_text)
    all_calls = []
    seen_keys = set()

    for i, chunk in enumerate(chunks):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=3000,
                system=[{
                    "type": "text",
                    "text": CALL_INTEL_SYSTEM,
                    "cache_control": {"type": "ephemeral"}
                }],
                messages=[{"role": "user", "content": CALL_INTEL_PROMPT.format(
                    vname=vname, sprint=sprint,
                    n=i+1, total=len(chunks),
                    sprint_gaps=sprint_gaps_context,
                    text=chunk
                )}]
            )
            raw = re.sub(r"```json|```", "", resp.content[0].text.strip()).strip()
            s = raw.find("["); e = raw.rfind("]") + 1
            if s == -1 or e == 0:
                continue
            calls = json.loads(raw[s:e])
            if not isinstance(calls, list):
                continue

            for call in calls:
                expert = call.get("expert_name", "Not Available")
                date   = call.get("session_date", "Not Available")
                dedup_key = f"{expert.lower().strip()}_{date}"
                if dedup_key in seen_keys:
                    continue
                seen_keys.add(dedup_key)

                # Enrich with feedback file data and context
                call["venture_name"]            = vname
                call["venture_feedback_rating"] = call.get("venture_feedback_rating") or feedback_rating
                call["venture_feedback_text"]   = call.get("venture_feedback_text") or feedback_summary
                if call["venture_feedback_text"] in ["Not Available", None, ""]:
                    call["venture_feedback_text"] = feedback_summary
                call["sprint_gaps_context"] = sprint_gaps_context
                call["sources_used"] = ["Transcript", "Feedback", "Sprint Plan", "Journey Doc"]

                # Ensure all 7 fields exist with defaults
                for field in ["call_objective", "company_updates", "venture_partner_guidance"]:
                    if field not in call or not call[field]:
                        call[field] = "Not Available"
                for field in ["discussion_topics", "key_decisions", "action_items", "risks_challenges"]:
                    if field not in call or not isinstance(call[field], list):
                        call[field] = []

                all_calls.append(call)

        except Exception:
            continue

    # Fallback if transcript had content but nothing parsed
    if not all_calls:
        all_calls.append({
            "venture_name":             vname,
            "expert_name":              "Not Available",
            "expert_role":              "Not Available",
            "session_date":             "Not Available",
            "call_type":                "Not Available",
            "call_objective":           "Could not parse call data from transcript",
            "discussion_topics":        [],
            "company_updates":          "Not Available",
            "venture_partner_guidance": "Not Available",
            "key_decisions":            [],
            "action_items":             [],
            "risks_challenges":         [],
            "venture_feedback_rating":  feedback_rating,
            "venture_feedback_text":    feedback_summary,
            "sprint_gaps_context":      sprint_gaps_context,
            "sources_used":             ["Transcript", "Feedback", "Sprint Plan", "Journey Doc"],
        })

    return all_calls
