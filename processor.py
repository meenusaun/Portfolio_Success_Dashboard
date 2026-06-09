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

    PROMPT = """Venture:{vname}|Sprint:{sprint}|Chunk {n}/{total}

Extract ALL signals below into TWO categories. Assign GREEN, AMBER or RED to each.

SPRINT MOMENTUM SIGNALS — look specifically for:
- Session attendance / meetings attended
- Tasks completed or milestones achieved
- Export orders / deals / contracts won
- Founder engagement (positive or negative)
- Progress or lack of progress toward sprint objectives
- Any evidence sprint is on track or stalled

SELF INVESTMENT SIGNALS — look specifically for (must relate to sprint '{sprint}'):
- Staff hired or not hired for sprint-relevant roles
- Equipment / tools / software purchased or withheld
- Capital invested or refused
- New market entry or channel established
- Any concrete financial commitment or withdrawal

CATEGORY RULES:
GREEN momentum: strongly engaged, tasks done, orders won, attendance confirmed, clear progress.
AMBER momentum: partial progress, mixed engagement, delayed but still active.
RED momentum: disengaged, no progress, missing sessions, exit intent, sprint stalled.
GREEN investment: hired, spent, bought, concretely committed for sprint '{sprint}'.
AMBER investment: intent shown but not yet committed, planning to invest soon.
RED investment: no intent, not ready, unsure, withdrawing resources.

One signal per line, EXACT format:
MOMENTUM:[type]|EVIDENCE:[exact quote]|SOURCE:[doc name]|CATEGORY:GREEN
INVESTMENT:[type]|EVIDENCE:[exact quote]|SOURCE:[doc name]|CATEGORY:AMBER

Be THOROUGH — extract every signal including subtle or indirect ones.
Do NOT skip negative signals. If genuinely none: MOMENTUM:None found

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
        return f"NPS {nps:+d} — {g} Green, {a} Amber, {r} Red of {tot} signals → {rag}."

    m_reason = _reason(m_rag, m_nps, m_g, m_a, m_r, m_tot, "momentum")
    i_reason = _reason(i_rag, i_nps, i_g, i_a, i_r, i_tot, "investment")

    score_matrix = {
        ("Green","Green"):10, ("Green","Amber"):8, ("Green","Red"):5, ("Green","ZERO"):5,
        ("Amber","Green"):8,  ("Amber","Amber"):7, ("Amber","Red"):3, ("Amber","ZERO"):3,
        ("Red",  "Green"):5,  ("Red",  "Amber"):3, ("Red",  "Red"):1, ("Red",  "ZERO"):0,
        ("ZERO", "Green"):5,  ("ZERO", "Amber"):3, ("ZERO", "Red"):0, ("ZERO", "ZERO"):0,
    }
    numeric_score = score_matrix.get((m_rag, i_rag), 0)

    return {
        "momentum_rag":      m_rag,
        "investment_rag":    i_rag,
        "overall_rag":       overall,
        "momentum_reason":   m_reason,
        "investment_reason": i_reason,
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
