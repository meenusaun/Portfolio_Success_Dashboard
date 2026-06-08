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
    Each chunk is processed independently — results merged and deduplicated.
    No character limit — processes documents of any size.
    """
    chunks = chunk_text(full_text)
    result = {"momentum": [], "investment": []}
    seen   = set()

    PROMPT = """Venture: {vname} | Sprint Type: {sprint} | Chunk {n} of {total}

Extract ALL signals from the text below into TWO categories:

SPRINT MOMENTUM SIGNALS — evidence of founder engagement and sprint progress:
- Session attendance / meetings attended
- Tasks completed or milestones achieved  
- Export orders / deals / contracts won
- Positive founder engagement or feedback
- Progress toward sprint objectives
- Any evidence the sprint is on track

SELF INVESTMENT SIGNALS — evidence of money/resources committed, SPECIFIC to sprint type '{sprint}':
- Staff hired (especially sprint-relevant roles)
- Equipment / tools / software purchased
- Capital invested in sprint-related activities
- Self-funded sprint continuation
- New market entry or channel established
- Any concrete financial commitment to growth

Format EXACTLY (one signal per line):
MOMENTUM: [signal type] | EVIDENCE: [exact quote from text] | SOURCE: [document name]
INVESTMENT: [signal type] | EVIDENCE: [exact quote from text] | SOURCE: [document name]

IMPORTANT: Be THOROUGH. Find EVERY signal, including subtle ones.
Do NOT miss signals just because they are indirect or implied.
If genuinely none in a category: MOMENTUM: None found

--- DOCUMENTS (Chunk {n}/{total}) ---
{text}"""

    for i, chunk in enumerate(chunks):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-5", max_tokens=2000,
                messages=[{"role":"user","content":
                    PROMPT.format(vname=vname, sprint=sprint,
                                  n=i+1, total=len(chunks), text=chunk)}])
            for line in resp.content[0].text.splitlines():
                line = line.strip()
                if line.startswith("MOMENTUM:") and "None" not in line:
                    parts = [p.strip() for p in line.split("|")]
                    t  = parts[0].replace("MOMENTUM:","").strip()
                    e  = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s  = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    dk = f"m_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["momentum"].append({"type":t,"evidence":e,"source":s})
                elif line.startswith("INVESTMENT:") and "None" not in line:
                    parts = [p.strip() for p in line.split("|")]
                    t  = parts[0].replace("INVESTMENT:","").strip()
                    e  = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s  = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    dk = f"i_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["investment"].append({"type":t,"evidence":e,"source":s})
        except Exception as e:
            result.setdefault("errors",[]).append(f"Chunk {i+1}: {e}")

    return result, len(chunks)


def score_rag_from_signals(client, vname, sprint, notes, att_summary,
                            signals, pct_raw):
    """
    Score RAG using COMPLETE signals extracted from ALL documents.
    This is always in sync with signals — same data, same result.
    """
    try:
        pct = float(str(pct_raw).replace("%","").strip())
        if pct <= 1: pct *= 100
    except: pct = 0

    m_count = len(signals.get("momentum",[]))
    i_count = len(signals.get("investment",[]))
    m_sigs  = "\n".join(f"  - {s['type']}: {s['evidence']}"
                        for s in signals.get("momentum",[]))
    i_sigs  = "\n".join(f"  - {s['type']}: {s['evidence']}"
                        for s in signals.get("investment",[]))

    prompt = f"""You are scoring a venture in an accelerator program.
Score based ONLY on the signals extracted from ALL their documents below.

Venture: {vname} | Sprint: {sprint} | Completion: {pct:.0f}%
Attendance: {att_summary}
Notes: {notes}

MOMENTUM SIGNALS FOUND ({m_count} total):
{m_sigs or "  None found"}

INVESTMENT SIGNALS FOUND ({i_count} total):
{i_sigs or "  None found"}

SCORING RULES — read carefully:

Sprint Momentum RAG:
- Green: Founder ENGAGED, completing sprint objectives. Multiple positive signals.
  Strong evidence = orders won, tasks done, active engagement, positive progress.
- Amber: Some progress but DELAYED or INCONSISTENT. Mixed signals.
- Red: Founder DISENGAGED, sprint unlikely to reach objective. No progress signals.
- ZERO: Genuinely no data from any source.

CRITICAL: If there are MULTIPLE strong momentum signals (orders, tasks, engagement),
score MUST be Green. Do not score Amber if strong positive evidence exists.

Self Investment RAG (must be specific to sprint '{sprint}'):
- Green: Founder HAS ALREADY invested sprint-relevant resources.
  Strong evidence = hired relevant staff, bought tools, spent money on sprint goals.
- Amber: Showing INTENT to invest but not yet committed.
- Red: No investment intent, not ready to commit.
- ZERO: No data.

CRITICAL: If investment signals clearly show money spent or staff hired for
sprint-related goals, score MUST be Green. Do not downgrade without reason.

Scoring matrix:
Green+Green=10, Green+Amber=8, Green+Red=5, Green+ZERO=5
Amber+Green=8, Amber+Amber=7, Amber+Red=3, Amber+ZERO=3
Red+anything low=1-3, ZERO+ZERO=0

Return ONLY this JSON:
{{"momentum_rag":"Green/Amber/Red/ZERO",
  "momentum_reason":"one sentence citing specific signals",
  "investment_rag":"Green/Amber/Red/ZERO",
  "investment_reason":"one sentence citing specific sprint-relevant signals",
  "momentum_score":0,
  "investment_score":0}}"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5", max_tokens=400,
            messages=[{"role":"user","content":prompt}])
        raw  = re.sub(r"```json|```","",resp.content[0].text.strip()).strip()
        data = json.loads(raw)
        m    = data.get("momentum_rag","Unknown")
        i    = data.get("investment_rag","Unknown")
        order = {"Red":0,"Amber":1,"Green":2,"ZERO":3,"Unknown":4}
        present = [x for x in [m,i] if x not in ["ZERO","Unknown"]]
        overall = min(present, key=lambda x: order.get(x,4)) if present else "ZERO"
        return {
            "momentum_rag":     m,
            "investment_rag":   i,
            "overall_rag":      overall,
            "momentum_reason":  data.get("momentum_reason","—"),
            "investment_reason":data.get("investment_reason","—"),
            "momentum_score":   data.get("momentum_score",0),
            "investment_score": data.get("investment_score",0),
        }
    except Exception as e:
        return {"momentum_rag":"Unknown","investment_rag":"Unknown",
                "overall_rag":"Unknown","momentum_reason":str(e),
                "investment_reason":"—","momentum_score":0,"investment_score":0}


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
