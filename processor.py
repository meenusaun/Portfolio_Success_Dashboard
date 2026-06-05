"""
Batch processor for Portfolio Success Dashboard
Processes ventures in batches, reads all documents, scores RAG
"""
import json, re, os, tempfile
from pathlib import Path
from anthropic import Anthropic

CHUNK_SIZE = 120000  # safe Claude API limit per call

def chunk_text(text, size=CHUNK_SIZE):
    """Split text at file boundaries."""
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
    """Extract signals from full_text using chunking."""
    chunks  = chunk_text(full_text)
    result  = {"momentum": [], "investment": []}
    seen    = set()

    PROMPT = """Venture: {vname} | Sprint: {sprint} | Chunk {n}/{total}

Extract ALL signals into TWO categories:

SPRINT MOMENTUM SIGNALS: attendance, task completion, orders won, milestones, positive feedback, founder engagement, progress on sprint goals.
SELF INVESTMENT SIGNALS: money spent, staff hired, tools/subscriptions, new markets entered, capital invested, self-funded activities specific to sprint '{sprint}'.

Format EXACTLY (one per line):
MOMENTUM: [signal type] | EVIDENCE: [exact quote] | SOURCE: [doc name]
INVESTMENT: [signal type] | EVIDENCE: [exact quote] | SOURCE: [doc name]

Be THOROUGH. Find every signal including subtle ones.
If none in a category: MOMENTUM: None found

--- DOCUMENTS ---
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
                    t = parts[0].replace("MOMENTUM:","").strip()
                    e = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    dk = f"m_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["momentum"].append({"type":t,"evidence":e,"source":s})
                elif line.startswith("INVESTMENT:") and "None" not in line:
                    parts = [p.strip() for p in line.split("|")]
                    t = parts[0].replace("INVESTMENT:","").strip()
                    e = parts[1].replace("EVIDENCE:","").strip() if len(parts)>1 else ""
                    s = parts[2].replace("SOURCE:","").strip()   if len(parts)>2 else ""
                    dk = f"i_{t}_{e[:40]}"
                    if dk not in seen:
                        seen.add(dk)
                        result["investment"].append({"type":t,"evidence":e,"source":s})
        except Exception as e:
            result["errors"] = result.get("errors",[]) + [f"Chunk {i+1}: {e}"]

    return result, len(chunks)

def score_rag_from_signals(client, vname, sprint, notes, att_summary,
                            signals, pct_raw):
    """Score RAG using full signals extracted from all documents."""
    try:
        pct = float(str(pct_raw).replace("%","").strip())
        if pct <= 1: pct *= 100
    except: pct = 0

    m_sigs = "\n".join(f"- {s['type']}: {s['evidence']}" for s in signals.get("momentum",[]))
    i_sigs = "\n".join(f"- {s['type']}: {s['evidence']}" for s in signals.get("investment",[]))

    prompt = f"""Score this venture's RAG based on ALL signals extracted from their documents.

Venture: {vname} | Sprint: {sprint} | Completion: {pct:.0f}%
Attendance: {att_summary}
Notes: {notes}

MOMENTUM SIGNALS FOUND:
{m_sigs or "None"}

INVESTMENT SIGNALS FOUND:
{i_sigs or "None"}

SCORING RULES:
Sprint Momentum:
- Green: Founder engaged, completing sprint objectives. Strong signals of progress.
- Amber: Some progress but delayed. Inconsistent engagement.
- Red: Founder disengaged, sprint unlikely to complete.
- ZERO: No data at all.

Self Investment (specific to sprint '{sprint}'):
- Green: Already invested sprint-related resources (hired, bought tools, spent money).
- Amber: Plans to invest, showing intent but not yet committed.
- Red: No investment intent, not ready to commit.
- ZERO: No data.

Scoring matrix: Green+Green=10, Green+Amber=8, Amber+Amber=7, Red+anything=1-5, ZERO+ZERO=0

Return ONLY JSON:
{{"momentum_rag":"Green/Amber/Red/ZERO","momentum_reason":"one sentence","investment_rag":"Green/Amber/Red/ZERO","investment_reason":"one sentence","momentum_score":0,"investment_score":0}}"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5", max_tokens=300,
            messages=[{"role":"user","content":prompt}])
        raw  = re.sub(r"```json|```","",resp.content[0].text.strip()).strip()
        data = json.loads(raw)
        m    = data.get("momentum_rag","Unknown")
        i    = data.get("investment_rag","Unknown")
        # Overall = worst of two
        order = {"Red":0,"Amber":1,"Green":2,"ZERO":3,"Unknown":4}
        scores_present = [x for x in [m,i] if x not in ["ZERO","Unknown"]]
        overall = min(scores_present, key=lambda x: order.get(x,4)) if scores_present else "ZERO"
        return {
            "momentum_rag":    m,
            "investment_rag":  i,
            "overall_rag":     overall,
            "momentum_reason": data.get("momentum_reason","—"),
            "investment_reason":data.get("investment_reason","—"),
            "momentum_score":  data.get("momentum_score",0),
            "investment_score":data.get("investment_score",0),
        }
    except Exception as e:
        return {"momentum_rag":"Unknown","investment_rag":"Unknown",
                "overall_rag":"Unknown","momentum_reason":str(e),
                "investment_reason":"—","momentum_score":0,"investment_score":0}

def process_venture(client, vname, venture_data, load_v_files_fn,
                    get_text_fn, extract_common_fn, load_common_fn,
                    get_attendance_fn, attendance_data, notes, sprint, pct_raw):
    """Fully process one venture — read all docs, extract signals, score RAG."""
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

        # 2. Load common docs venture sections
        common_text    = load_common_fn()
        venture_common = extract_common_fn(vname, common_text)

        # 3. Build full combined text
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

        full_text = "\n\n".join(
            f"=== {k} ===\n{v}" for k,v in sources.items() if v
        )
        total_chars  = len(full_text)
        sources_used = [k for k,v in sources.items() if v]

        # 4. Attendance
        att          = get_attendance_fn(vname, attendance_data)
        att_sessions = att["sessions"]     if att else 0
        att_dates    = att["dates"]        if att else []
        att_weeks    = att["weeks_active"] if att else 0
        att_summary  = f"{att_sessions} sessions ({', '.join(att_dates)})" if att_sessions else "No attendance data"

        # 5. Extract signals with chunking
        signals, num_chunks = extract_signals_from_text(client, vname, sprint, full_text)

        # 6. Score RAG from signals
        rag = score_rag_from_signals(client, vname, sprint, notes,
                                      att_summary, signals, pct_raw)

        result.update({
            "status":          "done",
            "signals":         signals,
            "rag":             rag,
            "total_chars":     total_chars,
            "num_chunks":      num_chunks,
            "sources_used":    sources_used,
            "att_sessions":    att_sessions,
            "att_dates":       att_dates,
        })

    except Exception as e:
        result.update({"status":"error","error":str(e),
                       "rag":{"momentum_rag":"Unknown","investment_rag":"Unknown",
                               "overall_rag":"Unknown","momentum_reason":str(e),
                               "investment_reason":"—","momentum_score":0,"investment_score":0},
                       "signals":{"momentum":[],"investment":[]}})
    return result
