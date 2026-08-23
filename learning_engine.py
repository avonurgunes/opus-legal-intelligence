import json, os, re, hashlib
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from difflib import SequenceMatcher
from io import BytesIO
from docx import Document

LEARNING_FILE = Path(os.getenv("OLI_LEARNING_FILE", "learning_memory.json"))

def _now():
    return datetime.now(ZoneInfo("Europe/Istanbul")).isoformat()

def _norm(s):
    return re.sub(r"\s+"," ",(s or "")).strip()

def _similarity(a,b):
    return SequenceMatcher(None,_norm(a).lower(),_norm(b).lower(),autojunk=False).ratio()

def load_memory():
    if not LEARNING_FILE.exists():
        return {"version":1,"records":[]}
    try:
        data=json.loads(LEARNING_FILE.read_text(encoding="utf-8"))
        if not isinstance(data,dict): raise ValueError()
        data.setdefault("version",1); data.setdefault("records",[])
        return data
    except Exception:
        return {"version":1,"records":[]}

def save_memory(data):
    LEARNING_FILE.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")

def read_docx_paragraphs(raw):
    doc=Document(BytesIO(raw))
    return [_norm(p.text) for p in doc.paragraphs if _norm(p.text)]

def pair_paragraphs(original_raw, final_raw):
    """Conservative paragraph alignment. Returns changed pairs and additions."""
    old=read_docx_paragraphs(original_raw)
    new=read_docx_paragraphs(final_raw)
    used=set(); pairs=[]
    for i,o in enumerate(old):
        best_j=None; best=0.0
        for j,n in enumerate(new):
            if j in used: continue
            s=_similarity(o,n)
            if s>best:
                best=s; best_j=j
        if best_j is not None and best>=0.38:
            used.add(best_j)
            n=new[best_j]
            if _norm(o)!=_norm(n):
                pairs.append({"original":o,"final":n,"similarity":round(best,3),"kind":"CHANGE"})
    for j,n in enumerate(new):
        if j not in used:
            # Only substantial additions become candidates.
            if len(n)>=35:
                pairs.append({"original":"","final":n,"similarity":0.0,"kind":"ADDITION"})
    return pairs

def infer_edit_style(original, final):
    if not original: return "NEW_CLAUSE"
    ratio=1-_similarity(original,final)
    if ratio<=0.20: return "MICRO"
    if ratio<=0.58: return "PHRASE"
    return "BLOCK"

def make_candidate(pair, source_name=""):
    original=pair.get("original",""); final=pair.get("final","")
    digest=hashlib.sha1((original+"\n"+final).encode("utf-8")).hexdigest()[:14]
    return {
        "candidate_id":digest,
        "source_name":source_name,
        "topic":"SINIFLANDIRILMADI",
        "original_text":original,
        "final_text":final,
        "edit_style":infer_edit_style(original,final),
        "kind":pair.get("kind","CHANGE"),
        "similarity":pair.get("similarity",0),
        "scope":"REVIEW_REQUIRED",
        "confidence":0.35,
        "approved":False,
        "created_at":_now()
    }

def build_candidates(original_raw, final_raw, source_name=""):
    return [make_candidate(x,source_name) for x in pair_paragraphs(original_raw,final_raw)]

def approve_candidates(candidates, decisions):
    """
    decisions: candidate_id -> {approve, topic, scope}
    Only explicit approvals persist.
    """
    mem=load_memory()
    existing={r.get("candidate_id"):r for r in mem["records"]}
    added=0
    for c in candidates:
        d=decisions.get(c["candidate_id"],{})
        if not d.get("approve"): continue
        rec=dict(c)
        rec["topic"]=(d.get("topic") or rec["topic"]).strip()
        rec["scope"]=d.get("scope") or "GENERAL"
        rec["approved"]=True
        rec["approved_at"]=_now()
        # Similar approved precedents increase confidence, but never change Rule Library.
        same=[r for r in mem["records"] if r.get("approved") and r.get("topic")==rec["topic"]]
        rec["precedent_count"]=len(same)+1
        rec["confidence"]=round(min(0.95,0.45+0.10*len(same)),2)
        existing[rec["candidate_id"]]=rec
        added+=1
    mem["records"]=list(existing.values())
    save_memory(mem)
    return added, mem

def relevant_learnings(contract_text, limit=12):
    mem=load_memory()
    approved=[r for r in mem["records"] if r.get("approved") and r.get("scope")!="FILE_ONLY"]
    scored=[]
    for r in approved:
        probes=[r.get("original_text",""),r.get("final_text",""),r.get("topic","")]
        score=max((_similarity(contract_text,p) for p in probes if p),default=0)
        score=max(score,float(r.get("confidence",0))*0.55)
        if score>=0.25: scored.append((score,r))
    scored.sort(key=lambda x:x[0],reverse=True)
    return [r for _,r in scored[:limit]]

def learning_prompt_block(contract_text):
    rows=relevant_learnings(contract_text)
    if not rows: return "Henüz onaylanmış geçmiş drafting öğrenimi yok."
    parts=[]
    for r in rows:
        parts.append(
            f"- Konu: {r.get('topic')} | Stil: {r.get('edit_style')} | Güven: {r.get('confidence')}\n"
            f"  Önce: {r.get('original_text')}\n"
            f"  Nihai tercih: {r.get('final_text')}"
        )
    return "\n".join(parts)
