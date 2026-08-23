import html
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


def _doc_signature(raw):
    pars=read_docx_paragraphs(raw)
    head=" ".join(pars[:12])
    body=" ".join(pars[:80])
    return head, body

def match_batch_documents(raw_files, revised_files):
    """
    Match raw/revised DOCX files using filename + document-text similarity.
    Returns confident pairs and uncertain pairs for manual review.
    """
    raw_meta=[]
    rev_meta=[]
    for name,data in raw_files:
        h,b=_doc_signature(data); raw_meta.append({"name":name,"data":data,"head":h,"body":b})
    for name,data in revised_files:
        h,b=_doc_signature(data); rev_meta.append({"name":name,"data":data,"head":h,"body":b})

    candidates=[]
    for i,r in enumerate(raw_meta):
        for j,v in enumerate(rev_meta):
            fn=_similarity(Path(r["name"]).stem,Path(v["name"]).stem)
            head=_similarity(r["head"],v["head"])
            body=_similarity(r["body"][:12000],v["body"][:12000])
            score=0.20*fn+0.35*head+0.45*body
            candidates.append((score,i,j))
    candidates.sort(reverse=True)
    used_r=set(); used_v=set(); pairs=[]
    for score,i,j in candidates:
        if i in used_r or j in used_v: continue
        used_r.add(i); used_v.add(j)
        pairs.append({
            "raw_name":raw_meta[i]["name"],"raw_data":raw_meta[i]["data"],
            "revised_name":rev_meta[j]["name"],"revised_data":rev_meta[j]["data"],
            "match_score":round(score,3),
            "status":"CONFIDENT" if score>=0.62 else "REVIEW"
        })
    unmatched_raw=[x["name"] for i,x in enumerate(raw_meta) if i not in used_r]
    unmatched_revised=[x["name"] for j,x in enumerate(rev_meta) if j not in used_v]
    return pairs,unmatched_raw,unmatched_revised

def batch_candidates(pairs):
    allc=[]
    for p in pairs:
        cs=build_candidates(p["raw_data"],p["revised_data"],source_name=p["revised_name"])
        cs=filter_meaningful_candidates(cs)
        for c in cs:
            c["pair_raw_name"]=p["raw_name"]
            c["pair_revised_name"]=p["revised_name"]
            c["pair_match_score"]=p["match_score"]
        allc.extend(cs)
    return allc

def cluster_candidates(candidates):
    """
    Lightweight clusters by inferred legal topic keywords + edit style.
    AI-free first pass, intended for review/approval UI.
    """
    topics=[
        ("Münhasırlık",["münhasır","başka proje","başka bir proje","rakip"]),
        ("Ücret / Ödeme",["ücret","ödeme","kdv","fatura","serbest meslek","bölüm başı"]),
        ("Cezai Şart",["cezai şart","ceza koşulu"]),
        ("Opsiyon / Sezon",["opsiyon","sezon","yeni sezon","devam sezon"]),
        ("Tanıtım / PR",["tanıtım","pr ","röportaj","gala","basın"]),
        ("Sosyal Medya",["sosyal medya","instagram","paylaşım"]),
        ("Ürün Yerleştirme",["ürün yerleştirme","ürün kullan","marka"]),
        ("Yapay Zeka / Dijital Kullanım",["yapay zeka","machine learning","data mining","sentetik","dijital kopya"]),
        ("Fikri Haklar",["mali hak","fsek","telif","icracı","bağlantılı hak"]),
        ("Fesih",["fesih","sona er","sözleşmeyi sona"]),
        ("Çalışma / Set",["çekim","set ","çalışma saati","iş günü","takvim"]),
    ]
    groups={}
    for c in candidates:
        hay=(" "+c.get("original_text","")+" "+c.get("final_text","")+" ").lower()
        topic="Diğer"
        for label,keys in topics:
            if any(k in hay for k in keys):
                topic=label; break
        key=(topic,c.get("edit_style","UNKNOWN"))
        groups.setdefault(key,[]).append(c)
    out=[]
    for (topic,style),items in groups.items():
        out.append({
            "cluster_id":f"{topic}|{style}",
            "topic":topic,"edit_style":style,"count":len(items),
            "confidence":round(min(0.95,0.40+0.08*len(items)),2),
            "examples":items[:5],
            "candidate_ids":[x["candidate_id"] for x in items]
        })
    return sorted(out,key=lambda x:(-x["count"],x["topic"]))


def semantic_diff(original, final):
    def clean(s):
        s=(s or "").replace("\u00a0"," ")
        return re.sub(r"\s+"," ",s).strip()
    def norm_semantic(s):
        s=clean(s).lower()
        s=re.sub(r"[“”\"'’`´]", "", s)
        s=re.sub(r"\s*([,;:.!?()\-/])\s*", r"\1", s)
        return s
    o=clean(original); f=clean(final)
    if norm_semantic(o)==norm_semantic(f):
        return {"meaningful":False,"changes":[],"old_html":html.escape(o),"new_html":html.escape(f)}
    tok_re=re.compile(r"\w+|[^\w\s]|\s+", re.UNICODE)
    ot=tok_re.findall(o); ft=tok_re.findall(f)
    sm=SequenceMatcher(None,ot,ft,autojunk=False)
    changes=[]; old_parts=[]; new_parts=[]
    for tag,i1,i2,j1,j2 in sm.get_opcodes():
        os="".join(ot[i1:i2]); ns="".join(ft[j1:j2])
        if tag=="equal":
            old_parts.append(html.escape(os)); new_parts.append(html.escape(ns))
        else:
            if norm_semantic(os)==norm_semantic(ns):
                old_parts.append(html.escape(os)); new_parts.append(html.escape(ns)); continue
            changes.append({"type":tag,"old":os,"new":ns})
            if os: old_parts.append(f'<span class="oli-del">{html.escape(os)}</span>')
            if ns: new_parts.append(f'<span class="oli-ins">{html.escape(ns)}</span>')
    return {"meaningful":bool(changes),"changes":changes,"old_html":"".join(old_parts),"new_html":"".join(new_parts)}

def compact_change_summary(original, final, max_items=4):
    d=semantic_diff(original,final)
    if not d["meaningful"]: return ""
    parts=[]
    for ch in d["changes"][:max_items]:
        old=re.sub(r"\s+"," ",ch.get("old","")).strip()
        new=re.sub(r"\s+"," ",ch.get("new","")).strip()
        if old and new: parts.append(f'"{old}" → "{new}"')
        elif old: parts.append(f'Silindi: "{old}"')
        elif new: parts.append(f'Eklendi: "{new}"')
    return " • ".join(parts)

def filter_meaningful_candidates(candidates):
    out=[]
    for c in candidates:
        if c.get("kind")=="ADDITION":
            if len(_norm(c.get("final_text","")))>=35: out.append(c)
            continue
        if semantic_diff(c.get("original_text",""),c.get("final_text",""))["meaningful"]:
            cc=dict(c)
            cc["diff_summary"]=compact_change_summary(cc.get("original_text",""),cc.get("final_text",""))
            out.append(cc)
    return out
