from io import BytesIO
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

AUTHOR="Av. Onur Güneş"

def _track_replace(paragraph, old, new, cid):
    full=paragraph.text
    idx=full.find(old)
    if idx<0: return cid, False
    before, after=full[:idx], full[idx+len(old):]
    for r in paragraph.runs: r._element.getparent().remove(r._element)
    if before: paragraph.add_run(before)
    dele=OxmlElement("w:del"); dele.set(qn("w:id"),str(cid)); dele.set(qn("w:author"),AUTHOR)
    dr=OxmlElement("w:r"); dt=OxmlElement("w:delText"); dt.text=old; dr.append(dt); dele.append(dr); paragraph._p.append(dele); cid+=1
    ins=OxmlElement("w:ins"); ins.set(qn("w:id"),str(cid)); ins.set(qn("w:author"),AUTHOR)
    ir=OxmlElement("w:r"); it=OxmlElement("w:t"); it.text=new; ir.append(it); ins.append(ir); paragraph._p.append(ins); cid+=1
    if after: paragraph.add_run(after)
    return cid, True

def _yellow_first_word(paragraph):
    text=paragraph.text.strip()
    if not text: return
    first=text.split()[0]
    for run in paragraph.runs:
        if first in run.text:
            shd=OxmlElement("w:shd"); shd.set(qn("w:fill"),"FFF2CC")
            run._r.get_or_add_rPr().append(shd)
            return

def create_revised_word(raw, revisions):
    doc=Document(BytesIO(raw)); cid=1
    attention=[]
    for rev in revisions:
        excerpt=(rev.get("current_excerpt") or "").strip()
        suggestion=(rev.get("suggested_revision") or "").strip()
        matched=False
        if excerpt and suggestion and rev.get("revision_type")!="EK HÜKÜM":
            for p in doc.paragraphs:
                if excerpt in p.text:
                    cid, matched=_track_replace(p,excerpt,suggestion,cid)
                    _yellow_first_word(p)
                    break
        if not matched:
            # Unknown/new/blank/manual-attention items: mark only first word of the referenced paragraph.
            ref=(rev.get("reference") or "").strip()
            for p in doc.paragraphs:
                if ref and p.text.strip().startswith(ref):
                    _yellow_first_word(p); matched=True; break
        if not matched: attention.append(rev)
    bio=BytesIO(); doc.save(bio)
    return bio.getvalue(), attention
