from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower()


def _set_space_preserve(node):
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _first_run_properties(paragraph):
    """Copy the first visible run formatting so inserted text inherits font/size/etc."""
    for r in paragraph._p.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is not None:
            return deepcopy(rPr)
    # Also inspect tracked runs if the source already contains redlines.
    for change_tag in ("w:ins", "w:del"):
        for change in paragraph._p.findall(qn(change_tag)):
            for r in change.findall(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    return deepcopy(rPr)
    return None


def _tracked_run(tag: str, text: str, change_id: int, author: str, rPr=None):
    change = OxmlElement(tag)
    change.set(qn("w:id"), str(change_id))
    change.set(qn("w:author"), author)
    change.set(qn("w:date"), datetime.now(timezone.utc).isoformat())

    r = OxmlElement("w:r")
    if rPr is not None:
        r.append(deepcopy(rPr))
    t_tag = "w:delText" if tag == "w:del" else "w:t"
    t = OxmlElement(t_tag)
    _set_space_preserve(t)
    t.text = text
    r.append(t)
    change.append(r)
    return change


def _clear_paragraph_content(paragraph):
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def replace_paragraph_tracked(paragraph, new_text: str, change_id: int, author: str):
    old_text = paragraph.text
    rPr = _first_run_properties(paragraph)
    _clear_paragraph_content(paragraph)
    p = paragraph._p
    if old_text:
        p.append(_tracked_run("w:del", old_text, change_id, author, rPr))
        change_id += 1
    p.append(_tracked_run("w:ins", new_text, change_id, author, rPr))
    return change_id + 1


def insert_after_tracked(paragraph, new_text: str, change_id: int, author: str):
    new_p = OxmlElement("w:p")
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        new_p.append(deepcopy(pPr))
    rPr = _first_run_properties(paragraph)
    new_p.append(_tracked_run("w:ins", new_text, change_id, author, rPr))
    paragraph._p.addnext(new_p)
    return change_id + 1


def append_end_tracked(doc, new_text: str, change_id: int, author: str):
    # Last meaningful paragraph is the formatting donor.
    donor = next((p for p in reversed(doc.paragraphs) if p.text.strip()), None)
    p = doc.add_paragraph()
    if donor is not None:
        pPr = donor._p.find(qn("w:pPr"))
        if pPr is not None:
            p._p.insert(0, deepcopy(pPr))
        rPr = _first_run_properties(donor)
    else:
        rPr = None
    _clear_paragraph_content(p)
    p._p.append(_tracked_run("w:ins", new_text, change_id, author, rPr))
    return change_id + 1


def _all_paragraphs(doc):
    items = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                items.extend(cell.paragraphs)
    return items


def find_best_paragraph(doc, anchor_text: str):
    anchor = _norm(anchor_text)
    if not anchor:
        return None, 0.0

    paragraphs = _all_paragraphs(doc)
    for p in paragraphs:
        txt = _norm(p.text)
        if anchor in txt or (txt and txt in anchor and len(txt) > 25):
            return p, 1.0

    best_p, best_score = None, 0.0
    for p in paragraphs:
        txt = _norm(p.text)
        if not txt:
            continue
        score = SequenceMatcher(None, anchor, txt).ratio()
        if score > best_score:
            best_p, best_score = p, score
    return best_p, best_score


def apply_revisions_to_docx(
    original_bytes: bytes,
    revisions: list[dict],
    author: str = "Av. Onur Güneş",
):
    doc = Document(BytesIO(original_bytes))
    change_id = 1000
    applied, skipped = [], []

    for rev in revisions:
        action = (rev.get("action") or "REPLACE_PARAGRAPH").upper()
        new_text = (rev.get("replacement_text") or "").strip()
        anchor = (rev.get("anchor_text") or "").strip()

        if not new_text:
            skipped.append({**rev, "reason": "Boş revizyon metni"})
            continue

        if action == "APPEND_END":
            change_id = append_end_tracked(doc, new_text, change_id, author)
            applied.append({**rev, "match_score": 1.0})
            continue

        paragraph, score = find_best_paragraph(doc, anchor)
        if paragraph is None or score < 0.42:
            skipped.append({**rev, "reason": f"Anchor bulunamadı (eşleşme {score:.2f})"})
            continue

        if action == "APPEND_AFTER":
            change_id = insert_after_tracked(paragraph, new_text, change_id, author)
        else:
            change_id = replace_paragraph_tracked(paragraph, new_text, change_id, author)

        applied.append({**rev, "match_score": score})

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue(), applied, skipped
