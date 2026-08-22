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


def _tracked_run(tag: str, text: str, change_id: int, author: str):
    change = OxmlElement(tag)
    change.set(qn("w:id"), str(change_id))
    change.set(qn("w:author"), author)
    change.set(qn("w:date"), datetime.now(timezone.utc).isoformat())

    r = OxmlElement("w:r")
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
    _clear_paragraph_content(paragraph)
    p = paragraph._p
    if old_text:
        p.append(_tracked_run("w:del", old_text, change_id, author))
        change_id += 1
    p.append(_tracked_run("w:ins", new_text, change_id, author))
    return change_id + 1


def insert_after_tracked(paragraph, new_text: str, change_id: int, author: str):
    new_p = OxmlElement("w:p")
    # Preserve paragraph style/numbering where sensible.
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        new_p.append(deepcopy(pPr))
    new_p.append(_tracked_run("w:ins", new_text, change_id, author))
    paragraph._p.addnext(new_p)
    return change_id + 1


def append_end_tracked(doc, new_text: str, change_id: int, author: str):
    p = doc.add_paragraph()
    _clear_paragraph_content(p)
    p._p.append(_tracked_run("w:ins", new_text, change_id, author))
    return change_id + 1


def _all_paragraphs(doc):
    # Main document paragraphs plus table-cell paragraphs.
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

    # Exact/contains first.
    for p in paragraphs:
        txt = _norm(p.text)
        if anchor in txt or (txt and txt in anchor and len(txt) > 25):
            return p, 1.0

    # Fuzzy fallback.
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
    author: str = "Opus Legal Intelligence",
):
    doc = Document(BytesIO(original_bytes))
    change_id = 1000
    applied = []
    skipped = []

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
    bio.seek(0)
    return bio.getvalue(), applied, skipped
