import streamlit as st
import json
import re
from io import BytesIO
from pathlib import Path

from docx import Document
from pypdf import PdfReader
from openai import OpenAI

from revision_engine import apply_revisions_to_docx

st.set_page_config(
    page_title="Opus Legal Intelligence",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

RULES = json.loads(Path(__file__).with_name("rules.json").read_text(encoding="utf-8"))
REVISION_LIBRARY = json.loads(
    Path(__file__).with_name("revision_library.json").read_text(encoding="utf-8")
)

POWER_GUIDANCE = {
    "Düşük": "Sadece kritik konuları masaya taşı; ikincil risklerde pazarlık sermayesini koru.",
    "Orta": "Kritik riskleri ve önemli ticari risklerin çoğunu müzakere et.",
    "Yüksek": "Opus standartlarına geniş ölçüde yaklaş; yüksek ve orta risklerde güçlü revizyon iste.",
    "Çok Yüksek": "İdeal Opus pozisyonuna mümkün olduğunca yaklaş; gereksiz tek taraflı hükümleri kabul etme.",
}

REV_MODE_BY_POWER = {
    "Düşük": "MINIMUM",
    "Orta": "STANDARD",
    "Yüksek": "STANDARD",
    "Çok Yüksek": "STRONG",
}

CSS = """
<style>
:root{
  --opus-black:#111214;
  --opus-panel:#1a1b1f;
  --opus-gold:#b89550;
  --opus-gold2:#d2b36f;
  --opus-cream:#f4efe5;
}
.stApp{
  background:
    radial-gradient(circle at 10% 0%, rgba(184,149,80,.10), transparent 28%),
    linear-gradient(180deg,#0f1012 0%,#15161a 100%);
}
[data-testid="stHeader"]{background:rgba(15,16,18,.75);}
.block-container{max-width:1420px;padding-top:2rem;}
h1,h2,h3{letter-spacing:-.02em}
h1{color:var(--opus-cream)!important;font-weight:800!important}
h2,h3{color:#f3f1ec!important}
p, label, .stMarkdown, [data-testid="stCaptionContainer"]{color:#d9d5cc!important}
.opus-kicker{color:var(--opus-gold2);letter-spacing:.16em;font-size:.78rem;font-weight:700}
.opus-hero{
  padding:1.7rem 1.9rem;border:1px solid rgba(184,149,80,.35);
  background:linear-gradient(135deg,rgba(184,149,80,.13),rgba(255,255,255,.025));
  border-radius:18px;margin-bottom:1.2rem
}
.opus-module{
  min-height:125px;padding:1.1rem;border-radius:16px;
  border:1px solid rgba(184,149,80,.20);background:rgba(255,255,255,.035)
}
.opus-module.active{border-color:rgba(184,149,80,.65);background:rgba(184,149,80,.09)}
.opus-module-title{font-size:1.12rem;color:#fff;font-weight:750;margin-bottom:.4rem}
.opus-module-sub{font-size:.87rem;color:#aaa59b}
div[data-testid="stMetric"]{
  border:1px solid rgba(184,149,80,.22);border-radius:14px;padding:.8rem;
  background:rgba(255,255,255,.025)
}
div.stButton > button, div.stDownloadButton > button{
  border-radius:10px;border:1px solid rgba(184,149,80,.6);
}
div.stButton > button[kind="primary"]{
  background:linear-gradient(90deg,#a98543,#c5a563);color:#0f1012;border:none;font-weight:800
}
[data-testid="stFileUploader"]{border-radius:14px}
details{border:1px solid rgba(184,149,80,.15)!important;border-radius:12px!important}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def read_docx(file_bytes: bytes) -> str:
    doc = Document(BytesIO(file_bytes))
    parts = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def read_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(BytesIO(file_bytes))
    parts = []
    for i, page in enumerate(reader.pages, 1):
        txt = page.extract_text() or ""
        if txt.strip():
            parts.append(f"\n--- SAYFA {i} ---\n{txt}")
    return "\n".join(parts)


def extract_text(uploaded) -> str:
    data = uploaded.getvalue()
    name = uploaded.name.lower()
    if name.endswith(".docx"):
        return read_docx(data)
    if name.endswith(".pdf"):
        return read_pdf(data)
    raise ValueError("Desteklenmeyen dosya türü.")


def get_api_key():
    try:
        return st.secrets["OPENAI_API_KEY"]
    except Exception:
        return None


def get_model():
    default = "gpt-5.6-terra"
    try:
        return st.secrets.get("OPENAI_MODEL", default)
    except Exception:
        return default


def clean_json(text: str):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
    text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start:end+1]
    return json.loads(text)


def analyse_contract(contract_text: str, negotiation_power: str):
    rules_text = "\n".join(
        f"{r['id']} | {r['title']} | Öncelik: {r['priority']} | Opus standardı: {r['standard']}"
        for r in RULES
    )

    system = """Sen Opus Legal Intelligence (OLI) sözleşme analiz motorusun.
Ana akım TV dizisi oyuncu sözleşmesini yalnız verilen Opus kural setine göre incele.

İLKELER:
- Müvekkil OYUNCU/AJANS.
- Avantajlı hüküm GREEN/KORU olabilir.
- Hüküm yoksa NOT_FOUND; fakat kural koruyucu bir hükmün sözleşmede bulunmasını gerektiriyorsa assessment içinde bunun eksiklik olduğunu açıkla.
- 'Geçmişte revize edilmemiş = kabul edilmiş' varsayımı yapma.
- Pazarlık gücü hukuki riski değil müzakere önceliğini etkiler.
- Metinde desteklenmeyen bilgi üretme.
- Yanıt yalnız geçerli JSON.

ŞEMA:
{
 "overall_risk":"LOW|MEDIUM|HIGH|VERY_HIGH",
 "executive_summary":"kısa özet",
 "top_negotiation_points":["en fazla 5"],
 "findings":[{
   "rule_id":"OLI-TV-001",
   "title":"...",
   "status":"RED|ORANGE|YELLOW|GREEN|NOT_FOUND",
   "contract_reference":"madde no/bölüm",
   "clause_excerpt":"kısa hüküm özeti",
   "assessment":"neden",
   "recommended_revision":"kısa prensip",
   "negotiation_priority":"MUST|SHOULD|OPTIONAL|KEEP",
   "confidence":"HIGH|MEDIUM|LOW"
 }]
}
Her 30 kural için findings üret.
"""

    user = f"""PROFİL: ACTOR_TV_MAINSTREAM
PAZARLIK GÜCÜ: {negotiation_power}
STRATEJİ: {POWER_GUIDANCE[negotiation_power]}

RULE LIBRARY:
{rules_text}

SÖZLEŞME:
{contract_text[:140000]}
"""

    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return clean_json(resp.output_text)


def revision_context_for(rule_id: str):
    for entry in REVISION_LIBRARY.get("entries", []):
        if entry.get("rule_id") == rule_id:
            return entry
    return None


def build_revision_drafts(contract_text, selected_findings, negotiation_power):
    lib_parts = []
    for f in selected_findings:
        entry = revision_context_for(f.get("rule_id"))
        if entry:
            lib_parts.append(json.dumps(entry, ensure_ascii=False))

    system = """Sen OLI Revision Engine'sin.
Görevin sözleşmedeki seçilmiş bulgular için OPUS tarzında uygulanabilir revizyon metni üretmektir.

ÖNCELİK:
1) Varsa REVISION LIBRARY'deki doğrulanmış Opus kalıbı.
2) Rule Library'deki standardın özü.
3) Mevcut sözleşmenin terminolojisi, taraf tanımları, proje adı ve madde dili.

KURALLAR:
- Yeni hukuki politika icat etme.
- Geçmiş Opus metnini körü körüne kopyalama; mevcut sözleşmeye uyarla.
- Mevcut hüküm varsa action=REPLACE_PARAGRAPH ve anchor_text sözleşmeden birebir kısa bir parça olsun.
- Koruyucu hüküm tamamen eksikse NOT_FOUND diye bırakma: mutlaka yeni hüküm üret. Sözleşmede konu bakımından en uygun mevcut maddeyi anchor seç ve APPEND_AFTER kullan. APPEND_END yalnız gerçekten uygun bölüm bulunamıyorsa son çaredir.
- Eksik hükmün ekleneceği yeri konu bütünlüğüne göre seç: ücret/ödeme hükümleri ücret bölümüne; mali hak/telif hükümleri mali haklar bölümüne; fesih/cezai şart hükümleri ilgili fesih/ceza bölümüne; vergi/damga vergisi diğer hükümler/vergi bölümüne.
- replacement_text sadece sözleşmeye girecek nihai madde/paragraf metni olsun.
- Pazarlık gücü sadece sertlik düzeyini belirler.
- Yanıt yalnız geçerli JSON.

ŞEMA:
{
 "revisions":[{
   "rule_id":"...",
   "title":"...",
   "action":"REPLACE_PARAGRAPH|APPEND_AFTER|APPEND_END",
   "anchor_text":"sözleşmeden birebir 15-120 karakter veya boş",
   "replacement_text":"Word'e işlenecek revize hüküm",
   "reason":"kısa açıklama",
   "confidence":"HIGH|MEDIUM|LOW"
 }]
}
"""

    mode = REV_MODE_BY_POWER[negotiation_power]
    user = f"""PAZARLIK GÜCÜ: {negotiation_power}
REVİZYON MODU: {mode}

SEÇİLMİŞ BULGULAR:
{json.dumps(selected_findings, ensure_ascii=False, indent=2)}

REVISION LIBRARY:
{chr(10).join(lib_parts) if lib_parts else "Bu kurallar için özel geçmiş kalıp yok."}

SÖZLEŞME:
{contract_text[:140000]}
"""

    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return clean_json(resp.output_text)


def icon(status):
    return {
        "RED": "🔴", "ORANGE": "🟠", "YELLOW": "🟡",
        "GREEN": "🟢", "NOT_FOUND": "⚪"
    }.get(status, "⚪")


st.markdown("""
<div class="opus-hero">
  <div class="opus-kicker">OPUS • PRIVATE COUNSEL SYSTEM</div>
  <h1 style="margin:.2rem 0 .35rem 0">OPUS LEGAL INTELLIGENCE</h1>
  <div style="color:#aaa59b">Contract intelligence • Negotiation strategy • Revision memory • Word redline</div>
</div>
""", unsafe_allow_html=True)

cols = st.columns(4)
module_data = [
    ("Sözleşmeler", "Analyse • Negotiate • Revise", True),
    ("Arabuluculuk", "Prepare • Generate", False),
    ("KVKK", "Privacy • Compliance", False),
    ("Dava Dosyaları", "Litigation Workspace", False),
]
for col, (title, sub, active) in zip(cols, module_data):
    with col:
        cls = "opus-module active" if active else "opus-module"
        st.markdown(
            f'<div class="{cls}"><div class="opus-module-title">{title}</div>'
            f'<div class="opus-module-sub">{sub if active else "Yakında • " + sub}</div></div>',
            unsafe_allow_html=True
        )

st.divider()
st.header("Yeni Sözleşme Analizi")

c1,c2,c3 = st.columns(3)
with c1:
    contract_type = st.selectbox("Sözleşme türü", ["Oyuncu Sözleşmesi"])
with c2:
    project_type = st.selectbox("Proje türü", ["Ana Akım TV", "Dijital", "Sinema"])
with c3:
    negotiation_power = st.select_slider(
        "Pazarlık gücü", ["Düşük","Orta","Yüksek","Çok Yüksek"], value="Orta"
    )

uploaded = st.file_uploader("Sözleşmeyi yükle", type=["docx","pdf"])
if project_type != "Ana Akım TV":
    st.info("Aktif profil şu an ACTOR_TV_MAINSTREAM.")

if uploaded:
    try:
        original_bytes = uploaded.getvalue()
        text = extract_text(uploaded)
        st.session_state["contract_text"] = text
        st.session_state["original_bytes"] = original_bytes
        st.session_state["uploaded_name"] = uploaded.name
        st.success(f"{uploaded.name} okundu • {len(text):,} karakter")

        with st.expander("Belgeden çıkarılan metni kontrol et"):
            st.text_area("Belge metni", text[:20000], height=220)

        if not get_api_key():
            st.warning("OPENAI_API_KEY tanımlı değil.")
        elif project_type == "Ana Akım TV":
            if st.button("⚖️ OLI Analizini Çalıştır", type="primary", use_container_width=True):
                with st.spinner("OLI 30 kontrol noktasını inceliyor..."):
                    result = analyse_contract(text, negotiation_power)
                    st.session_state["oli_result"] = result
                    st.session_state.pop("revision_drafts", None)
    except Exception as e:
        st.error(f"Belge/analiz hatası: {e}")

result = st.session_state.get("oli_result")
if result:
    st.divider()
    st.header("Analiz Sonucu")

    findings = result.get("findings", [])
    a,b,c,d = st.columns(4)
    a.metric("Genel Risk", result.get("overall_risk","-"))
    b.metric("Kritik", sum(f.get("status")=="RED" for f in findings))
    c.metric("Müzakere", sum(f.get("status")=="ORANGE" for f in findings))
    d.metric("Korunacak", sum(f.get("status")=="GREEN" for f in findings))

    st.subheader("Yönetici Özeti")
    st.write(result.get("executive_summary",""))

    st.subheader("Masaya Getirilecek Konular")
    for item in result.get("top_negotiation_points", [])[:5]:
        st.write("• " + item)

    st.subheader("30 Kural Analizi")
    order={"RED":0,"ORANGE":1,"YELLOW":2,"GREEN":3,"NOT_FOUND":4}
    findings=sorted(findings,key=lambda f:order.get(f.get("status"),9))

    selectable=[]
    for f in findings:
        with st.expander(f"{icon(f.get('status'))} {f.get('rule_id','')} — {f.get('title','')}"):
            st.caption(
                f"Durum: {f.get('status','-')} • Müzakere: {f.get('negotiation_priority','-')} "
                f"• Güven: {f.get('confidence','-')}"
            )
            st.write("**Referans:**", f.get("contract_reference","-"))
            st.write("**Mevcut hüküm:**", f.get("clause_excerpt","-"))
            st.write("**OLI değerlendirmesi:**", f.get("assessment",""))
            lib = revision_context_for(f.get("rule_id"))
            if lib:
                st.write("**Opus Revision Library:**", lib.get("preferred_drafting",""))
            st.write("**İlk revizyon yaklaşımı:**", f.get("recommended_revision",""))

            default = f.get("status") in ("RED","ORANGE")
            if f.get("status") not in ("GREEN",):
                chosen = st.checkbox(
                    "Word revizyonuna dahil et",
                    value=default,
                    key=f"pick_{f.get('rule_id')}"
                )
                if chosen:
                    selectable.append(f)

    if uploaded and uploaded.name.lower().endswith(".docx"):
        st.divider()
        st.header("Word Revizyon Motoru")
        st.caption("OLI mevcut hükümleri revize eder; eksik koruyucu hükümleri uygun bölüme ekler. Word biçimi mümkün olduğunca korunur ve değişiklik yazarı Av. Onur Güneş olarak işlenir.")

        if st.button("✍️ Revizyon Metinlerini Hazırla", use_container_width=True):
            if not selectable:
                st.warning("En az bir bulguyu Word revizyonuna dahil et.")
            else:
                with st.spinner("Opus revizyon kalıpları sözleşmeye uyarlanıyor..."):
                    drafts = build_revision_drafts(
                        st.session_state["contract_text"],
                        selectable,
                        negotiation_power
                    )
                    st.session_state["revision_drafts"] = drafts.get("revisions", [])

        drafts = st.session_state.get("revision_drafts", [])
        edited=[]
        if drafts:
            st.subheader("Revizyonları Kontrol Et")
            for i, rev in enumerate(drafts):
                with st.expander(f"✏️ {rev.get('rule_id')} — {rev.get('title')}", expanded=True):
                    st.caption(f"Eylem: {rev.get('action')} • Güven: {rev.get('confidence')}")
                    if rev.get("anchor_text"):
                        st.write("**Eşleşme noktası:**", rev.get("anchor_text"))
                    st.write("**Neden:**", rev.get("reason",""))
                    new_text = st.text_area(
                        "Word'e uygulanacak metin",
                        value=rev.get("replacement_text",""),
                        height=160,
                        key=f"revtext_{i}"
                    )
                    include = st.checkbox("Bu revizyonu uygula", value=True, key=f"apply_{i}")
                    if include:
                        edited.append({**rev, "replacement_text": new_text})

            if st.button("📄 Revize Word'ü Oluştur", type="primary", use_container_width=True):
                try:
                    revised_bytes, applied, skipped = apply_revisions_to_docx(
                        st.session_state["original_bytes"],
                        edited,
                        author="Av. Onur Güneş"
                    )
                    st.session_state["revised_docx"] = revised_bytes
                    st.session_state["applied_revisions"] = applied
                    st.session_state["skipped_revisions"] = skipped
                except Exception as e:
                    st.error(f"Word oluşturulamadı: {e}")

            if st.session_state.get("revised_docx"):
                name = Path(st.session_state.get("uploaded_name","sozlesme.docx")).stem
                st.success(
                    f"{len(st.session_state.get('applied_revisions',[]))} revizyon Word'e işlendi."
                )
                skipped = st.session_state.get("skipped_revisions", [])
                if skipped:
                    st.warning(f"{len(skipped)} revizyon eşleşme bulunamadığı için uygulanamadı.")
                st.download_button(
                    "⬇️ Revize Word'ü İndir",
                    data=st.session_state["revised_docx"],
                    file_name=f"{name}_OLI_revize.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
    elif uploaded:
        st.info("Word üzerinde revizyon özelliği şu an yalnız .docx dosyalarında aktif.")

st.divider()
with st.expander("OLI Bilgi Tabanı"):
    st.write(f"**Rule Library:** {len(RULES)} kontrol noktası")
    st.write(f"**Revision Library:** {len(REVISION_LIBRARY.get('entries',[]))} doğrulanmış revizyon kalıbı")

st.caption("OLI v0.4 • Prototip. Gerçek müvekkil belgeleri için erişim, veri güvenliği, saklama ve meslek sırrı mimarisi ayrıca tamamlanmalıdır.")
