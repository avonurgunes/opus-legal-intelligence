import streamlit as st
import json
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path

from docx import Document
from pypdf import PdfReader
from openai import OpenAI

from revision_engine import apply_revisions_to_docx
from mediation import render_mediation

st.set_page_config(
    page_title="Opus Legal Intelligence",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
:root{color-scheme:light}
html,body,[data-testid="stAppViewContainer"],.stApp{background:#F7F7F5!important;color:#17191C!important}
[data-testid="stHeader"]{background:#0B1118!important;border-bottom:1px solid #202833!important}
[data-testid="stAppViewBlockContainer"]{max-width:1180px!important;padding-top:1.25rem!important;padding-bottom:3rem!important}
h1,h2,h3,h4{color:#111418!important;letter-spacing:-.025em!important}
p,label,[data-testid="stCaptionContainer"]{color:#5E6670!important}
.oli-brand{font-size:1.45rem;font-weight:850;color:#D6A13A!important;letter-spacing:.03em;margin-top:.1rem}
.oli-brand-sub{font-size:.72rem;color:#8A929B!important;margin-top:.15rem}
div[role="radiogroup"]{justify-content:flex-end!important;gap:.35rem!important}
div[role="radiogroup"] label{background:#121923!important;border:1px solid #26303B!important;padding:.42rem .78rem!important;border-radius:8px!important}
div[role="radiogroup"] label p{color:#E8E8E5!important}
div[role="radiogroup"] label:has(input:checked){border-color:#D39A2E!important;box-shadow:inset 0 -2px 0 #D39A2E!important}
.oli-workspace{background:#FFF;border:1px solid #E2E4E7;border-radius:14px;padding:22px 24px;margin-top:18px;box-shadow:0 4px 18px rgba(17,24,39,.055)}
.oli-workspace-title{font-size:1.35rem;font-weight:800;color:#121519;margin-bottom:4px;border-left:4px solid #D49A2A;padding-left:12px}
.oli-workspace-sub{font-size:.9rem;color:#69717A;margin:5px 0 4px 16px}
.oli-resultbar{background:#FFF;border:1px solid #E2E4E7;border-left:4px solid #D49A2A;border-radius:11px;padding:15px 17px;margin:14px 0 10px;color:#2B3137;box-shadow:0 2px 10px rgba(17,24,39,.035)}
div[data-testid="stFileUploaderDropzone"],div[data-testid="stTextInput"] input,div[data-testid="stTextArea"] textarea,div[data-testid="stSelectbox"] div[data-baseweb="select"]>div{background:#FFF!important;border-color:#D8DDE2!important;color:#1F2328!important}
div[data-testid="stExpander"] details{background:#FFF!important;border:1px solid #E2E4E7!important;border-radius:10px!important}
.stButton>button,.stDownloadButton>button{min-height:46px!important;border-radius:8px!important;font-weight:750!important}
.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{background:linear-gradient(90deg,#C68B22,#E1A735)!important;color:#FFF!important;border:0!important;box-shadow:0 3px 10px rgba(190,132,29,.20)!important}
hr{border-color:#E3E6E8!important}
</style>
""",unsafe_allow_html=True)



RULES = json.loads(Path(__file__).with_name("rules.json").read_text(encoding="utf-8"))
REVISION_LIBRARY = json.loads(
    Path(__file__).with_name("revision_library.json").read_text(encoding="utf-8")
)
CLAUSE_BANK = json.loads(
    Path(__file__).with_name("clause_bank.json").read_text(encoding="utf-8")
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


def analyse_contract(contract_text: str, negotiation_power: str, initial_note: str = ""):
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

MICRO JSON ÖRNEĞİ:
{"revisions":[{"rule_id":"OLI-TV-XXX","title":"Münhasırlık","action":"MICRO","anchor_text":"görev almayacaktır","replacement_text":"görev alabilir","reason":"Münhasırlık kaldırılıyor","confidence":"HIGH"}]}

PHRASE JSON ÖRNEĞİ:
{"revisions":[{"rule_id":"OLI-TV-XXX","title":"Onay","action":"PHRASE","anchor_text":"YAPIMCI'nın önceden yazılı onayı ile","replacement_text":"YAPIMCI'ya önceden yazılı bilgi verilmesi kaydıyla","reason":"Onay bilgilendirmeye çevriliyor","confidence":"HIGH"}]}

NOT: MICRO/PHRASE çıktısında replacement_text TAM PARAGRAF OLAMAZ.
}
Her 30 kural için findings üret.
"""

    user = f"""PROFİL: ACTOR_TV_MAINSTREAM
PAZARLIK GÜCÜ: {negotiation_power}
STRATEJİ: {POWER_GUIDANCE[negotiation_power]}
DOSYAYA ÖZGÜ İLK NOT / AJANS NOTU: {initial_note or "Yok"}
Bu not yalnız bu dosyanın analiz ve müzakere önceliklerini etkiler; Opus standardı değildir.

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


def clause_bank_for(rule_id: str):
    return [e for e in CLAUSE_BANK.get("entries", []) if e.get("rule_id") == rule_id]


def build_revision_drafts(contract_text, selected_findings, negotiation_power):
    lib_parts = []
    bank_parts = []
    for f in selected_findings:
        entry = revision_context_for(f.get("rule_id"))
        if entry:
            lib_parts.append(json.dumps(entry, ensure_ascii=False))
        for bank_entry in clause_bank_for(f.get("rule_id")):
            bank_parts.append(json.dumps(bank_entry, ensure_ascii=False))

    system = """Sen OLI Revision Engine'sin.

ZORUNLU DRAFTING PRENSİBİ — MICRO FIRST:
1. Mevcut cümlenin terminolojisini, kelime sırasını ve iskeletini mümkün olduğunca koru.
2. Hukuki sonucu yalnız gerekli kelime/ibare/sayı/istisnayı değiştirerek elde edebiliyorsan tüm cümleyi yeniden yazma.
3. Örnek: 'Oyuncu dizi süresi boyunca başka bir projede görev almayacaktır.' → hedef serbestlik ise yalnız 'almayacaktır' kısmını 'alabilir' olarak değiştir.
4. MICRO yetmezse PHRASE; PHRASE yetmezse ancak o zaman BLOCK/REPLACE_PARAGRAPH kullan.
5. Stil güzelleştirmek veya kalıp metin kullanmak REPLACE_PARAGRAPH gerekçesi değildir.
6. original_text mümkün olan en küçük güvenli eşleşme parçası; revised_text yalnız onun doğal karşılığı olmalıdır.
7. REPLACE_PARAGRAPH son çaredir.

Görevin sözleşmedeki seçilmiş bulgular için OPUS tarzında uygulanabilir revizyon metni üretmektir.

ÖNCELİK:
1) Varsa MADDE BANKASI'ndaki Av. Onur Güneş tarafından gözden geçirilmiş hazır cümle.
2) Varsa REVISION LIBRARY'deki doğrulanmış Opus yaklaşımı.
3) Rule Library'deki standardın özü.
4) Mevcut sözleşmenin terminolojisi, taraf tanımları, proje adı ve madde dili.

KURALLAR:
- Yeni hukuki politika icat etme.
- MADDE BANKASI'nda uygun cümle varsa esas metin olarak onu kullan; yalnız sözleşmenin tanımlarına, proje adına, madde numarasına ve mevcut cümlenin gramerine uyacak kadar değiştir.
- Madde Bankası metnini gereksiz yere uzatma, yeni şartlar ekleme veya daha ayrıntılı hale getirme.
- Mevcut hüküm kısmen uygunsa paragrafı baştan yazmak yerine mümkün olan en küçük kelime/cümlecik değişikliğiyle Madde Bankası standardını mevcut cümleye yedir.
- Geçmiş Opus metnini körü körüne kopyalama; mevcut sözleşmeye uyarla.
- Mevcut hüküm varsa ÖNCE action=MICRO düşün. Yalnız birkaç kelime/cümlecik yetmiyorsa PHRASE kullan. REPLACE_PARAGRAPH yalnız mevcut paragrafın iskeleti korunarak güvenli sonuç elde edilemiyorsa kullanılabilir.
- MICRO için anchor_text yalnız değişecek birebir kelime/ibare/cümlecik olsun; replacement_text yalnız onun yerine gelecek metin olsun.
- PHRASE için anchor_text değişecek birebir cümlecik olsun; replacement_text yalnız o cümleciğin yeni hali olsun.
- REPLACE_PARAGRAPH için anchor_text paragraftan ayırt edici bir parça, replacement_text ise tam yeni paragraf olabilir.
- Koruyucu hüküm tamamen eksikse NOT_FOUND diye bırakma: mutlaka yeni hüküm üret. Sözleşmede konu bakımından en uygun mevcut maddeyi anchor seç ve APPEND_AFTER kullan. APPEND_END yalnız gerçekten uygun bölüm bulunamıyorsa son çaredir.
- Eksik hükmün ekleneceği yeri konu bütünlüğüne göre seç: ücret/ödeme hükümleri ücret bölümüne; mali hak/telif hükümleri mali haklar bölümüne; fesih/cezai şart hükümleri ilgili fesih/ceza bölümüne; vergi/damga vergisi diğer hükümler/vergi bölümüne.
- APPEND_AFTER kullanırken anchor_text mutlaka sözleşmede BİREBİR bulunan ve hedef bölümdeki son uygun paragraftan 25-100 karakterlik bir parça olmalı. Madde numarasını uydurma. Yeni metnin numaralandırması belirsizse numarasız koruyucu paragraf üret; Word'e yerleştirildikten sonra kullanıcı kontrol eder.
- replacement_text sadece sözleşmeye girecek nihai madde/paragraf metni olsun.
- Pazarlık gücü sadece sertlik düzeyini belirler.
- Yanıt yalnız geçerli JSON.

ŞEMA:
{
 "revisions":[{
   "rule_id":"...",
   "title":"...",
   "action":"MICRO|PHRASE|REPLACE_PARAGRAPH|APPEND_AFTER|APPEND_END",
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

MADDE BANKASI:
{chr(10).join(bank_parts) if bank_parts else "Bu kurallar için onaylı hazır madde yok."}

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



def analyse_extra_risks(contract_text: str):
    system = """Sen OLI Extra Risk Scanner'sın.
30 maddelik Opus Rule Library DIŞINDA kalan sözleşmesel riskleri ara.
Rule Library'deki konuları tekrar etme. Yalnız gerçekten yeni ve anlamlı hukuki/ticari bulguları yaz.
Bulguları otomatik revize etme; kullanıcıya not olarak sun.
Yanıt yalnız JSON:
{"extra_findings":[{"title":"...","risk":"HIGH|MEDIUM|LOW","reference":"madde/bölüm","assessment":"...","suggested_action":"..."}]}
Hiç ek bulgu yoksa boş liste döndür."""
    user = f"""MEVCUT 30 KURAL BAŞLIĞI:
{chr(10).join(r.get('title','') for r in RULES)}

SÖZLEŞME:
{contract_text[:140000]}"""
    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[{"role":"system","content":system},{"role":"user","content":user}]
    )
    return clean_json(resp.output_text)


def compare_final_revision(oli_text: str, final_text: str):
    system = """Sen OLI Learning Review motorusun.
OLI'nin ürettiği revize metin ile Av. Onur Güneş'in nihai revize metnini karşılaştır.
Yalnız anlamlı hukuki/drafting farklarını çıkar.
Bir hükme dokunulmamasını asla otomatik olarak 'Opus standardı' kabul etme.
Her fark için bunun kütüphaneye alınabilecek genel bir tercih mi, yoksa dosyaya özgü mü olabileceğini öner.
Yanıt yalnız JSON:
{"learning_candidates":[{"title":"...","oli_position":"...","final_position":"...","difference":"...","recommendation":"LIBRARY_CANDIDATE|CASE_SPECIFIC","confidence":"HIGH|MEDIUM|LOW"}]}"""
    user = f"""OLI REVİZE METİN:
{oli_text[:100000]}

AV. ONUR GÜNEŞ NİHAİ METİN:
{final_text[:100000]}"""
    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[{"role":"system","content":system},{"role":"user","content":user}]
    )
    return clean_json(resp.output_text)



def build_initial_review_note(result):
    system = """Sen OLI Ajans Bilgilendirme Notu motorusun.
Verilen sözleşme analizinden yalnız müdahale edilmesi, ajansa bildirilmesi veya karar alınması gereken noktaları kısa Türkçe not halinde yaz.
Her satır mümkünse sözleşme madde numarasıyla başlasın.
Dil kısa ve pratik olsun: '5.4. Münhasırlık Yapımcı onayına bağlanmış. Bilgilendirmeye çevirelim.' gibi.
Uzun hukuki açıklama yapma. GREEN/KORU maddelerini yazma. Yanıt yalnız JSON:
{"note_items":[{"reference":"5.4","title":"Münhasırlık","note":"..."}]}"""
    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[{"role":"system","content":system},
               {"role":"user","content":json.dumps(result, ensure_ascii=False)[:100000]}]
    )
    return clean_json(resp.output_text)


def classify_word_flags(contract_text: str, result: dict, extra_risks: dict | None = None):
    rules_text = "\n".join(f"{r['id']} | {r['title']} | {r['standard']}" for r in RULES)
    system = """Sen OLI Word Flag motorusun.
Sözleşmedeki yalnız iki tür hükmü işaretle:
ORANGE: hukuki/ticari risk veya manuel dikkat gerektiren hüküm; fakat Word revizyonu ile zaten doğrudan değiştirilecek parça olmak zorunda değil.
BLUE: mevcut Rule Library ile ANLAMSAL olarak eşleşmeyen, gerçekten yeni/öğrenilmemiş bir hukuki düzenleme.
Aynı konunun farklı kelimelerle yazılması BLUE değildir. Örneğin farklı yazılmış cezai şart yine cezai şarttır.
Her flag için anchor_text sözleşmede BİREBİR geçen 25-100 karakterlik kısa parça olmalı.
Yanıt yalnız JSON:
{"flags":[{"color":"orange|blue","title":"...","reference":"...","anchor_text":"...","reason":"..."}]}"""
    user = f"""RULE LIBRARY:
{rules_text}

30 KURAL ANALİZİ:
{json.dumps(result, ensure_ascii=False)}

EK RİSKLER:
{json.dumps(extra_risks or {}, ensure_ascii=False)}

SÖZLEŞME:
{contract_text[:140000]}"""
    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(model=get_model(), input=[{"role":"system","content":system},{"role":"user","content":user}])
    return clean_json(resp.output_text)


def compare_counterparty_return(our_text: str, returned_text: str):
    system = """Sen OLI Negotiation Compare motorusun.
Av. Onur Güneş tarafından gönderilen revize sözleşme ile karşı taraftan dönen sözleşmeyi hukuki anlam bakımından karşılaştır.
Sınıflar:
ACCEPTED = bizim değişiklik kabul edilmiş.
PARTIAL = kısmen kabul edilmiş/değiştirilmiş.
REJECTED = bizim değişiklik kaldırılmış/eski pozisyona dönülmüş.
NEW = karşı taraf yeni bir hüküm eklemiş.
Kısa, ajansa aktarılabilir dil kullan. Yanıt yalnız JSON:
{"items":[{"reference":"...","title":"...","status":"ACCEPTED|PARTIAL|REJECTED|NEW","our_position":"...","counterparty_position":"...","note":"..."}]}"""
    client = OpenAI(api_key=get_api_key())
    resp = client.responses.create(
        model=get_model(),
        input=[{"role":"system","content":system},
               {"role":"user","content":f"BİZİM GÖNDERDİĞİMİZ:\\n{our_text[:100000]}\\n\\nKARŞI TARAF DÖNÜŞÜ:\\n{returned_text[:100000]}"}]
    )
    return clean_json(resp.output_text)




def _norm_ref(ref):
    return (ref or "").strip().rstrip(".").replace("MADDE ","").replace("Madde ","").strip()

def extract_full_original_clause(contract_text, item):
    text=contract_text or ""; ref=_norm_ref(item.get("reference") or item.get("rule_id") or "")
    lines=text.splitlines()
    if ref and re.fullmatch(r"\d+(?:\.\d+)*",ref):
        pat=re.compile(rf"^\s*(?:MADDE\s+)?{re.escape(ref)}(?:[\.\-\):]|\s)",re.I)
        start=next((i for i,x in enumerate(lines) if pat.search(x)),None)
        if start is not None:
            depth=ref.count("."); collected=[lines[start]]
            nxt=re.compile(r"^\s*(?:MADDE\s+)?(\d+(?:\.\d+)*)(?:[\.\-\):]|\s)",re.I)
            for line in lines[start+1:]:
                m=nxt.search(line)
                if m and m.group(1)!=ref and m.group(1).count(".")<=depth: break
                collected.append(line)
            full="\n".join(collected).strip()
            if full: return full
    target=(item.get("target_text") or item.get("anchor_text") or "").strip()
    if target:
        for line in lines:
            if target in line or (len(target)>35 and target[:35] in line): return line.strip()
    return target or "Orijinal madde bulunamadı."

def build_full_revised_clause(original,item):
    if (item.get("mode") or "REPLACE").upper()=="APPEND_AFTER":
        return (original.rstrip()+"\n\n"+(item.get("append_text") or "").strip()).strip()
    repl=(item.get("replacement_text") or "").strip(); target=(item.get("target_text") or "").strip()
    if target and target in original: return original.replace(target,repl,1)
    return repl or original

def _preview_default_mode(item):
    req = (item.get("redline_mode") or "AUTO").upper()
    if item.get("mode") == "APPEND_AFTER":
        return "➕ Yeni hüküm"
    if req == "MICRO":
        return "🩹 Mikro değişiklik"
    if req == "PHRASE":
        return "✂️ Cümlecik / doğal ekleme"
    if req == "BLOCK":
        return "📝 Tam blok değişikliği"
    return "🤖 Otomatik seç"


def render_revision_preview(items, contract_text):
    st.subheader("Revizyon Önizleme")
    st.caption("Gerekçe → tam orijinal madde → tam revize madde. Word yalnız kabul ettiğin son metinlerden oluşturulur.")
    edited=[]; accepted=rejected=manual=0
    mm={"🤖 Otomatik seç":"AUTO","🩹 Mikro değişiklik":"MICRO","✂️ Cümlecik / doğal ekleme":"PHRASE","📝 Tam blok değişikliği":"BLOCK","➕ Yeni hüküm":"AUTO"}
    for i,item in enumerate(items):
        rid=item.get("reference") or item.get("rule_id") or f"REV-{i+1}"
        title=item.get("title") or item.get("rule_title") or rid
        original=extract_full_original_clause(contract_text,item)
        proposed=build_full_revised_clause(original,item)
        with st.expander(f"{i+1}. {rid} — {title}",expanded=(i<3)):
            st.markdown("**Neden revize ediyorum?**")
            st.info(item.get("reason") or item.get("problem") or "Revizyon gerekçesi belirtilmemiş.")
            st.markdown("**ORİJİNAL MADDE**")
            st.text_area("Orijinal",original,height=150,disabled=True,key=f"orig_{i}",label_visibility="collapsed")
            st.markdown("**REVİZE MADDE**")
            final=st.text_area("Revize",proposed,height=180,key=f"revised_{i}",label_visibility="collapsed")
            c1,c2=st.columns(2)
            decision=c1.radio("Karar",["✅ Kabul","❌ Reddet"],horizontal=True,key=f"decision_{i}")
            choices=["🤖 Otomatik seç","🩹 Mikro değişiklik","✂️ Cümlecik / doğal ekleme","📝 Tam blok değişikliği"]
            if item.get("mode")=="APPEND_AFTER": choices=["➕ Yeni hüküm","🤖 Otomatik seç"]
            default=_preview_default_mode(item); idx=choices.index(default) if default in choices else 0
            vm=c2.selectbox("Word'de uygulanma biçimi",choices,index=idx,key=f"vmode_{i}")
            if decision=="❌ Reddet": rejected+=1; continue
            accepted+=1; ni=dict(item); ni["redline_mode"]=mm[vm]
            ni["preview_original_clause"]=original; ni["preview_final_clause"]=final
            if (item.get("mode") or "REPLACE").upper()=="APPEND_AFTER":
                ni["append_text"]=final[len(original):].strip() if final.startswith(original) else final
            else:
                ni["target_text"]=original; ni["replacement_text"]=final
            if final.strip()!=proposed.strip(): manual+=1; ni["edited_by_user"]=True
            edited.append(ni)
    st.info(f"☑ {accepted} kabul • ☒ {rejected} reddet • ✏️ {manual} manuel düzenleme")
    return edited




# Compact top navigation
nav1, nav2 = st.columns([1.2, 4.8])
with nav1:
    st.markdown('<div class="oli-brand">OLI</div><div class="oli-brand-sub">Opus Legal Intelligence</div>', unsafe_allow_html=True)
with nav2:
    selected_module = st.radio(
        "Modül",
        ["Sözleşmeler", "Arabuluculuk", "Madde Bankası"],
        horizontal=True,
        label_visibility="collapsed",
        key="oli_module_nav_v59"
    )
st.divider()
if selected_module == "Arabuluculuk":
    render_mediation()
    st.stop()
elif selected_module == "Madde Bankası":
    st.header("Madde Bankası")
    st.caption("Opus'un kontrollü revizyon kalıpları ve drafting tercihleri.")
    with st.expander("Aktif Madde Bankası", expanded=True):
        for rid, item in REVISION_LIBRARY.items():
            st.markdown(f"**{rid} — {item.get('title','')}**")
            st.write(item.get("preferred_drafting",""))
    st.stop()

st.markdown("""
<div class="oli-workspace">
  <div class="oli-workspace-title">Sözleşme Revizyonu</div>
  <div class="oli-workspace-sub">Belgeyi yükle, OLI analiz etsin ve revize Word dosyasını hazırlasın.</div>
</div>
""", unsafe_allow_html=True)

c1,c2,c3 = st.columns([1,1,1])
with c1:
    contract_type = st.selectbox("Sözleşme türü", ["Oyuncu Sözleşmesi"])
with c2:
    project_type = st.selectbox("Proje türü", ["Ana Akım TV", "Dijital", "Sinema"])
with c3:
    negotiation_power = st.select_slider(
        "Pazarlık gücü", ["Düşük","Orta","Yüksek","Çok Yüksek"], value="Orta"
    )

initial_note = st.text_area(
    "Kısa Not / Ajans Notu",
    placeholder="Dosyaya özgü talimatı yaz. Örn. ücret tamam; münhasırlık kaldırılacak; bölüm garantisi 8 bölüm.",
    height=72
)

uploaded = st.file_uploader("Sözleşmeyi yükle", type=["docx","pdf"], key="contract_upload_v58")

if project_type != "Ana Akım TV":
    st.info("Aktif revizyon profili şu an ACTOR_TV_MAINSTREAM. Dijital ve Sinema modülleri sonraki aşamada detaylandırılacak.")

if uploaded:
    try:
        original_bytes = uploaded.getvalue()
        text = extract_text(uploaded)
        st.session_state["contract_text"] = text
        st.session_state["original_bytes"] = original_bytes
        st.session_state["uploaded_name"] = uploaded.name
        st.success(f"{uploaded.name} okundu")

        if not get_api_key():
            st.warning("OPENAI_API_KEY tanımlı değil.")
        elif not uploaded.name.lower().endswith(".docx"):
            st.warning("Otomatik Word revizyonu için .docx yükle. PDF şu an yalnız metin analizi için okunabilir.")
        elif project_type == "Ana Akım TV":
            if st.button("OLI ANALİZİNİ ÇALIŞTIR VE WORD'E AKTAR", type="primary", use_container_width=True):
                try:
                    with st.spinner("OLI sözleşmeyi analiz ediyor ve revize Word'ü hazırlıyor..."):
                        # 1) Main legal analysis
                        result = analyse_contract(text, negotiation_power, initial_note)
                        st.session_state["oli_result"] = result
                        try:
                            st.session_state["initial_review_note"] = build_initial_review_note(result)
                        except Exception:
                            st.session_state["initial_review_note"] = {"note_items": []}


                        # 2) Extra-risk layer automatically
                        try:
                            extra_risks = analyse_extra_risks(text)
                        except Exception:
                            extra_risks = {"extra_findings":[]}
                        st.session_state["extra_risks"] = extra_risks

                        # 3) Auto-select material findings
                        findings = result.get("findings", [])
                        selected = [
                            f for f in findings
                            if f.get("status") in ("RED","ORANGE","YELLOW")
                            and f.get("status") != "GREEN"
                        ]

                        # 4) Build natural revisions automatically
                        drafts_obj = build_revision_drafts(
                            text,
                            selected,
                            negotiation_power
                        )
                        drafts = drafts_obj.get("revisions", [])
                        st.session_state["revision_drafts"] = drafts

                        # 5) Color flags automatically — no separate button
                        try:
                            flags = classify_word_flags(
                                text,
                                result,
                                extra_risks
                            ).get("flags", [])
                        except Exception:
                            flags = []
                        st.session_state["word_flags"] = flags

                        # 6) Apply directly to Word
                        revised_bytes, applied, skipped, placeholder_count, flag_stats = apply_revisions_to_docx(
                            original_bytes,
                            drafts,
                            author="Av. Onur Güneş",
                            flags=flags
                        )

                        st.session_state["revised_docx"] = revised_bytes
                        st.session_state["applied_revisions"] = applied
                        st.session_state["skipped_revisions"] = skipped
                        st.session_state["placeholder_count"] = placeholder_count
                        st.session_state["flag_stats"] = flag_stats

                    st.success("Revizyon tamamlandı.")
                except Exception as e:
                    st.error(f"Belge/analiz hatası: {e}")

    except Exception as e:
        st.error(f"Belge okunamadı: {e}")

if st.session_state.get("revised_docx"):
    applied = st.session_state.get("applied_revisions", [])
    skipped = st.session_state.get("skipped_revisions", [])
    ph = st.session_state.get("placeholder_count", 0)
    fs = st.session_state.get("flag_stats", {})
    result = st.session_state.get("oli_result", {})
    findings = result.get("findings", [])

    critical = sum(f.get("status") == "RED" for f in findings)
    orange = fs.get("orange", 0)
    blue = fs.get("blue", 0)

    st.markdown(f"""
    <div class="oli-resultbar">
      <strong>Revize Word hazır.</strong><br>
      {len(applied)} değişiklik uygulandı · {critical} kritik bulgu · 🟨 {ph} boş alan · 🟧 {orange} risk işareti · 🟦 {blue} yeni/standart dışı hüküm
    </div>
    """, unsafe_allow_html=True)

    note_items = st.session_state.get("initial_review_note", {}).get("note_items", [])
    if note_items:
        st.markdown("#### Kısa Notlar")
        for ni in note_items[:12]:
            ref = (ni.get("reference") or "").strip()
            title = (ni.get("title") or "").strip()
            note = (ni.get("note") or "").strip()
            lead = " — ".join(x for x in [ref, title] if x)
            st.write(f"**{lead}**: {note}" if lead else note)

    if skipped:
        st.warning(f"{len(skipped)} revizyon Word üzerinde eşleşme bulunamadığı için uygulanamadı.")

    name = Path(st.session_state.get("uploaded_name","sozlesme.docx")).stem
    today_tr = datetime.now(ZoneInfo("Europe/Istanbul")).strftime("%d.%m.%Y")
    st.download_button(
        "REVİZE WORD'Ü İNDİR",
        data=st.session_state["revised_docx"],
        file_name=f"{name} - {today_tr} REVİZE.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        type="primary"
    )

    # Optional details only; normal workflow does not require rereading.
    with st.expander("Analiz Özeti"):
        st.write(result.get("executive_summary",""))
        if result.get("top_negotiation_points"):
            st.markdown("**Öne çıkan konular**")
            for x in result.get("top_negotiation_points",[])[:6]:
                st.write("• " + x)
        if skipped:
            st.markdown("**Uygulanamayan revizyonlar**")
            for x in skipped:
                st.write(f"• {x.get('rule_id','')} — {x.get('reason','Eşleşmedi')}")

st.divider()
st.header("Karşı Taraf Dönüşü")
st.caption("Bizim gönderdiğimiz revize Word ile karşı taraftan dönen Word'ü karşılaştırır ve ajans bilgilendirme notunu çıkarır.")
returned_upload = st.file_uploader("Karşı taraftan dönen Word", type=["docx"], key="counterparty_return_upload")
if returned_upload and st.session_state.get("revised_docx"):
    if st.button("🔄 Karşı Taraf Revizyonunu Karşılaştır", use_container_width=True):
        try:
            our_text = read_docx(st.session_state["revised_docx"])
            returned_text = read_docx(returned_upload.getvalue())
            with st.spinner("Müzakere farkları sınıflandırılıyor..."):
                st.session_state["counterparty_review"] = compare_counterparty_return(our_text, returned_text)
        except Exception as e:
            st.error(f"Karşılaştırma yapılamadı: {e}")

    cp_items = st.session_state.get("counterparty_review", {}).get("items", [])
    if cp_items:
        icons = {"ACCEPTED":"🟢","PARTIAL":"🟡","REJECTED":"🔴","NEW":"🔵"}
        lines=[]
        for x in cp_items:
            label = icons.get(x.get("status"),"⚪")
            st.write(f"{label} **{x.get('reference','')} {x.get('title','')}** — {x.get('note','')}")
            lines.append(f"{x.get('reference','')} {x.get('note','')}".strip())
        st.text_area("Revizyon Dönüş Notu", "\n".join(lines), height=180)

        st.text_area(
            "Ajans Geri Bildirimi",
            key="agency_feedback",
            placeholder="Örn. 5.4 kabul. KDV'de ısrar. Cezai şart maksimum 3 olabilir.",
            height=100
        )
        st.caption("Ajans geri bildirimi dosyaya özgüdür; Rule Library veya Madde Bankası'na otomatik eklenmez.")




st.divider()

st.markdown("""
<style>
/* v0.6.1 final contrast override — intentionally LAST */
html,body,[data-testid="stAppViewContainer"],.stApp{background:#F7F7F5!important;color:#17191C!important}
[data-testid="stAppViewBlockContainer"] h1,
[data-testid="stAppViewBlockContainer"] h2,
[data-testid="stAppViewBlockContainer"] h3,
[data-testid="stAppViewBlockContainer"] h4{color:#15181B!important;opacity:1!important}
[data-testid="stAppViewBlockContainer"] p,
[data-testid="stAppViewBlockContainer"] label,
[data-testid="stAppViewBlockContainer"] span,
[data-testid="stAppViewBlockContainer"] small,
[data-testid="stAppViewBlockContainer"] [data-testid="stCaptionContainer"]{opacity:1!important}
[data-testid="stAppViewBlockContainer"] label p{color:#3F464D!important}
[data-testid="stAppViewBlockContainer"] [data-testid="stCaptionContainer"] p{color:#687078!important}
[data-testid="stAppViewBlockContainer"] .stMarkdown p{color:#4D555D!important}
[data-testid="stFileUploaderDropzone"] *{opacity:1!important;color:#59616A!important}
[data-testid="stFileUploaderDropzone"] button *{color:#20252A!important}
div[data-testid="stExpander"] summary *{opacity:1!important;color:#343A40!important}
.oli-footer{color:#858B91!important;font-size:.78rem!important}
</style>
""",unsafe_allow_html=True)

st.markdown("<div class=\"oli-footer\">OLI • Sözleşmeler v0.6.1 Contrast + Micro Engine + Notes + Arabuluculuk v1.3.1 • Prototip.</div>", unsafe_allow_html=True)
