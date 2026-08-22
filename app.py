import streamlit as st
import json
import re
from io import BytesIO
from pathlib import Path

from docx import Document
from pypdf import PdfReader
from openai import OpenAI

st.set_page_config(
    page_title="Opus Legal Intelligence",
    page_icon="⚖️",
    layout="wide"
)

RULES = json.loads(
    Path(__file__).with_name("rules.json").read_text(encoding="utf-8")
)

POWER_GUIDANCE = {
    "Düşük": "Sadece kritik konuları masaya taşı; ikincil risklerde pazarlık sermayesini koru.",
    "Orta": "Kritik riskleri ve önemli ticari risklerin çoğunu müzakere et.",
    "Yüksek": "Opus standartlarına geniş ölçüde yaklaş; yüksek ve orta risklerde güçlü revizyon iste.",
    "Çok Yüksek": "İdeal Opus pozisyonuna mümkün olduğunca yaklaş; gereksiz tek taraflı hükümleri kabul etme.",
}

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
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY tanımlı değil.")

    rules_text = "\n".join(
        f"{r['id']} | {r['title']} | Öncelik: {r['priority']} | Opus standardı: {r['standard']}"
        for r in RULES
    )

    max_chars = 140_000
    if len(contract_text) > max_chars:
        contract_text = contract_text[:max_chars]

    system = """Sen Opus Legal Intelligence (OLI) sözleşme analiz motorusun.
Görevin ana akım TV dizisi oyuncu sözleşmesini yalnızca verilen Opus kural setine göre incelemektir.

TEMEL İLKELER:
- Müvekkil tarafı OYUNCU/AJANS'tır.
- Kural setindeki standardı esas al.
- Bir hüküm avantajlıysa sırf konu başlığı geçti diye risk üretme; GREEN/KORU diyebilirsin.
- Hüküm yoksa bunu "NOT_FOUND" olarak işaretle; yoktan madde uydurma.
- Sözleşmeden kısa ve doğru bir alıntı/parafraz ver; madde numarası görünüyorsa belirt.
- "Revize edilmemiş = kabul edilebilir" şeklinde bir varsayım yapma.
- Pazarlık gücü hukuki riski değiştirmez; yalnızca müzakere önceliğini etkiler.
- Metinde desteklenmeyen bilgi üretme.
- Yanıt SADECE geçerli JSON olsun; markdown kullanma.

JSON ŞEMASI:
{
  "overall_risk": "LOW|MEDIUM|HIGH|VERY_HIGH",
  "executive_summary": "kısa özet",
  "top_negotiation_points": ["en fazla 5 kısa madde"],
  "findings": [
    {
      "rule_id": "OLI-TV-001",
      "title": "Münhasırlık",
      "status": "RED|ORANGE|YELLOW|GREEN|NOT_FOUND",
      "contract_reference": "madde no veya bölüm",
      "clause_excerpt": "çok kısa sözleşme özeti/alinti",
      "assessment": "neden",
      "recommended_revision": "kısa öneri",
      "negotiation_priority": "MUST|SHOULD|OPTIONAL|KEEP",
      "confidence": "HIGH|MEDIUM|LOW"
    }
  ]
}

Her 30 kural için findings içinde bir kayıt üret.
"""

    user = f"""PROFİL: ACTOR_TV_MAINSTREAM
PAZARLIK GÜCÜ: {negotiation_power}
PAZARLIK STRATEJİSİ: {POWER_GUIDANCE[negotiation_power]}

OPUS RULE LIBRARY:
{rules_text}

SÖZLEŞME:
{contract_text}
"""

    client = OpenAI(api_key=api_key)
    model = "gpt-5.6-terra"
    try:
        model = st.secrets.get("OPENAI_MODEL", model)
    except Exception:
        pass

    response = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return clean_json(response.output_text)

def status_icon(status):
    return {
        "RED": "🔴",
        "ORANGE": "🟠",
        "YELLOW": "🟡",
        "GREEN": "🟢",
        "NOT_FOUND": "⚪",
    }.get(status, "⚪")

st.title("OPUS LEGAL INTELLIGENCE")
st.caption("Private Legal Intelligence • Prototype v0.2")

modules = st.columns(4)
with modules[0]:
    st.subheader("Sözleşmeler")
    st.write("Analyse & Negotiate")
with modules[1]:
    st.subheader("Arabuluculuk")
    st.write("Yakında")
with modules[2]:
    st.subheader("KVKK")
    st.write("Yakında")
with modules[3]:
    st.subheader("Dava Dosyaları")
    st.write("Yakında")

st.divider()
st.header("Yeni Sözleşme Analizi")

c1, c2, c3 = st.columns(3)
with c1:
    contract_type = st.selectbox("Sözleşme türü", ["Oyuncu Sözleşmesi"])
with c2:
    project_type = st.selectbox("Proje türü", ["Ana Akım TV", "Dijital", "Sinema"])
with c3:
    negotiation_power = st.select_slider(
        "Pazarlık gücü",
        ["Düşük", "Orta", "Yüksek", "Çok Yüksek"],
        value="Orta"
    )

uploaded = st.file_uploader("Sözleşmeyi yükle", type=["docx", "pdf"])

if project_type != "Ana Akım TV":
    st.info("Aktif hukuk profili şu an ACTOR_TV_MAINSTREAM. Dijital ve Sinema sonraki sürümlerde.")

if uploaded:
    try:
        text = extract_text(uploaded)
        st.success(f"{uploaded.name} okundu • {len(text):,} karakter")
        if len(text.strip()) < 300:
            st.warning("Belgeden yeterli metin çıkarılamadı. Taranmış/görüntü tabanlı PDF olabilir.")
        with st.expander("Çıkarılan metni kontrol et"):
            st.text_area("Belge metni", text[:20_000], height=250)

        api_key = get_api_key()
        if not api_key:
            st.warning("AI analizi için Streamlit Secrets altında OPENAI_API_KEY eklenmesi gerekiyor.")
        elif project_type == "Ana Akım TV":
            if st.button("⚖️ OLI Analizini Çalıştır", type="primary", use_container_width=True):
                with st.spinner("OLI 30 kontrol noktasını inceliyor..."):
                    try:
                        result = analyse_contract(text, negotiation_power)
                        st.session_state["oli_result"] = result
                    except Exception as e:
                        st.error(f"Analiz çalıştırılamadı: {e}")

    except Exception as e:
        st.error(f"Belge okunamadı: {e}")

result = st.session_state.get("oli_result")
if result:
    st.divider()
    st.header("OLI Analiz Sonucu")

    risk = result.get("overall_risk", "-")
    findings = result.get("findings", [])
    reds = sum(1 for f in findings if f.get("status") == "RED")
    oranges = sum(1 for f in findings if f.get("status") == "ORANGE")
    greens = sum(1 for f in findings if f.get("status") == "GREEN")

    a, b, c, d = st.columns(4)
    a.metric("Genel Risk", risk)
    b.metric("Kritik", reds)
    c.metric("Müzakere", oranges)
    d.metric("Korunacak", greens)

    st.subheader("Yönetici Özeti")
    st.write(result.get("executive_summary", ""))

    st.subheader("Masaya Getirilecek Konular")
    for item in result.get("top_negotiation_points", [])[:5]:
        st.write(f"• {item}")

    st.subheader("30 Kural Analizi")
    order = {"RED": 0, "ORANGE": 1, "YELLOW": 2, "GREEN": 3, "NOT_FOUND": 4}
    findings = sorted(findings, key=lambda x: order.get(x.get("status"), 9))

    for f in findings:
        icon = status_icon(f.get("status"))
        title = f"{icon} {f.get('rule_id','')} — {f.get('title','')}"
        with st.expander(title):
            st.caption(
                f"Durum: {f.get('status','-')} • "
                f"Müzakere: {f.get('negotiation_priority','-')} • "
                f"Güven: {f.get('confidence','-')}"
            )
            if f.get("contract_reference"):
                st.write("**Sözleşme referansı:**", f.get("contract_reference"))
            if f.get("clause_excerpt"):
                st.write("**Mevcut hüküm:**", f.get("clause_excerpt"))
            st.write("**OLI değerlendirmesi:**", f.get("assessment", ""))
            st.write("**Önerilen revizyon:**", f.get("recommended_revision", ""))

    st.download_button(
        "Analiz JSON'unu indir",
        data=json.dumps(result, ensure_ascii=False, indent=2),
        file_name="oli_analysis.json",
        mime="application/json"
    )

st.divider()
with st.expander("Aktif OLI Rule Library"):
    st.metric("Doğrulanmış kontrol noktası", len(RULES))
    for r in RULES:
        st.write(f"**{r['id']} — {r['title']}** · {r['priority']}")
        st.caption(r["standard"])

st.caption(
    "v0.2 prototip notu: Gerçek müvekkil belgeleriyle kullanmadan önce erişim, veri güvenliği, "
    "saklama ve meslek sırrı mimarisi ayrıca tamamlanmalıdır."
)
