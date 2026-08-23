from __future__ import annotations

import streamlit as st

from oli_contract_analysis import analyze_contract, extract_text, get_api_key
from oli_contract_profiles import CONTRACT_TYPES, PROJECT_TYPES, NEGOTIATION_LEVELS, PROFILE_STATUS


def _severity_icon(level):
    return {"KRİTİK":"🔴","ÖNEMLİ":"🟠","DİKKAT":"🟡"}.get(level,"⚪")


def render_contracts():
    st.markdown("""
    <div class="work-card intro-card">
      <div>
        <div class="eyebrow">SÖZLEŞMELER</div>
        <h1>Sözleşme İncelemesi</h1>
        <p>Sözleşmeyi yükle. OLI müdahale edilmesi gereken maddeleri, nedenini ve önerilen revizyon metnini çıkarsın.</p>
      </div>
      <div class="status-pill">Clean v0.1</div>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        c1,c2,c3=st.columns([1.2,1,1])
        with c1:
            contract_type=st.selectbox("Sözleşme türü",CONTRACT_TYPES)
        with c2:
            project_type=st.selectbox("Proje türü",PROJECT_TYPES)
        with c3:
            negotiation_power=st.segmented_control(
                "Pazarlık gücü", NEGOTIATION_LEVELS, default="Orta"
            ) or "Orta"

        initial_note=st.text_area(
            "İlk Not / Ajans Notu (Opsiyonel)",
            placeholder="Örn. Ücret tamam. Münhasırlık önemli. Oyuncu başka dijital projelerde çalışabilmeli.",
            height=92,
            max_chars=1000,
        )

        uploaded=st.file_uploader("Sözleşmeyi yükle",type=["docx","pdf"])

        status=PROFILE_STATUS.get((contract_type,project_type),"BASIC")
        if status=="BASIC":
            st.caption("Bu kombinasyonda özel Rule Library henüz tamamlanmadı; OLI genel entertainment/media sözleşme analizi yapar.")

        run=st.button("OLI ANALİZİNİ ÇALIŞTIR",type="primary",use_container_width=True)

    if run:
        if not uploaded:
            st.warning("Önce sözleşmeyi yükle.")
            return
        if not get_api_key():
            st.error("OPENAI_API_KEY tanımlı değil.")
            return
        try:
            raw=uploaded.getvalue()
            text=extract_text(uploaded.name,raw)
            if len(text.strip())<100:
                st.error("Belgeden yeterli metin çıkarılamadı.")
                return
            with st.spinner("OLI sözleşmeyi inceliyor..."):
                result=analyze_contract(
                    text,contract_type,project_type,negotiation_power,initial_note
                )
            st.session_state["clean_contract_result"]=result
            st.session_state["clean_contract_name"]=uploaded.name
        except Exception as e:
            st.error(f"Analiz hatası: {e}")

    result=st.session_state.get("clean_contract_result")
    if not result:
        return

    st.markdown("---")
    a,b=st.columns([1,3])
    a.metric("Genel Risk",result.get("overall_risk","-"))
    b.markdown(
        f"""<div class="summary-card"><b>Yönetici Özeti</b><br>{result.get("executive_summary","")}</div>""",
        unsafe_allow_html=True,
    )

    notes=result.get("short_notes",[])
    if notes:
        st.subheader("OLI Kısa Notları")
        note_text="\n".join(
            f"{n.get('reference','')}: {n.get('note','')}".strip(": ")
            for n in notes
        )
        st.text_area("Ajansa aktarılabilir kısa not",note_text,height=min(220,70+len(notes)*32))

    revisions=result.get("revisions",[])
    st.subheader(f"Revizyon Önerileri ({len(revisions)})")
    if not revisions:
        st.success("Anlamlı bir revizyon önerisi bulunmadı.")
        return

    for i,item in enumerate(revisions,1):
        sev=item.get("severity","")
        ref=item.get("reference","")
        title=item.get("title","Revizyon")
        with st.expander(f"{_severity_icon(sev)} {i}. {ref} — {title}",expanded=(i<=3)):
            st.markdown(f"**Neden?** {item.get('problem','')}")
            if item.get("current_excerpt"):
                st.markdown("**Sözleşmedeki ifade**")
                st.code(item["current_excerpt"],language=None)
            st.markdown(
                f"<span class='type-badge'>{item.get('revision_type','')}</span>",
                unsafe_allow_html=True,
            )
            st.markdown("**Önerilen revizyon**")
            st.code(item.get("suggested_revision",""),language=None)

    st.caption("Clean v0.1'de OLI Word dosyasına otomatik müdahale etmez. Önce analiz ve revizyon önerisi kalitesini sabitliyoruz.")
