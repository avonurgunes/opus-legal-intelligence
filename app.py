from __future__ import annotations

import streamlit as st

from oli_contracts_ui import render_contracts
from mediation import render_mediation


st.set_page_config(
    page_title="OLI — Opus Legal Intelligence",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
:root{color-scheme:light}
html,body,[data-testid="stAppViewContainer"],.stApp{
    background:#F4F5F7!important;color:#171A1F!important;
}
[data-testid="stHeader"]{
    background:#0B1118!important;border-bottom:1px solid #202A35!important;
}
[data-testid="stAppViewBlockContainer"]{
    max-width:1120px!important;padding-top:1.2rem!important;padding-bottom:3rem!important;
}
h1,h2,h3{color:#15181C!important;letter-spacing:-.025em!important}
p,label,[data-testid="stCaptionContainer"]{color:#59636D!important}
.top-brand{display:flex;align-items:center;gap:12px;padding:.4rem 0 .8rem}
.logo{font-weight:900;font-size:1.65rem;color:#D6A13A;letter-spacing:.06em}
.brand-sub{font-size:.75rem;color:#87919C;line-height:1.15}
.work-card{
    background:#FFF;border:1px solid #DDE2E7;border-radius:16px;
    box-shadow:0 5px 20px rgba(16,24,40,.05);padding:22px 24px;margin:14px 0 18px;
}
.intro-card{display:flex;justify-content:space-between;align-items:flex-start}
.intro-card h1{font-size:1.55rem;margin:.15rem 0 .25rem}
.intro-card p{margin:0;max-width:700px}
.eyebrow{font-size:.72rem;font-weight:800;color:#A97924;letter-spacing:.12em}
.status-pill{
    border:1px solid #E4D4B4;background:#FCF8EF;color:#8A641E;
    padding:6px 10px;border-radius:999px;font-size:.75rem;font-weight:800;
}
.summary-card{
    background:#FFF;border:1px solid #DDE2E7;border-left:4px solid #D6A13A;
    border-radius:11px;padding:14px 16px;min-height:86px;color:#303840;
}
.type-badge{
    display:inline-block;background:#F4EFE4;color:#75551A;border:1px solid #E3D3B4;
    border-radius:999px;padding:3px 8px;font-size:.75rem;font-weight:800;margin-bottom:8px;
}
div[data-testid="stVerticalBlockBorderWrapper"]{
    background:#FFF!important;border-color:#DDE2E7!important;border-radius:14px!important;
    box-shadow:0 2px 12px rgba(16,24,40,.035)!important;
}
div[data-testid="stFileUploaderDropzone"]{
    background:#FBFCFD!important;border:1px dashed #B9C2CB!important;border-radius:11px!important;
}
div[data-testid="stTextArea"] textarea,
div[data-testid="stSelectbox"] div[data-baseweb="select"]>div{
    background:#FFF!important;color:#1F252B!important;border-color:#C9D0D7!important;
}
.stButton>button{
    min-height:46px!important;border-radius:9px!important;font-weight:800!important;
}
.stButton>button[kind="primary"]{
    background:linear-gradient(90deg,#B97E18,#DCA23A)!important;color:#FFF!important;border:none!important;
}
div[data-testid="stExpander"] details{
    background:#FFF!important;border:1px solid #DDE2E7!important;border-radius:11px!important;
}
.nav-wrap div[role="radiogroup"]{justify-content:flex-end!important}
</style>
""",unsafe_allow_html=True)

b1,b2=st.columns([1,3.5])
with b1:
    st.markdown(
        '<div class="top-brand"><div class="logo">OLI</div><div class="brand-sub">OPUS LEGAL<br>INTELLIGENCE</div></div>',
        unsafe_allow_html=True,
    )
with b2:
    module=st.radio(
        "Modül",
        ["Sözleşmeler","Arabuluculuk","KVKK","Dava Dosyaları"],
        horizontal=True,
        label_visibility="collapsed",
    )

if module=="Sözleşmeler":
    render_contracts()
elif module=="Arabuluculuk":
    render_mediation()
elif module=="KVKK":
    st.markdown('<div class="work-card"><div class="eyebrow">KVKK</div><h1>KVKK</h1><p>Bu modül sonraki aşamada geliştirilecek.</p></div>',unsafe_allow_html=True)
else:
    st.markdown('<div class="work-card"><div class="eyebrow">DAVA DOSYALARI</div><h1>Dava Dosyaları</h1><p>Bu modül sonraki aşamada geliştirilecek.</p></div>',unsafe_allow_html=True)

st.markdown("---")
st.caption("OLI Clean v0.2.4 • Sözleşmeler sıfırdan • Arabuluculuk v1.3.1 korunmuştur.")
