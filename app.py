import streamlit as st
import json
from pathlib import Path

st.set_page_config(page_title='Opus Legal Intelligence', page_icon='⚖️', layout='wide')

RULES = json.loads(Path(__file__).with_name('rules.json').read_text(encoding='utf-8'))

st.title('OPUS LEGAL INTELLIGENCE')
st.caption('Private Legal Intelligence • Prototype v0.1')

modules = st.columns(4)
with modules[0]:
    st.subheader('Sözleşmeler')
    st.write('Analyse & Negotiate')
with modules[1]:
    st.subheader('Arabuluculuk')
    st.write('Yakında')
with modules[2]:
    st.subheader('KVKK')
    st.write('Yakında')
with modules[3]:
    st.subheader('Dava Dosyaları')
    st.write('Yakında')

st.divider()
st.header('Yeni Sözleşme Analizi')

c1,c2,c3 = st.columns(3)
with c1:
    contract_type = st.selectbox('Sözleşme türü', ['Oyuncu Sözleşmesi'])
with c2:
    project_type = st.selectbox('Proje türü', ['Ana Akım TV', 'Dijital', 'Sinema'])
with c3:
    negotiation_power = st.select_slider('Pazarlık gücü', ['Düşük','Orta','Yüksek','Çok Yüksek'], value='Orta')

uploaded = st.file_uploader('Sözleşmeyi yükle', type=['docx','pdf'])

if project_type != 'Ana Akım TV':
    st.info('Bu prototipte aktif hukuk profili: ACTOR_TV_MAINSTREAM. Dijital ve Sinema profilleri sonraki sürümlerde eklenecek.')

if uploaded:
    st.success(f'{uploaded.name} yüklendi.')
    st.warning('v0.1 arayüz prototipi: belge metni/AI analizi bir sonraki teknik adımda bağlanacak.')

st.divider()
st.subheader('Aktif OLI Rule Library')
st.metric('Doğrulanmış kontrol noktası', len(RULES))

for r in RULES:
    with st.expander(f"{r['id']} — {r['title']} · {r['priority']}"):
        st.write(r['standard'])
        if r.get('negotiation_sensitive'):
            st.caption('Pazarlık gücüne duyarlı')
