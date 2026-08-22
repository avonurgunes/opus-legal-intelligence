# OLI v0.4

v0.4 patch:
- Word Track Changes yazarı: **Av. Onur Güneş**
- Eksik koruyucu hükümler artık NOT_FOUND'da bırakılmaz; konu bakımından uygun bölüme yeni madde olarak eklenir.
- Revize edilen metin, kaynak paragrafın run biçimini (font, punto, bold/italic vb.) miras alır.
- Yeni eklenen paragraf, yerleştirildiği komşu paragrafın paragraf ve yazı biçimini miras alır.
- v0.3 Revision Library ve Word redline akışı korunur.

## GitHub
ZIP içindeki dosyaları repository köküne yükleyip mevcutların üzerine yazın:
- app.py
- revision_engine.py
- revision_library.json
- rules.json
- requirements.txt

Commit sonrası Streamlit yeniden deploy olur.
