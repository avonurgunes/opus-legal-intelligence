# OLI - Arabuluculuk v1.3.1

v1.2 korunmuştur; toplu indirme mantığı düzeltildi.

## Toplu indirme
ZIP yoktur.

Üç arabuluculuk belgesi:
1. Bilgilendirme ve Belirleme Tutanağı
2. Anlaşma Belgesi
3. Son Tutanak

tek bir dosya içinde art arda birleştirilerek indirilebilir:

- TEK WORD → `tum_tutanaklar.docx`
- TEK PDF → `tum_tutanaklar.pdf`
- TEK UDF → `tum_tutanaklar.udf`

Word ve PDF'de her tutanak yeni sayfadan başlar.
Ayrıca her tutanağın ayrı Word/PDF/UDF indirme butonları korunmuştur.

## GitHub
ZIP içindeki 8 dosyanın tamamını repository köküne yükleyip mevcut dosyaların üzerine yazın ve Commit changes yapın.


## v1.3.1
- Streamlit açılışını engelleyen `mediation.py` SyntaxError düzeltildi.
- `app.py` ve `mediation.py` Python derleme kontrolünden geçirildi.
