# OLI — Sözleşmeler v0.5 Final Patch

Bu paket mevcut Arabuluculuk v1.3.1'i korur ve Sözleşmeler modülüne son konuşulan geliştirmeleri ekler.

## Sözleşmeler
- İlk Not / Ajans Notu: yalnız dosyaya özgü analiz girdisi.
- İlk İnceleme Notu: ajansa gönderilebilir kısa müdahale özeti.
- Word redline yazarı: Av. Onur Güneş.
- Word redline saati: Europe/Istanbul.
- Revize dosya adı: `<orijinal ad> - GG.AA.YYYY REVİZE.docx`.
- Mevcut font ve puntoyu mümkün olduğunca korur; yeni gövde metnini zorla kalın yapmaz.
- Revizyonlar mümkün olduğunca minimal token/cümlecik düzeyinde Track Changes ile işlenir.
- Eksik koruyucu hükümler uygun bölüme APPEND_AFTER ile eklenebilir.
- 🟨 Sarı: doldurulacak/boş alan.
- 🟧 Turuncu: risk/dikkat.
- 🟦 Mavi: Rule Library ile anlamsal olarak eşleşmeyen yeni/öğrenilmemiş hüküm. Sadece farklı ifade edilmiş mevcut konu mavi sayılmaz.
- Karşı taraf dönüşü: bizim revize Word ile dönen Word'ü ACCEPTED / PARTIAL / REJECTED / NEW olarak sınıflandırır.
- Revizyon Dönüş Notu: kısa ajans bilgilendirme dili.
- Ajans Geri Bildirimi: dosyaya özgü talimat alanı.
- Nihai Revizyondan Öğren: mevcut kontrollü öğrenme akışı korunur.
- Madde Bankası mevcut haliyle korunur; otomatik öğrenme yapılmaz.

## Not
Word turuncu işaretlemede Word'ün yerleşik highlight paletindeki en yakın renk olan darkYellow kullanılır; mavi için cyan kullanılır.
