# OLI Sözleşmeler v0.5.6.1 — Batch Learning Fix

## Düzeltmeler
- "1 çift bulundu" ifadesi artık "1 sözleşme çifti eşleştirildi" olarak gösterilir.
- Eşleştirme ile revizyon sayısı birbirinden ayrıldı.
- Yüklenen DOCX byte'ları session_state içinde kalıcı kopyaya çevrilir; ikinci buton Streamlit rerun sonrasında dosyaları kaybetmez.
- "Tüm Çiftlerden Öğrenme Özeti Çıkar" artık doğrudan session_state'teki eşleşmiş dosya çiftlerini işler.
- Analiz sonrası toplam revizyon/öğrenme adayı sayısı gösterilir.
- MICRO / PHRASE / BLOCK+YENİ dağılımı metrik olarak gösterilir.
- Küme onay/kaydet akışı korunur.

## Beklenen akış
1. Ham + revize Word'leri yükle.
2. Sözleşmeleri Eşleştir.
3. "1 sözleşme çifti" vb. sonucu gör.
4. Tüm Çiftlerden Öğrenme Özeti Çıkar.
5. Örn. "34 revizyon/öğrenme adayı bulundu" sonucunu gör.
6. Kümeleri onayla.
