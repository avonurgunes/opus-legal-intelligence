# OLI v0.6.1 — Contrast + Real Micro Engine + Kısa Notlar

Bu sürüm üç gerçek bugı düzeltir:

1. Sayfanın altındaki silik görünüm:
   Eski CSS override'larının opacity/text color etkisi son katmanda iptal edildi.
   Karşı Taraf Dönüşü, uploader, footer ve expander metinleri okunur hale getirildi.

2. Kısa Notlar:
   Analiz sırasında `build_initial_review_note()` otomatik çalışır.
   Sonuç altında kısa ajans/inceleme notları tekrar gösterilir.

3. Gerçek MICRO düzeltme:
   Önceki promptta çelişkili biçimde "mevcut hüküm varsa REPLACE_PARAGRAPH" yazıyordu.
   Bu kaldırıldı.
   Şema artık MICRO | PHRASE | REPLACE_PARAGRAPH destekler.
   MICRO/PHRASE için model yalnız değişecek parçayı döndürür.
   Word motoru bu parçayı mevcut paragrafın içinde lokal olarak değiştirir.
   REPLACE_PARAGRAPH artık teknik olarak BLOCK kabul edilir; son çaredir.

Word işaret renkleri de daha açık pastel tonlarda tutulur.
