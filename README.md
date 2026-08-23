# OLI — Sözleşmeler v0.5.1 Preview Patch

## Bu sürümün odağı
1. Word indirmeden önce gerçek Revizyon Önizleme / Onay paneli.
2. Kelime kelime kırmızı "konfeti"yi azaltan AUTO / MICRO / BLOCK Track Changes motoru.
3. Turuncu ve mavi işaretlerin Word'de tracked formatting change (`w:rPrChange`) olarak kaydedilmesi.

## Önizleme
Her revizyon kartında:
- Kabul / Reddet
- Mevcut/referans metin
- OLI önerisi
- Öneriyi elle düzenleme
- Word uygulama biçimi:
  - Otomatik
  - Mikro değişiklik
  - Cümle/blok değişikliği
  - Yeni hüküm

Word yalnız onaylanan ve önizlemede son hali verilen revizyon planından üretilir.

## Track Changes
AUTO modu yalnız gerçekten küçük/lokal değişikliklerde MICRO kullanır.
Metnin yapısı belirgin değişiyorsa BLOCK kullanır: eski metin tek tracked deletion, yeni metin tek tracked insertion olarak görünür.
Bu, Word'deki kelime kelime dağılmayı ciddi ölçüde azaltmak içindir.

## Renkler
- Sarı: doldurulacak alan (mevcut placeholder sistemi korunur).
- Turuncu: risk/dikkat.
- Mavi: Rule Library ile anlamsal eşleşmeyen yeni hüküm.
Turuncu ve mavi artık `w:rPrChange` ile tracked formatting olarak kaydedilir.

## GitHub
ZIP içindeki 8 dosyanın tamamını repository köküne yükleyip mevcutların üzerine yazın ve Commit changes yapın.
