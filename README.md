# OLI Sözleşmeler v0.5.4 — Hybrid Redline

## Word bug
`next_change_id is not defined` düzeltildi.
Word revision id'leri artık belgedeki mevcut tracked changes taranarak güvenli şekilde üretilir.

## Yeni redline mantığı
Dört uygulama seviyesi vardır:
- AUTO
- MICRO: birkaç kelimelik gerçek küçük değişiklik.
- PHRASE: mevcut cümle yapısı korunur, değişen ifade/cümlecik tek blok silinir ve yenisi tek blok eklenir.
- BLOCK: hüküm gerçekten baştan yazılıyorsa tam blok değişimi.

AUTO küçük değişiklikte MICRO, orta değişiklikte PHRASE, esaslı yeniden yazımda BLOCK seçer.

## Drafting ilkesi
Madde Bankası kalıp cümle dayatmaz.
OLI, gelen sözleşmenin mevcut terminolojisini ve cümle yapısını koruyarak mümkün olan en küçük doğal hukuki müdahaleyi yapar.
Örn. "başka projede görev almayacaktır" → gerekiyorsa yalnız "görev alabilir" gibi lokal değişiklik.

## Test
Paketleme öncesinde:
- app.py / revision_engine.py / mediation.py syntax kontrolü,
- gerçek DOCX üzerinde MICRO/PHRASE,
- mavi/turuncu tracked formatting,
- oluşan DOCX'in tekrar açılabilmesi
test edilir.
