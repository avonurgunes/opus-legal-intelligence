# OLI Sözleşmeler v0.5.5 — Learning Engine v1

## Amaç
Nihai revizyonlardan kontrollü biçimde öğrenmeye başlamak.

## İki Word ile geçmiş öğrenme
- İlk gelen sözleşme (.docx)
- Av. Onur Güneş'in nihai/revize sözleşmesi (.docx)
- Sistem değişen paragrafları ve yeni eklemeleri eşleştirir.
- MICRO / PHRASE / BLOCK / NEW_CLAUSE drafting stili çıkarır.
- Her fark önce Öğrenme Adayı olur.
- Kullanıcı açıkça `✓ Öğren` demeden hiçbir kayıt kalıcı hafızaya girmez.
- Dosyaya özgü tavizler `FILE_ONLY` seçilebilir ve yeni sözleşmelerde genel emsal olarak kullanılmaz.

## Yeni sözleşmede kullanım
Onaylı geçmiş öğrenimler analiz/drafting promptuna bağlanır.
Bunlar kalıp cümle olarak değil, drafting davranışı ve tercih örneği olarak kullanılır.
Madde Bankası cümlesini körlemesine yapıştırmak yerine gelen sözleşmenin dili korunur.

## Güven
Aynı konu onaylı örneklerde tekrar ettikçe precedent_count ve confidence artar.
Rule Library otomatik değiştirilmez.

## ÖNEMLİ — kalıcılık
Bu v1 `learning_memory.json` kullanır. Lokal/GitHub çalışma kopyasında kalıcıdır; Streamlit Cloud runtime dosya sistemi deploy/restart sonrası kalıcılık garantisi vermez.
Bu nedenle bu sürüm öğrenme davranışını ve onay akışını kurar; gerçek production kalıcılığı için sonraki adım harici persistent storage (örn. Supabase/Postgres) bağlamaktır.

## Korunanlar
v0.5.4 Hybrid Redline ve çalışan Word üretim motoru değiştirilmeden korunmuştur.
