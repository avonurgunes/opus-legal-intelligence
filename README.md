# OLI Sözleşmeler v0.5.6 — Batch Learning

v0.5.5 Learning Engine üzerine Toplu Öğrenme eklenmiştir.

Akış:
1. Birden fazla ham DOCX yükle.
2. Birden fazla Av. Onur Güneş revizeli DOCX yükle.
3. OLI dosya adı + içerik benzerliğiyle çiftleri eşleştirir.
4. Güçlü eşleşmeler yeşil, kontrol gerektirenler sarı gösterilir.
5. Tüm çiftlerden değişiklik adayları çıkarılır.
6. Adaylar konu + drafting stili (MICRO/PHRASE/BLOCK/NEW_CLAUSE) bazında kümelenir.
7. Kullanıcı tek tek yüzlerce değişiklik yerine kümeleri onaylayabilir.
8. Yalnız açıkça onaylanan kümelerin tekil emsalleri Learning Memory'ye girer.
9. FILE_ONLY kayıtlar genel drafting emsali olarak kullanılmaz.

Not: v0.5.6 da learning_memory.json kullanır. Streamlit Cloud deploy/restart kalıcılığı garanti etmez. Production kalıcılığı için harici veritabanı sonraki adımdır.

v0.5.4 Word/Hybrid Redline motoruna dokunulmamıştır.
