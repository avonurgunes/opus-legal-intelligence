# OLI - Arabuluculuk v1

Mevcut Sözleşmeler v0.5 korunmuştur ve Arabuluculuk modülü aktif edilmiştir.

## Arabuluculuk v1
- Arabulucu seçimi: Av. Arb. Onur GÜNEŞ / Av. Arb. Serap TAŞDELEN GÜNEŞ
- Arabulucu sicil ve adres bilgileri otomatik
- 1-5 taraf
- Taraf arama / kayıt / otomatik getirme (prototype SQLite)
- Dosya no
- Dosya açılış, anlaşma, son tutanak tarihleri
- Uyuşmazlık: İşçilik Alacağı, Kira Tespit, Kiralananın Tahliyesi, birleşik kira
- İşten çıkış/ayrılış bildirgesi PDF/DOCX üzerinden AI ile taraf bilgisi önerme
- Kira ve işçilik için dinamik anlaşma alanları
- Rakam ve tarihleri boş bırakabilme
- Üç belgenin taslağını ekranda tamamen düzenleme
- Bilgilendirme ve Belirleme Tutanağı: Word + PDF
- Anlaşma Belgesi: Word + PDF
- Son Tutanak: Word + PDF

Şablon metinleri kullanıcının sağladığı gerçek kira ve işçilik arabuluculuk örnekleri esas alınarak kurulmuştur.

## Güvenlik notu
Taraf hafızası şimdilik uygulama dizinindeki SQLite dosyasını kullanır. Community Cloud yeniden deploy/restart süreçlerinde bu kayıtların kalıcılığı garanti değildir ve bu prototip depolama şifreli değildir. Gerçek kalıcı taraf veritabanı, erişim ve şifreleme katmanıyla daha sonra kurulmalıdır.

## GitHub
ZIP içindeki dosyaların tamamını repository köküne yükleyip mevcut dosyaların üzerine yazın ve Commit changes yapın.
