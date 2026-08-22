# OLI v0.5

Bu sürüm v0.4.1 üzerine aşağıdaki geliştirmeleri getirir:

- **Madde Bankası v0.1 entegrasyonu:** Kullanıcının kısmen gözden geçirdiği mevcut 15 hazır cümle `clause_bank.json` olarak motora bağlandı. Madde Bankası, Revision Library'den önce gelir.
- **Minimal Track Changes:** OLI artık mümkün olduğunda paragrafın tamamını silip yeniden yazmak yerine yalnız değişen kelime/cümlecikleri silme-ekleme olarak işler.
- **Daha az AI genişletmesi:** Onaylı Madde Bankası cümlesi varsa AI'nın metni gereksiz ayrıntılandırmaması için prompt sıkılaştırıldı.
- **Bold temizliği:** Yeni/revize edilen gövde metninde kaynak paragraftan gelen bold zorlaması kaldırılır; font ve punto gibi diğer temel biçimler korunur.
- **Türkiye saati:** Track Changes zaman damgası Europe/Istanbul duvar saati olarak yazılır; önceki 3 saatlik kaymayı önlemek için offset XML'e eklenmez.
- v0.4.1'deki tarihli `REVİZE` dosya adı, ek risk taraması, nihai revizyondan öğrenme ve sarı placeholder işaretleme korunur.

## GitHub
ZIP içindeki dosyaların hepsini repository köküne yükleyip mevcutların üzerine yazın ve commit edin.
