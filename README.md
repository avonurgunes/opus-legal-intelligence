# OLI Clean v0.1.2 — No Package Conflict

Bu sürüm Streamlit Cloud'daki `contracts.py` isim çakışmasını ortadan kaldırır.

Sorun:
Repo'da eski sürümlerden kalan `contracts.py`, Python'ın `contracts/` paketini gölgeliyordu.
Bu nedenle `from contracts.analysis ...` importu hata veriyordu.

Çözüm:
Sözleşme modülü artık benzersiz root-level dosya adları kullanır:
- oli_contracts_ui.py
- oli_contract_analysis.py
- oli_contract_profiles.py
- oli_actor_tv_rules.json

`app.py` yalnız `oli_contracts_ui` üzerinden sözleşme ekranını çağırır.
Eski `contracts.py` dosyası repo'da kalsa bile artık bu import zincirini bozamaz.

Arabuluculuk korunmuştur.
