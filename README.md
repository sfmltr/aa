# NCR Panel

MRB mühendisleri için Python (PySide6) ile yazılmış, üç kolonlu genel bilgi ekranı.
NCR takip Excel'ini okur, raporları kullanıcıya göre filtreler, ilgili Outlook maillerini ve
klasörlerdeki mail/PDF dosyalarını bulur, CATIA V5 modellerini açar ve tekrarlayan NCR'ları işaretler.

## Kurulum (Windows)

1. Python 3.10+ kurulu olmalı (`py -3 --version`).
2. `run_panel.bat` dosyasına çift tıklayın. İlk çalıştırmada `.venv` oluşturulur ve bağımlılıklar
   (`requirements.txt`) kurulur, sonra panel açılır.

Elle kurulum:

```bat
py -3 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\pythonw run_panel.py
```

Outlook, Excel makrosu ve CATIA entegrasyonu `pywin32` ile COM üzerinden çalışır; bu özellikler
yalnızca Windows'ta aktiftir. Diğer işlevler (Excel okuma, dosya arama, tekrarlayan NCR) her
platformda çalışır.

## İlk açılış

Panel ilk açıldığında **Ayarlar** penceresi gelir. En azından **Excel** sekmesinden NCR listesini
seçin. Ayarlar `%APPDATA%\NCRPanel\config.json` dosyasında saklanır; dosya indeksi önbelleği
`%APPDATA%\NCRPanel\cache\` altındadır.

## Ekran düzeni

| Kolon | Bölüm | İşlev |
|---|---|---|
| Sol üst | MRB-ENG Raporları | H kolonu `MRB-ENG` olan satırlar: NCR NO (D), Part No (J), Discrepancy (G), CN/CSI (Q). Kullanıcı (C kolonu) seçimine göre filtrelenir. Satırlar Excel'deki renkleriyle gelir. |
| Sol alt | EXTERNAL APPROVER | H kolonu `EXTERNAL APPROVER` olan satırlar + Note (A). Note içeriği Consultation / Stress / SNN gibi kategorilere ayrılır. **Makro ile güncelle** düğmesi Excel içindeki `B03_Update_NCRs` makrosunu çalıştırıp listeyi yeniler. |
| Orta | Rapor Detayı | Seçili NCR'ın tüm alanları, discrepancy metni ve içinden çıkarılan parça numaraları. |
| Orta | E-postalar | Outlook'ta NCR NO geçen mailler. "Approved", "Concurred" vb. geçenler **yeşil**, red kelimeleri geçenler kırmızı. Çift tık ile Outlook'ta açılır. |
| Orta | İlgili Dosyalar | Seçtiğiniz klasörlerdeki .msg / .eml / .pdf / .txt / .docx dosyalarında discrepancy'deki parça numaraları aranır. Eşleşme varsa sol listede **sarı ⚠** rozeti görünür. Part No kolonu bu aramada **kullanılmaz** (son assembly numarası olduğu için). |
| Sağ üst | CATIA V5 Modelleri | Model klasöründeki dosyalar; çift tık önce CATIA'yı başlatır (çalışmıyorsa), sonra modeli açar. |
| Sağ orta | Hızlı Klasörler | Seçtiğiniz klasörler; çift tık ile Explorer'da açılır. |
| Sağ alt | Tekrarlayan NCR | Her alt klasördeki not dosyası (.txt) discrepancy ile karşılaştırılır. Eşleşme varsa **REPETITIVE!** uyarısı ve klasör listesi gelir; sol listede 🔁 rozeti görünür. |

Kısayollar: **F5** Excel'i yenile, **F6** makro, **F7** dosya indeksi, **F8** notları tara, **Ctrl+,** ayarlar.
Sağ tık menülerinde kopyalama, klasörde gösterme ve "Excel'de satıra git" seçenekleri vardır.

## Renkler

Satır renkleri Excel'den okunur:

* `openpyxl` (varsayılan): hücrelere elle verilmiş dolgu renkleri (tema/indeksli renkler dahil).
* `Excel COM`: Excel'in gerçekten **görüntülediği** renk; koşullu biçimlendirme ile gelen renkler de
  dahildir (Windows). Sayfada koşullu biçimlendirme varsa `auto` modu otomatik olarak buna geçer.

**Ayarlar → Renkler** sekmesi Excel'de bulunan renkleri sayılarıyla listeler; her renge "Proje X",
"Onaydan döndü" gibi etiket verebilirsiniz. Etiketler listenin üstünde çip olarak ve satır ipuçlarında görünür.

## Ayarlar özeti

* **Excel**: dosya, sayfa, başlık satırı (otomatik tespit), kolon harfleri, durum metinleri, makro adı, renk kaynağı.
* **Outlook**: aranacak klasörler (varsayılan Inbox + Sent Items), kaç gün geriye, gövdede arama, onay/red kelimeleri.
* **Dosya arama**: klasörler, uzantılar, parça no deseni (regex, varsayılan `3 rakam + 1 harf + 4 rakam [+ -ekler]`), kök numara eşleştirme, NCR NO ile de listeleme, açılışta otomatik indeksleme.
* **CATIA**: `CNEXT.exe` yolu, model klasörü, uzantılar, COM ile açma, açılış beklemesi.
* **Hızlı klasörler**, **Tekrarlayan NCR** (kök klasörler, not deseni, benzerlik eşiği, tarama derinliği).
* **Genel**: açık/koyu tema, yazı boyutu, otomatik Excel yenileme, Note kategorileri (anahtar kelime listesi).

## Geliştirme

```bash
pip install -r requirements.txt pytest
python -m pytest tests -q
python -m ncr_panel
```

Kod yapısı: `ncr_panel/config.py` (ayarlar), `excel_data.py` (Excel + renk + makro), `outlook_mail.py`,
`file_search.py` (indeks ve parça no eşleme), `repetitive.py`, `catia.py`, `ui/` (PySide6 arayüz).
İstek metni `docs/PANEL_istek.txt` dosyasındadır.
