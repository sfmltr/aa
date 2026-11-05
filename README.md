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
| Orta | Rapor Detayı | Seçili NCR'ın tüm alanları, discrepancy metni ve içinden çıkarılan parça numaraları. Parça numarasına tıklamak onu panoya kopyalar; **Tümünü kopyala** hepsini alt alta kopyalar. |
| Orta | E-postalar | Outlook'ta NCR NO geçen mailler. "Approved", "Concurred" vb. geçenler **yeşil**, red kelimeleri geçenler kırmızı. Çift tık ile Outlook'ta açılır. Onay maili olan NCR'ların yanında sol listede **✅** görünür. |
| Orta | İlgili Dosyalar | Seçtiğiniz klasörlerdeki .msg / .eml / .pdf / .txt / .docx dosyalarında discrepancy'deki parça numaraları aranır. Eşleşme varsa sol listede **sarı ⚠** rozeti görünür. Part No kolonu bu aramada **kullanılmaz** (son assembly numarası olduğu için). |
| Sağ üst | CATIA V5 Modelleri | Model klasöründeki dosyalar. NCR seçilince discrepancy'deki parça numaralarını içeren modeller **★ ile en üstte** vurgulanır ("Sadece NCR eşleşenleri" ile yalnız onlar listelenir). Önce **CATIA'yı başlat**, sonra modele çift tıklayın; model açılırken CATIA asla yeniden başlatılmaz. |
| Sağ orta | Hızlı Klasörler | Seçtiğiniz klasörler (ağ yolları, ör. `\\vds\...` dahil); çift tık ile Explorer'da açılır. |
| Sağ alt | Tekrarlayan NCR | Her alt klasördeki not dosyası (.txt) discrepancy ile karşılaştırılır. Eşleşme varsa **REPETITIVE!** uyarısı ve klasör listesi gelir; sol listede 🔁 rozeti görünür. |

Sol listedeki rozetler iki kutuda da aynıdır: **⚠** ilgili mail/PDF dosyası, **🔁** tekrarlayan NCR, **✅** onay maili.
Kutu başlığındaki sayaç her rozetten kaç kayıt olduğunu gösterir.

Kısayollar: **F5** Excel'i yenile, **F6** makro, **F7** dosya indeksi, **F8** notları tara, **F9** mail taraması, **Ctrl+,** ayarlar.

## E-posta araması nasıl çalışır?

1. **Excel listesi yüklenince / yenilenince** listedeki tüm NCR'lar arka planda Outlook'ta aranır.
   NCR'lar 40'lık gruplar halinde tek sorguda aranır (NCR başına ayrı sorgu yapılmaz).
2. Daha önce taranmış NCR'lar için yalnızca **son taramadan sonra gelen** mailler aranır.
   Sonuçlar `%APPDATA%\NCRPanel\cache\mail_cache.json` dosyasında saklanır, program kapanıp açılınca da geçerlidir.
3. **Bir NCR'a tıklayınca** önbellekteki mailler anında gösterilir, ardından sadece o NCR için son
   kontrolden sonra gelen yeni mailler aranır. Bu istek toplu taramanın arasına girer, beklemez.
4. "⟳ Yeni mailler" düğmesi aynı kontrolü elle yapar; sağ tıkla **tüm geçmişi yeniden tarayabilirsiniz**.

## Outlook / CATIA donmalarına karşı alınan önlemler

* Tüm Outlook çağrıları **tek bir arka plan iş parçacığından**, sırayla yapılır. Eskiden her tıklama ayrı
  bir iş parçacığında Outlook'a eşzamanlı ağır sorgu gönderiyordu; Outlook'un nesne modeli tek iş
  parçacıklı olduğu için bu Outlook arayüzünü donduruyordu.
* Tarama sırasında düzenli aralıklarla Outlook'a nefes aldırılır, ortak (public) klasörler taranmaz,
  klasör listesi bir kez bulunup saklanır, gereksiz maillerin içeriği yüklenmez (Folder.GetTable).
* Outlook kapalıysa panel onu gizlice **başlatmaz** (Ayarlar'dan açılabilir).
* CATIA'ya yalnızca **çalışan** örneğe bağlanılır. Eski sürümde kullanılan bağlantı yöntemi CATIA
  bulunamazsa ortam ayarları olmadan ikinci bir CATIA açabiliyordu. Zaten açık model tekrar açılmaz,
  penceresi öne getirilir; açma sırasında CATIA'nın modal uyarı pencereleri kapatılır.
* CATIA, kısayolunuzdaki ortam argümanlarıyla (`-env`, `-direnv`) başlatılır: Ayarlar → CATIA'da
  doğrudan masaüstü kısayolunu (.lnk) seçebilir veya argümanları yapıştırabilirsiniz.
* Excel renkleri COM ile okunurken açık Excel'iniz kullanılmaz; dosyanın kopyası ayrı, gizli bir Excel'de okunur.
* Model listesi ve ağ klasörü kontrolleri arka planda yapılır; ağ yavaş olsa da arayüz takılmaz.
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
* **Outlook**: aranacak klasörler (varsayılan Inbox + Sent Items), kaç gün geriye, gövdede arama, onay/red kelimeleri, toplu tarama, tıklamada yeni mail kontrol aralığı, sorgu başına NCR sayısı.
* **Dosya arama**: klasörler, uzantılar, parça no deseni (regex, varsayılan `3 rakam + 1 harf + 4 rakam [+ -ekler]`), kök numara eşleştirme, NCR NO ile de listeleme, açılışta otomatik indeksleme.
* **CATIA**: `CNEXT.exe` veya CATIA kısayolu (.lnk), başlatma argümanları, model klasörü, uzantılar.
* **Hızlı klasörler**, **Tekrarlayan NCR** (kök klasörler, not deseni, benzerlik eşiği, tarama derinliği).
* **Genel**: açık/koyu tema, yazı boyutu, otomatik Excel yenileme, Note kategorileri (anahtar kelime listesi).

## Geliştirme

```bash
pip install -r requirements.txt pytest
python -m pytest tests -q
python -m ncr_panel
```

Kod yapısı: `ncr_panel/config.py` (ayarlar), `excel_data.py` (Excel + renk + makro), `outlook_mail.py`
(mail önbelleği + toplu/artımlı arama), `ui/com_worker.py` (Outlook/CATIA için tek COM iş parçacığı), `ui/mail_service.py`,
`file_search.py` (indeks ve parça no eşleme), `repetitive.py`, `catia.py`, `ui/` (PySide6 arayüz).
İstek metni `docs/PANEL_istek.txt` dosyasındadır.
