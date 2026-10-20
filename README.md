# ⚖️ UYAP Evrak Asistanı (PRO)

[![Sürüm](https://img.shields.io/badge/s%C3%BCr%C3%BCm-v2.1.0-blue.svg)](https://github.com/tolunayakay/uyap-evrak-asistani/releases)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![Yerel AI](https://img.shields.io/badge/AI-Gemma%204%20E2B%20%28yerel%29-orange.svg)](https://huggingface.co/google/gemma-4-E2B-it)
[![License](https://img.shields.io/badge/license-GPL--3.0-green.svg)](LICENSE)

**UYAP Evrak Asistanı**, avukatlar, hukuk büroları ve adliye personeli için tasarlanmış modern, hızlı ve yerel çalışan bir evrak yönetim masaüstü uygulamasıdır. UYAP dosya boyutu sınırlarını, UDF dönüştürme ihtiyaçlarını, KVKK kapsamındaki veri gizleme gereksinimlerini ve evrak özetlemeyi tek bir arayüzde çözer.

**v2.0.0 ile birlikte** uygulama, evrakları **bilgisayarınızda yerel olarak çalışan bir yapay zekâ** ile özetleyebiliyor. Belgeleriniz hiçbir zaman internete gönderilmez.

<img width="1471" height="986" alt="1" src="https://github.com/user-attachments/assets/4a49ab53-cea5-490f-b43d-4e6f63306735" />

---

## ✨ Temel Özellikler

### 📑 1. Birleştirme, Sıkıştırma & Ek Listesi
* **Çoklu Format Desteği:** `PDF`, `JPG`, `PNG`, `WEBP` ve `BMP` dosyalarını sürükle-bırak ile listeye ekleme, sıraya dizme ve tek bir PDF'te birleştirme.
* **Akıllı A4 Uyarlama:** Farklı boyut ve çözünürlükteki görseller, en-boy oranı bozulmadan standart A4 sayfaya yerleştirilir.
* **Çift Motorlu Sıkıştırma:**
  * **Akıllı Mod:** Metin katmanlarını ve vektör kalitesini korur; gömülü görselleri optimize eder.
  * **Agresif Raster (150 DPI):** Taranmış ağır evrakları okunabilirliği koruyarak küçültür.
* **Bates Numaralandırma:** Sayfa altına `EK-1 | Sayfa 1 / 10` biçiminde numara basma.
* **Ek Listesi (Fihrist):** <sup>v2</sup> Birleştirilen belgenin **başına veya sonuna**, her ekin adını, sayfa aralığını (`EK-1: Fatura S. 1-3`) ve sayfa sayısını gösteren bir **Ek Listesi** sayfası eklenir. Satırlar tıklanabilir; her ek için PDF yer imi (bookmark) oluşturulur.
* **UYAP Akıllı MB Bölücüsü:** Büyük PDF'leri UYAP yükleme sınırına (varsayılan 9,5 MB) göre, sayfa bütünlüğünü bozmadan parçalara ayırma.

### 🖼️ 2. Görsel Sayfa Önizleme & Yönetim
* **Sayfa Matrisi:** PDF sayfalarını küçük resimler hâlinde inceleme; sürükleyerek sıralama.
* **Döndürme & Silme:** Tekli veya çoklu sayfa seçerek döndürme ya da istenmeyen sayfaları silme; sonucu yeni PDF olarak kaydetme.
* **Boş Sayfa Temizleme:** Çift taraflı taramalardaki boş sayfaları tespit edip kullanıcı onayıyla temizleme.
* **Sayfa Sayfa Bölme & Aralık Çıkarma:** Her sayfayı ayrı dosya yapma veya `1-3, 5` biçiminde aralık seçip yeni belge üretme.

### 🔄 3. Format Dönüştürücü (UDF / Word / PDF)
* **PDF ➔ UDF:** PDF evraklarını metin, tablo ve görselleriyle UYAP Şablon Editörü'nde açılabilen `.udf` biçimine dönüştürme.
* **PDF ➔ Word (.docx):** Sütun yapısını, tabloları ve başlıkları koruyarak düzenlenebilir Word belgesi üretme.
* **Word (.docx) ➔ UDF:** Word dilekçelerini resim ve tablolarıyla UYAP Doküman Formatı'na çevirme.
* **UDF ➔ Word (.docx):** UDF dosyalarını biçimlendirmeyi bozmadan düzenlenebilir Word belgesine çevirme.
* **Word (.docx) ➔ PDF:** Dilekçeleri Microsoft Word altyapısıyla PDF'e dönüştürme.
* **Metin Önizleme:** PDF, Word veya UDF belgelerindeki metni görüntüleme ve tek tıkla panoya kopyalama.

### 🛡️ 4. Araçlar: e-İmza, Kaşe, KVKK & Gizlilik
* **e-İmza & Zaman Damgası Ön Kontrolü:** `.udf` ve `.pdf` evraklardaki CAdES / PAdES imzalarını çözümler. <sup>v2</sup> İmzalayan **kişilerin gerçek adları ve unvanları**, **sertifika sağlayıcısı**, imza zamanı ve **zaman damgası saati** gösterilir; birden fazla imzalayan (ör. hâkim ve zabıt kâtibi) ayrı ayrı listelenir. T.C. kimlik numarası KVKK gereği maskelenir.
* **Durum Rozetleri:** Dosya listesinde imzalı belgelerin yanında `✓ e-İmzalı` ve `⏱ Z.Damgalı` rozetleri.
* **"Aslı Gibidir" Dijital Kaşesi:** Belgenin yönüne duyarlı olarak seçilen köşeye avukat adı, sicil ve tarih bilgisiyle kaşe basma.
* **KVKK Akıllı Karartma (Kalıcı Redaksiyon):**
  * **Algoritmik T.C. Kimlik Tespiti:** Yalnızca resmî doğrulama algoritmasından geçen T.C. kimlik numaralarını kalıcı olarak karartır.
  * **IBAN ve Telefon Maskeleme:** `TR` ile başlayan IBAN'ları ve `05xx` / `+90` telefon numaralarını gizler.
  * **Özel Kelime Redaksiyonu:** Belirttiğiniz isim, unvan, adres vb. ifadeleri kalıcı olarak bantlar.
* **Filigran Bas:** <sup>v2</sup> Sayfaların arkasına 45° açıyla yarı saydam **"GİZLİDİR"**, **"TASLAKTIR"**, **"DELİL NİTELİĞİNDEDİR"** veya elle yazılan herhangi bir metin. Renk, saydamlık ve sayfa aralığı seçilebilir; taranmış sayfalarda filigran görünür kalması için otomatik öne alınır.
* **Metaveri Kazıma:** <sup>v2</sup> Seçili belgede veya listedeki **tüm PDF ve Word belgelerinde** yazar adı, son düzenleyen kullanıcı / bilgisayar adı, şirket, tarihler, düzenleme süresi, şablon yolu, revizyon geçmişi ve değişiklik / yorum yazarlarını tek tıkla temizler. İçerik değişmez; temiz kopya oluşturulabilir veya orijinalin üzerine yazılabilir.

### ✨ 5. AI Evrak Özeti (Yerel Yapay Zekâ) <sup>v2</sup>
* **Tamamen Yerel:** Özetler, bilgisayarınızda çalışan Google **Gemma 4 E2B** modeliyle (Apache 2.0) üretilir. Belge içeriği internete, buluta veya üçüncü taraf bir servise **gönderilmez**.
* **Tek Tıkla Özet:** Listedeki PDF, Word ve UDF evraklarının yanındaki **✨ AI Özet** butonuyla; özet ekrana kelime kelime yazılır ve panoya kopyalanabilir.
* **Belge Türüne Göre Özet:** Sözleşme, dilekçe, mahkeme kararı, ihtarname, bilirkişi raporu ve icra belgeleri otomatik algılanır; özet o türe uygun başlıklarla çıkarılır. Örneğin sözleşmelerde *davacı / davalı* yerine *kiraya veren / kiracı* gibi taraf sıfatları kullanılır. Tür yanlış algılanırsa özet penceresinden değiştirilebilir.
* **Uzun Evraklar:** 30.000 karaktere kadar olan evraklar bütün olarak; daha uzun evraklar giriş, gelişme ve sonuç / hüküm bölümleri birleştirilerek özetlenir.
* **Ekran Kartı Hızlandırma:** Ayarlar'daki **İşlem birimi** (*Otomatik / İşlemci / Ekran Kartı*) ile uygun bir harici ekran kartında özetler **3-5 kat daha hızlı** üretilir. Kart ilk kullanımda doğrulanır; sorun çıkarsa otomatik olarak işlemciye dönülür.
* **Model Yönetimi:** Model (~3,1 GB) Ayarlar'dan bir kez indirilir; indirme yarıda kalırsa kaldığı yerden devam eder. **Modeli Sil** ile kaldırılabilir.

> ⚠️ AI özeti hata içerebilir. Hukuki işlemlerde daima asıl belgeyi esas alın.

### ⚙️ 6. Ayarlar & Tercih Hafızası
* **Tema:** Açık / Koyu tema; seçim hatırlanır.
* **Kaşe ve Filigran Hafızası:** Avukat bilgisi, kaşe unvanı, konum, renk ve filigran tercihleri saklanır.
* **Otomatik Güncelleme Denetimi:** Açılışta GitHub'da yeni sürüm olup olmadığını kontrol eder (kapatılabilir).
* **Sessiz Yapılandırma:** Tercihler, çalışma klasöründe dosya kalabalığı yaratmadan Windows'un yerel uygulama veri alanında (`%APPDATA%\UyapAsistani`), macOS'ta `~/Library/Application Support/UyapAsistani` klasöründe saklanır.

---

## 🚀 Sürüm Geçmişi

### 🆕 v2.1.0: macOS Desteği
* **🍎 macOS Sürümü:** Uygulama artık Apple Silicon (M1 ve sonrası) Mac'lerde de çalışıyor. Windows ve Mac aynı koddan derleniyor; tüm özellikler iki sistemde de kullanılabilir.
* **Mac'te AI Evrak Özeti:** Özetler Mac'in ekran kartında (**Metal**) üretilir; ayrıca bir kurulum gerekmez.
* **Windows:** Davranış ve görünüm değişmedi.

### v2.0.1 Düzeltmeleri
* **Tema Geçişi:** Ayarlar penceresi açıldıktan sonra üst kısımdaki ☀/🌙 anahtarıyla koyu temadan açık temaya dönülememesi düzeltildi.
* **Geliştiriciler İçin:** `pip install -r requirements.txt` artık AI kütüphanesini (llama-cpp-python) derleyici gerektirmeden, hazır paketle kuruyor.

### v2.0.0 Yenilikleri
* **✨ AI Evrak Özeti:** Listedeki PDF, Word ve UDF evrakları, bilgisayarda yerel olarak çalışan **Gemma 4 E2B** yapay zekâ modeliyle özetlenir; belge internete gönderilmez. Model (~3,1 GB) Ayarlar'dan bir kez indirilir.
* **Belge Türüne Göre Özet:** Sözleşme, dilekçe, mahkeme kararı, ihtarname, bilirkişi raporu ve icra belgeleri otomatik algılanır; özet türe uygun başlıklarla çıkarılır (ör. sözleşmelerde davacı/davalı yerine kiraya veren/kiracı).
* **Ekran Kartı Hızlandırma:** Ayarlar'daki *İşlem birimi* (Otomatik / İşlemci / Ekran Kartı) ile uygun bir harici ekran kartında özetler 3-5 kat hızlı üretilir; sorun çıkarsa otomatik olarak işlemciye dönülür.
* **Ek Listesi (Fihrist):** Birleştirilen PDF'in başına veya sonuna, her ekin adını ve sayfa aralığını gösteren tıklanabilir bir Ek Listesi eklenir.
* **Filigran Bas:** Sayfalara 45° açıyla yarı saydam "GİZLİDİR", "TASLAKTIR", "DELİL NİTELİĞİNDEDİR" veya serbest metin filigranı.
* **Metaveri Kazıma:** PDF ve Word belgelerindeki yazar, bilgisayar adı, revizyon geçmişi gibi gizli bilgiler tek tıkla, toplu olarak temizlenir.
* **e-İmza Ön Kontrolü:** İmzalayan kişilerin gerçek adları, unvanları, sertifika sağlayıcısı ve zaman damgası saati gösterilir.
* **Arayüz:** "KVKK & Kaşe" sekmesi **Araçlar** oldu; "Nasıl Kullanılır?" rehberi yenilendi; emeği geçenler Ayarlar'a taşındı.

### v1.0.1 Yenilikleri
* **DOCX ➔ UDF Dönüştürücü:** Hazırlanan Word dilekçelerini doğrudan UYAP Editör formatına (`.udf`) çevirme özelliği eklendi.
* **Görsel Grid Panel:** PDF sayfalarını küçük resim matrisiyle inceleme, döndürme ve silme yeteneği getirildi.
* **e-İmza & Zaman Damgası Modülü:** CAdES/PAdES bloklarını okuyan imza doğrulama ve liste rozetleri eklendi.
* **Kalıcı Ayarlar Altyapısı:** Tema ve kaşe profili tercihleri oturumlar arasında hatırlanacak şekilde güncellendi.
* **Performans İyileştirmeleri:** Evrak birleştirme ve dönüştürme aşamalarındaki thread senkronizasyonu optimize edildi.

---

## 📦 Kurulum & Çalıştırma

Uygulama **taşınabilir (portable)** yapıdadır; kurulum, yönetici yetkisi veya ek program gerektirmez.

**Windows**
1. Sağ taraftaki **Releases** sekmesine gidin.
2. En güncel sürümün altındaki **Assets** bölümünden `UYAP.Evrak.Asistani.exe` dosyasını indirin.
3. Çift tıklayarak doğrudan kullanmaya başlayın.

> İlk açılışta Windows SmartScreen "Windows bilgisayarınızı korudu" uyarısı gösterebilir. **Ek bilgi ➔ Yine de çalıştır** ile açabilirsiniz.

**macOS (Apple Silicon)**
1. **Releases** sekmesinden `UYAP.Evrak.Asistani-macOS.zip` dosyasını indirin ve çift tıklayarak açın.
2. Çıkan **UYAP Evrak Asistanı** uygulamasını **Uygulamalar** klasörüne sürükleyin.
3. Uygulamayı açın.

> İlk açılışta macOS "doğrulanamadı" uyarısı verir ve Word ➔ PDF gibi bazı özellikler ilk kullanımda izin ister. Adım adım anlatım için aşağıdaki **🍎 macOS'ta İlk Kullanım ve İzinler** bölümüne bakın.

### Sistem Gereksinimleri
| | Gereksinim |
|---|---|
| **İşletim sistemi** | Windows 10 / 11 (64 bit) veya macOS 12 ve üzeri (Apple Silicon: M1 ve sonrası) |
| **Bellek (RAM)** | 4 GB; AI Evrak Özeti için en az 8 GB önerilir |
| **Disk** | ~85 MB (uygulama) + ~3,1 GB (isteğe bağlı AI modeli) |
| **Ekran kartı** | Gerekli değil. AI özetini hızlandırmak için 4 GB+ belleğe sahip harici bir kart (NVIDIA / AMD / Intel, Vulkan destekli güncel sürücü) kullanılabilir |
| **Microsoft Word** | Yalnızca *Word ➔ PDF* dönüştürme için gereklidir |

**İnternetsiz kullanım:** AI modelini başka bir bilgisayardan kopyalamak için `gemma-4-E2B-it-Q4_K_M.gguf` dosyasını hedef bilgisayarda Windows'ta `%APPDATA%\UyapAsistani\models\`, Mac'te `~/Library/Application Support/UyapAsistani/models/` klasörüne koymanız yeterlidir.

### 🍎 macOS'ta İlk Kullanım ve İzinler

Uygulama Apple Silicon (M1 ve sonrası) Mac'lerde, macOS 12 ve üzerinde çalışır. AI Evrak Özeti, Mac'in ekran kartını **Metal** ile kullanır; Ayarlar'daki *İşlem birimi* ile işlemciye geçilebilir.

macOS, güvenlik gereği bazı adımlarda sizden onay ister. Aşağıdaki izinlerin her biri **yalnızca bir kez** verilir.

**1. İlk açılış: "Apple doğrulayamadı" uyarısı**

Uygulama Apple'ın ücretli geliştirici imzasını taşımadığı için macOS ilk açılışta *"UYAP Evrak Asistanı" açılamadı / Apple, kötü amaçlı yazılım içermediğini doğrulayamadı* uyarısı gösterir.
1. Uyarıda **Bitti**'ye basın (**Çöp Sepeti'ne Taşı**'ya değil).
2. **Sistem Ayarları ➔ Gizlilik ve Güvenlik** bölümünü açın ve sayfanın altına kaydırın.
3. *"UYAP Evrak Asistanı" engellendi* yazısının yanındaki **Yine de Aç** düğmesine basın.
4. Mac parolanızı veya Touch ID'yi isteyebilir; ardından açılan pencerede tekrar **Yine de Aç**'a basın.

Sonraki açılışlarda bu uyarı çıkmaz.

**2. Word ➔ PDF: Microsoft Word'ü denetleme izni**

Word ➔ PDF dönüştürme, Mac'inizdeki Microsoft Word'ü kullanır (Word yüklü olmalıdır). İlk dönüştürmede macOS *"UYAP Evrak Asistanı", "Microsoft Word" uygulamasını denetlemek için erişim izni istiyor* diye sorar: **İzin Ver**'e basın.

Yanlışlıkla **İzin Verme** dediyseniz: **Sistem Ayarları ➔ Gizlilik ve Güvenlik ➔ Otomasyon ➔ UYAP Evrak Asistanı** altında **Microsoft Word** anahtarını açın ve dönüştürmeyi tekrar deneyin.

**3. Word ➔ PDF: Word'ün "Dosyaya Erişim İzni Ver" penceresi**

Word, kendi güvenlik kuralı gereği bir klasördeki belgelere ilk kez erişirken **Dosyaya Erişim İzni Ver** penceresi açar. Bu pencere Word'e aittir ve arka planda açılabilir; dönüştürme bekliyor gibi görünürse Word penceresine bakın.
1. **Seçin...** düğmesine basın.
2. Açılan pencerede belgenin bulunduğu **klasörü** seçin ve **Erişim İzni Ver**'e basın.

Aynı dönüştürmede bu pencere bir kez daha (PDF'in kaydedileceği yer için) çıkabilir. İzin verilen klasörlerde tekrar sorulmaz. Dönüştürme "Dönüşüm Hatası" ile biterse izni verdikten sonra yeniden deneyin.

**4. Masaüstü, Belgeler ve İndirilenler klasörleri**

macOS, uygulamanın bu klasörlerdeki dosyalara erişmesi için onay isteyebilir: **İzin Ver**'e basın. Sonradan değiştirmek için: **Sistem Ayarları ➔ Gizlilik ve Güvenlik ➔ Dosyalar ve Klasörler ➔ UYAP Evrak Asistanı**.

**Ayarlar, AI modeli ve kaldırma**
* Ayarlar ve AI modeli `~/Library/Application Support/UyapAsistani` klasöründe saklanır (Finder'da **Git ➔ Klasöre Git...** ile bu adresi yazarak açabilirsiniz).
* Uygulamayı kaldırmak için **Uygulamalar** klasöründen Çöp Sepeti'ne taşıyın; AI modeliyle birlikte tüm verileri silmek için yukarıdaki klasörü de silin.

**Bilinen sınırlar**
* Intel işlemcili Mac'ler desteklenmez.
* Uyarı pencerelerindeki düğmeler macOS'ta İngilizce (**OK**, **Yes**, **No**) görünür; dosya açma ve kaydetme pencereleri Türkçedir.

---

## 🛠️ Geliştiriciler İçin

```bash
pip install -r requirements.txt
python uyap_asistan.py
```

**EXE derleme (PyInstaller):**

```bash
python ai_motor_hazirla.py
python -m PyInstaller --noconfirm "UYAP Evrak Asistanı.spec"
```

* `ai_motor_hazirla.py`, ekran kartı desteği için resmî llama.cpp Vulkan DLL'lerini (sürümü sabit, SHA256 doğrulamalı) `ai_motor/` klasörüne indirir. Bu adım atlanırsa EXE yalnızca işlemciyle çalışır.
* EXE'nin sürüm bilgisi, `uyap_asistan.py` içindeki `APP_VERSION` değerinden otomatik okunur.
* Derlenen EXE'nin AI altyapısını doğrulamak için: `"UYAP Evrak Asistanı.exe" --ai-test sonuc.json`

**macOS'ta çalıştırma ve .app derleme:**

Xcode Command Line Tools gerekir (`xcode-select --install`). AI kütüphanesi kaynaktan, Metal desteğiyle derlenir.

```bash
python3 -m venv .venv && source .venv/bin/activate
CMAKE_ARGS="-DGGML_METAL=on" pip install -r requirements-macos.txt
python uyap_asistan.py
```

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm "UYAP Evrak Asistanı (macOS).spec"
```

* Çıktı: `dist/UYAP Evrak Asistanı.app`. AI doğrulaması: `"dist/UYAP Evrak Asistanı.app/Contents/MacOS/UYAP Evrak Asistanı" --ai-test sonuc.json`
* Uygulama imzasız olduğu için ilk açılışta macOS uyarı verir; **Sistem Ayarları ➔ Gizlilik ve Güvenlik ➔ Yine de Aç** ile açılabilir.

**Sürüm yayınlama (otomatik):** `uyap_asistan.py` içindeki `APP_VERSION` güncellenip `v` ile başlayan bir etiket (ör. `v2.1.0`) gönderildiğinde, GitHub Actions Windows `.exe`'yi ve macOS `.app`'i derler, test eder ve ikisini bir **taslak** sürüm olarak Releases'e yükler. Taslak, kontrol edildikten sonra GitHub'da yayınlanır.

```bash
git tag v2.1.0
git push origin v2.1.0
```

---

## 🔒 Gizlilik ve Güvenlik
* Bu yazılım **yerel olarak çalışır.** İnternet yalnızca AI modelinin ilk indirilmesi ve (isteğe bağlı, Ayarlar'dan kapatılabilen) güncelleme kontrolü için kullanılır; bu bağlantılarda belge verisi gönderilmez.
* Belgeleriniz hiçbir uzak sunucuya, bulut servisine veya üçüncü taraf API'ye iletilmez. **AI Evrak Özeti de tamamen bilgisayarınızda çalışır.**
* Tüm PDF işleme, redaksiyon, dönüştürme ve özetleme işlemleri doğrudan yerel bilgisayarınızda gerçekleşir.

---

## 🙏 Emeği Geçenler
* Zafer Kaya
* yanu34

## 📄 Lisans
Bu proje [GNU General Public License v3.0](LICENSE) ile lisanslanmıştır.
