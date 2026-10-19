# ⚖️ UYAP Evrak Asistanı (PRO)

[![Sürüm](https://img.shields.io/badge/s%C3%BCr%C3%BCm-v2.0.0-blue.svg)](https://github.com/tolunayakay/uyap-evrak-asistani/releases)
[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![Yerel AI](https://img.shields.io/badge/AI-Gemma%204%20E2B%20%28yerel%29-orange.svg)](https://huggingface.co/google/gemma-4-E2B-it)
[![License](https://img.shields.io/badge/license-GPL--3.0-green.svg)](LICENSE)

**UYAP Evrak Asistanı**, avukatlar, hukuk büroları ve adliye personeli için tasarlanmış modern, hızlı ve yerel çalışan bir evrak yönetim masaüstü uygulamasıdır. UYAP dosya boyutu sınırlarını, UDF dönüştürme ihtiyaçlarını, KVKK kapsamındaki veri gizleme gereksinimlerini ve evrak özetlemeyi tek bir arayüzde çözer.

**v2.0.0 ile birlikte** uygulama, evrakları **bilgisayarınızda yerel olarak çalışan bir yapay zekâ** ile özetleyebiliyor. Belgeleriniz hiçbir zaman internete gönderilmez.

<img width="1317" height="836" alt="UYAP Evrak Asistanı" src="https://github.com/user-attachments/assets/c7f01b9b-0492-4ac6-81f0-9d6e46266532" />

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
* **Sessiz Yapılandırma:** Tercihler, çalışma klasöründe dosya kalabalığı yaratmadan Windows'un yerel uygulama veri alanında (`%APPDATA%\UyapAsistani`) saklanır.

---

## 🚀 Sürüm Geçmişi

### 🆕 v2.0.0 Yenilikleri
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

1. Sağ taraftaki **Releases** sekmesine gidin.
2. En güncel sürümün altındaki **Assets** bölümünden `UYAP Evrak Asistanı.exe` dosyasını indirin.
3. Çift tıklayarak doğrudan kullanmaya başlayın.

> İlk açılışta Windows SmartScreen "Windows bilgisayarınızı korudu" uyarısı gösterebilir. **Ek bilgi ➔ Yine de çalıştır** ile açabilirsiniz.

### Sistem Gereksinimleri
| | Gereksinim |
|---|---|
| **İşletim sistemi** | Windows 10 / 11 (64 bit) |
| **Bellek (RAM)** | 4 GB; AI Evrak Özeti için en az 8 GB önerilir |
| **Disk** | ~85 MB (uygulama) + ~3,1 GB (isteğe bağlı AI modeli) |
| **Ekran kartı** | Gerekli değil. AI özetini hızlandırmak için 4 GB+ belleğe sahip harici bir kart (NVIDIA / AMD / Intel, Vulkan destekli güncel sürücü) kullanılabilir |
| **Microsoft Word** | Yalnızca *Word ➔ PDF* dönüştürme için gereklidir |

**İnternetsiz kullanım:** AI modelini başka bir bilgisayardan kopyalamak için `gemma-4-E2B-it-Q4_K_M.gguf` dosyasını hedef bilgisayarda `%APPDATA%\UyapAsistani\models\` klasörüne koymanız yeterlidir.

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
