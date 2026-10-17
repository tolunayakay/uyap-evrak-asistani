# ⚖️ UYAP Evrak Asistanı (PRO)

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

**UYAP Evrak Asistanı**, avukatlar, hukuk büroları ve adliye personeli için tasarlanmış modern, hızlı ve yerel çalışan bir evrak yönetim ve optimizasyon masaüstü uygulamasıdır. UYAP dosya boyutu kısıtlamalarını (MB limitleri), UDF format dönüştürme ihtiyaçlarını ve KVKK kapsamındaki veri gizleme gereksinimlerini tek bir modern arayüzde çözer.

<img width="1317" height="836" alt="UYAP Evrak Asistanı" src="https://github.com/user-attachments/assets/c7f01b9b-0492-4ac6-81f0-9d6e46266532" />

---

## ✨ Temel Özellikler

### 📑 1. PDF & Görsel Birleştirme & Sıkıştırma
* **Çoklu Format Desteği:** `PDF`, `JPG`, `PNG`, `WEBP` ve `BMP` formatındaki dosyaları sürükle-bırak yöntemiyle listeye ekleme, sıraya dizme ve tek bir PDF dosyasında birleştirme.
* **Akıllı A4 Uyarlama:** Eklenen farklı çözünürlük ve boyutlardaki görseller, en-boy oranı (aspect ratio) bozulmadan standart A4 sayfasına yerleştirilir.
* **Çift Motorlu Sıkıştırma:**
  * **Akıllı Mod (Vektörel):** Metin tabakalarını ve vektör kalitesini korur; gömülü görselleri kayıpsız optimize eder.
  * **Agresif Raster (150 DPI):** Taranmış ağır evrakları 150 DPI JPEG seviyesine indirgeyerek okunabilirliği korurken dosya boyutunu küçültür.
* **Bates Numaralandırma:** Evrak altbilgisine otomatik olarak `EK-1 | Sayfa 1 / 10` formatında numaralandırma ve ek etiketi basma.
* **UYAP Akıllı MB Bölücüsü:** UYAP yükleme sınırına (varsayılan 9.5 MB) göre büyük PDF'leri sayfa bütünlüğünü bozmadan otomatik parçalara ayırma.

### 🖼️ 2. Görsel Sayfa Önizleme & Yönetim (Grid Panel)
* **Sayfa Matrisi (Küçük Resimler):** PDF belgelerini doğrudan görsel kartlar halinde inceleyebilme.
* **Sayfa Döndürme & Ayıklama:** Tekli veya çoklu sayfa seçerek 90°, 180°, 270° döndürme ya da istenmeyen sayfaları doğrudan belgeden silme.
* **Anında Dışa Aktarma:** Görsel matris üzerinden ayıklanan veya yönü düzeltilen sayfaları yeni bir PDF olarak kaydetme.
* **Boş Sayfa Temizleme:** Çift taraflı tarayıcılardan kaynaklanan boş/arkası görünen sayfaları histogram analiziyle tespit edip kullanıcı onayıyla temizleme.
* **Sayfa Sayfa Parçalama & Aralık Çıkarma:** Her sayfayı bağımsız bir dosya yapma veya `1-3, 5` formatında aralık seçip yeni belge üretme.

### 🔄 3. Format Dönüştürücü (UDF / Word / PDF)
* **Word (.docx) ➔ UDF:** Microsoft Word belgelerini doğrudan UYAP Doküman Formatı (`.udf`) standardına dönüştürme (yazı tipi, stil ve paragraflar korunur).
* **UDF ➔ Word (.docx):** UYAP Doküman Editörü (`.udf`) dosyalarının XML yapısını çözerek biçimlendirmeyi bozmadan düzenlenebilir Word belgesine çevirme.
* **Word (.docx) ➔ PDF:** Dilekçeleri doğrudan Microsoft Word motoru altyapısıyla resmi PDF formatına dönüştürme.
* **Metin Önizleme ve Kopyalama:** Taranmamış PDF, Word veya UDF belgelerindeki metinleri ayıklama, anlık önizleme ve tek tıkla panoya kopyalama.

### 🔏 4. e-İmza & Zaman Damgası (TSA) Analizi
* **CAdES / PAdES Taraması:** `.udf` ve `.pdf` formatındaki evraklarda yer alan nitelikli elektronik imzaları ve zaman damgalarını otomatik tespit etme.
* **Durum Rozetleri:** Dosya listesinde belgelerin yanına anlık durum rozetleri ekleme (`[✓ e-İmzalı]`, `[⏱ Z.Damgalı]`).

### 🛡️ 5. Hukuki Güvenlik & KVKK Modülü
* **"Aslı Gibidir" Dijital Kaşesi:** Belgenin yönüne duyarlı olarak sağ/sol/orta alt veya üst köşelere avukat sicil, isim ve tarih bilgisiyle dijital kaşe uygulama.
* **KVKK Akıllı Karartma (Kalıcı Redaksiyon):**
  * **Algoritmik T.C. Kimlik Tespiti:** Yalnızca resmi 11 haneli matematiksel algoritma kontrolünden geçen gerçek T.C. kimlik numaralarını tespit eder ve üzerini kalıcı olarak çizer.
  * **IBAN Maskeleme:** TR ile başlayan banka IBAN numaralarını otomatik tespit edip karartır.
  * **Telefon Numarası Maskeleme:** Mobil hat formatlarını (`05xx`, `+90`) yakalayarak gizler.
  * **Özel Kelime Redaksiyonu:** Kullanıcının belirteceği isim, unvan, adres vb. özel kelimeleri arayıp kalıcı olarak bantlar.
* **Metaveri Kazıma (doc.scrub):** Karartma yapılan dosyalardaki gizli katmanları, metadata bilgilerini ve ekleri tamamen temizler; karartılan metinlerin geri getirilmesini imkânsız kılar.

### ⚙️ 6. Akıllı Tercih Hafızası
* **Tema Hafızası:** Son seçilen Gece / Gündüz modu (Dark / Light) otomatik kaydedilir.
* **Kaşe Hafızası:** Girilen avukat adı, sicil, kaşe unvanı, konum ve renk seçimleri arka planda saklanır.
* **Sessiz Yapılandırma:** Tercihler çalışma dizininde dosya kalabalığı yaratmadan sistemin güvenli yerel uygulama veri alanında saklanır.

---

## 🚀 Sürüm Geçmişi

### 🆕 v1.0.1 Yenilikleri
* **DOCX ➔ UDF Dönüştürücü:** Hazırlanan Word dilekçelerini doğrudan UYAP Editör formatına (`.udf`) çevirme özelliği eklendi.
* **Görsel Grid Panel:** PDF sayfalarını küçük resim matrisiyle inceleme, döndürme ve silme yeteneği getirildi.
* **e-İmza & Zaman Damgası Modülü:** CAdES/PAdES bloklarını okuyan imza doğrulama ve liste rozetleri eklendi.
* **Kalıcı Ayarlar Altyapısı:** Tema ve kaşe profili tercihleri oturumlar arasında hatırlanacak şekilde güncellendi.
* **Performans İyileştirmeleri:** Evrak birleştirme ve dönüştürme aşamalarındaki thread senkronizasyonu optimize edildi.

---

## 📦 Kurulum & Çalıştırma

Uygulama tamamen **taşınabilir (portable)** yapıdadır; bilgisayarınızda yönetici yetkisi veya ilave bağımlılık gerektirmez.

1. Sağ taraftaki **Releases** sekmesine gidin.
2. En güncel sürümün altındaki **Assets** bölümünden `UYAP Evrak Asistanı.exe` dosyasını indirin.
3. Çift tıklayarak doğrudan kullanmaya başlayın.

---

## 🔒 Gizlilik ve Güvenlik
* Bu yazılım **tamamen çevrim dışı (offline)** çalışır.
* Belgeleriniz hiçbir uzak sunucuya, bulut servisine veya üçüncü taraf API'ye iletilmez.
* Tüm PDF işleme, OCR, redaksiyon ve dönüştürme süreçleri doğrudan yerel bilgisayarınızın işlemcisi üzerinde gerçekleşir.
