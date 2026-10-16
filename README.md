# ⚖️ UYAP Evrak Asistanı (PRO)

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

**UYAP Evrak Asistanı**, avukatlar, hukuk büroları ve adliye personeli için tasarlanmış modern, hızlı ve yerel çalışan bir evrak yönetim ve optimizasyon masaüstü uygulamasıdır. UYAP dosya boyutu kısıtlamalarını (MB limitleri), UDF format dönüştürme ihtiyaçlarını ve KVKK kapsamındaki veri gizleme gereksinimlerini tek bir modern arayüzde çözer.

Tüm işlemler **çevrimdışı (offline)** yürütülür; gizli müvekkil ve dava dosyalarınız üçüncü taraf sunuculara gönderilmez.

---

## ✨ Temel Özellikler

### 1. 📑 PDF & Görsel Birleştirme & Sıkıştırma
* **Çoklu Format Desteği:** PDF, JPG, PNG, WEBP, BMP dosyalarını sürükle-bırak yöntemiyle sıraya dizme ve birleştirme.
* **Akıllı A4 Uyarlama:** Eklenen görseller en-boy oranı korunarak otomatik A4 sayfasına yerleştirilir.
* **İki Farklı Sıkıştırma Modu:**
  * **Akıllı Mod (Vektörel):** Metin netliğini korur, gömülü görselleri kayıpsız optimize eder.
  * **Agresif Raster (150 DPI):** Taranmış ağır evrakları 150 DPI JPEG kalitesine indirgeyerek dosya boyutunu küçültür.
* **Bates Numaralandırma:** Evrak altlarına `EK-1 | Sayfa 1 / 10` formatında numaralandırma ve ek etiketi basma.
* **UYAP Akıllı MB Bölücüsü:** UYAP yükleme sınırına (varsayılan 9.5 MB) göre PDF'leri sayfa bütünlüğünü bozmadan otomatik parçalara ayırma.

### 2. ✂️ Sayfa Yönetimi
* **Sayfa Döndürme:** Belirli sayfaları veya tüm belgeyi $90^\circ$, $180^\circ$ veya $270^\circ$ döndürme.
* **Akıllı Boş Sayfa Temizleme:** Çift taraflı tarayıcılardan kaynaklanan beyaz/boş sayfaları histogram analiziyle tespit edip onay alarak temizler.
* **Sayfa Sayfa Parçalama:** Her sayfayı bağımsız bir PDF dosyası olarak kaydetme.
* **Aralık Ayırma:** Belirli sayfaları (`1-3, 5` formatında) seçip yeni bir PDF'e çıkartma.

### 3. 🔄 Format Dönüştürücü
* **UDF ➔ Word (.docx):** UYAP Doküman Editörü (.udf) dosyalarının içindeki XML yapısını çözüp doğrudan Word belgesine çevirme.
* **Word (.docx) ➔ PDF:** Hazırlanan dilekçeleri Microsoft Word motoru/docx2pdf aracılığıyla tek tıkla resmi PDF formatına dönüştürme.
* **Metin Önizleme ve Kopyalama:** Taranmamış PDF, Word veya UDF belgelerindeki metinleri ayrıştırıp önizleme ve panoya kopyalama.

### 4. 🛡️ Hukuki Güvenlik & KVKK Modülü
* **"Aslı Gibidir" Dijital Kaşesi:** Belgenin yönüne duyarlı olarak sağ/sol/orta alt veya üst köşelere avukat sicil, isim ve tarih bilgisiyle dijital kaşe uygulama.
* **KVKK Akıllı Karartma (Kalıcı Redaksiyon):**
  * **T.C. Kimlik Algoritması:** Yalnızca resmi 11 haneli matematiksel doğrulama algoritmasından geçen gerçek T.C. kimlik numaralarını tespit eder ve üzerini kalıcı olarak siyah bantla çizer.
  * **IBAN Tespiti:** `TR` ile başlayan IBAN numaralarını otomatik bulur ve maskeler.
  * **Telefon Maskeleme:** Mobil hat formatlarını (05xx / +90) tespit edip gizler.
  * **Özel Kelime Maskeleme:** Kullanıcının belirteceği isim, adres vb. kelimeleri sayfa içinde arayıp kalıcı olarak redakte eder.
  * **Metaveri Kazıma (`doc.scrub`):** Maskelenen dosyanın arkasındaki gizli katmanları, metadata bilgilerini ve ekleri tamamen temizler (metin geri döndürülemez).

---

## 🛠️ Gereksinimler ve Kurulum

### Sistem Gereksinimleri
* **İşletim Sistemi:** Windows 10/11 (macOS ve Linux üzerinde temel özellikler çalışır; Word dönüştürme Windows Word gerektirir)
* **Python:** Python 3.8 veya üzeri

### Kurulum Adımları

1. Depoyu klonlayın veya indirin:
   ```bash
   git clone https://github.com/kullanici-adi/uyap-evrak-asistani.git
   cd uyap-evrak-asistani
   ```

2. Gerekli Python kütüphanelerini yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

3. Uygulamayı çalıştırın:
   ```bash
   python uyap_asistan.py
   ```

---

## 📦 Bağımlılıklar (`requirements.txt`)

```text
customtkinter>=5.2.0
PyMuPDF>=1.23.0
Pillow>=10.0.0
python-docx>=1.0.0
docx2pdf>=0.1.8
tkinterdnd2>=0.3.0
```

> **İpucu (Windows Word Dönüşümü):** Word (.docx) belgelerini PDF'e dönüştürmek için bilgisayarınızda Microsoft Word yüklü olmalıdır.

---

## 🚀 Tek Dosya Halinde Çalıştırılabilir Dosya (EXE) Yapma

Uygulamayı Python yüklü olmayan bilgisayarlarda çalıştırmak için PyInstaller ile derleyebilirsiniz:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onedir --windowed --add-data "logo.ico;." uyap_asistan.py
```

*(Eğer tek dosya `.exe` istiyorsanız `--onedir` yerine `--onefile` kullanabilirsiniz.)*

---

## 🔒 Güvenlik ve Gizlilik

* **Tamamen Yerel:** Belgeler harici bir API'ye veya buluta yüklenmez; tüm kütüphaneler yerel bellek ve disk üzerinde çalışır.
* **Gerçek Redaksiyon:** KVKK karartması bir görsel katman eklemekle kalmaz; PyMuPDF'in `apply_redactions` ve `scrub` işlevlerini kullanarak veriyi PDF nesne ağacından kalıcı olarak yok eder.

---

## 📄 Lisans

Bu proje [MIT Lisansı](LICENSE) kapsamında lisanslanmıştır.