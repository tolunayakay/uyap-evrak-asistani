# ⚖️ UYAP Evrak Asistanı

UYAP kullanıcıları, avukatlar ve hukuk çalışanları için geliştirilmiş; Word, PDF ve UDF formatındaki belgelerin düzenlenmesini, dönüştürülmesini ve sisteme uygun hale getirilmesini kolaylaştıran masaüstü uygulaması.

---

<img width="1022" height="730" alt="Ekran görüntüsü 2026-10-13 203444" src="https://github.com/user-attachments/assets/a46de8b8-0680-47e4-8413-1c724d529233" />



## 🚀 Özellikler

* **Belge Düzenleme & Dönüştürme:** Word (`.docx`) ve PDF belgeleri üzerinde pratik işlemler.
* **Sürükle & Bırak Desteği:** Dosyaları kolayca arayüze sürükleyip işleme alma (`tkinterdnd2`).
* **Hızlı ve Güvenli:** Belgeler tamamen yerel makinede işlenir, harici sunuculara veri aktarılmaz.

---

## 🛠️ Kurulum

Projeyi bilgisayarınızda yerel olarak çalıştırmak için adımları takip edin:

### 1. Depoyu İndirin (Clone)
```bash
git clone https://github.com/tolunayakay/uyap-evrak-asistani.git
cd uyap-evrak-asistani
```

### 2. Gerekli Kütüphaneleri Yükleyin
```bash
pip install -r requirements.txt
```

### 3. Uygulamayı Başlatın
```bash
python uyap_asistan.py
```

---

## 📦 Kullanılan Başlıca Kütüphaneler

* **PyMuPDF (`fitz`):** Yüksek performanslı PDF işleme ve dönüştürme.
* **python-docx:** Word belgeleri oluşturma ve düzenleme.
* **Pillow (`PIL`):** Görsel formatlandırma ve optimizasyon.
* **tkinterdnd2:** Tkinter arayüzü için sürükle-bırak entegrasyonu.

---

## 📄 Lisans

Bu proje **GNU General Public License v3.0 (GPL-3.0)** altında lisanslanmıştır. Detaylar için [LICENSE](LICENSE) dosyasına göz atabilirsiniz.
