import os
import io
import re
import sys
import datetime
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
import pymupdf  # PyMuPDF API
from PIL import Image, ImageTk
import docx

DESTEKLENEN_UZANTILAR = ('.pdf', '.jpg', '.jpeg', '.png', '.webp', '.bmp', '.udf')

# Türkçe karakter destekli sistem fontunu belirle
TURKCE_FONT_YOLU = "C:/Windows/Fonts/arial.ttf" if os.path.exists("C:/Windows/Fonts/arial.ttf") else None


def resource_path(relative_path):
    """PyInstaller --onefile derlemesinde geçici klasördeki kaynakları bulur."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class UyapEvrakAsistaniApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UYAP Evrak Asistanı")
        self.root.geometry("930x630")
        self.root.resizable(False, False)
        self.root.configure(bg="#F4F6F9")

        # Otomatik Logo & İkon Entegrasyonu
        self._uygulama_ikonu_yukle()

        self.dosyalar = []

        # --- ÜST BAŞLIK ALANI ---
        baslik_cerceve = tk.Frame(root, bg="#1B263B", pady=8)
        baslik_cerceve.pack(fill="x")
        
        baslik_ic = tk.Frame(baslik_cerceve, bg="#1B263B")
        baslik_ic.pack(fill="x", padx=16)

        tk.Label(
            baslik_ic,
            text="⚖ UYAP EVRAK ASİSTANI",
            font=("Segoe UI", 15, "bold"),
            fg="#FFFFFF",
            bg="#1B263B"
        ).pack(side="left")

        tk.Label(
            baslik_ic,
            text="by Tol",
            font=("Segoe UI", 9, "italic", "bold"),
            fg="#DDA15E",
            bg="#1B263B"
        ).pack(side="left", padx=(10, 0))

        tk.Button(
            baslik_ic,
            text="❓ Nasıl Kullanılır? (Detaylı Rehber)",
            command=self.rehber_penceresi_ac,
            bg="#3A506B",
            fg="#FFFFFF",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=10,
            pady=3
        ).pack(side="right")

        # --- ANA GÖVDE (İKİ SÜTUN) ---
        govde = tk.Frame(root, bg="#F4F6F9", padx=12, pady=8)
        govde.pack(fill="both", expand=True)

        # SOL SÜTUN: LİSTE
        sol_sutun = tk.Frame(govde, bg="#F4F6F9")
        sol_sutun.pack(side="left", fill="both", expand=True, padx=(0, 10))

        liste_cerceve = tk.LabelFrame(
            sol_sutun,
            text=" 📂 İşlenecek Dosyalar (PDF, Görsel, UDF Sürükleyebilirsiniz) ",
            font=("Segoe UI", 9, "bold"),
            bg="#F4F6F9",
            fg="#1B263B",
            padx=8,
            pady=6
        )
        liste_cerceve.pack(fill="both", expand=True)

        self.dosya_listebox = tk.Listbox(
            liste_cerceve,
            selectmode=tk.SINGLE,
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            relief="solid",
            bd=1
        )
        self.dosya_listebox.pack(side="left", fill="both", expand=True)

        self.dosya_listebox.drop_target_register(DND_FILES)
        self.dosya_listebox.dnd_bind('<<Drop>>', self.dosyalari_surukle_birak)

        kaydirma = tk.Scrollbar(liste_cerceve, orient="vertical", command=self.dosya_listebox.yview)
        kaydirma.pack(side="right", fill="y")
        self.dosya_listebox.config(yscrollcommand=kaydirma.set)

        # Liste Altı Butonları
        liste_btn_cerceve = tk.Frame(sol_sutun, bg="#F4F6F9")
        liste_btn_cerceve.pack(fill="x", pady=(6, 0))

        tk.Button(
            liste_btn_cerceve,
            text="➕ Dosya Ekle",
            command=self.dosya_ekle_dialog,
            bg="#0D6EFD",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=4
        ).pack(side="left", padx=2)

        tk.Button(
            liste_btn_cerceve,
            text="➖ Seçileni Kaldır",
            command=self.secileni_sil,
            bg="#6C757D",
            fg="white",
            font=("Segoe UI", 9),
            relief="flat",
            padx=8,
            pady=4
        ).pack(side="left", padx=2)

        tk.Button(
            liste_btn_cerceve,
            text="🗑 Listeyi Temizle",
            command=self.listeyi_temizle,
            bg="#DC3545",
            fg="white",
            font=("Segoe UI", 9),
            relief="flat",
            padx=8,
            pady=4
        ).pack(side="right", padx=2)

        # SAĞ SÜTUN: AYARLAR VE İŞLEMLER
        sag_sutun = tk.Frame(govde, bg="#F4F6F9", width=380)
        sag_sutun.pack(side="right", fill="y")
        sag_sutun.pack_propagate(False)

        # 1. Sıkıştırma Seviyesi
        ayar_cerceve = tk.LabelFrame(
            sag_sutun,
            text=" ⚙ PDF Kalite & Sıkıştırma ",
            font=("Segoe UI", 8, "bold"),
            bg="#F4F6F9",
            fg="#1B263B",
            padx=8,
            pady=2
        )
        ayar_cerceve.pack(fill="x", pady=(0, 3))

        self.mod_secimi = tk.StringVar(value="standart")
        tk.Radiobutton(
            ayar_cerceve,
            text="Dengeli (150 DPI)",
            variable=self.mod_secimi,
            value="standart",
            bg="#F4F6F9",
            font=("Segoe UI", 8)
        ).pack(side="left", padx=5)

        tk.Radiobutton(
            ayar_cerceve,
            text="Maksimum Sıkıştırma (100 DPI)",
            variable=self.mod_secimi,
            value="maksimum",
            bg="#F4F6F9",
            font=("Segoe UI", 8)
        ).pack(side="left", padx=5)

        # 2. Bates Numaralandırma
        bates_cerceve = tk.LabelFrame(
            sag_sutun,
            text=" 🏷 Ek Etiketi / Bates ",
            font=("Segoe UI", 8, "bold"),
            bg="#F4F6F9",
            fg="#1B263B",
            padx=8,
            pady=2
        )
        bates_cerceve.pack(fill="x", pady=(0, 3))

        self.bates_aktif = tk.BooleanVar(value=False)
        self.bates_check = tk.Checkbutton(
            bates_cerceve,
            text="Sayfa Altına Numara Bas",
            variable=self.bates_aktif,
            command=self._bates_durum_guncelle,
            bg="#F4F6F9",
            font=("Segoe UI", 8, "bold"),
            fg="#0D6EFD"
        )
        self.bates_check.pack(side="left")

        tk.Label(bates_cerceve, text="Etiket:", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left", padx=(10, 2))
        self.bates_on_ek = tk.Entry(bates_cerceve, font=("Segoe UI", 8), width=8)
        self.bates_on_ek.insert(0, "EK-1")
        self.bates_on_ek.config(state="disabled")
        self.bates_on_ek.pack(side="left")

        # 3. İşlem Butonları Paneli
        islem_cerceve = tk.LabelFrame(
            sag_sutun,
            text=" ⚡ İşlemler ",
            font=("Segoe UI", 8, "bold"),
            bg="#F4F6F9",
            fg="#1B263B",
            padx=8,
            pady=3
        )
        islem_cerceve.pack(fill="both", expand=True)

        # Birleştirme & Sıkıştırma Satırı
        birlestir_bar = tk.Frame(islem_cerceve, bg="#F4F6F9")
        birlestir_bar.pack(fill="x", pady=1)

        tk.Button(
            birlestir_bar,
            text="📑 PDF'leri Birleştir",
            command=self.tumunu_birlestir_pdf_yap,
            bg="#198754",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="left", fill="x", expand=True, padx=(0, 2))

        tk.Button(
            birlestir_bar,
            text="🗜 PDF Sıkıştır",
            command=self.pdf_tekil_sikistir,
            bg="#FD7E14",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="right", fill="x", expand=True, padx=(2, 0))

        # Bölme Satırı
        bolme_bar = tk.Frame(islem_cerceve, bg="#F4F6F9")
        bolme_bar.pack(fill="x", pady=1)

        tk.Button(
            bolme_bar,
            text="✂ Sayfa Böl",
            command=self.pdf_sayfa_sayfa_bol,
            bg="#6F42C1",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="left", fill="x", expand=True, padx=(0, 2))

        tk.Button(
            bolme_bar,
            text="🎯 Aralık Ayır",
            command=self.pdf_aralik_bol,
            bg="#495057",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="right", fill="x", expand=True, padx=(2, 0))

        # Akıllı UYAP Boyut Bölücü
        tk.Button(
            islem_cerceve,
            text="⚖ Akıllı UYAP Boyut Bölücüsü (MB Sınırına Göre)",
            command=self.pdf_akilli_boyut_bol,
            bg="#028090",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(fill="x", pady=1)

        # Sayfa Döndürme & Boş Sayfa Temizleme
        don_bos_bar = tk.Frame(islem_cerceve, bg="#F4F6F9")
        don_bos_bar.pack(fill="x", pady=1)

        tk.Button(
            don_bos_bar,
            text="🔄 Sayfa Döndür (90°)",
            command=self.pdf_dondur_penceresi,
            bg="#0582CA",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="left", fill="x", expand=True, padx=(0, 2))

        tk.Button(
            don_bos_bar,
            text="🧹 Boş Sayfaları Temizle",
            command=self.pdf_bos_sayfalari_temizle,
            bg="#588157",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="right", fill="x", expand=True, padx=(2, 0))

        # Aslı Gibidir Kaşesi (Tek başına tam genişlikte buton)
        tk.Button(
            islem_cerceve,
            text="🛡 'Aslı Gibidir' Kaşesi",
            command=self.pdf_kase_ekle_penceresi,
            bg="#C1121F",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(fill="x", pady=1)

        # KVKK Karartma
        tk.Button(
            islem_cerceve,
            text="🕶 KVKK Akıllı Karartma (T.C., IBAN Gizle)",
            command=self.akilli_karartma_penceresi,
            bg="#212529",
            fg="#FFC107",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(fill="x", pady=1)

        # UDF ➔ Word & OCR
        udf_ocr_bar = tk.Frame(islem_cerceve, bg="#F4F6F9")
        udf_ocr_bar.pack(fill="x", pady=1)

        tk.Button(
            udf_ocr_bar,
            text="📄 UDF ➔ Word (.docx)",
            command=self.udf_to_word,
            bg="#157347",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="left", fill="x", expand=True, padx=(0, 2))

        tk.Button(
            udf_ocr_bar,
            text="🔍 Metin Kopyala",
            command=self.metin_oku_penceresi,
            bg="#023E8A",
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            pady=3
        ).pack(side="right", fill="x", expand=True, padx=(2, 0))

        # Durum Çubuğu
        self.durum_etiketi = tk.Label(
            root,
            text="Hazır. PDF, Görsel veya UDF dosyalarını sol alana sürükleyebilirsiniz.",
            font=("Segoe UI", 9),
            bg="#E9ECEF",
            fg="#495057",
            anchor="w",
            padx=12,
            pady=4
        )
        self.durum_etiketi.pack(fill="x", side="bottom")

    def _uygulama_ikonu_yukle(self):
        """Pencere simgesini .ico veya .png dosyasından yükler."""
        ico_yolu = resource_path("logo.ico")
        png_yolu = resource_path("logo.png")

        if os.path.exists(ico_yolu):
            try:
                self.root.iconbitmap(ico_yolu)
                return
            except Exception:
                pass

        if os.path.exists(png_yolu):
            try:
                img = Image.open(png_yolu)
                self._icon_img = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, self._icon_img)
            except Exception:
                pass

    def _bates_durum_guncelle(self):
        if self.bates_aktif.get():
            self.bates_on_ek.config(state="normal")
        else:
            self.bates_on_ek.config(state="disabled")

    def dosya_ekle_listeye(self, dosya_yolu):
        dosya_yolu = dosya_yolu.strip('{}')
        uzanti = os.path.splitext(dosya_yolu)[1].lower()
        if uzanti in DESTEKLENEN_UZANTILAR and dosya_yolu not in self.dosyalar:
            self.dosyalar.append(dosya_yolu)
            boyut_mb = os.path.getsize(dosya_yolu) / (1024 * 1024)
            simgeler = {'.pdf': '📄', '.udf': '⚖'}
            tur_simge = simgeler.get(uzanti, '🖼')
            dosya_adi = os.path.basename(dosya_yolu)
            self.dosya_listebox.insert(tk.END, f"{tur_simge} {dosya_adi} ({boyut_mb:.2f} MB)")

    def dosyalari_surukle_birak(self, event):
        dosya_yollari = self.root.tk.splitlist(event.data)
        for dosya in dosya_yollari:
            self.dosya_ekle_listeye(dosya)
        self.durum_etiketi.config(text=f"Toplam {len(self.dosyalar)} dosya eklendi.")

    def dosya_ekle_dialog(self):
        secilenler = filedialog.askopenfilenames(
            title="Dosyaları Seçin",
            filetypes=[
                ("Tüm Desteklenen Dosyalar", "*.pdf *.udf *.jpg *.jpeg *.png *.webp *.bmp"),
                ("UDF Dosyaları", "*.udf"),
                ("PDF Dosyaları", "*.pdf"),
                ("Görsel Dosyaları", "*.jpg *.jpeg *.png *.webp *.bmp")
            ]
        )
        for dosya in secilenler:
            self.dosya_ekle_listeye(dosya)
        self.durum_etiketi.config(text=f"Toplam {len(self.dosyalar)} dosya eklendi.")

    def secileni_sil(self):
        secili = self.dosya_listebox.curselection()
        if secili:
            indeks = secili[0]
            del self.dosyalar[indeks]
            self.dosya_listebox.delete(indeks)
            self.durum_etiketi.config(text="Seçilen dosya kaldırıldı.")

    def listeyi_temizle(self):
        self.dosyalar.clear()
        self.dosya_listebox.delete(0, tk.END)
        self.durum_etiketi.config(text="Liste temizlendi.")

    def _secili_dosya_al(self):
        secili = self.dosya_listebox.curselection()
        if not secili:
            messagebox.showwarning("Seçim Yapılmadı", "Lütfen sol listeden bir dosya seçin.")
            return None
        return self.dosyalar[secili[0]]

    # ========================================================
    # 1. SAYFA DÖNDÜRME & YÖN DÜZELTME MODÜLÜ
    # ========================================================
    def pdf_dondur_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Sayfa Döndürme yalnızca PDF dosyaları içindir.")
            return

        doc = pymupdf.open(kaynak_yol)
        toplam_sayfa = len(doc)

        pencere = tk.Toplevel(self.root)
        pencere.title("Sayfa Döndür & Yön Düzelt")
        pencere.geometry("400x260")
        pencere.resizable(False, False)
        pencere.configure(bg="#F4F6F9")

        tk.Label(
            pencere,
            text="🔄 Sayfa Döndürme Modülü",
            font=("Segoe UI", 11, "bold"),
            bg="#1B263B",
            fg="white",
            pady=6
        ).pack(fill="x")

        icerik = tk.Frame(pencere, bg="#F4F6F9", padx=14, pady=8)
        icerik.pack(fill="both", expand=True)

        tk.Label(icerik, text=f"Toplam Sayfa: {toplam_sayfa}", font=("Segoe UI", 9, "bold"), bg="#F4F6F9").pack(anchor="w")

        derece_var = tk.IntVar(value=90)
        yon_cerceve = tk.Frame(icerik, bg="#F4F6F9")
        yon_cerceve.pack(fill="x", pady=4)
        tk.Radiobutton(yon_cerceve, text="90° Sağa", variable=derece_var, value=90, bg="#F4F6F9").pack(side="left")
        tk.Radiobutton(yon_cerceve, text="90° Sola", variable=derece_var, value=270, bg="#F4F6F9").pack(side="left", padx=10)
        tk.Radiobutton(yon_cerceve, text="180° Baş Aşağı", variable=derece_var, value=180, bg="#F4F6F9").pack(side="left")

        tk.Label(icerik, text="Döndürülecek Sayfalar (Boş bırakırsanız hepsi döner):", font=("Segoe UI", 8), bg="#F4F6F9").pack(anchor="w", pady=(6, 2))
        aralik_entry = tk.Entry(icerik, font=("Segoe UI", 9))
        aralik_entry.pack(fill="x")
        tk.Label(icerik, text="Örn: 1-3, 5 (veya tümü için boş bırakın)", font=("Segoe UI", 7, "italic"), fg="#6C757D", bg="#F4F6F9").pack(anchor="w")

        def dondur_kaydet():
            aralik_metni = aralik_entry.get().strip()
            secilen_indeksler = set()
            if aralik_metni:
                try:
                    for parca in aralik_metni.split(','):
                        parca = parca.strip()
                        if '-' in parca:
                            b, s = map(int, parca.split('-'))
                            for i in range(b, s + 1):
                                if 1 <= i <= toplam_sayfa:
                                    secilen_indeksler.add(i - 1)
                        else:
                            i = int(parca)
                            if 1 <= i <= toplam_sayfa:
                                secilen_indeksler.add(i - 1)
                except Exception:
                    messagebox.showerror("Hata", "Geçersiz sayfa formatı! Örn: 1-3, 5")
                    return
            else:
                secilen_indeksler = set(range(toplam_sayfa))

            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Döndürülmüş PDF'i Kaydet",
                initialfile="dondurulmus_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol:
                return

            try:
                derece = derece_var.get()
                for idx in secilen_indeksler:
                    sayfa = doc[idx]
                    sayfa.set_rotation((sayfa.rotation + derece) % 360)

                doc.save(hedef_yol, garbage=4, deflate=True)
                doc.close()
                pencere.destroy()
                self.durum_etiketi.config(text=f"{len(secilen_indeksler)} sayfa döndürüldü.")
                messagebox.showinfo("Başarılı", f"{len(secilen_indeksler)} sayfa başarıyla döndürüldü ve kaydedildi!")
            except Exception as e:
                messagebox.showerror("Hata", f"Döndürme hatası:\n{str(e)}")

        tk.Button(
            icerik,
            text="🔄 Döndür ve Kaydet",
            command=dondur_kaydet,
            bg="#0582CA",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=5
        ).pack(fill="x", pady=(10, 0))

    # ========================================================
    # 2. GELİŞMİŞ 'ASLI GİBİDİR' KAŞESİ
    # ========================================================
    def pdf_kase_ekle_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Kaşe Ekleme yalnızca PDF dosyaları içindir.")
            return

        doc = pymupdf.open(kaynak_yol)
        toplam_sayfa = len(doc)
        doc.close()

        bugun = datetime.datetime.now().strftime("%d.%m.%Y")

        pencere = tk.Toplevel(self.root)
        pencere.title("Otomatik 'Aslı Gibidir' Kaşesi")
        pencere.geometry("450x450")
        pencere.resizable(False, False)
        pencere.configure(bg="#F4F6F9")

        tk.Label(
            pencere,
            text="🛡 Dijital Aslı Gibidir / Suret Kaşesi",
            font=("Segoe UI", 11, "bold"),
            bg="#1B263B",
            fg="white",
            pady=6
        ).pack(fill="x")

        icerik = tk.Frame(pencere, bg="#F4F6F9", padx=14, pady=6)
        icerik.pack(fill="both", expand=True)

        tk.Label(icerik, text="Kaşe Metni Başlığı:", font=("Segoe UI", 8, "bold"), bg="#F4F6F9").pack(anchor="w")
        baslik_entry = tk.Entry(icerik, font=("Segoe UI", 9))
        baslik_entry.insert(0, "ASLI GİBİDİR")
        baslik_entry.pack(fill="x", pady=(0, 2))

        tk.Label(icerik, text="Avukat İsim / Unvan:", font=("Segoe UI", 8, "bold"), bg="#F4F6F9").pack(anchor="w")
        avukat_entry = tk.Entry(icerik, font=("Segoe UI", 9))
        avukat_entry.insert(0, "Av. Tolunay AKAY")
        avukat_entry.pack(fill="x", pady=(0, 2))

        tk.Label(icerik, text="Tarih / Ek Bilgi:", font=("Segoe UI", 8, "bold"), bg="#F4F6F9").pack(anchor="w")
        ek_entry = tk.Entry(icerik, font=("Segoe UI", 9))
        ek_entry.insert(0, f"Tarih: {bugun}")
        ek_entry.pack(fill="x", pady=(0, 2))

        # Renk Seçimi
        renk_var = tk.StringVar(value="kirmizi")
        renk_cerceve = tk.Frame(icerik, bg="#F4F6F9")
        renk_cerceve.pack(fill="x", pady=2)
        tk.Label(renk_cerceve, text="Renk:", font=("Segoe UI", 8, "bold"), bg="#F4F6F9").pack(side="left")
        tk.Radiobutton(renk_cerceve, text="Kırmızı", variable=renk_var, value="kirmizi", bg="#F4F6F9", fg="#C1121F", font=("Segoe UI", 8, "bold")).pack(side="left", padx=5)
        tk.Radiobutton(renk_cerceve, text="Lacivert", variable=renk_var, value="lacivert", bg="#F4F6F9", fg="#003049", font=("Segoe UI", 8, "bold")).pack(side="left", padx=5)

        # Kaşe Konumu Seçimi
        konum_var = tk.StringVar(value="sag_alt")
        konum_cerceve = tk.LabelFrame(icerik, text=" Kaşe Konumu ", font=("Segoe UI", 8, "bold"), bg="#F4F6F9", fg="#1B263B", padx=6, pady=2)
        konum_cerceve.pack(fill="x", pady=3)

        konum_ust = tk.Frame(konum_cerceve, bg="#F4F6F9")
        konum_ust.pack(fill="x")
        tk.Radiobutton(konum_ust, text="↘ Sağ Alt", variable=konum_var, value="sag_alt", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left", expand=True)
        tk.Radiobutton(konum_ust, text="↙ Sol Alt", variable=konum_var, value="sol_alt", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left", expand=True)

        konum_alt = tk.Frame(konum_cerceve, bg="#F4F6F9")
        konum_alt.pack(fill="x")
        tk.Radiobutton(konum_alt, text="⬇ Alt Orta", variable=konum_var, value="alt_orta", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left", expand=True)
        tk.Radiobutton(konum_alt, text="↗ Sağ Üst", variable=konum_var, value="sag_ust", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left", expand=True)

        # Sayfa Kapsamı Seçimi
        sayfa_kapsam_var = tk.StringVar(value="hepsi")
        kapsam_cerceve = tk.LabelFrame(icerik, text=" Sayfa Kapsamı ", font=("Segoe UI", 8, "bold"), bg="#F4F6F9", fg="#1B263B", padx=6, pady=2)
        kapsam_cerceve.pack(fill="x", pady=3)

        tk.Radiobutton(kapsam_cerceve, text="Tüm Sayfalara Bas", variable=sayfa_kapsam_var, value="hepsi", bg="#F4F6F9", font=("Segoe UI", 8)).pack(anchor="w")
        
        aralik_secim_satir = tk.Frame(kapsam_cerceve, bg="#F4F6F9")
        aralik_secim_satir.pack(fill="x")
        tk.Radiobutton(aralik_secim_satir, text="Yalnızca Belirtilen Sayfalara:", variable=sayfa_kapsam_var, value="aralik", bg="#F4F6F9", font=("Segoe UI", 8)).pack(side="left")
        kase_sayfa_entry = tk.Entry(aralik_secim_satir, font=("Segoe UI", 8), width=14)
        kase_sayfa_entry.insert(0, "1")
        kase_sayfa_entry.pack(side="left", padx=4)
        tk.Label(aralik_secim_satir, text="(Örn: 1-3, 5)", font=("Segoe UI", 7, "italic"), fg="#6C757D", bg="#F4F6F9").pack(side="left")

        def kaseyi_bas():
            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Kaşeli PDF'i Kaydet",
                initialfile="kaseli_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol:
                return

            try:
                doc = pymupdf.open(kaynak_yol)
                toplam_s = len(doc)
                renk_kodu = (0.75, 0.05, 0.05) if renk_var.get() == "kirmizi" else (0.0, 0.18, 0.38)

                b_metin = baslik_entry.get().strip()
                a_metin = avukat_entry.get().strip()
                e_metin = ek_entry.get().strip()
                tam_kase_metni = f"{b_metin}\n{a_metin}\n{e_metin}"

                hedef_indeksler = set()
                if sayfa_kapsam_var.get() == "hepsi":
                    hedef_indeksler = set(range(toplam_s))
                else:
                    girdi = kase_sayfa_entry.get().strip()
                    try:
                        for parca in girdi.split(','):
                            parca = parca.strip()
                            if '-' in parca:
                                b, son = map(int, parca.split('-'))
                                for s in range(b, son + 1):
                                    if 1 <= s <= toplam_s:
                                        hedef_indeksler.add(s - 1)
                            else:
                                s = int(parca)
                                if 1 <= s <= toplam_s:
                                    hedef_indeksler.add(s - 1)
                    except Exception:
                        doc.close()
                        messagebox.showerror("Hata", "Geçersiz sayfa formatı girdiniz! (Örn: 1-3, 5)")
                        return

                if not hedef_indeksler:
                    doc.close()
                    messagebox.showwarning("Uyarı", "Kaşe basılacak geçerli hiçbir sayfa seçilmedi.")
                    return

                font_param = {"fontfile": TURKCE_FONT_YOLU, "fontname": "arial_tr"} if TURKCE_FONT_YOLU else {"fontname": "helv"}

                secili_konum = konum_var.get()
                kutu_w, kutu_h = 165, 52
                kenar_payi = 25

                tahmini_metin_yuksekligi = 35.0
                dikey_bosluk = max(0.0, (kutu_h - tahmini_metin_yuksekligi) / 2.0)

                for idx in hedef_indeksler:
                    sayfa = doc[idx]
                    w = sayfa.rect.width
                    h = sayfa.rect.height

                    if secili_konum == "sag_alt":
                        cerceve_rect = pymupdf.Rect(w - kutu_w - kenar_payi, h - kutu_h - kenar_payi, w - kenar_payi, h - kenar_payi)
                    elif secili_konum == "sol_alt":
                        cerceve_rect = pymupdf.Rect(kenar_payi, h - kutu_h - kenar_payi, kenar_payi + kutu_w, h - kenar_payi)
                    elif secili_konum == "alt_orta":
                        orta_x = (w - kutu_w) / 2
                        cerceve_rect = pymupdf.Rect(orta_x, h - kutu_h - kenar_payi, orta_x + kutu_w, h - kenar_payi)
                    elif secili_konum == "sag_ust":
                        cerceve_rect = pymupdf.Rect(w - kutu_w - kenar_payi, kenar_payi, w - kenar_payi, kenar_payi + kutu_h)
                    else:
                        cerceve_rect = pymupdf.Rect(w - kutu_w - kenar_payi, h - kutu_h - kenar_payi, w - kenar_payi, h - kenar_payi)

                    sayfa.draw_rect(cerceve_rect, color=renk_kodu, width=1.5)

                    metin_rect = pymupdf.Rect(
                        cerceve_rect.x0 + 4,
                        cerceve_rect.y0 + dikey_bosluk,
                        cerceve_rect.x1 - 4,
                        cerceve_rect.y1
                    )

                    sayfa.insert_textbox(
                        metin_rect,
                        tam_kase_metni,
                        fontsize=9,
                        color=renk_kodu,
                        align=pymupdf.TEXT_ALIGN_CENTER,
                        **font_param
                    )

                doc.save(hedef_yol, garbage=4, deflate=True)
                doc.close()
                pencere.destroy()
                self.durum_etiketi.config(text=f"Kaşe seçilen {len(hedef_indeksler)} sayfaya başarıyla basıldı.")
                messagebox.showinfo("Başarılı", f"Kaşe seçilen {len(hedef_indeksler)} sayfaya başarıyla basıldı!\n\nDosya: {os.path.basename(hedef_yol)}")
            except Exception as e:
                messagebox.showerror("Hata", f"Kaşe basılırken hata oluştu:\n{str(e)}")

        tk.Button(
            icerik,
            text="🛡 Kaşeyi Bas ve Kaydet",
            command=kaseyi_bas,
            bg="#C1121F",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=5
        ).pack(fill="x", pady=(6, 0))

    # ========================================================
    # 3. BOŞ SAYFALARI OTOMATİK TEMİZLEME MODÜLÜ
    # ========================================================
    def pdf_bos_sayfalari_temizle(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Boş sayfa temizleme yalnızca PDF dosyaları içindir.")
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Temizlenmiş PDF'i Kaydet",
            initialfile="temiz_" + os.path.basename(kaynak_yol)
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.config(text="Sayfalar piksel ve metin analizinden geçiriliyor...")
            self.root.update()

            doc = pymupdf.open(kaynak_yol)
            yeni_doc = pymupdf.open()
            silinen_sayfa_sayisi = 0

            for i, sayfa in enumerate(doc):
                metin = sayfa.get_text().strip()
                if metin:
                    yeni_doc.insert_pdf(doc, from_page=i, to_page=i)
                    continue

                pix = sayfa.get_pixmap(dpi=50)
                samples = pix.samples
                toplam_piksel = pix.width * pix.height
                beyaz_olmayan = 0
                step = pix.n

                for p in range(0, len(samples), step):
                    val = samples[p]
                    if val < 240:
                        beyaz_olmayan += 1

                oran = beyaz_olmayan / toplam_piksel
                if oran > 0.002:
                    yeni_doc.insert_pdf(doc, from_page=i, to_page=i)
                else:
                    silinen_sayfa_sayisi += 1

            if len(yeni_doc) == 0:
                yeni_doc.insert_pdf(doc, from_page=0, to_page=0)
                silinen_sayfa_sayisi -= 1

            yeni_doc.save(hedef_yol, garbage=4, deflate=True)
            yeni_doc.close()
            doc.close()

            self.durum_etiketi.config(text=f"Temizlik tamamlandı: {silinen_sayfa_sayisi} boş sayfa silindi.")
            messagebox.showinfo(
                "Temizlik Başarılı",
                f"Toplam {silinen_sayfa_sayisi} adet boş/beyaz sayfa tespit edilip silindi!\n\nDosya: {os.path.basename(hedef_yol)}"
            )
        except Exception as e:
            messagebox.showerror("Hata", f"Boş sayfalar temizlenirken hata oluştu:\n{str(e)}")

    # ========================================================
    # 4. AKILLI UYAP BOYUT BÖLÜCÜSÜ (CHUNK SPLITTER)
    # ========================================================
    def pdf_akilli_boyut_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Akıllı Boyut Bölücüsü yalnızca PDF dosyaları içindir.")
            return

        dosya_mb = os.path.getsize(kaynak_yol) / (1024 * 1024)
        
        limit_str = simpledialog.askstring(
            "Hedef Boyut Limiti (MB)",
            f"Seçili Dosya Boyutu: {dosya_mb:.2f} MB\n\nHer bir parçanın maksimum boyutunu (MB) girin:\n(UYAP ek yükleme limiti için 9.5 MB önerilir)",
            initialvalue="9.5",
            parent=self.root
        )
        if not limit_str:
            return

        try:
            hedef_limit_mb = float(limit_str.replace(',', '.'))
            if hedef_limit_mb <= 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("Hata", "Lütfen geçerli bir megabayt (MB) değeri girin! (Örn: 9.5)")
            return

        hedef_limit_bytes = hedef_limit_mb * 1024 * 1024

        hedef_klasor = filedialog.askdirectory(title="Parçaların Kaydedileceği Klasörü Seçin")
        if not hedef_klasor:
            return

        try:
            self.durum_etiketi.config(text="Akıllı boyut analizi ve bölme işlemi yapılıyor...")
            self.root.update()

            ana_doc = pymupdf.open(kaynak_yol)
            toplam_sayfa = len(ana_doc)
            dosya_kok = os.path.splitext(os.path.basename(kaynak_yol))[0]

            parca_sayisi = 1
            mevcut_parca_doc = pymupdf.open()
            mevcut_parca_sayfa_sayisi = 0
            uretilen_dosyalar = []

            for i in range(toplam_sayfa):
                mevcut_parca_doc.insert_pdf(ana_doc, from_page=i, to_page=i)
                mevcut_parca_sayfa_sayisi += 1

                test_bytes = mevcut_parca_doc.tobytes(garbage=4, deflate=True)
                mevcut_boyut = len(test_bytes)

                if mevcut_boyut > hedef_limit_bytes and mevcut_parca_sayfa_sayisi > 1:
                    mevcut_parca_doc.delete_page(mevcut_parca_sayfa_sayisi - 1)
                    cikti_yol = os.path.join(hedef_klasor, f"{dosya_kok}_parca_{parca_sayisi}.pdf")
                    mevcut_parca_doc.save(cikti_yol, garbage=4, deflate=True)
                    mevcut_parca_doc.close()
                    
                    uretilen_boyut_mb = os.path.getsize(cikti_yol) / (1024 * 1024)
                    uretilen_dosyalar.append((os.path.basename(cikti_yol), uretilen_boyut_mb))

                    parca_sayisi += 1
                    mevcut_parca_doc = pymupdf.open()
                    mevcut_parca_doc.insert_pdf(ana_doc, from_page=i, to_page=i)
                    mevcut_parca_sayfa_sayisi = 1

            if len(mevcut_parca_doc) > 0:
                cikti_yol = os.path.join(hedef_klasor, f"{dosya_kok}_parca_{parca_sayisi}.pdf")
                mevcut_parca_doc.save(cikti_yol, garbage=4, deflate=True)
                mevcut_parca_doc.close()
                uretilen_boyut_mb = os.path.getsize(cikti_yol) / (1024 * 1024)
                uretilen_dosyalar.append((os.path.basename(cikti_yol), uretilen_boyut_mb))

            ana_doc.close()

            ozet_metin = "\n".join([f"• {ad}: {b:.2f} MB" for ad, b in uretilen_dosyalar])
            self.durum_etiketi.config(text=f"Bölme tamamlandı: {len(uretilen_dosyalar)} parça oluşturuldu.")
            messagebox.showinfo(
                "Akıllı Bölme Başarılı",
                f"Toplam {dosya_mb:.2f} MB'lık evrak {hedef_limit_mb} MB sınırına göre {len(uretilen_dosyalar)} parçaya bölündü!\n\n"
                f"{ozet_metin}\n\nKlasör: {hedef_klasor}"
            )
        except Exception as e:
            messagebox.showerror("Hata", f"Akıllı bölme sırasında hata oluştu:\n{str(e)}")

    # ========================================================
    # AYRINTILI KULLANIM REHBERİ
    # ========================================================
    def rehber_penceresi_ac(self):
        pencere = tk.Toplevel(self.root)
        pencere.title("Kullanım Kılavuzu & Özellik Rehberi")
        pencere.geometry("680x580")
        pencere.resizable(False, False)
        pencere.configure(bg="#F4F6F9")

        tk.Label(
            pencere,
            text="📖 UYAP Evrak Asistanı - Detaylı Kullanım Kılavuzu",
            font=("Segoe UI", 12, "bold"),
            bg="#1B263B",
            fg="white",
            pady=8
        ).pack(fill="x")

        metin_alani = scrolledtext.ScrolledText(
            pencere,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#212529",
            padx=14,
            pady=10
        )
        metin_alani.pack(fill="both", expand=True, padx=12, pady=8)

        rehber_icerik = """
UYAP EVRAK ASİSTANI - DETAYLI KULLANIM REHBERİ

Bu yazılım; adliye, icra ve UYAP evrak süreçlerinde karşılaşılan teknik kısıtlamaları (dosya boyutu, sayfa yönü, format uyumsuzlukları, veri gizliliği) en pratik ve güvenli şekilde çözmek amacıyla tasarlanmıştır.

==================================================
1. DOSYA EKLEME & LİSTE YÖNETİMİ
==================================================
• Sürükle-Bırak: PDF, görsel (JPG, PNG, WEBP, BMP) veya UDF dosyalarını doğrudan sol taraftaki listeye sürükleyip bırakabilirsiniz.
• Dosya Ekle: Pencere açarak bilgisayarınızdan tekli ya da toplu dosya seçebilirsiniz.
• Seçileni Kaldır / Listeyi Temizle: Yanlış eklenen evrakları listeden çıkarabilir veya tüm listeyi sıfırlayabilirsiniz.

==================================================
2. BİRLEŞTİRME & SIKIŞTIRMA MODÜLLERİ
==================================================
• 📑 PDF'leri Birleştir:
  - Listedeki tüm evrakları (PDF'ler ve fotoğraflar) sırasıyla tek bir PDF haline getirir.
  - Görselleri (dekont, senet, tutanak fotoğrafları) sayfayı bozmadan otomatik olarak A4 boyutuna oranlar, kenar boşluklarını ayarlar ve ortalar.
  - Sıkıştırma seviyesine göre çıktıyı UYAP'a uygun hafif boyutta üretir.

• 🗜 PDF Sıkıştır:
  - Listeden seçtiğiniz tek bir PDF'in kalitesini okunabilir seviyede tutarak boyutunu küçültür (örn: 15 MB ➔ 2 MB).
  - 'Dengeli (150 DPI)': Standart evraklar ve dilekçeler için ideal netlik ve boyut tasarrufu sağlar.
  - 'Maksimum Sıkıştırma (100 DPI)': Çok ağır taranmış dosyaları UYAP limitine sokmak için yüksek oranda sıkıştırır.

• 🏷 Ek Etiketi / Bates Numaralandırma:
  - 'Sayfa Altına Numara Bas' kutusunu işaretlerseniz, evraklar birleştirilirken veya sıkıştırılırken her sayfanın alt ortasına numara basılır.
  - Etiket kutusuna 'EK-1', 'DELİL-2' veya dilediğiniz başlığı yazabilirsiniz. Çıktı formatı: 'EK-1 | Sayfa 1/15' şeklinde yerleştirilir.

==================================================
3. SAYFA BÖLME & PARÇALAMA İŞLEMLERİ
==================================================
• ✂ Sayfa Böl:
  - Seçtiğiniz PDF'in her bir sayfasını tek tek ayırır.
  - Belirlediğiniz klasörün içine 'dosya_sayfa_1.pdf', 'dosya_sayfa_2.pdf' şeklinde kaydeder.

• 🎯 Aralık Ayır:
  - Belgeden yalnızca istediğiniz belirli sayfaları çekip yeni bir PDF yapar.
  - Giriş Formatı: '1, 3, 5' veya '1-4, 7-10'.

• ⚖ Akıllı UYAP Boyut Bölücüsü (MB Sınırına Göre):
  - Çok hacimli dosyaları hedef MB sınırına göre akıllıca parçalar.
  - Varsayılan olarak UYAP ek sınırına tam uyan '9.5 MB' önerilir.
  - Sayfaları sırayla tartarak 'dosya_parca_1.pdf', 'dosya_parca_2.pdf' üretir; hiçbir parça limiti aşmaz.

==================================================
4. EVRAK DÜZENLEME & İMZA ÖNCESİ HAZIRLIK
==================================================
• 🔄 Sayfa Döndür (90°):
  - Ters veya yan taranmış sayfaların yönünü (90° Sağa/Sola veya 180°) düzeltir.
  - İster belgenin tamamını döndürün, ister sadece '1-3, 5' gibi belirli sayfaları.

• 🛡 'Aslı Gibidir' Kaşesi:
  - Belgenin üzerine şık çerçeveli dijital kaşe basar.
  - Konum: Sağ Alt, Sol Alt, Alt Orta veya Sağ Üst.
  - Kapsam: 'Tüm Sayfalara' veya 'Yalnızca Belirtilen Sayfalara (Örn: 1)'.
  - Renk: Kırmızı veya Lacivert resmi mühür rengi.
  - Türkçe karakterler tam desteklidir.

• 🧹 Boş Sayfaları Temizle:
  - Çift taraflı tarayıcılardan geçen evrakların arkasındaki bomboş beyaz sayfaları piksel analiziyle tespit edip siler.

==================================================
5. METİN İŞLEMLERİ & VERİ GÜVENLİĞİ
==================================================
• 🕶 KVKK Akıllı Karartma (Redaction):
  - T.C. Kimlik Numaralarını, IBAN'ları ve telefon numaralarını kalıcı olarak maskeler.
  - Özel isim veya kelimeleri de gizleyebilirsiniz.
  - Metin katmanı PDF belleğinden tamamen kazınır, geri alınamaz.

• 📄 UDF ➔ Word (.docx):
  - UYAP '.udf' evraklarını Microsoft Word (.docx) formatına dönüştürür.

• 🔍 Metin Kopyala:
  - Belgedeki tüm metinleri önizleme ekranında listeler ve tek tıkla panoya kopyalar.
"""
        metin_alani.insert(tk.END, rehber_icerik.strip())
        metin_alani.config(state="disabled")

        tk.Button(
            pencere,
            text="Anladım, Kapat",
            command=pencere.destroy,
            bg="#1B263B",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=5
        ).pack(fill="x", padx=12, pady=(0, 8))

    # ========================================================
    # KVKK AKILLI KARARTMA (REDACTION) MODÜLÜ
    # ========================================================
    def akilli_karartma_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Akıllı Karartma yalnızca PDF dosyaları için geçerlidir.")
            return

        pencere = tk.Toplevel(self.root)
        pencere.title("KVKK Akıllı Karartma")
        pencere.geometry("450x390")
        pencere.resizable(False, False)
        pencere.configure(bg="#F4F6F9")

        tk.Label(
            pencere,
            text="🕶 Kalıcı Veri Karartma / Maskeleme",
            font=("Segoe UI", 11, "bold"),
            bg="#1B263B",
            fg="white",
            pady=6
        ).pack(fill="x")

        icerik = tk.Frame(pencere, bg="#F4F6F9", padx=16, pady=8)
        icerik.pack(fill="both", expand=True)

        var_tc = tk.BooleanVar(value=True)
        tk.Checkbutton(icerik, text="11 Haneli T.C. Kimlik Numaralarını Gizle", variable=var_tc, bg="#F4F6F9", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=2)

        var_iban = tk.BooleanVar(value=True)
        tk.Checkbutton(icerik, text="IBAN Numaralarını Gizle (TR...)", variable=var_iban, bg="#F4F6F9", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=2)

        var_tel = tk.BooleanVar(value=False)
        tk.Checkbutton(icerik, text="Telefon Numaralarını Gizle (05xx / +90)", variable=var_tel, bg="#F4F6F9", font=("Segoe UI", 8)).pack(anchor="w", pady=2)

        tk.Label(icerik, text="Özel İsim / Kelimeleri Gizle (Virgülle ayırın):", bg="#F4F6F9", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(8, 2))
        ozel_entry = tk.Entry(icerik, font=("Segoe UI", 9), width=45)
        ozel_entry.pack(fill="x", pady=2)

        def islemi_baslat():
            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Karartılmış PDF'i Kaydet",
                initialfile="karartilmis_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol:
                return

            try:
                doc = pymupdf.open(kaynak_yol)
                toplam_maskeleme = 0
                regex_listesi = []
                if var_tc.get():
                    regex_listesi.append(r'\b[1-9][0-9]{10}\b')
                if var_iban.get():
                    regex_listesi.append(r'TR\s*[0-9]{2}(?:\s*[0-9]{4}){5}\s*[0-9]{2}')
                if var_tel.get():
                    regex_listesi.append(r'(?:\+90\s*|0\s*)?5[0-9]{2}[\s\-]?[0-9]{3}[\s\-]?[0-9]{2}[\s\-]?[0-9]{2}')

                ozel_kelimeler = [k.strip() for k in ozel_entry.get().split(',') if k.strip()]

                for sayfa in doc:
                    sayfa_metni = sayfa.get_text()
                    for desen in regex_listesi:
                        for eslesme in re.finditer(desen, sayfa_metni):
                            aranan = eslesme.group()
                            for rect in sayfa.search_for(aranan):
                                sayfa.add_redact_annot(rect, fill=(0, 0, 0))
                                toplam_maskeleme += 1

                    for kelime in ozel_kelimeler:
                        for rect in sayfa.search_for(kelime):
                            sayfa.add_redact_annot(rect, fill=(0, 0, 0))
                            toplam_maskeleme += 1

                    sayfa.apply_redactions()

                doc.save(hedef_yol, garbage=4, deflate=True)
                doc.close()
                pencere.destroy()
                self.durum_etiketi.config(text=f"Karartma tamamlandı: {toplam_maskeleme} alan maskelendi.")
                messagebox.showinfo("Başarılı", f"Kalıcı Karartma Uygulandı!\nToplam Maskelenen Veri: {toplam_maskeleme} adet")
            except Exception as e:
                messagebox.showerror("Hata", f"Karartma hatası:\n{str(e)}")

        tk.Button(
            icerik,
            text="🛡 Karart ve Yeni PDF Olarak Kaydet",
            command=islemi_baslat,
            bg="#212529",
            fg="#FFC107",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            pady=6
        ).pack(fill="x", pady=(12, 0))

    # ========================================================
    # UDF ➔ WORD (.docx) DÖNÜŞTÜRÜCÜ
    # ========================================================
    def udf_to_word(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.udf'):
            messagebox.showwarning("Hatalı Dosya", "Lütfen listeden bir UDF (.udf) dosyası seçin.")
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Belgesi", "*.docx")],
            title="Word (.docx) Olarak Kaydet",
            initialfile=os.path.splitext(os.path.basename(kaynak_yol))[0] + ".docx"
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.config(text="UDF içeriği Word'e aktarılıyor...")
            self.root.update()

            with zipfile.ZipFile(kaynak_yol, 'r') as zf:
                if 'content.xml' not in zf.namelist():
                    raise ValueError("Geçersiz UDF dosyası: 'content.xml' bulunamadı.")
                xml_raw = zf.read('content.xml')

            root_xml = ET.fromstring(xml_raw)
            yeni_docx = docx.Document()
            paragraf_bulundu = False

            content_node = root_xml.find('content')
            if content_node is not None and content_node.text:
                for satir in content_node.text.splitlines():
                    yeni_docx.add_paragraph(satir.replace('\u200B', ''))
                paragraf_bulundu = True

            if not paragraf_bulundu:
                elements_node = root_xml.find('elements')
                if elements_node is not None:
                    for elem in elements_node.findall('.//paragraph'):
                        p_metin = "".join(elem.itertext()).strip().replace('\u200B', '')
                        yeni_docx.add_paragraph(p_metin)
                    paragraf_bulundu = True

            if not paragraf_bulundu:
                tum_metin = "".join(root_xml.itertext()).replace('\u200B', '')
                for satir in tum_metin.splitlines():
                    if satir.strip():
                        yeni_docx.add_paragraph(satir)

            yeni_docx.save(hedef_yol)
            self.durum_etiketi.config(text="UDF ➔ Word dönüşümü tamamlandı.")
            messagebox.showinfo("Başarılı", f"UDF belgesi Word formatına aktarıldı!\n\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Word belgesi oluşturulurken hata oluştu:\n{str(e)}")

    # ========================================================
    # DİĞER PDF İŞLEMLERİ (BİRLEŞTİR, SIKIŞTIR, BÖL)
    # ========================================================
    def _gorseli_a4_sayfaya_ekle(self, gorsel_yolu, hedef_doc, kalite):
        a4_w, a4_h = 595.0, 842.0
        yeni_sayfa = hedef_doc.new_page(width=a4_w, height=a4_h)

        with Image.open(gorsel_yolu) as img:
            if img.mode in ("RGBA", "P", "CMYK"):
                img = img.convert("RGB")
            
            margin = 36.0
            kullanilabilir_w = a4_w - (2 * margin)
            kullanilabilir_h = a4_h - (2 * margin)

            img_w, img_h = img.size
            oran = min(kullanilabilir_w / img_w, kullanilabilir_h / img_h)
            cizim_w = img_w * oran
            cizim_h = img_h * oran

            x0 = (a4_w - cizim_w) / 2
            y0 = (a4_h - cizim_h) / 2
            rect = pymupdf.Rect(x0, y0, x0 + cizim_w, y0 + cizim_h)

            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format="JPEG", quality=kalite, optimize=True)
            yeni_sayfa.insert_image(rect, stream=img_byte_arr.getvalue())

    def _pdf_sayfalarini_optimize_et(self, kaynak_doc, hedef_doc, dpi, kalite):
        for sayfa in kaynak_doc:
            pix = sayfa.get_pixmap(dpi=dpi)
            resim_bytes = pix.pil_tobytes(format="JPEG", quality=kalite, optimize=True)
            yeni_sayfa = hedef_doc.new_page(width=sayfa.rect.width, height=sayfa.rect.height)
            yeni_sayfa.insert_image(sayfa.rect, stream=resim_bytes)

    def _bates_numaralari_bas(self, doc):
        if not self.bates_aktif.get():
            return
        on_ek = self.bates_on_ek.get().strip()
        toplam_sayfa = len(doc)

        font_param = {"fontfile": TURKCE_FONT_YOLU, "fontname": "arial_tr"} if TURKCE_FONT_YOLU else {"fontname": "helv"}

        for i, sayfa in enumerate(doc):
            sayfa_no = i + 1
            metin = f"{on_ek} | Sayfa {sayfa_no} / {toplam_sayfa}" if on_ek else f"Sayfa {sayfa_no} / {toplam_sayfa}"
            w, h = sayfa.rect.width, sayfa.rect.height
            rect = pymupdf.Rect(0, h - 26, w, h - 8)
            sayfa.insert_textbox(
                rect,
                metin,
                fontsize=8.5,
                align=pymupdf.TEXT_ALIGN_CENTER,
                color=(0.2, 0.2, 0.2),
                **font_param
            )

    def tumunu_birlestir_pdf_yap(self):
        pdf_ve_resimler = [d for d in self.dosyalar if d.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png', '.webp', '.bmp'))]
        if not pdf_ve_resimler:
            messagebox.showwarning("Geçerli Dosya Yok", "Birleştirmek için listede PDF veya görsel dosyası bulunmalıdır.")
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Oluşturulacak PDF'i Kaydet",
            initialfile="uyap_birlestirilmis_evrak.pdf"
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.config(text="PDF oluşturuluyor...")
            self.root.update()

            toplam_eski_boyut = sum(os.path.getsize(d) for d in pdf_ve_resimler) / (1024 * 1024)
            dpi = 100 if self.mod_secimi.get() == "maksimum" else 150
            kalite = 50 if self.mod_secimi.get() == "maksimum" else 70

            yeni_doc = pymupdf.open()
            for dosya in pdf_ve_resimler:
                uzanti = os.path.splitext(dosya)[1].lower()
                if uzanti == '.pdf':
                    kaynak_pdf = pymupdf.open(dosya)
                    self._pdf_sayfalarini_optimize_et(kaynak_pdf, yeni_doc, dpi, kalite)
                    kaynak_pdf.close()
                else:
                    self._gorseli_a4_sayfaya_ekle(dosya, yeni_doc, kalite)

            self._bates_numaralari_bas(yeni_doc)
            yeni_doc.save(hedef_yol, garbage=4, deflate=True)
            yeni_doc.close()

            yeni_boyut = os.path.getsize(hedef_yol) / (1024 * 1024)
            self.durum_etiketi.config(text=f"Tamamlandı: {yeni_boyut:.2f} MB")
            messagebox.showinfo("Başarılı", f"PDF'ler birleştirildi!\n\nEski: {toplam_eski_boyut:.2f} MB\nYeni: {yeni_boyut:.2f} MB")
        except Exception as e:
            messagebox.showerror("Hata", f"İşlem hatası:\n{str(e)}")

    def pdf_tekil_sikistir(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Bu özellik yalnızca PDF dosyaları içindir.")
            return

        eski_boyut = os.path.getsize(kaynak_yol) / (1024 * 1024)
        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Sıkıştırılmış PDF'i Kaydet",
            initialfile="uyap_kucuk_" + os.path.basename(kaynak_yol)
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.config(text="Sıkıştırılıyor...")
            self.root.update()

            dpi = 100 if self.mod_secimi.get() == "maksimum" else 150
            kalite = 50 if self.mod_secimi.get() == "maksimum" else 70

            doc = pymupdf.open(kaynak_yol)
            yeni_doc = pymupdf.open()
            self._pdf_sayfalarini_optimize_et(doc, yeni_doc, dpi, kalite)
            self._bates_numaralari_bas(yeni_doc)
            yeni_doc.save(hedef_yol, garbage=4, deflate=True)
            yeni_doc.close()
            doc.close()

            yeni_boyut = os.path.getsize(hedef_yol) / (1024 * 1024)
            kazanc = ((eski_boyut - yeni_boyut) / eski_boyut) * 100 if eski_boyut > 0 else 0
            self.durum_etiketi.config(text=f"Tamamlandı: {eski_boyut:.2f} MB ➔ {yeni_boyut:.2f} MB")
            messagebox.showinfo("Başarılı", f"Eski: {eski_boyut:.2f} MB\nYeni: {yeni_boyut:.2f} MB\nKazanılan: %{kazanc:.1f}")
        except Exception as e:
            messagebox.showerror("Hata", f"Sıkıştırma hatası:\n{str(e)}")

    def pdf_sayfa_sayfa_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Bu özellik yalnızca PDF dosyaları içindir.")
            return

        hedef_klasor = filedialog.askdirectory(title="Sayfaların Kaydedileceği Klasörü Seçin")
        if not hedef_klasor:
            return

        try:
            doc = pymupdf.open(kaynak_yol)
            toplam_sayfa = len(doc)
            dosya_kok = os.path.splitext(os.path.basename(kaynak_yol))[0]

            for i in range(toplam_sayfa):
                tekil = pymupdf.open()
                tekil.insert_pdf(doc, from_page=i, to_page=i)
                tekil.save(os.path.join(hedef_klasor, f"{dosya_kok}_sayfa_{i + 1}.pdf"), garbage=4, deflate=True)
                tekil.close()

            doc.close()
            messagebox.showinfo("Başarılı", f"{toplam_sayfa} sayfa başarıyla ayrıldı!\nKlasör: {hedef_klasor}")
        except Exception as e:
            messagebox.showerror("Hata", f"Sayfalara bölerken hata:\n{str(e)}")

    def pdf_aralik_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Bu özellik yalnızca PDF dosyaları içindir.")
            return

        doc = pymupdf.open(kaynak_yol)
        toplam_sayfa = len(doc)

        aralik_metni = simpledialog.askstring("Sayfa Seçimi", f"Toplam Sayfa: {toplam_sayfa}\nAlmak istediğiniz sayfalar (Örn: 1-3, 5):", parent=self.root)
        if not aralik_metni:
            doc.close()
            return

        secilen_indeksler = set()
        try:
            for parca in aralik_metni.split(','):
                parca = parca.strip()
                if '-' in parca:
                    b, s = map(int, parca.split('-'))
                    for i in range(b, s + 1):
                        if 1 <= i <= toplam_sayfa:
                            secilen_indeksler.add(i - 1)
                else:
                    i = int(parca)
                    if 1 <= i <= toplam_sayfa:
                        secilen_indeksler.add(i - 1)
        except Exception:
            doc.close()
            messagebox.showerror("Hata", "Geçersiz format! Örn: 1-3, 5")
            return

        hedef_yol = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Dosyası", "*.pdf")], title="Ayrılan Sayfaları Kaydet", initialfile=f"ayrilan_{os.path.basename(kaynak_yol)}")
        if not hedef_yol:
            doc.close()
            return

        try:
            yeni = pymupdf.open()
            for idx in sorted(list(secilen_indeksler)):
                yeni.insert_pdf(doc, from_page=idx, to_page=idx)
            yeni.save(hedef_yol, garbage=4, deflate=True)
            yeni.close()
            doc.close()
            messagebox.showinfo("Başarılı", f"Seçilen {len(secilen_indeksler)} sayfa kaydedildi!")
        except Exception as e:
            messagebox.showerror("Hata", f"İşlem hatası:\n{str(e)}")

    def metin_oku_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol:
            return

        self.durum_etiketi.config(text="Metinler okunuyor...")
        self.root.update()

        metin_icerik = ""
        uzanti = os.path.splitext(kaynak_yol)[1].lower()

        try:
            if uzanti == '.pdf':
                doc = pymupdf.open(kaynak_yol)
                for i, sayfa in enumerate(doc):
                    t = sayfa.get_text()
                    if t.strip():
                        metin_icerik += f"\n--- [ Sayfa {i + 1} ] ---\n" + t
                doc.close()
            elif uzanti == '.udf':
                with zipfile.ZipFile(kaynak_yol, 'r') as zf:
                    if 'content.xml' in zf.namelist():
                        root_xml = ET.fromstring(zf.read('content.xml'))
                        c_node = root_xml.find('content')
                        if c_node is not None and c_node.text:
                            metin_icerik = c_node.text.replace('\u200B', '')
                        else:
                            metin_icerik = "".join(root_xml.itertext()).replace('\u200B', '')

            if not metin_icerik.strip():
                metin_icerik = "Kopyalanabilir metin bulunamadı."

            pencere = tk.Toplevel(self.root)
            pencere.title(f"Metin Önizleme: {os.path.basename(kaynak_yol)}")
            pencere.geometry("640x480")

            # İkincil pencere ikonu
            ico_yolu = resource_path("logo.ico")
            if os.path.exists(ico_yolu):
                try:
                    pencere.iconbitmap(ico_yolu)
                except Exception:
                    pass

            txt = scrolledtext.ScrolledText(pencere, wrap=tk.WORD, font=("Consolas", 10))
            txt.pack(fill="both", expand=True, padx=10, pady=10)
            txt.insert(tk.END, metin_icerik)

            def kopyala():
                pencere.clipboard_clear()
                pencere.clipboard_append(txt.get("1.0", tk.END))
                messagebox.showinfo("Kopyalandı", "Metin panoya kopyalandı!")

            tk.Button(pencere, text="📋 Panoya Kopyala", command=kopyala, bg="#0D6EFD", fg="white", font=("Segoe UI", 9, "bold")).pack(pady=5)
            self.durum_etiketi.config(text="Metin okundu.")
        except Exception as e:
            messagebox.showerror("Hata", f"Metin okunamadı:\n{str(e)}")


if __name__ == "__main__":
    pencere = TkinterDnD.Tk()
    uygulama = UyapEvrakAsistaniApp(pencere)
    pencere.mainloop()