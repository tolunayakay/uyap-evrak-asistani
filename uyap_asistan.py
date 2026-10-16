import os
import io
import re
import sys
import datetime
import threading
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageOps
import docx
import pymupdf  # PyMuPDF
import customtkinter as ctk

# Word -> PDF dönüştürücü desteği
try:
    from docx2pdf import convert as docx_to_pdf_convert
except ImportError:
    docx_to_pdf_convert = None

# TkinterDnD2 CustomTkinter Entegrasyonu
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.TkdndVersion = TkinterDnD._require(self)
except ImportError:
    CTkDnD = ctk.CTk
    DND_FILES = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

DESTEKLENEN_UZANTILAR = ('.pdf', '.jpg', '.jpeg', '.png', '.webp', '.bmp', '.udf', '.docx')


def sistem_turkce_font_bul():
    windir = os.environ.get("WINDIR", "C:\\Windows")
    font_adaylari = [
        os.path.join(windir, "Fonts", "arial.ttf"),
        os.path.join(windir, "Fonts", "calibri.ttf"),
        os.path.join(windir, "Fonts", "segoeui.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc"
    ]
    for yol in font_adaylari:
        if os.path.exists(yol):
            return yol
    return None

TURKCE_FONT_YOLU = sistem_turkce_font_bul()


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def tc_kimlik_dogrula(tc_no: str) -> bool:
    if not re.match(r'^[1-9]\d{10}$', tc_no):
        return False
    b = [int(ch) for ch in tc_no]
    tekler = sum(b[0:9:2])
    ciftler = sum(b[1:8:2])
    if ((tekler * 7) - ciftler) % 10 != b[9]:
        return False
    if sum(b[:10]) % 10 != b[10]:
        return False
    return True


def sayfa_araligi_ayristir(metin: str, toplam_sayfa: int):
    if not metin or not metin.strip():
        return set(range(toplam_sayfa))

    secilenler = set()
    parcalar = [p.strip() for p in metin.split(',') if p.strip()]
    if not parcalar:
        return set(range(toplam_sayfa))

    for p in parcalar:
        if '-' in p:
            alt = p.split('-')
            if len(alt) != 2:
                raise ValueError(f"Geçersiz aralık formatı: '{p}'")
            b, s = int(alt[0].strip()), int(alt[1].strip())
            if b > s:
                raise ValueError(f"Aralık başlangıcı bitişten büyük olamaz: '{p}'")
            if b < 1 or s > toplam_sayfa:
                raise ValueError(f"Sayfa aralığı 1 ile {toplam_sayfa} arasında olmalıdır: '{p}'")
            for i in range(b, s + 1):
                secilenler.add(i - 1)
        else:
            s = int(p)
            if s < 1 or s > toplam_sayfa:
                raise ValueError(f"Geçersiz sayfa numarası: {s} (Belge 1-{toplam_sayfa} arasıdır)")
            secilenler.add(s - 1)
    return secilenler


class ModernUyapAsistani(CTkDnD):
    def __init__(self):
        super().__init__()
        self.title("UYAP Evrak Asistanı")
        self.geometry("1100x700")
        self.minsize(1040, 650)
        self.configure(fg_color="#F8FAFC")

        self.dosyalar = []
        self._uygulama_ikonu_yukle()
        self._arayuz_olustur()

    def _uygulama_ikonu_yukle(self):
        ico = resource_path("logo.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(default=ico)
            except Exception:
                pass

    def _arayuz_olustur(self):
        # 1. ÜST HEADER
        header = ctk.CTkFrame(self, fg_color="#FFFFFF", height=64, corner_radius=0, border_width=1, border_color="#E2E8F0")
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        logo_alani = ctk.CTkFrame(header, fg_color="transparent")
        logo_alani.pack(side="left", padx=24, pady=12)

        ctk.CTkLabel(
            logo_alani,
            text="⚖ UYAP Evrak Asistanı",
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color="#0F172A"
        ).pack(side="left")

        ctk.CTkLabel(
            logo_alani,
            text="PRO",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            fg_color="#EEF2F6",
            text_color="#475569",
            corner_radius=6,
            padx=7,
            pady=2
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            header,
            text="Rehber & Kılavuz",
            command=self.rehber_penceresi_ac,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#334155",
            corner_radius=8,
            width=120,
            height=32
        ).pack(side="right", padx=24)

        # 2. ALT DURUM ÇUBUĞU
        alt_bar = ctk.CTkFrame(self, fg_color="#FFFFFF", height=38, corner_radius=0, border_width=1, border_color="#E2E8F0")
        alt_bar.pack(fill="x", side="bottom")
        alt_bar.pack_propagate(False)

        self.progress_bar = ctk.CTkProgressBar(alt_bar, width=140, height=8, mode="indeterminate", corner_radius=4)
        self.progress_bar.pack(side="right", padx=16, pady=14)
        self.progress_bar.set(0)

        self.durum_etiketi = ctk.CTkLabel(
            alt_bar,
            text="Hazır. Evraklarınızı sol alana sürükleyebilir veya 'Dosya Ekle' ile seçebilirsiniz.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#64748B"
        )
        self.durum_etiketi.pack(side="left", padx=18)

        # 3. ANA GÖVDE
        govde = ctk.CTkFrame(self, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=20, pady=16)

        # SOL KART
        sol_kart = ctk.CTkFrame(govde, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        sol_kart.pack(side="left", fill="both", expand=True, padx=(0, 10))

        sol_baslik_satir = ctk.CTkFrame(sol_kart, fg_color="transparent")
        sol_baslik_satir.pack(fill="x", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            sol_baslik_satir,
            text="İŞLENECEK DOSYALAR",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#94A3B8"
        ).pack(side="left")

        self.dosya_sayisi_badge = ctk.CTkLabel(
            sol_baslik_satir,
            text="0 dosya",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#64748B"
        )
        self.dosya_sayisi_badge.pack(side="right")

        self.liste_frame = ctk.CTkScrollableFrame(sol_kart, fg_color="#F8FAFC", corner_radius=8)
        self.liste_frame.pack(fill="both", expand=True, padx=16, pady=6)

        self.bos_bilgi_etiketi = ctk.CTkLabel(
            self.liste_frame,
            text="Belgeleri buraya sürükleyin\nveya aşağıdaki butondan ekleyin\n(PDF, Word, Görsel, UDF)",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94A3B8"
        )
        self.bos_bilgi_etiketi.pack(pady=120)

        if DND_FILES:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self.dosyalari_surukle_birak)
            except Exception:
                pass

        liste_alt_bar = ctk.CTkFrame(sol_kart, fg_color="transparent")
        liste_alt_bar.pack(fill="x", padx=16, pady=(10, 16))

        ctk.CTkButton(
            liste_alt_bar,
            text="➕ Dosya Ekle",
            command=self.dosya_ekle_dialog,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0F172A",
            hover_color="#1E293B",
            corner_radius=8,
            height=34
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            liste_alt_bar,
            text="▲",
            command=lambda: self._dosya_sirala(-1),
            width=36,
            height=34,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#334155",
            corner_radius=8
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            liste_alt_bar,
            text="▼",
            command=lambda: self._dosya_sirala(1),
            width=36,
            height=34,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#334155",
            corner_radius=8
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            liste_alt_bar,
            text="Kaldır",
            command=self.secileni_sil,
            width=65,
            height=34,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#334155",
            corner_radius=8
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            liste_alt_bar,
            text="Temizle",
            command=self.listeyi_temizle,
            width=65,
            height=34,
            fg_color="#FEE2E2",
            hover_color="#FECACA",
            text_color="#991B1B",
            corner_radius=8
        ).pack(side="right")

        # SAĞ KART: Yüksek Kontrastlı, Dinamik Renk Senkronizasyonlu Sekme Alanı
        sag_kart = ctk.CTkFrame(govde, fg_color="#FFFFFF", width=470, corner_radius=12, border_width=1, border_color="#E2E8F0")
        sag_kart.pack(side="right", fill="both")
        sag_kart.pack_propagate(False)

        self.tabview = ctk.CTkTabview(
            sag_kart,
            fg_color="transparent",
            segmented_button_fg_color="#E2E8F0",
            segmented_button_selected_color="#1E293B",
            segmented_button_selected_hover_color="#0F172A",
            segmented_button_unselected_color="#F8FAFC",
            segmented_button_unselected_hover_color="#E2E8F0",
            corner_radius=8
        )
        self.tabview.pack(fill="both", expand=True, padx=14, pady=10)

        # Seçili sekme metnini BEYAZ, seçili olmayanları KOYU yapan senkronizasyon fonksiyonu
        def _sekme_renklerini_guncelle(secilen_sekme):
            try:
                for ad, btn in self.tabview._segmented_button._buttons_dict.items():
                    if ad == secilen_sekme:
                        btn.configure(text_color="#FFFFFF")
                    else:
                        btn.configure(text_color="#334155")
            except Exception:
                pass

        self._sekme_renklerini_guncelle = _sekme_renklerini_guncelle

        try:
            self.tabview._segmented_button.configure(
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                height=38,
                command=lambda val: (self._sekme_renklerini_guncelle(val), self.tabview.set(val))
            )
        except Exception:
            pass

        self.tab_birlestir = self.tabview.add("Birleştir")
        self.tab_duzenle = self.tabview.add("Sayfalar")
        self.tab_donusturucu = self.tabview.add("Dönüştür")
        self.tab_guvenlik = self.tabview.add("KVKK & Kaşe")

        self._sekme_birlestir_olustur()
        self._sekme_duzenle_olustur()
        self._sekme_donusturucu_olustur()
        self._sekme_guvenlik_olustur()

        # İlk açılışta seçili olan sekmenin metin rengini beyaz yap
        self._sekme_renklerini_guncelle(self.tabview.get())

    # ========================================================
    # SEKME 1: BİRLEŞTİR
    # ========================================================
    def _sekme_birlestir_olustur(self):
        tab = self.tab_birlestir

        ctk.CTkLabel(tab, text="SIKIŞTIRMA MODU", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(4, 4))
        self.mod_secimi = ctk.StringVar(value="akilli_vektor")

        ctk.CTkRadioButton(
            tab,
            text="Akıllı Mod (Metinleri Koru, Görselleri Sıkıştır)",
            variable=self.mod_secimi,
            value="akilli_vektor",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#334155"
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            tab,
            text="Agresif Raster (150 DPI Taranmış Evrak Küçültme)",
            variable=self.mod_secimi,
            value="maksimum_raster",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#334155"
        ).pack(anchor="w", pady=3)

        ctk.CTkLabel(tab, text="BATES NUMARALANDIRMA", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(14, 4))
        bates_frame = ctk.CTkFrame(tab, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        bates_frame.pack(fill="x", pady=2)

        self.bates_aktif = ctk.BooleanVar(value=False)
        self.bates_switch = ctk.CTkSwitch(
            bates_frame,
            text="Sayfa Altına Numara Bas",
            variable=self.bates_aktif,
            command=self._bates_toggle,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#334155"
        )
        self.bates_switch.pack(side="left", padx=12, pady=10)

        self.bates_on_ek = ctk.CTkEntry(bates_frame, width=90, height=28, placeholder_text="EK-1", state="disabled")
        self.bates_on_ek.insert(0, "EK-1")
        self.bates_on_ek.pack(side="right", padx=12)

        ctk.CTkLabel(tab, text="EYLEMLER", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(16, 6))

        ctk.CTkButton(
            tab,
            text="📑 Listedeki Belgeleri Birleştir (PDF Yap)",
            command=lambda: self._thread_calistir(self.tumunu_birlestir_pdf_yap),
            fg_color="#0F172A",
            hover_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            height=40,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🗜 Seçili PDF'i Sıkıştır & Küçült",
            command=lambda: self._thread_calistir(self.pdf_tekil_sikistir),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="⚖ UYAP Akıllı MB Bölücüsü (Hızlı Parçala)",
            command=lambda: self._thread_calistir(self.pdf_akilli_boyut_bol),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 2: SAYFALAR
    # ========================================================
    def _sekme_duzenle_olustur(self):
        tab = self.tab_duzenle

        ctk.CTkLabel(tab, text="SAYFA AYIRMA VE DÖNDÜRME", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🔄 Sayfa Döndür (90° / 180°)",
            command=self.pdf_dondur_penceresi,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🧹 Boş Sayfaları Temizle (Akıllı Filtre)",
            command=lambda: self._thread_calistir(self.pdf_bos_sayfalari_temizle),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="✂ Sayfa Sayfa Böl (Her Sayfayı Tek Dosya Yap)",
            command=lambda: self._thread_calistir(self.pdf_sayfa_sayfa_bol),
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🎯 Belirli Sayfa Aralığını Ayır",
            command=self.pdf_aralik_bol,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 3: DÖNÜŞTÜR
    # ========================================================
    def _sekme_donusturucu_olustur(self):
        tab = self.tab_donusturucu

        ctk.CTkLabel(tab, text="FORMAT DÖNÜŞÜMLERİ", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(4, 6))

        bilgi_kutusu = ctk.CTkFrame(tab, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        bilgi_kutusu.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            bilgi_kutusu,
            text="💡 Dönüştürmek istediğiniz dosyayı sol listeden seçin\nveya doğrudan ilgili butona basıp dosya seçin.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#64748B",
            justify="left"
        ).pack(anchor="w", padx=12, pady=8)

        ctk.CTkButton(
            tab,
            text="📄 UDF ➔ Word (.docx)",
            command=lambda: self._thread_calistir(self.udf_to_word),
            fg_color="#0F172A",
            hover_color="#1E293B",
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="📑 Word (.docx) ➔ PDF",
            command=lambda: self._thread_calistir(self.word_to_pdf),
            fg_color="#0F172A",
            hover_color="#1E293B",
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🔍 Evraktaki Metinleri Görüntüle / Kopyala",
            command=self.metin_oku_penceresi,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 4: KVKK & KAŞE
    # ========================================================
    def _sekme_guvenlik_olustur(self):
        tab = self.tab_guvenlik

        ctk.CTkLabel(tab, text="HUKUKİ İŞLEMLER VE GÜVENLİK", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color="#94A3B8").pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🛡 'Aslı Gibidir' Dijital Kaşesi Bas",
            command=self.pdf_kase_ekle_penceresi,
            fg_color="#0F172A",
            hover_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🕶 KVKK Akıllı Karartma (T.C., IBAN, Tel)",
            command=self.akilli_karartma_penceresi,
            fg_color="#F1F5F9",
            hover_color="#E2E8F0",
            text_color="#1E293B",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    def _bates_toggle(self):
        if self.bates_aktif.get():
            self.bates_on_ek.configure(state="normal")
        else:
            self.bates_on_ek.configure(state="disabled")

    # ========================================================
    # LİSTE VE DOSYA YÖNETİMİ
    # ========================================================
    def _liste_arayuzunu_yenile(self):
        for widget in self.liste_frame.winfo_children():
            widget.destroy()

        if not self.dosyalar:
            self.bos_bilgi_etiketi = ctk.CTkLabel(
                self.liste_frame,
                text="Belgeleri buraya sürükleyin\nveya aşağıdaki butondan ekleyin\n(PDF, Word, Görsel, UDF)",
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color="#94A3B8"
            )
            self.bos_bilgi_etiketi.pack(pady=120)
            self.dosya_sayisi_badge.configure(text="0 dosya")
            self.secili_dosya_indeksi = None
            return

        self.secili_dosya_indeksi = getattr(self, "secili_dosya_indeksi", None)
        if self.secili_dosya_indeksi is not None and self.secili_dosya_indeksi >= len(self.dosyalar):
            self.secili_dosya_indeksi = len(self.dosyalar) - 1

        self.dosya_satir_widgetleri = []
        for i, yol in enumerate(self.dosyalar):
            uzanti = os.path.splitext(yol)[1].lower()
            boyut_mb = os.path.getsize(yol) / (1024 * 1024)
            simge_haritasi = {'.pdf': '📄', '.udf': '⚖', '.docx': '📝'}
            simge = simge_haritasi.get(uzanti, '🖼')
            ad = os.path.basename(yol)

            is_selected = (self.secili_dosya_indeksi == i)
            arkaplan = "#E2E8F0" if is_selected else "#FFFFFF"

            satir = ctk.CTkFrame(self.liste_frame, fg_color=arkaplan, corner_radius=8, height=36)
            satir.pack(fill="x", pady=2)
            satir.pack_propagate(False)

            def secim_yap(idx=i):
                self.secili_dosya_indeksi = idx
                self._liste_arayuzunu_yenile()

            satir.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))

            lbl_ad = ctk.CTkLabel(
                satir,
                text=f"{simge}  {ad}",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold" if is_selected else "normal"),
                text_color="#0F172A"
            )
            lbl_ad.pack(side="left", padx=10)
            lbl_ad.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))

            lbl_boyut = ctk.CTkLabel(
                satir,
                text=f"{boyut_mb:.2f} MB",
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color="#64748B"
            )
            lbl_boyut.pack(side="right", padx=10)
            lbl_boyut.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))

            self.dosya_satir_widgetleri.append(satir)

        self.dosya_sayisi_badge.configure(text=f"{len(self.dosyalar)} dosya")

    def dosya_ekle_listeye(self, dosya_yolu):
        dosya_yolu = dosya_yolu.strip('{}').strip('"')
        if not os.path.exists(dosya_yolu):
            return
        if os.path.isdir(dosya_yolu):
            for root, _, files in os.walk(dosya_yolu):
                for f in sorted(files):
                    self.dosya_ekle_listeye(os.path.join(root, f))
            return

        uzanti = os.path.splitext(dosya_yolu)[1].lower()
        if uzanti in DESTEKLENEN_UZANTILAR and dosya_yolu not in self.dosyalar:
            self.dosyalar.append(dosya_yolu)
            self._liste_arayuzunu_yenile()

    def dosyalari_surukle_birak(self, event):
        yollar = self.tk.splitlist(event.data)
        for y in yollar:
            self.dosya_ekle_listeye(y)
        self.durum_etiketi.configure(text=f"Toplam {len(self.dosyalar)} dosya hazır.")

    def dosya_ekle_dialog(self):
        secilenler = filedialog.askopenfilenames(
            title="Dosyaları Seçin",
            filetypes=[
                ("Tüm Desteklenen Evraklar", "*.pdf *.docx *.udf *.jpg *.jpeg *.png *.webp *.bmp"),
                ("PDF Dosyaları", "*.pdf"),
                ("Word Belgeleri", "*.docx"),
                ("UDF Dosyaları", "*.udf"),
                ("Görsel Dosyaları", "*.jpg *.jpeg *.png *.webp *.bmp")
            ]
        )
        for d in secilenler:
            self.dosya_ekle_listeye(d)
        self.durum_etiketi.configure(text=f"Toplam {len(self.dosyalar)} dosya hazır.")

    def secileni_sil(self):
        if hasattr(self, "secili_dosya_indeksi") and self.secili_dosya_indeksi is not None:
            if 0 <= self.secili_dosya_indeksi < len(self.dosyalar):
                del self.dosyalar[self.secili_dosya_indeksi]
                self.secili_dosya_indeksi = None
                self._liste_arayuzunu_yenile()
                self.durum_etiketi.configure(text="Seçilen belge kaldırıldı.")

    def listeyi_temizle(self):
        self.dosyalar.clear()
        self.secili_dosya_indeksi = None
        self._liste_arayuzunu_yenile()
        self.durum_etiketi.configure(text="Liste temizlendi.")

    def _dosya_sirala(self, yon):
        if not hasattr(self, "secili_dosya_indeksi") or self.secili_dosya_indeksi is None:
            return
        eski = self.secili_dosya_indeksi
        yeni = eski + yon
        if 0 <= yeni < len(self.dosyalar):
            self.dosyalar[eski], self.dosyalar[yeni] = self.dosyalar[yeni], self.dosyalar[eski]
            self.secili_dosya_indeksi = yeni
            self._liste_arayuzunu_yenile()

    def _secili_dosya_al(self):
        if not hasattr(self, "secili_dosya_indeksi") or self.secili_dosya_indeksi is None:
            if self.dosyalar:
                self.secili_dosya_indeksi = 0
                self._liste_arayuzunu_yenile()
                return self.dosyalar[0]
            messagebox.showwarning("Seçim Yapılmadı", "Lütfen sol taraftaki listeden bir evrak seçin.")
            return None
        return self.dosyalar[self.secili_dosya_indeksi]

    def _ayni_dosya_kontrolu(self, kaynak_yol, hedef_yol):
        if not hedef_yol:
            return False
        if os.path.abspath(kaynak_yol) == os.path.abspath(hedef_yol):
            messagebox.showerror("Hata", "Hedef dosya ile kaynak dosya aynı olamaz! Lütfen farklı bir isimle kaydedin.")
            return False
        return True

    def _thread_calistir(self, fn):
        def gorev():
            self.after(0, self.progress_bar.start)
            try:
                fn()
            finally:
                self.after(0, self.progress_bar.stop)
                self.after(0, lambda: self.progress_bar.set(0))
        threading.Thread(target=gorev, daemon=True).start()

    # ========================================================
    # 1. BİRLEŞTİRME VE OPTİMİZASYON İŞLEMLERİ
    # ========================================================
    def _gorseli_a4_sayfaya_ekle(self, gorsel_yolu, hedef_doc):
        with Image.open(gorsel_yolu) as img:
            img = ImageOps.exif_transpose(img)
            if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                arka_plan = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                mask = img.split()[-1]
                arka_plan.paste(img, mask=mask)
                img = arka_plan
            elif img.mode != "RGB":
                img = img.convert("RGB")

            img_w, img_h = img.size
            if img_w > img_h:
                a4_w, a4_h = 842.0, 595.0
            else:
                a4_w, a4_h = 595.0, 842.0

            yeni_sayfa = hedef_doc.new_page(width=a4_w, height=a4_h)
            margin = 28.0
            kullanilabilir_w = a4_w - (2 * margin)
            kullanilabilir_h = a4_h - (2 * margin)

            oran = min(kullanilabilir_w / img_w, kullanilabilir_h / img_h)
            cizim_w = img_w * oran
            cizim_h = img_h * oran

            x0 = (a4_w - cizim_w) / 2.0
            y0 = (a4_h - cizim_h) / 2.0
            rect = pymupdf.Rect(x0, y0, x0 + cizim_w, y0 + cizim_h)

            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75, optimize=True)
            yeni_sayfa.insert_image(rect, stream=buf.getvalue())

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
            rect = pymupdf.Rect(0, h - 24, w, h - 6)
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
            messagebox.showwarning("Evrak Yok", "Birleştirmek için listeye PDF veya görsel ekleyin.")
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Birleştirilmiş PDF'i Kaydet",
            initialfile="uyap_birlestirilmis_evrak.pdf"
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.configure(text="Evraklar birleştiriliyor...")
            toplam_eski_boyut = sum(os.path.getsize(d) for d in pdf_ve_resimler) / (1024 * 1024)

            yeni_doc = pymupdf.open()
            mod = self.mod_secimi.get()

            for dosya in pdf_ve_resimler:
                uzanti = os.path.splitext(dosya)[1].lower()
                if uzanti == '.pdf':
                    with pymupdf.open(dosya) as k_doc:
                        if mod == "maksimum_raster":
                            for sayfa in k_doc:
                                pix = sayfa.get_pixmap(dpi=150)
                                resim_bytes = pix.pil_tobytes(format="JPEG", quality=65, optimize=True)
                                s = yeni_doc.new_page(width=sayfa.rect.width, height=sayfa.rect.height)
                                s.insert_image(sayfa.rect, stream=resim_bytes)
                        else:
                            yeni_doc.insert_pdf(k_doc)
                else:
                    self._gorseli_a4_sayfaya_ekle(dosya, yeni_doc)

            self._bates_numaralari_bas(yeni_doc)
            yeni_doc.save(hedef_yol, garbage=4, deflate=True, clean=True)
            yeni_doc.close()

            yeni_boyut = os.path.getsize(hedef_yol) / (1024 * 1024)
            self.durum_etiketi.configure(text=f"Birleştirme tamamlandı: {yeni_boyut:.2f} MB")
            messagebox.showinfo("Başarılı", f"PDF'ler birleştirildi!\n\nEski Toplam: {toplam_eski_boyut:.2f} MB\nYeni Dosya: {yeni_boyut:.2f} MB")
        except Exception as e:
            messagebox.showerror("Hata", f"Birleştirme sırasında hata oluştu:\n{str(e)}")

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
            initialfile="kucuk_" + os.path.basename(kaynak_yol)
        )
        if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
            return

        try:
            self.durum_etiketi.configure(text="PDF optimize ediliyor...")
            mod = self.mod_secimi.get()

            with pymupdf.open(kaynak_yol) as doc:
                if mod == "maksimum_raster":
                    yeni_doc = pymupdf.open()
                    for sayfa in doc:
                        pix = sayfa.get_pixmap(dpi=150)
                        resim_bytes = pix.pil_tobytes(format="JPEG", quality=60, optimize=True)
                        s = yeni_doc.new_page(width=sayfa.rect.width, height=sayfa.rect.height)
                        s.insert_image(sayfa.rect, stream=resim_bytes)
                    self._bates_numaralari_bas(yeni_doc)
                    yeni_doc.save(hedef_yol, garbage=4, deflate=True)
                    yeni_doc.close()
                else:
                    self._bates_numaralari_bas(doc)
                    doc.save(hedef_yol, garbage=4, deflate=True, clean=True)

            yeni_boyut = os.path.getsize(hedef_yol) / (1024 * 1024)
            if yeni_boyut >= eski_boyut:
                messagebox.showinfo("Bilgi", f"Belge zaten optimize durumdaydı.\nEski: {eski_boyut:.2f} MB\nYeni: {yeni_boyut:.2f} MB")
            else:
                kazanc = ((eski_boyut - yeni_boyut) / eski_boyut) * 100
                messagebox.showinfo("Başarılı", f"Sıkıştırma tamamlandı!\nEski: {eski_boyut:.2f} MB\nYeni: {yeni_boyut:.2f} MB\nKazanılan: %{kazanc:.1f}")

            self.durum_etiketi.configure(text=f"İşlem tamamlandı: {yeni_boyut:.2f} MB")
        except Exception as e:
            messagebox.showerror("Hata", f"Sıkıştırma hatası:\n{str(e)}")

    def pdf_akilli_boyut_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Akıllı Boyut Bölücüsü yalnızca PDF dosyaları içindir.")
            return

        dosya_mb = os.path.getsize(kaynak_yol) / (1024 * 1024)
        limit_str = simpledialog.askstring(
            "Hedef Sınır (MB)",
            f"Dosya Boyutu: {dosya_mb:.2f} MB\n\nHer parçanın azami boyutunu (MB) girin:\n(UYAP için 9.5 MB önerilir)",
            initialvalue="9.5",
            parent=self
        )
        if not limit_str:
            return

        try:
            hedef_limit_mb = float(limit_str.replace(',', '.'))
            if hedef_limit_mb <= 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("Hata", "Lütfen geçerli bir MB değeri girin!")
            return

        hedef_limit_bytes = int(hedef_limit_mb * 1024 * 1024)
        hedef_klasor = filedialog.askdirectory(title="Parçaların Kaydedileceği Klasörü Seçin")
        if not hedef_klasor:
            return

        try:
            self.durum_etiketi.configure(text="Sayfalar tartılıyor ve bölünüyor...")
            uretilenler = []
            dosya_kok = os.path.splitext(os.path.basename(kaynak_yol))[0]

            with pymupdf.open(kaynak_yol) as ana_doc:
                toplam_s = len(ana_doc)
                parca_no = 1
                baslangic_idx = 0

                while baslangic_idx < toplam_s:
                    parca_doc = pymupdf.open()
                    parca_doc.insert_pdf(ana_doc, from_page=baslangic_idx, to_page=baslangic_idx)

                    ilk_sayfa_boyut = len(parca_doc.tobytes(garbage=3, deflate=True))
                    if ilk_sayfa_boyut > hedef_limit_bytes:
                        parca_doc.close()
                        messagebox.showwarning(
                            "Aşırı Büyük Sayfa",
                            f"Sayfa {baslangic_idx + 1} tek başına ({ilk_sayfa_boyut / (1024*1024):.2f} MB) seçilen {hedef_limit_mb} MB sınırını aşıyor!\nBu sayfa ayrı bir parça olarak kaydedilecektir."
                        )
                        parca_doc = pymupdf.open()
                        parca_doc.insert_pdf(ana_doc, from_page=baslangic_idx, to_page=baslangic_idx)
                        cikti = os.path.join(hedef_klasor, f"{dosya_kok}_parca_{parca_no}.pdf")
                        parca_doc.save(cikti, garbage=4, deflate=True)
                        parca_doc.close()
                        uretilenler.append((os.path.basename(cikti), os.path.getsize(cikti) / (1024 * 1024)))
                        parca_no += 1
                        baslangic_idx += 1
                        continue

                    son_gecerli = baslangic_idx
                    mevcut_idx = baslangic_idx + 1

                    while mevcut_idx < toplam_s:
                        parca_doc.insert_pdf(ana_doc, from_page=mevcut_idx, to_page=mevcut_idx)
                        boyut = len(parca_doc.tobytes(garbage=3, deflate=True))
                        if boyut > hedef_limit_bytes:
                            parca_doc.delete_page(len(parca_doc) - 1)
                            break
                        else:
                            son_gecerli = mevcut_idx
                            mevcut_idx += 1

                    cikti = os.path.join(hedef_klasor, f"{dosya_kok}_parca_{parca_no}.pdf")
                    parca_doc.save(cikti, garbage=4, deflate=True)
                    parca_doc.close()

                    uretilenler.append((os.path.basename(cikti), os.path.getsize(cikti) / (1024 * 1024)))
                    parca_no += 1
                    baslangic_idx = son_gecerli + 1

            ozet = "\n".join([f"• {ad}: {b:.2f} MB" for ad, b in uretilenler])
            self.durum_etiketi.configure(text=f"{len(uretilenler)} parça üretildi.")
            messagebox.showinfo("Bölme Tamamlandı", f"Belge {len(uretilenler)} parçaya bölündü:\n\n{ozet}\n\nKlasör: {hedef_klasor}")
        except Exception as e:
            messagebox.showerror("Hata", f"Boyut bölme hatası:\n{str(e)}")

    # ========================================================
    # 2. SAYFALAR İŞLEMLERİ
    # ========================================================
    def pdf_dondur_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Sayfa Döndürme yalnızca PDF dosyaları içindir.")
            return

        with pymupdf.open(kaynak_yol) as d_temp:
            toplam_sayfa = len(d_temp)

        dialog = ctk.CTkToplevel(self)
        dialog.title("Sayfa Döndür & Yön Düzelt")
        dialog.geometry("400x260")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text=f"Toplam Sayfa: {toplam_sayfa}", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")).pack(pady=(14, 6))

        derece_var = ctk.IntVar(value=90)
        radio_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        radio_bar.pack(pady=4)

        ctk.CTkRadioButton(radio_bar, text="90° Sağa", variable=derece_var, value=90).pack(side="left", padx=8)
        ctk.CTkRadioButton(radio_bar, text="90° Sola", variable=derece_var, value=270).pack(side="left", padx=8)
        ctk.CTkRadioButton(radio_bar, text="180°", variable=derece_var, value=180).pack(side="left", padx=8)

        ctk.CTkLabel(dialog, text="Döndürülecek Sayfalar (Örn: 1-3, 5 | Boşsa tümü):", font=ctk.CTkFont(family="Segoe UI", size=11)).pack(pady=(10, 2))
        aralik_entry = ctk.CTkEntry(dialog, width=280)
        aralik_entry.pack(pady=2)

        def dondur_kaydet():
            try:
                secilenler = sayfa_araligi_ayristir(aralik_entry.get().strip(), toplam_sayfa)
            except ValueError as ve:
                messagebox.showerror("Aralık Hatası", str(ve))
                return

            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Döndürülmüş PDF'i Kaydet",
                initialfile="dondurulmus_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
                return

            try:
                with pymupdf.open(kaynak_yol) as doc:
                    derece = derece_var.get()
                    for idx in secilenler:
                        sayfa = doc[idx]
                        sayfa.set_rotation((sayfa.rotation + derece) % 360)
                    doc.save(hedef_yol, garbage=4, deflate=True)

                dialog.destroy()
                self.durum_etiketi.configure(text=f"{len(secilenler)} sayfa döndürüldü.")
                messagebox.showinfo("Başarılı", f"{len(secilenler)} sayfa döndürüldü ve kaydedildi!")
            except Exception as e:
                messagebox.showerror("Hata", f"Döndürme hatası:\n{str(e)}")

        ctk.CTkButton(dialog, text="Uygula ve Kaydet", command=dondur_kaydet, fg_color="#0F172A", hover_color="#1E293B", corner_radius=8).pack(pady=16)

    def pdf_bos_sayfalari_temizle(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Boş sayfa temizleme yalnızca PDF dosyaları içindir.")
            return

        silinecekler = []
        try:
            self.durum_etiketi.configure(text="Sayfalar taranıyor...")
            with pymupdf.open(kaynak_yol) as doc:
                toplam_s = len(doc)
                for i in range(toplam_s):
                    sayfa = doc[i]
                    if sayfa.get_text().strip():
                        continue

                    pix = sayfa.get_pixmap(dpi=36, colorspace=pymupdf.csGRAY)
                    img = Image.frombytes("L", [pix.width, pix.height], pix.samples)
                    hist = img.histogram()
                    toplam_piksel = pix.width * pix.height
                    oran = sum(hist[:240]) / float(toplam_piksel)

                    if oran < 0.0008:
                        silinecekler.append(i)

            if not silinecekler:
                messagebox.showinfo("Bilgi", "Belgede boş sayfa tespit edilmedi.")
                self.durum_etiketi.configure(text="Boş sayfa bulunamadı.")
                return

            if len(silinecekler) == toplam_s:
                messagebox.showwarning("Uyarı", "Belgedeki tüm sayfalar boş görünüyor. Dosyanın tamamı silinemez.")
                return

            silinecek_str = ", ".join([str(s + 1) for s in silinecekler])
            onay = messagebox.askyesno(
                "Boş Sayfa Silme Onayı",
                f"Şu sayfalar boş tespit edildi:\nSayfa No: {silinecek_str}\n\nBu sayfalar silinip yeni PDF oluşturulsun mu?"
            )
            if not onay:
                return

            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Temizlenmiş PDF'i Kaydet",
                initialfile="temiz_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
                return

            with pymupdf.open(kaynak_yol) as doc:
                yeni_doc = pymupdf.open()
                silinecek_set = set(silinecekler)
                for idx in range(len(doc)):
                    if idx not in silinecek_set:
                        yeni_doc.insert_pdf(doc, from_page=idx, to_page=idx)
                yeni_doc.save(hedef_yol, garbage=4, deflate=True)
                yeni_doc.close()

            self.durum_etiketi.configure(text=f"{len(silinecekler)} boş sayfa temizlendi.")
            messagebox.showinfo("Başarılı", f"Toplam {len(silinecekler)} sayfa temizlendi!\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Boş sayfa temizleme hatası:\n{str(e)}")

    def pdf_sayfa_sayfa_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Bu özellik yalnızca PDF dosyaları içindir.")
            return

        hedef_klasor = filedialog.askdirectory(title="Sayfaların Kaydedileceği Klasörü Seçin")
        if not hedef_klasor:
            return

        try:
            with pymupdf.open(kaynak_yol) as doc:
                toplam = len(doc)
                kok = os.path.splitext(os.path.basename(kaynak_yol))[0]
                for i in range(toplam):
                    tekil = pymupdf.open()
                    tekil.insert_pdf(doc, from_page=i, to_page=i)
                    tekil.save(os.path.join(hedef_klasor, f"{kok}_sayfa_{i + 1}.pdf"), garbage=4, deflate=True)
                    tekil.close()

            messagebox.showinfo("Başarılı", f"{toplam} sayfa başarıyla ayrıldı!\nKlasör: {hedef_klasor}")
        except Exception as e:
            messagebox.showerror("Hata", f"Sayfa bölme hatası:\n{str(e)}")

    def pdf_aralik_bol(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Bu özellik yalnızca PDF dosyaları içindir.")
            return

        with pymupdf.open(kaynak_yol) as d:
            toplam_sayfa = len(d)

        aralik_metni = simpledialog.askstring("Sayfa Seçimi", f"Toplam Sayfa: {toplam_sayfa}\nAlmak istediğiniz sayfalar (Örn: 1-3, 5):", parent=self)
        if not aralik_metni:
            return

        try:
            secilenler = sayfa_araligi_ayristir(aralik_metni.strip(), toplam_sayfa)
        except ValueError as ve:
            messagebox.showerror("Hata", str(ve))
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Ayrılan Sayfaları Kaydet",
            initialfile=f"ayrilan_{os.path.basename(kaynak_yol)}"
        )
        if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
            return

        try:
            with pymupdf.open(kaynak_yol) as doc:
                yeni = pymupdf.open()
                for idx in sorted(list(secilenler)):
                    yeni.insert_pdf(doc, from_page=idx, to_page=idx)
                yeni.save(hedef_yol, garbage=4, deflate=True)
                yeni.close()

            messagebox.showinfo("Başarılı", f"Seçilen {len(secilenler)} sayfa ayrıldı ve kaydedildi!")
        except Exception as e:
            messagebox.showerror("Hata", f"İşlem hatası:\n{str(e)}")

    # ========================================================
    # 3. DÖNÜŞTÜR İŞLEMLERİ
    # ========================================================
    def udf_to_word(self):
        kaynak_yol = None
        if hasattr(self, "secili_dosya_indeksi") and self.secili_dosya_indeksi is not None and self.dosyalar:
            aday = self.dosyalar[self.secili_dosya_indeksi]
            if aday.lower().endswith('.udf'):
                kaynak_yol = aday

        if not kaynak_yol:
            kaynak_yol = filedialog.askopenfilename(
                title="Dönüştürülecek UDF Dosyasını Seçin",
                filetypes=[("UYAP UDF Dosyaları", "*.udf")]
            )

        if not kaynak_yol:
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
            self.durum_etiketi.configure(text="UDF içeriği Word'e aktarılıyor...")
            with zipfile.ZipFile(kaynak_yol, 'r') as zf:
                if 'content.xml' not in zf.namelist():
                    raise ValueError("Geçersiz UDF: 'content.xml' bulunamadı.")
                xml_raw = zf.read('content.xml')

            root_xml = ET.fromstring(xml_raw)
            yeni_docx = docx.Document()
            paragraf_eklendi = False

            content_node = root_xml.find('content')
            if content_node is not None and content_node.text and content_node.text.strip():
                for satir in content_node.text.splitlines():
                    yeni_docx.add_paragraph(satir.replace('\u200B', ''))
                paragraf_eklendi = True

            if not paragraf_eklendi:
                elements_node = root_xml.find('elements')
                if elements_node is not None:
                    p_list = elements_node.findall('.//paragraph')
                    if p_list:
                        for elem in p_list:
                            p_metin = "".join(elem.itertext()).strip().replace('\u200B', '')
                            yeni_docx.add_paragraph(p_metin)
                        paragraf_eklendi = True

            if not paragraf_eklendi:
                tum_metin = "".join(root_xml.itertext()).replace('\u200B', '')
                for satir in tum_metin.splitlines():
                    if satir.strip():
                        yeni_docx.add_paragraph(satir)

            yeni_docx.save(hedef_yol)
            self.durum_etiketi.configure(text="UDF ➔ Word aktarımı tamamlandı.")
            messagebox.showinfo("Başarılı", f"UDF belgesi Word formatına dönüştürüldü!\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Word dönüşüm hatası:\n{str(e)}")

    def word_to_pdf(self):
        kaynak_yol = None
        if hasattr(self, "secili_dosya_indeksi") and self.secili_dosya_indeksi is not None and self.dosyalar:
            aday = self.dosyalar[self.secili_dosya_indeksi]
            if aday.lower().endswith('.docx'):
                kaynak_yol = aday

        if not kaynak_yol:
            kaynak_yol = filedialog.askopenfilename(
                title="Dönüştürülecek Word (.docx) Dosyasını Seçin",
                filetypes=[("Word Belgeleri", "*.docx")]
            )

        if not kaynak_yol:
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="PDF Olarak Kaydet",
            initialfile=os.path.splitext(os.path.basename(kaynak_yol))[0] + ".pdf"
        )
        if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
            return

        try:
            self.durum_etiketi.configure(text="Word belgesi PDF'e dönüştürülüyor...")
            if docx_to_pdf_convert is not None:
                docx_to_pdf_convert(kaynak_yol, hedef_yol)
            else:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(os.path.abspath(kaynak_yol))
                doc.SaveAs(os.path.abspath(hedef_yol), FileFormat=17)
                doc.Close()
                word.Quit()

            self.durum_etiketi.configure(text="Word ➔ PDF dönüşümü tamamlandı.")
            messagebox.showinfo("Başarılı", f"Word belgesi başarıyla PDF'e dönüştürüldü!\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror(
                "Dönüşüm Hatası",
                f"Word ➔ PDF dönüştürme hatası:\n{str(e)}\n\n"
                "Not: Bilgisayarınızda Microsoft Word yüklü olmalı ve 'docx2pdf' kütüphanesi kurulu olmalıdır."
            )

    def metin_oku_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol:
            return

        self.durum_etiketi.configure(text="Metinler okunuyor...")
        metin_icerik = ""
        uzanti = os.path.splitext(kaynak_yol)[1].lower()

        try:
            if uzanti == '.pdf':
                with pymupdf.open(kaynak_yol) as doc:
                    for i, sayfa in enumerate(doc):
                        t = sayfa.get_text()
                        if t.strip():
                            metin_icerik += f"\n--- [ Sayfa {i + 1} ] ---\n" + t
            elif uzanti == '.docx':
                doc = docx.Document(kaynak_yol)
                metin_icerik = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
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
                metin_icerik = "Kopyalanabilir metin katmanı bulunamadı.\n(Belge taranmış bir görsel olabilir.)"

            pencere = ctk.CTkToplevel(self)
            pencere.title(f"Metin Önizleme: {os.path.basename(kaynak_yol)}")
            pencere.geometry("640x500")
            pencere.after(100, pencere.lift)

            txt = ctk.CTkTextbox(pencere, font=ctk.CTkFont(family="Consolas", size=12))
            txt.pack(fill="both", expand=True, padx=16, pady=(16, 10))
            txt.insert("0.0", metin_icerik)

            def kopyala():
                pencere.clipboard_clear()
                pencere.clipboard_append(txt.get("0.0", "end"))
                messagebox.showinfo("Kopyalandı", "Metin panoya kopyalandı!")

            ctk.CTkButton(pencere, text="📋 Panoya Kopyala", command=kopyala, fg_color="#0F172A", hover_color="#1E293B", corner_radius=8).pack(pady=(0, 16))
            self.durum_etiketi.configure(text="Metin hazır.")
        except Exception as e:
            messagebox.showerror("Hata", f"Metin okunamadı:\n{str(e)}")

    # ========================================================
    # 4. KVKK & KAŞE İŞLEMLERİ
    # ========================================================
    def pdf_kase_ekle_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Kaşe Ekleme yalnızca PDF dosyaları içindir.")
            return

        with pymupdf.open(kaynak_yol) as doc_temp:
            toplam_sayfa = len(doc_temp)

        bugun = datetime.datetime.now().strftime("%d.%m.%Y")

        dialog = ctk.CTkToplevel(self)
        dialog.title("Dijital 'Aslı Gibidir' Kaşesi")
        dialog.geometry("440x480")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text="Kaşe Başlığı:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(12, 2))
        baslik_entry = ctk.CTkEntry(dialog)
        baslik_entry.insert(0, "ASLI GİBİDİR")
        baslik_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Avukat İsim / Sicil No / Unvan:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        avukat_entry = ctk.CTkEntry(dialog, placeholder_text="Av. Ad Soyad - Baro Sicil No")
        avukat_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Tarih / Ek Açıklama:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        ek_entry = ctk.CTkEntry(dialog)
        ek_entry.insert(0, f"Tarih: {bugun}")
        ek_entry.pack(fill="x", padx=20)

        konum_var = ctk.StringVar(value="sag_alt")
        ctk.CTkLabel(dialog, text="Konum:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        konum_opt = ctk.CTkOptionMenu(
            dialog,
            values=["Sağ Alt", "Sol Alt", "Alt Orta", "Sağ Üst"],
            command=lambda v: konum_var.set({"Sağ Alt": "sag_alt", "Sol Alt": "sol_alt", "Alt Orta": "alt_orta", "Sağ Üst": "sag_ust"}[v])
        )
        konum_opt.pack(fill="x", padx=20)

        renk_var = ctk.StringVar(value="kirmizi")
        renk_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        renk_frame.pack(fill="x", padx=20, pady=8)
        ctk.CTkRadioButton(renk_frame, text="Kırmızı", variable=renk_var, value="kirmizi").pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(renk_frame, text="Lacivert", variable=renk_var, value="lacivert").pack(side="left")

        aralik_entry = ctk.CTkEntry(dialog, placeholder_text="Sayfa No (Boşsa Tümü | Örn: 1, 3-5)")
        aralik_entry.pack(fill="x", padx=20, pady=8)

        def uygula():
            try:
                secilenler = sayfa_araligi_ayristir(aralik_entry.get().strip(), toplam_sayfa)
            except ValueError as ve:
                messagebox.showerror("Hata", str(ve))
                return

            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Kaşeli PDF'i Kaydet",
                initialfile="kaseli_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
                return

            try:
                with pymupdf.open(kaynak_yol) as doc:
                    renk = (0.75, 0.05, 0.05) if renk_var.get() == "kirmizi" else (0.0, 0.18, 0.38)
                    satirlar = [baslik_entry.get().strip(), avukat_entry.get().strip(), ek_entry.get().strip()]
                    metin = "\n".join([s for s in satirlar if s])

                    font_param = {"fontfile": TURKCE_FONT_YOLU, "fontname": "arial_tr"} if TURKCE_FONT_YOLU else {"fontname": "helv"}
                    kutu_w, kutu_h = 175.0, 56.0
                    margin = 20.0
                    knm = konum_var.get()

                    for idx in secilenler:
                        sayfa = doc[idx]
                        w, h = sayfa.rect.width, sayfa.rect.height

                        if knm == "sag_alt":
                            x0, y0 = w - kutu_w - margin, h - kutu_h - margin
                        elif knm == "sol_alt":
                            x0, y0 = margin, h - kutu_h - margin
                        elif knm == "alt_orta":
                            x0, y0 = (w - kutu_w) / 2.0, h - kutu_h - margin
                        else:
                            x0, y0 = w - kutu_w - margin, margin

                        kutu_rect = pymupdf.Rect(x0, y0, x0 + kutu_w, y0 + kutu_h)
                        sayfa.draw_rect(kutu_rect, color=renk, width=1.5)

                        ic_rect = pymupdf.Rect(kutu_rect.x0 + 4, kutu_rect.y0 + 4, kutu_rect.x1 - 4, kutu_rect.y1 - 4)
                        f_size = 9.0
                        while f_size >= 6.0:
                            if sayfa.insert_textbox(ic_rect, metin, fontsize=f_size, color=renk, align=pymupdf.TEXT_ALIGN_CENTER, **font_param) >= 0:
                                break
                            f_size -= 0.5

                    doc.save(hedef_yol, garbage=4, deflate=True)

                dialog.destroy()
                self.durum_etiketi.configure(text=f"Kaşe seçilen {len(secilenler)} sayfaya uygulandı.")
                messagebox.showinfo("Başarılı", f"Kaşe basıldı!\nDosya: {os.path.basename(hedef_yol)}")
            except Exception as e:
                messagebox.showerror("Hata", f"Kaşe basma hatası:\n{str(e)}")

        ctk.CTkButton(dialog, text="🛡 Kaşeyi Bas ve Kaydet", command=uygula, fg_color="#0F172A", hover_color="#1E293B", corner_radius=8).pack(pady=16)

    def akilli_karartma_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Akıllı Karartma yalnızca PDF dosyaları içindir.")
            return

        with pymupdf.open(kaynak_yol) as d_test:
            toplam_metin = "".join([s.get_text() for s in d_test]).strip()

        if not toplam_metin:
            messagebox.showwarning(
                "Metin Katmanı Yok",
                "Bu belge taranmış görsel içeriyor ve metin katmanı barındırmıyor.\n"
                "Otomatik karartma seçilebilir metinlerde uygulanabilir."
            )
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("KVKK Akıllı Karartma")
        dialog.geometry("440x360")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text="Otomatik Karartma Seçenekleri", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")).pack(anchor="w", padx=20, pady=(14, 6))

        var_tc = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="Resmi T.C. Kimlik Algoritmasıyla Doğrulanan Numaralar", variable=var_tc).pack(anchor="w", padx=20, pady=4)

        var_iban = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="IBAN Numaraları (TR...)", variable=var_iban).pack(anchor="w", padx=20, pady=4)

        var_tel = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(dialog, text="Telefon Numaraları (05xx / +90)", variable=var_tel).pack(anchor="w", padx=20, pady=4)

        ctk.CTkLabel(dialog, text="Özel Maskelenecek Kelimeler (Virgülle ayırın):", font=ctk.CTkFont(family="Segoe UI", size=11)).pack(anchor="w", padx=20, pady=(12, 2))
        ozel_entry = ctk.CTkEntry(dialog, placeholder_text="Örn: Ahmet Yılmaz, Gizli Konu")
        ozel_entry.pack(fill="x", padx=20)

        def uygula():
            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Karartılmış PDF'i Kaydet",
                initialfile="karartilmis_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
                return

            try:
                toplam_mask = 0
                with pymupdf.open(kaynak_yol) as doc:
                    ozel_kelimeler = [k.strip() for k in ozel_entry.get().split(',') if k.strip()]

                    for sayfa in doc:
                        txt = sayfa.get_text()

                        if var_tc.get():
                            for match in re.finditer(r'\b[1-9]\d{10}\b', txt):
                                if tc_kimlik_dogrula(match.group()):
                                    for r in sayfa.search_for(match.group()):
                                        sayfa.add_redact_annot(r, fill=(0, 0, 0))
                                        toplam_mask += 1

                        if var_iban.get():
                            for match in re.finditer(r'TR\s*[0-9]{2}(?:\s*[0-9]{4}){5}\s*[0-9]{2}', txt):
                                for r in sayfa.search_for(match.group()):
                                    sayfa.add_redact_annot(r, fill=(0, 0, 0))
                                    toplam_mask += 1

                        if var_tel.get():
                            for match in re.finditer(r'(?:\+90\s*|0\s*)?5[0-9]{2}[\s\-]?[0-9]{3}[\s\-]?[0-9]{2}[\s\-]?[0-9]{2}', txt):
                                for r in sayfa.search_for(match.group()):
                                    sayfa.add_redact_annot(r, fill=(0, 0, 0))
                                    toplam_mask += 1

                        for k in ozel_kelimeler:
                            for r in sayfa.search_for(k):
                                sayfa.add_redact_annot(r, fill=(0, 0, 0))
                                toplam_mask += 1

                        sayfa.apply_redactions()

                    doc.scrub(attached_files=True, clean_pages=True, metadata=True)
                    doc.save(hedef_yol, garbage=4, deflate=True)

                dialog.destroy()
                self.durum_etiketi.configure(text=f"Karartma tamamlandı: {toplam_mask} veri maskelendi.")
                messagebox.showinfo("Başarılı", f"Kalıcı Karartma Uygulandı!\nToplam Maskelenen Veri: {toplam_mask} adet\nMetaveriler kazındı.")
            except Exception as e:
                messagebox.showerror("Hata", f"Karartma hatası:\n{str(e)}")

        ctk.CTkButton(dialog, text="🛡 Kalıcı Karart ve Kaydet", command=uygula, fg_color="#0F172A", hover_color="#1E293B", corner_radius=8).pack(pady=18)

    def rehber_penceresi_ac(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Kullanım Rehberi")
        pencere.geometry("640x520")
        pencere.resizable(False, False)
        pencere.after(100, pencere.lift)

        txt = ctk.CTkTextbox(pencere, font=ctk.CTkFont(family="Segoe UI", size=12))
        txt.pack(fill="both", expand=True, padx=16, pady=16)

        rehber = """UYAP EVRAK ASİSTANI - KULLANIM REHBERİ

1. Birleştir:
• Belgeleri sol alana sürükleyip 'PDF Yap' ile tek evrak haline getirebilirsiniz.
• Görsellerin (dekont, tutanak) yönü bozulmaz, otomatik A4 sayfasına oranlanır.
• Bates numaralandırma işaretlenirse sayfaların altına 'EK-1 | Sayfa 1/5' formatında numara basılır.

2. Sayfalar:
• Boş sayfaları temizleme butonu çift taraflı taramalardaki boş sayfaları tespit eder ve onayınızı alarak temizler.
• UYAP Akıllı MB Bölücüsü büyük dosyaları UYAP limitine (9.5 MB) göre sayfa bütünlüğünü bozmadan ayırır.

3. Dönüştür:
• UDF (.udf) dosyalarını tek tıkla Microsoft Word (.docx) formatına çevirebilirsiniz.
• Hazırladığınız Word (.docx) dilekçelerini doğrudan resmi UYAP uyumlu PDF'e dönüştürebilirsiniz.

4. KVKK & Kaşe:
• 'Aslı Gibidir' kaşesi belgenin yönüne duyarlı şekilde otomatik yerleşir.
• KVKK Karartma modülü yalnızca geçerli T.C. Kimlik numaralarını algoritmayla doğrulayarak ve IBAN'ları tespit ederek maskeler; ardından metaverileri PDF'ten tamamen kazır."""

        txt.insert("0.0", rehber.strip())
        txt.configure(state="disabled")


if __name__ == "__main__":
    app = ModernUyapAsistani()
    app.mainloop()