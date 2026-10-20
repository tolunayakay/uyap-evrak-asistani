# -*- coding: utf-8 -*-
import os
import io
import re
import sys
import json
import datetime
import threading
import zipfile
import base64
import webbrowser
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageOps, ImageDraw, ImageFilter
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pymupdf  # PyMuPDF
import customtkinter as ctk

try:
    import pythoncom
except ImportError:
    pythoncom = None

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


# ============================================================================
# PLATFORM (WINDOWS / MACOS) FARKLARI
# ============================================================================
IS_MAC = sys.platform == "darwin"

# Arayüz yazı tipleri: Windows'ta Segoe UI / Consolas, macOS'ta sistem yazı tipi / Menlo
ARAYUZ_FONTU = "Helvetica Neue" if IS_MAC else "Segoe UI"
SABIT_GENISLIK_FONTU = "Menlo" if IS_MAC else "Consolas"


if IS_MAC:
    # macOS kaydetme penceresi, adı uzantısıyla yazılınca uzantıyı ikinci kez ekliyor (dilekce.pdf.pdf)
    _asil_asksaveasfilename = filedialog.asksaveasfilename

    def _mac_asksaveasfilename(**secenekler):
        yol = _asil_asksaveasfilename(**secenekler)
        uzanti = secenekler.get("defaultextension") or ""
        if yol and uzanti and yol.lower().endswith((uzanti + uzanti).lower()):
            yol = yol[:-len(uzanti)]
        return yol

    filedialog.asksaveasfilename = _mac_asksaveasfilename

    # Bazı butonlar işlemin tamamını arka plan iş parçacığında çalıştırıyor ve pencereleri (kaydet, klasör,
    # onay, değer girişi) oradan açıyor. macOS'ta Tk pencereleri ana iş parçacığı dışında açılınca hata
    # veriyor; bu çağrılar ana iş parçacığına aktarılır ve sonuç beklenir.
    def _ana_is_parcaciginda(fonksiyon):
        def sarmal(*argumanlar, **secenekler):
            kok = tk._default_root
            if threading.current_thread() is threading.main_thread() or kok is None:
                return fonksiyon(*argumanlar, **secenekler)
            sonuc, bitti = {}, threading.Event()

            def calistir():
                try:
                    sonuc["deger"] = fonksiyon(*argumanlar, **secenekler)
                except BaseException as e:
                    sonuc["hata"] = e
                finally:
                    bitti.set()

            kok.after(0, calistir)
            bitti.wait()
            if "hata" in sonuc:
                raise sonuc["hata"]
            return sonuc.get("deger")
        return sarmal

    for _modul, _adlar in ((filedialog, ("asksaveasfilename", "askopenfilename", "askopenfilenames", "askdirectory")),
                           (messagebox, ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel", "askyesnocancel")),
                           (simpledialog, ("askstring", "askinteger", "askfloat"))):
        for _ad in _adlar:
            setattr(_modul, _ad, _ana_is_parcaciginda(getattr(_modul, _ad)))


def dosyayi_ac(yol: str):
    """Dosyayı işletim sisteminin varsayılan uygulamasıyla açar."""
    if os.name == "nt":
        os.startfile(yol)
    else:
        import subprocess
        subprocess.Popen(["open" if IS_MAC else "xdg-open", yol])


# ============================================================================
# GİZLİ AYAR YÖNETİCİSİ (APPDATA / CONFIG DİZİNİNDE SAKLAR)
# ============================================================================
def get_ayar_dosya_yolu() -> str:
    if os.name == 'nt':
        taban_dizin = os.environ.get('APPDATA', os.path.expanduser('~'))
        klasor = os.path.join(taban_dizin, 'UyapAsistani')
    elif IS_MAC:
        klasor = os.path.expanduser('~/Library/Application Support/UyapAsistani')
    else:
        klasor = os.path.expanduser('~/.config/uyap_asistani')

    try:
        os.makedirs(klasor, exist_ok=True)
    except Exception:
        klasor = os.path.expanduser('~')
    return os.path.join(klasor, 'app_settings.dat')


AYAR_DOSYA_YOLU = get_ayar_dosya_yolu()


def ayarlari_yukle() -> dict:
    varsayilan = {
        "tema": "Light",
        "avukat_bilgisi": "",
        "kase_baslik": "ASLI GİBİDİR",
        "kase_konum": "Sağ Alt",
        "kase_renk": "kirmizi",
        "filigran_metin": "TASLAKTIR",
        "filigran_renk": "Gri",
        "filigran_opaklik": 0.25,
        "filigran_arkada": True,
        "fihrist_aktif": False,
        "fihrist_baslik": "EK LİSTESİ",
        "fihrist_konum": "Başa",
        "ai_ozet_aktif": False,
        "ai_islem_birimi": "otomatik",
        "otomatik_guncelleme_kontrol": True
    }
    if os.path.exists(AYAR_DOSYA_YOLU):
        try:
            with open(AYAR_DOSYA_YOLU, "r", encoding="utf-8") as f:
                kayitli = json.load(f)
                varsayilan.update(kayitli)
        except Exception:
            pass
    return varsayilan


def ayar_kaydet(anahtar: str, deger):
    ayarlar = ayarlari_yukle()
    ayarlar[anahtar] = deger
    try:
        with open(AYAR_DOSYA_YOLU, "w", encoding="utf-8") as f:
            json.dump(ayarlar, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


_BASLANGIC_AYARLARI = ayarlari_yukle()
ctk.set_appearance_mode(_BASLANGIC_AYARLARI.get("tema", "Light"))
ctk.set_default_color_theme("blue")

# ============================================================================
# UYGULAMA VERSİYONU VE GITHUB GÜNCELLEME DENETLEYİCİSİ
# ============================================================================
APP_VERSION = "2.0.1"

# GitHub bilgileri — kendi repo adresinize göre güncelleyin
GITHUB_REPO_OWNER = "tolunayakay"          # GitHub kullanıcı adınız
GITHUB_REPO_NAME  = "uyap-evrak-asistani"  # Repository adı


def _versiyon_tuple(v: str):
    """'1.2.3' → (1, 2, 3) dönüşümü; karşılaştırma için."""
    try:
        return tuple(int(x) for x in v.lstrip("vV").split("."))
    except Exception:
        return (0,)


def guncelleme_kontrol_et() -> dict | None:
    """
    GitHub Releases API'sini sorgular.
    Yeni sürüm varsa {'surum': str, 'notlar': str, 'url': str} döner.
    Hata veya güncel ise None döner.
    """
    api_url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
    )
    try:
        req = urllib.request.Request(
            api_url,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"UyapEvrakAsistani/{APP_VERSION}"
            }
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            veri = json.loads(resp.read().decode("utf-8"))

        en_son_tag = veri.get("tag_name", "").lstrip("vV")
        notlar      = veri.get("body", "").strip() or "—"
        html_url    = veri.get("html_url", "")

        if _versiyon_tuple(en_son_tag) > _versiyon_tuple(APP_VERSION):
            return {"surum": en_son_tag, "notlar": notlar, "url": html_url}
    except Exception:
        pass
    return None


DESTEKLENEN_UZANTILAR = ('.pdf', '.jpg', '.jpeg', '.png', '.webp', '.bmp', '.udf', '.docx')


def sistem_turkce_font_bul():
    windir = os.environ.get("WINDIR", "C:\\Windows")
    font_adaylari = [
        os.path.join(windir, "Fonts", "arial.ttf"),
        os.path.join(windir, "Fonts", "calibri.ttf"),
        os.path.join(windir, "Fonts", "segoeui.ttf"),
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
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


# ============================================================================
# METAVERİ KAZIMA (PDF & WORD)
# ============================================================================
def pdf_metaveri_kazi(kaynak: str, hedef: str) -> list:
    """PDF'in yazar, üretici, tarih, XMP ve not yazarı bilgilerini siler.
    Bulunan (silinen) alanları 'Alan: değer' listesi olarak döndürür."""
    alan_adlari = {"author": "Yazar", "creator": "Oluşturan Program", "producer": "Üretici",
                   "title": "Başlık", "subject": "Konu", "keywords": "Anahtar Kelimeler",
                   "creationDate": "Oluşturma Tarihi", "modDate": "Değiştirme Tarihi"}
    bulunanlar = []
    with pymupdf.open(kaynak) as doc:
        if doc.needs_pass:
            raise ValueError("Belge şifreli; önce şifresini kaldırın.")

        for anahtar, etiket in alan_adlari.items():
            deger = (doc.metadata or {}).get(anahtar)
            if deger:
                bulunanlar.append(f"{etiket}: {deger}")
        if doc.get_xml_metadata().strip():
            bulunanlar.append("XMP metaverisi (gizli XML)")

        not_yazarlari = set()
        for sayfa in doc:
            # Açılır (popup) notlar dahil tüm not nesnelerinin yazar (/T) alanını sil
            for xref, _, _ in sayfa.annot_xrefs():
                tur, deger = doc.xref_get_key(xref, "T")
                if tur != "null":
                    if tur == "string" and deger:
                        not_yazarlari.add(deger)
                    doc.xref_set_key(xref, "T", "null")
        if not_yazarlari:
            bulunanlar.append("Not/Yorum Yazarları: " + ", ".join(sorted(not_yazarlari)))

        doc.set_metadata({})
        doc.del_xml_metadata()
        # garbage=4 ile eski artımlı kayıtlar (revizyon geçmişi) ve sahipsiz nesneler atılır
        doc.save(hedef, garbage=4, deflate=True, clean=True)
    return bulunanlar


def docx_metaveri_kazi(kaynak: str, hedef: str) -> list:
    """Word belgesinin yazar, son düzenleyen, şirket, düzenleme süresi, şablon yolu,
    revizyon kimlikleri (rsid) ve değişiklik/yorum yazarlarını siler veya anonimleştirir."""
    bulunanlar = []
    simdi = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def ilk(desen, metin):
        m = re.search(desen, metin, re.S)
        return m.group(1).strip() if m and m.group(1).strip() else None

    with zipfile.ZipFile(kaynak, "r") as zin:
        parcalar = {ad: zin.read(ad) for ad in zin.namelist()}
        bilgiler = {ad: zin.getinfo(ad) for ad in zin.namelist()}

    # --- docProps/core.xml : yazar, son düzenleyen, revizyon sayısı, tarihler ---
    if "docProps/core.xml" in parcalar:
        core = parcalar["docProps/core.xml"].decode("utf-8", errors="ignore")
        for etiket, desen in (("Yazar", r"<dc:creator>(.*?)</dc:creator>"),
                              ("Son Düzenleyen", r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>"),
                              ("Revizyon Sayısı", r"<cp:revision>(.*?)</cp:revision>"),
                              ("Başlık", r"<dc:title>(.*?)</dc:title>"),
                              ("Konu", r"<dc:subject>(.*?)</dc:subject>"),
                              ("Anahtar Kelimeler", r"<cp:keywords>(.*?)</cp:keywords>"),
                              ("Açıklama", r"<dc:description>(.*?)</dc:description>"),
                              ("Oluşturma Tarihi", r"<dcterms:created[^>]*>(.*?)</dcterms:created>"),
                              ("Son Yazdırma", r"<cp:lastPrinted>(.*?)</cp:lastPrinted>")):
            deger = ilk(desen, core)
            if deger and not (etiket == "Revizyon Sayısı" and deger == "1"):
                bulunanlar.append(f"{etiket}: {deger}")
        parcalar["docProps/core.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<cp:revision>1</cp:revision>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{simdi}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{simdi}</dcterms:modified>'
            '</cp:coreProperties>'
        ).encode("utf-8")

    # --- docProps/app.xml : şirket, yönetici, toplam düzenleme süresi, şablon ---
    if "docProps/app.xml" in parcalar:
        app = parcalar["docProps/app.xml"].decode("utf-8", errors="ignore")
        for etiket, tag in (("Şirket", "Company"), ("Yönetici", "Manager"),
                            ("Toplam Düzenleme Süresi (dk)", "TotalTime"), ("Köprü Tabanı", "HyperlinkBase")):
            deger = ilk(rf"<{tag}>(.*?)</{tag}>", app)
            if deger and deger != "0":
                bulunanlar.append(f"{etiket}: {deger}")
        app = re.sub(r"<(Company|Manager|HyperlinkBase)>.*?</\1>", "", app, flags=re.S)
        app = re.sub(r"<TotalTime>.*?</TotalTime>", "<TotalTime>0</TotalTime>", app, flags=re.S)
        app = re.sub(r"<Template>.*?</Template>", "<Template>Normal.dotm</Template>", app, flags=re.S)
        parcalar["docProps/app.xml"] = app.encode("utf-8")

    # --- docProps/custom.xml : özel özellikler ---
    if "docProps/custom.xml" in parcalar:
        custom = parcalar["docProps/custom.xml"].decode("utf-8", errors="ignore")
        adlar = re.findall(r'<property[^>]*\bname="([^"]*)"', custom)
        if adlar:
            bulunanlar.append("Özel Özellikler: " + ", ".join(adlar))
            custom = re.sub(r"<property\b.*?</property>", "", custom, flags=re.S)
            parcalar["docProps/custom.xml"] = custom.encode("utf-8")

    # --- word/*.xml : rsid revizyon kimlikleri, değişiklik/yorum yazarları, şablon yolu ---
    yazarlar = set()
    rsid_bulundu = False
    for ad in list(parcalar):
        if not (ad.startswith("word/") and ad.endswith(".xml")):
            continue
        xml = parcalar[ad].decode("utf-8", errors="ignore")
        orijinal = xml

        yazarlar.update(a for a in re.findall(r'\b(?:w|w15):author="([^"]*)"', xml) if a and a != "Anonim")
        xml = re.sub(r'\b((?:w|w15):author)="[^"]*"', r'\1="Anonim"', xml)
        xml = re.sub(r'w:initials="[^"]*"', 'w:initials="A"', xml)
        # Kişi kartlarındaki e-posta / kullanıcı kimliği (people.xml)
        xml = re.sub(r"<w15:presenceInfo\b[^>]*/>", "", xml)

        if re.search(r'\sw:rsid\w*="', xml) or "<w:rsid" in xml:
            rsid_bulundu = True
        xml = re.sub(r'\sw:rsid\w*="[^"]*"', "", xml)
        xml = re.sub(r"<w:rsids>.*?</w:rsids>", "", xml, flags=re.S)
        xml = re.sub(r"<w:rsid\b[^>]*/>", "", xml)

        if ad == "word/settings.xml":
            if "<w:attachedTemplate" in xml:
                bulunanlar.append("Bağlı Şablon Yolu (bilgisayar/kullanıcı klasörü)")
            xml = re.sub(r"<w:attachedTemplate\b[^>]*/>", "", xml)

        if xml != orijinal:
            parcalar[ad] = xml.encode("utf-8")

    rels = "word/_rels/settings.xml.rels"
    if rels in parcalar:
        r = parcalar[rels].decode("utf-8", errors="ignore")
        r = re.sub(r'<Relationship\b[^>]*attachedTemplate[^>]*/>', "", r)
        parcalar[rels] = r.encode("utf-8")

    if yazarlar:
        bulunanlar.append("Değişiklik/Yorum Yazarları: " + ", ".join(sorted(yazarlar)))
    if rsid_bulundu:
        bulunanlar.append("Revizyon geçmişi kimlikleri (rsid)")

    with zipfile.ZipFile(hedef, "w", zipfile.ZIP_DEFLATED) as zout:
        # [Content_Types].xml ilk sırada kalmalı
        for ad in sorted(parcalar, key=lambda a: a != "[Content_Types].xml"):
            bilgi = zipfile.ZipInfo(ad, date_time=(1980, 1, 1, 0, 0, 0))
            bilgi.compress_type = zipfile.ZIP_DEFLATED
            bilgi.external_attr = bilgiler[ad].external_attr
            zout.writestr(bilgi, parcalar[ad])
    return bulunanlar


# ============================================================================
# AI EVRAK ÖZETİ (GÖMÜLÜ / IN-PROCESS LLM — llama-cpp-python, CPU)
# ============================================================================
# Google Gemma 4 E2B (Apache 2.0 lisanslı; ticari kullanım serbest)
AI_MODEL_ADI = "Gemma 4 E2B"
AI_MODEL_DOSYA_ADI = "gemma-4-E2B-it-Q4_K_M.gguf"
AI_MODEL_URL = "https://huggingface.co/unsloth/gemma-4-E2B-it-GGUF/resolve/main/gemma-4-E2B-it-Q4_K_M.gguf"
AI_MODEL_BOYUT_METNI = "3,1 GB"
# Önceki sürümlerin indirdiği, artık kullanılmayan modeller (açılışta silinir)
AI_ESKI_MODEL_DOSYALARI = ("Qwen2.5-1.5B-Instruct-Q4_K_M.gguf",)
AI_BAGLAM_TOKEN = 16384          # n_ctx: ~15-20 A4 sayfa
AI_CIKTI_TOKEN = 1024            # özet için ayrılan azami token
AI_TEK_PARCA_SINIR = 30000       # bu uzunluğa kadar metin olduğu gibi verilir
AI_GIRIS_KARAKTER = 10000
AI_ORTA_KARAKTER = 5000
AI_SONUC_KARAKTER = 15000


def ai_model_klasoru() -> str:
    klasor = os.path.join(os.path.dirname(AYAR_DOSYA_YOLU), "models")
    os.makedirs(klasor, exist_ok=True)
    return klasor


def ai_model_yolu() -> str:
    return os.path.join(ai_model_klasoru(), AI_MODEL_DOSYA_ADI)


def ai_eski_modelleri_temizle() -> int:
    """Model değiştiğinde eski sürümün dosyalarını (ve yarım indirmelerini) siler. Kazanılan MB'yi döndürür."""
    kazanilan = 0
    klasor = os.path.join(os.path.dirname(AYAR_DOSYA_YOLU), "models")
    for ad in AI_ESKI_MODEL_DOSYALARI:
        for yol in (os.path.join(klasor, ad), os.path.join(klasor, ad + ".part")):
            try:
                if os.path.exists(yol):
                    boyut = os.path.getsize(yol)
                    os.remove(yol)
                    kazanilan += boyut
            except OSError:
                pass
    return kazanilan // (1024 * 1024)


def ai_model_indirildi() -> bool:
    yol = ai_model_yolu()
    try:
        if os.path.getsize(yol) < 100 * 1024 * 1024:
            return False
        with open(yol, "rb") as f:
            return f.read(4) == b"GGUF"
    except OSError:
        return False


def ai_kutuphane_mevcut() -> bool:
    import importlib.util
    return importlib.util.find_spec("llama_cpp") is not None


def belge_metni_cikar(yol: str) -> str:
    """PDF, Word ve UDF belgelerinden düz metin çıkarır (görsellerde boş döner)."""
    uzanti = os.path.splitext(yol)[1].lower()
    if uzanti == '.pdf':
        with pymupdf.open(yol) as doc:
            return "\n".join(s.get_text() for s in doc)
    if uzanti == '.docx':
        d = docx.Document(yol)
        parcalar = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for satir in t.rows:
                hucreler = []
                for h in satir.cells:
                    if h.text.strip() and (not hucreler or hucreler[-1] != h.text.strip()):
                        hucreler.append(h.text.strip())
                if hucreler:
                    parcalar.append(" | ".join(hucreler))
        return "\n".join(parcalar)
    if uzanti == '.udf':
        with zipfile.ZipFile(yol, 'r') as zf:
            if 'content.xml' in zf.namelist():
                kok = ET.fromstring(zf.read('content.xml'))
                c = kok.find('content')
                metin = c.text if c is not None and c.text else "".join(kok.itertext())
                return metin.replace('​', '')
    return ""


def ai_metni_hazirla(metin: str):
    """Token stratejisi: ≤30.000 karakter tek parça; daha uzunsa
    giriş (ilk 10.000) + gelişme (ortadan 5.000) + sonuç/hüküm (son 15.000).
    (hazir_metin, kirpildi_mi) döndürür."""
    metin = re.sub(r"[ \t ]+", " ", metin)
    metin = re.sub(r" *\n *", "\n", metin)
    metin = re.sub(r"\n{3,}", "\n\n", metin).strip()

    if len(metin) <= AI_TEK_PARCA_SINIR:
        return metin, False

    giris = metin[:AI_GIRIS_KARAKTER]
    sonuc = metin[-AI_SONUC_KARAKTER:]
    # Gelişme bölümü giriş ile sonuç arasındaki boşluğun tam ortasından alınır (çakışma olmaz)
    bosluk_bas, bosluk_son = AI_GIRIS_KARAKTER, len(metin) - AI_SONUC_KARAKTER
    bosluk = bosluk_son - bosluk_bas
    orta_bas = bosluk_bas + max(0, (bosluk - AI_ORTA_KARAKTER) // 2)
    orta = metin[orta_bas:min(bosluk_son, orta_bas + AI_ORTA_KARAKTER)]

    ayrac = "\n\n[... belgenin bu kısmı uzunluk nedeniyle atlanmıştır ...]\n\n"
    return giris + ayrac + orta + ayrac + sonuc, True


AI_SISTEM_MESAJI = (
    "Sen Türk hukukunda deneyimli bir avukat asistanısın. Görevin hukuki evrakı Türkçe, kısa ve doğru özetlemektir. "
    "Yalnızca belgede yazan bilgileri kullan; isim, tarih ve tutarları belgede yazdığı gibi aynen aktar, hiçbir bilgi uydurma. "
    "Belgede olmayan bilgi için 'Belirtilmemiş' yaz."
)

_DAVA_DISI = "- Bu evrak bir dava belgesi DEĞİLDİR: 'davacı', 'davalı', 'mahkeme', 'hüküm' veya 'karar verilmiştir' ifadelerini kullanma.\n"

# Her belge türü için: görünen ad, türe özel kurallar, özet biçimi
AI_BELGE_TURLERI = {
    "sozlesme": {
        "ad": "Sözleşme",
        "kurallar": _DAVA_DISI +
            "- Tarafları sözleşmedeki sıfatlarıyla yaz (kiraya veren / kiracı, satıcı / alıcı, işveren / işçi, "
            "yüklenici / iş sahibi, hizmet veren / hizmet alan vb.). Adres, kimlik ve vergi numarası yazma.\n"
            "- Sözleşme maddelerini aynen kopyalama; her bilgiyi en fazla 20 kelimelik kendi cümlenle özetle. "
            "Bir bilgiyi yalnızca bir başlık altında yaz, tekrar etme.\n"
            "- Sözleşmenin konusu, neyin karşılığında ne yapıldığıdır (örn. 'X adresindeki taşınmazın büro olarak kiralanması').\n"
            "- Bedel başlığında sözleşmedeki ana tutarı (kira bedeli, satış bedeli, hizmet bedeli) mutlaka yaz.\n",
        "bicim":
            "BELGE TÜRÜ: <sözleşmenin adı, örn. Kira Sözleşmesi>\n"
            "TARAFLAR:\n• <sıfatı>: <ad / unvan>\n• <sıfatı>: <ad / unvan>\n"
            "SÖZLEŞMENİN KONUSU: <tek cümle>\n"
            "SÜRE VE ÖNEMLİ TARİHLER:\n• <başlangıç, bitiş, yenileme, teslim tarihleri>\n"
            "BEDEL VE ÖDEME KOŞULLARI:\n• <ana tutar, ödeme günü, artış, depozito/teminat>\n"
            "TARAFLARIN TEMEL YÜKÜMLÜLÜKLERİ:\n• <en fazla 4 kısa madde>\n"
            "CEZAİ ŞART, FESİH VE UYUŞMAZLIK:\n• <cezai şart, fesih koşulları, yetkili mahkeme/icra dairesi veya tahkim>\n"
            "DİKKAT EDİLMESİ GEREKEN HUSUSLAR:\n• <en fazla 3 kısa madde>",
    },
    "dilekce": {
        "ad": "Dilekçe (Dava / Cevap / Beyan)",
        "kurallar":
            "- Davacı davayı açan kişi veya şirkettir; vekil onun avukatıdır. Bunları karıştırma.\n"
            "- Dilekçedeki talepler henüz karara bağlanmamıştır: 'talep edilmiştir' de, 'karar verilmiştir' deme.\n",
        "bicim":
            "BELGE TÜRÜ: <dava dilekçesi, cevap dilekçesi, beyan dilekçesi vb.>\n"
            "MAHKEME / DOSYA NO: <varsa>\n"
            "TARAFLAR:\n• Davacı: <ad>\n• Davalı: <ad>\n• Vekiller: <avukat adları>\n"
            "DAVA KONUSU: <tek cümle>\n"
            "ÖNEMLİ TARİHLER VE TUTARLAR:\n• <tarih veya tutar ve neyle ilgili olduğu>\n"
            "İDDİALAR / AÇIKLAMALAR:\n• <en fazla 6 kısa madde>\n"
            "DELİLLER: <kısaca>\n"
            "TALEP SONUCU: <dilekçede istenen>",
    },
    "karar": {
        "ad": "Mahkeme Kararı",
        "kurallar":
            "- Davacı davayı açan taraftır; vekiller tarafların avukatlarıdır. Davacı ile davacı vekilini karıştırma.\n"
            "- Hüküm bölümünü belgede yazdığı gibi, madde madde aktar.\n",
        "bicim":
            "BELGE TÜRÜ: <gerekçeli karar, ara karar vb.>\n"
            "MAHKEME / ESAS NO / KARAR NO: <varsa>\n"
            "TARAFLAR:\n• Davacı: <ad>\n• Davalı: <ad>\n• Vekiller: <avukat adları>\n"
            "DAVA KONUSU: <tek cümle>\n"
            "TARAFLARIN İDDİA VE SAVUNMALARI:\n• <kısaca>\n"
            "MAHKEMENİN GEREKÇESİ:\n• <en fazla 4 kısa madde>\n"
            "HÜKÜM:\n• <hüküm maddeleri>\n"
            "KANUN YOLU VE SÜRE: <istinaf/temyiz ve süresi>",
    },
    "ihtarname": {
        "ad": "İhtarname",
        "kurallar": _DAVA_DISI +
            "- İhtarnamede taraflar 'ihtar eden (keşideci)' ve 'muhatap'tır.\n",
        "bicim":
            "BELGE TÜRÜ: İhtarname\n"
            "İHTAR EDEN: <ad>\nMUHATAP: <ad>\nNOTERLİK / YEVMİYE NO: <varsa>\n"
            "İHTARIN KONUSU: <tek cümle>\n"
            "TALEP EDİLEN: <ödeme, tahliye, ifa vb.; tutar varsa yaz>\n"
            "VERİLEN SÜRE: <gün>\n"
            "AKSİ HALDE: <ihtar edenin başvuracağı yollar>",
    },
    "bilirkisi": {
        "ad": "Bilirkişi Raporu",
        "kurallar": "- Bilirkişinin tespitlerini ve vardığı sonucu belgede yazdığı gibi aktar; kendi yorumunu ekleme.\n",
        "bicim":
            "BELGE TÜRÜ: Bilirkişi Raporu\n"
            "MAHKEME / DOSYA NO: <varsa>\nBİLİRKİŞİ: <ad, uzmanlık>\n"
            "İNCELEME KONUSU: <tek cümle>\n"
            "YAPILAN İNCELEMELER:\n• <kısaca>\n"
            "TESPİTLER:\n• <en fazla 6 kısa madde, tutarlarla>\n"
            "SONUÇ VE KANAAT: <bilirkişinin vardığı sonuç>",
    },
    "icra": {
        "ad": "İcra Belgesi",
        "kurallar": "- İcra belgelerinde taraflar 'alacaklı' ve 'borçlu'dur; 'davacı' veya 'davalı' kullanma.\n",
        "bicim":
            "BELGE TÜRÜ: <ödeme emri, takip talebi, haciz tutanağı vb.>\n"
            "İCRA DAİRESİ / DOSYA NO: <varsa>\n"
            "ALACAKLI: <ad>\nBORÇLU: <ad>\nVEKİLLER: <varsa>\n"
            "TAKİP TÜRÜ: <ilamsız, ilamlı, kambiyo vb.>\n"
            "ALACAK KALEMLERİ:\n• <asıl alacak, faiz, masraf ve tutarları>\n"
            "SÜRELER VE UYARILAR:\n• <ödeme/itiraz süreleri>\n"
            "ÖZET: <kısaca>",
    },
    "genel": {
        "ad": "Diğer / Genel Belge",
        "kurallar": "- Belge bir dava belgesi değilse 'davacı' ve 'davalı' kelimelerini kullanma; kişileri belgedeki sıfatlarıyla yaz.\n",
        "bicim":
            "BELGE TÜRÜ: <belgenin ne olduğu>\n"
            "İLGİLİ KİŞİ VE KURUMLAR:\n• <sıfatı>: <ad / unvan>\n"
            "KONU: <tek cümle>\n"
            "ÖNEMLİ TARİHLER VE TUTARLAR:\n• <tarih veya tutar ve neyle ilgili olduğu>\n"
            "ÖZET:\n• <en fazla 6 kısa madde>\n"
            "SONUÇ / İSTENEN İŞLEM: <varsa>",
    },
}


def _tr_kucuk(metin: str) -> str:
    return metin.replace("İ", "i").replace("I", "ı").lower()


# (tür, desen, başlıkta bulunursa puan, gövdede bulunursa puan) — başlık: ilk 400 karakter
_TUR_IPUCLARI = [
    ("karar", r"gerekçeli karar", 7, 3),
    ("karar", r"\bhüküm\s*:", 4, 4),
    ("karar", r"karar no\s*:", 4, 2),
    ("karar", r"karar tarihi", 3, 2),
    ("karar", r"açıkça okunup usulen anlatıldı|oy birliği(yle)? karar verildi|oybirliğiyle karar", 4, 4),
    ("karar", r"(istinaf|temyiz) yolu açık", 2, 2),
    ("dilekce", r"hakimliğine|başkanlığına|mahkemesi'?ne\b", 5, 1),
    ("dilekce", r"(dava|cevap|beyan|itiraz) dilekçesi", 4, 1),
    ("dilekce", r"sonuç ve (istem|talep)", 3, 3),
    ("dilekce", r"arz ve talep ederi[mz]|talep ederi[mz]|arz ederi[mz]", 2, 3),
    ("dilekce", r"açıklamalar\s*:", 2, 2),
    ("sozlesme", r"sözleşme(si)?\b", 5, 0),
    ("sozlesme", r"işbu sözleşme", 3, 4),
    ("sozlesme", r"\bmadde\s*\d+\b|\b\d+\.\s*madde\b", 2, 3),
    ("sozlesme", r"kiraya veren|\bkiracı|\bsatıcı|\balıcı|\bişveren|\byüklenici|iş sahibi|hizmet (veren|alan)", 2, 2),
    ("sozlesme", r"imza altına alınmıştır|nüsha olarak|taraflarca (okunarak )?imza", 2, 3),
    ("ihtarname", r"ihtarname", 7, 2),
    ("ihtarname", r"ihtar eden|keşideci", 3, 3),
    ("ihtarname", r"\bmuhatap", 2, 2),
    ("bilirkisi", r"bilirkişi (kurulu )?raporu", 7, 2),
    ("bilirkisi", r"sonuç ve kanaat|kanaatimizdir|kanaatindeyim|takdiri sayın mahkemeye", 3, 3),
    ("bilirkisi", r"(inceleme|görev) konusu", 2, 2),
    ("icra", r"ödeme emri", 7, 3),
    ("icra", r"icra (müdürlüğü|dairesi)", 3, 2),
    ("icra", r"takip talebi|ilamsız takip|ilamlı takip|takip türü", 3, 2),
    ("icra", r"\balacaklı\s*:|\bborçlu\s*:", 2, 2),
]


def belge_turu_tespit(metin: str) -> str:
    """Anahtar kelime puanlamasıyla belge türünü belirler (küçük modele bırakmaktan daha güvenilir)."""
    govde = _tr_kucuk(metin[:60000])
    baslik = govde[:400]
    puan = {}
    for tur, desen, baslik_puani, govde_puani in _TUR_IPUCLARI:
        if re.search(desen, baslik):
            puan[tur] = puan.get(tur, 0) + baslik_puani
        elif re.search(desen, govde):
            puan[tur] = puan.get(tur, 0) + govde_puani
    if not puan:
        return "genel"
    en_iyi = max(puan, key=puan.get)
    return en_iyi if puan[en_iyi] >= 4 else "genel"


def ai_kullanici_mesaji(metin: str, tur: str) -> str:
    # Talimatlar metinden SONRA verilir: küçük modeller en son okuduklarına daha iyi uyar
    t = AI_BELGE_TURLERI.get(tur, AI_BELGE_TURLERI["genel"])
    return (
        f"=== EVRAK BAŞLANGICI ===\n{metin}\n=== EVRAK SONU ===\n\n"
        f"Bu evrakın türü: {t['ad'].upper()}.\n"
        "Yukarıdaki evrakı özetle. Dikkat edilecek kurallar:\n"
        f"{t['kurallar']}"
        "- İsim, tarih ve tutarları belgede yazdığı gibi aynen aktar.\n"
        "- Düz metin yaz; Markdown, ** veya # kullanma.\n\n"
        f"Özeti tam olarak şu biçimde yaz:\n{t['bicim']}"
    )


# ----------------------------------------------------------------------------
# AI İŞLEM BİRİMİ: İŞLEMCİ (CPU) / EKRAN KARTI (GPU, Vulkan)
# 'ai_motor' klasörü (ai_motor_hazirla.py ile üretilir) varsa resmi llama.cpp Vulkan
# DLL'leri kullanılır; yoksa llama-cpp-python'un yerleşik CPU kütüphanesine düşülür.
# Ekran kartı seçimi süreç başına bir kez yapılabildiği için kart algılama ve ilk
# doğrulama ayrı alt süreçte yapılır: bozuk bir sürücü çökerse ana uygulama etkilenmez.
# ----------------------------------------------------------------------------
AI_ISLEM_BIRIMLERI = {"otomatik": "Otomatik", "cpu": "İşlemci", "gpu": "Ekran Kartı"}
AI_GPU_ASGARI_BELLEK_MB = 3800      # Otomatik modda harici kartın modeli tamamen taşıyabilmesi için
_AI_MOTOR_DURUMU = None             # Bu süreçte hazırlanan motor: {"motor": "vulkan"|"yerlesik", "vk": int|None}
_AI_GPU_ONBELLEK = None             # Algılanan ekran kartları (süreç başına bir kez)
_GGML_TUR = {1: "harici", 2: "dahili"}   # ggml_backend_dev_type: 0=CPU, 1=GPU, 2=IGPU


def ai_motor_klasoru():
    taban = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    yol = os.path.join(taban, "ai_motor")
    gerekli = ("llama.dll", "ggml.dll", "ggml-base.dll", "ggml-vulkan.dll")
    return yol if all(os.path.isfile(os.path.join(yol, f)) for f in gerekli) else None


def _ggml_kutuphaneleri(motor):
    import ctypes
    reg = ctypes.CDLL(os.path.join(motor, "ggml.dll"))
    base = ctypes.CDLL(os.path.join(motor, "ggml-base.dll"))

    def fn(ad, restype, argtypes):
        for lib in (reg, base):
            try:
                f = getattr(lib, ad)
            except AttributeError:
                continue
            f.restype, f.argtypes = restype, argtypes
            return f
        raise AttributeError(ad)
    return fn


def _ai_motoru_hazirla(vk):
    """Arka uçları yükler. vk: görünür Vulkan cihazı (None = yalnızca işlemci). Süreç başına bir kez çalışır."""
    global _AI_MOTOR_DURUMU
    if _AI_MOTOR_DURUMU is not None:
        return _AI_MOTOR_DURUMU
    motor = ai_motor_klasoru()
    # llama_cpp önceden içe aktarıldıysa yerleşik DLL'leri yüklemiştir; farklı sürüm DLL'leri karıştırılmamalı
    if not motor or "llama_cpp.llama_cpp" in sys.modules:
        _AI_MOTOR_DURUMU = {"motor": "yerlesik", "vk": None}
        return _AI_MOTOR_DURUMU
    import ctypes
    os.environ["GGML_VK_VISIBLE_DEVICES"] = str(vk) if vk is not None else "-1"
    os.environ["LLAMA_CPP_LIB_PATH"] = motor
    os.add_dll_directory(motor)
    fn = _ggml_kutuphaneleri(motor)
    fn("ggml_backend_load_all_from_path", None, [ctypes.c_char_p])(motor.encode("mbcs"))
    _AI_MOTOR_DURUMU = {"motor": "vulkan", "vk": vk}
    return _AI_MOTOR_DURUMU


def _ai_alt_surec(argumanlar, zaman_asimi):
    """Uygulamanın kendisini alt süreç olarak çalıştırır; JSON sonucunu döndürür (çökme/zaman aşımında None)."""
    import subprocess, tempfile
    fd, cikti = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    komut = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(__file__)]
    try:
        subprocess.run(komut + list(argumanlar) + [cikti], timeout=zaman_asimi,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL,
                       creationflags=0x08000000 if os.name == "nt" else 0)   # CREATE_NO_WINDOW
        if os.path.getsize(cikti) > 0:
            with open(cikti, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    finally:
        try:
            os.remove(cikti)
        except OSError:
            pass
    return None


def ai_gpu_listele():
    """Ekran kartlarını alt süreçte algılar: [{'vk', 'ad', 'tur': 'harici'|'dahili', 'bellek_mb'}]."""
    global _AI_GPU_ONBELLEK
    if _AI_GPU_ONBELLEK is None:
        if not ai_motor_klasoru():
            _AI_GPU_ONBELLEK = []
        else:
            sonuc = _ai_alt_surec(["--ai-cihazlar"], 60) or []
            _AI_GPU_ONBELLEK = [
                {"vk": int(c["ad"][6:]), "ad": c["aciklama"], "tur": _GGML_TUR[c["tur"]], "bellek_mb": c["toplam_mb"]}
                for c in sonuc if c.get("tur") in _GGML_TUR and c.get("ad", "").startswith("Vulkan")
            ]
    return _AI_GPU_ONBELLEK


def ai_cihaz_sec(mod, gpular):
    """Ayara ve donanıma göre cihaz seçer: (vulkan_no veya None, açıklama)."""
    harici = sorted((g for g in gpular if g["tur"] == "harici"), key=lambda g: -g["bellek_mb"])
    dahili = [g for g in gpular if g["tur"] == "dahili"]
    if mod == "cpu":
        return None, "İşlemci (CPU)"
    if mod == "otomatik":
        # Dahili kartlar (Intel Iris Xe vb.) genelde işlemciden yavaştır; otomatikte atlanır
        if harici and harici[0]["bellek_mb"] >= AI_GPU_ASGARI_BELLEK_MB:
            return harici[0]["vk"], harici[0]["ad"]
        return None, "İşlemci (CPU)"
    aday = (harici or dahili or [None])[0]
    if aday is None:
        return None, "İşlemci (CPU) — ekran kartı bulunamadı"
    return aday["vk"], aday["ad"]


def ai_metal_destekli() -> bool:
    """macOS: kurulu llama-cpp-python Metal (ekran kartı) desteğiyle derlenmiş mi?"""
    if not IS_MAC:
        return False
    try:
        import llama_cpp
        return bool(llama_cpp.llama_supports_gpu_offload())
    except Exception:
        return False


def _ai_cihazlari_yaz(cikti):
    """Alt süreç: tüm Vulkan cihazlarını listeler."""
    import ctypes
    sonuc = []
    motor = ai_motor_klasoru()
    if motor:
        os.add_dll_directory(motor)
        fn = _ggml_kutuphaneleri(motor)
        fn("ggml_backend_load_all_from_path", None, [ctypes.c_char_p])(motor.encode("mbcs"))
        dev_get = fn("ggml_backend_dev_get", ctypes.c_void_p, [ctypes.c_size_t])
        ad = fn("ggml_backend_dev_name", ctypes.c_char_p, [ctypes.c_void_p])
        aciklama = fn("ggml_backend_dev_description", ctypes.c_char_p, [ctypes.c_void_p])
        tur = fn("ggml_backend_dev_type", ctypes.c_int, [ctypes.c_void_p])
        bellek = fn("ggml_backend_dev_memory", None, [ctypes.c_void_p, ctypes.POINTER(ctypes.c_size_t), ctypes.POINTER(ctypes.c_size_t)])
        for i in range(fn("ggml_backend_dev_count", ctypes.c_size_t, [])()):
            d = dev_get(i)
            bos, toplam = ctypes.c_size_t(), ctypes.c_size_t()
            bellek(d, ctypes.byref(bos), ctypes.byref(toplam))
            sonuc.append({"ad": ad(d).decode(errors="replace"), "aciklama": aciklama(d).decode(errors="replace"),
                          "tur": tur(d), "toplam_mb": toplam.value // 1048576})
    with open(cikti, "w", encoding="utf-8") as f:
        json.dump(sonuc, f, ensure_ascii=False)


def _ai_gpu_dogrula_yaz(vk, cikti):
    """Alt süreç: modeli verilen ekran kartında yükleyip kısa bir cevap üretir."""
    sonuc = {"basarili": False}
    try:
        _ai_motoru_hazirla(int(vk))
        from llama_cpp import Llama, LLAMA_SPLIT_MODE_NONE
        llm = Llama(model_path=ai_model_yolu(), n_ctx=2048, n_gpu_layers=-1, split_mode=LLAMA_SPLIT_MODE_NONE,
                    main_gpu=0, verbose=False)
        r = llm.create_chat_completion(messages=[{"role": "user", "content": "1+1 kaçtır? Yalnızca sayıyla cevap ver."}],
                                       max_tokens=4, temperature=0)
        sonuc = {"basarili": bool(r["choices"][0]["message"]["content"].strip())}
    except Exception as e:
        sonuc = {"basarili": False, "hata": f"{type(e).__name__}: {e}"[:300]}
    with open(cikti, "w", encoding="utf-8") as f:
        json.dump(sonuc, f, ensure_ascii=False)


class AiOzetMotoru:
    """llama-cpp-python modelini tek sefer yükler ve süreç içinde paylaşır."""
    _llm = None
    _kilit = threading.Lock()
    cihaz_aciklama = ""

    @classmethod
    def yuklu_mu(cls):
        return cls._llm is not None

    @classmethod
    def _cihaz_belirle(cls, durum_bildir=None):
        """(vulkan_no veya None, açıklama). Seçilen kart ilk kez kullanılacaksa alt süreçte doğrulanır."""
        ayarlar = ayarlari_yukle()
        mod = ayarlar.get("ai_islem_birimi", "otomatik")
        if IS_MAC:
            # Apple Silicon: llama.cpp Metal ile ekran kartını doğrudan kullanır (ayrı motor/doğrulama gerekmez)
            if mod != "cpu" and ai_metal_destekli():
                return "metal", "Apple ekran kartı (Metal)"
            return None, "İşlemci (CPU)"
        if mod == "cpu" or not ai_motor_klasoru():
            return None, "İşlemci (CPU)"
        if durum_bildir:
            durum_bildir("algiliyor")
        vk, aciklama = ai_cihaz_sec(mod, ai_gpu_listele())
        if vk is None:
            return None, aciklama
        dogrulama = dict(ayarlar.get("ai_gpu_dogrulama", {}))
        if aciklama not in dogrulama:
            if durum_bildir:
                durum_bildir("dogruluyor")
            sonuc = _ai_alt_surec(["--ai-gpu-dogrula", str(vk)], 240) or {"basarili": False}
            dogrulama[aciklama] = bool(sonuc.get("basarili"))
            ayar_kaydet("ai_gpu_dogrulama", dogrulama)
        if not dogrulama[aciklama]:
            return None, f"İşlemci (CPU) — {aciklama} doğrulanamadı"
        return vk, aciklama

    @classmethod
    def yukle(cls, durum_bildir=None):
        if cls._llm is None:
            vk, aciklama = cls._cihaz_belirle(durum_bildir)
            hazir = _ai_motoru_hazirla(None if vk == "metal" else vk)
            if vk is not None and vk != "metal" and hazir["vk"] != vk:
                # Arka uçlar bu oturumda başka ayarla yüklendi; kart değişikliği yeniden başlatınca geçerli olur
                vk, aciklama = None, "İşlemci (CPU) — kart değişikliği için uygulamayı yeniden başlatın"
            if durum_bildir:
                durum_bildir("yukleniyor")
            from llama_cpp import Llama, LLAMA_SPLIT_MODE_NONE
            cekirdek = os.cpu_count() or 4
            ortak = dict(
                model_path=ai_model_yolu(),
                n_ctx=AI_BAGLAM_TOKEN,
                n_threads=max(1, cekirdek // 2),     # üretim: fiziksel çekirdek sayısı civarı
                n_threads_batch=max(1, cekirdek),   # belge okuma (prefill): tüm çekirdekler
                n_batch=512,
                verbose=False,
            )
            if vk is not None:
                try:
                    cls._llm = Llama(n_gpu_layers=-1, split_mode=LLAMA_SPLIT_MODE_NONE, main_gpu=0, **ortak)
                except Exception as e:
                    # Ekran kartı belleği yetmedi vb. → işlemciye dön
                    cls._llm = None
                    aciklama = f"İşlemci (CPU) — ekran kartında hata: {str(e)[:80]}"
            if cls._llm is None:
                cls._llm = Llama(n_gpu_layers=0, **ortak)
            cls.cihaz_aciklama = aciklama
        return cls._llm

    @classmethod
    def bosalt(cls):
        """Modeli bellekten çıkarır (dosya kilidini bırakmak için silmeden önce gerekli)."""
        if cls._llm is not None:
            try:
                cls._llm.close()
            except Exception:
                pass
            cls._llm = None
            import gc
            gc.collect()

    @classmethod
    def mesajlari_hazirla(cls, llm, metin, tur="genel"):
        """Bağlam penceresine sığmayan metni orantılı kısaltır."""
        sinir = AI_BAGLAM_TOKEN - AI_CIKTI_TOKEN - 128
        for _ in range(6):
            mesajlar = [
                {"role": "system", "content": AI_SISTEM_MESAJI},
                {"role": "user", "content": ai_kullanici_mesaji(metin, tur)},
            ]
            kaba = "\n".join(m["content"] for m in mesajlar)
            token_sayisi = len(llm.tokenize(kaba.encode("utf-8"), add_bos=False)) + 32
            if token_sayisi <= sinir:
                return mesajlar, token_sayisi
            oran = sinir / token_sayisi * 0.95
            kes = int(len(metin) * oran)
            # Baştan ve sondan koru, ortadan kırp (hüküm kısmı sondadır)
            bas = int(kes * 0.4)
            metin = metin[:bas] + "\n\n[... kısaltıldı ...]\n\n" + metin[-(kes - bas):]
        return mesajlar, token_sayisi


# ============================================================================
# e-İMZA & ZAMAN DAMGASI DOĞRULAMA / TESPİT MOTORU
# ============================================================================
# --- Küçük ASN.1 / DER (BER) çözümleyici: PKCS#7 / CAdES imzalarından imzalayan sertifikayı okur ---
_OID_SIGNED_DATA = bytes.fromhex("2a864886f70d010702")
_OID_TSTINFO = bytes.fromhex("2a864886f70d0109100104")
_OID_IMZA_ZAMANI = bytes.fromhex("2a864886f70d010905")
_OID_ZAMAN_DAMGASI = bytes.fromhex("2a864886f70d010910020e")
_OID_AD = {bytes.fromhex("550403"): "CN", bytes.fromhex("550405"): "seri", bytes.fromhex("55040a"): "O",
           bytes.fromhex("55042a"): "ad", bytes.fromhex("550404"): "soyad", bytes.fromhex("55040c"): "unvan"}


def _der_oku(v, i):
    """(etiket, içerik_başı, içerik_sonu, sonraki_konum). BER belirsiz uzunluk da desteklenir."""
    etiket = v[i]
    i += 1
    if etiket & 0x1F == 0x1F:
        while v[i] & 0x80:
            i += 1
        i += 1
    uz = v[i]
    i += 1
    if uz == 0x80:
        bas = i
        while not (v[i] == 0 and v[i + 1] == 0):
            i = _der_oku(v, i)[3]
        return etiket, bas, i, i + 2
    if uz & 0x80:
        n = uz & 0x7F
        uz = int.from_bytes(v[i:i + n], "big")
        i += n
    if i + uz > len(v):
        raise ValueError("DER uzunluğu taşıyor")
    return etiket, i, i + uz, i + uz


def _der_cocuklar(v, bas, son):
    cocuklar, i = [], bas
    while i < son:
        if v[i] == 0 and i + 1 < son and v[i + 1] == 0:
            break
        etiket, b, e, i = _der_oku(v, i)
        cocuklar.append((etiket, b, e))
    return cocuklar


def _der_metin(etiket, ham):
    if etiket == 0x1E:
        return ham.decode("utf-16-be", errors="replace")
    if etiket == 0x1C:
        return ham.decode("utf-32-be", errors="replace")
    try:
        return ham.decode("utf-8")
    except UnicodeDecodeError:
        return ham.decode("iso-8859-9", errors="replace")


def _der_zaman(etiket, ham):
    try:
        s = ham.decode("ascii").rstrip("Z").split(".")[0].split("+")[0]
        if etiket == 0x17:      # UTCTime
            yil = int(s[:2])
            s = ("19" if yil >= 50 else "20") + s
        dt = datetime.datetime.strptime(s[:14], "%Y%m%d%H%M%S").replace(tzinfo=datetime.timezone.utc)
        return dt.astimezone().strftime("%d.%m.%Y %H:%M:%S")
    except Exception:
        return None


def _der_isim(v, b, e):
    """X.500 Name → {'CN': .., 'seri': .., ...}"""
    alanlar = {}
    for _, sb, se in _der_cocuklar(v, b, e):            # SET
        for _, ab, ae in _der_cocuklar(v, sb, se):      # AttributeTypeAndValue
            c = _der_cocuklar(v, ab, ae)
            if len(c) == 2 and c[0][0] == 0x06:
                anahtar = _OID_AD.get(v[c[0][1]:c[0][2]])
                if anahtar and anahtar not in alanlar:
                    alanlar[anahtar] = _der_metin(c[1][0], v[c[1][1]:c[1][2]]).strip()
    return alanlar


def _der_sertifika(v, b, e):
    tbs = _der_cocuklar(v, b, e)[0]
    c = _der_cocuklar(v, tbs[1], tbs[2])
    if c[0][0] == 0xA0:     # version [0]
        c = c[1:]
    # serialNumber, signature, issuer, validity, subject
    return {"seri": v[c[0][1]:c[0][2]], "yayinci_ham": v[c[2][1]:c[2][2]],
            "yayinci": _der_isim(v, c[2][1], c[2][2]), "konu": _der_isim(v, c[4][1], c[4][2])}


def _der_zaman_damgasi_saati(v, b, e):
    """Zaman damgası jetonu (ContentInfo) → TSTInfo.genTime"""
    try:
        ci = _der_cocuklar(v, b, e)
        sd = _der_cocuklar(v, ci[1][1], ci[1][2])[0]
        enc = _der_cocuklar(v, sd[1], sd[2])[2]
        enc_c = _der_cocuklar(v, enc[1], enc[2])
        oktet = _der_cocuklar(v, enc_c[1][1], enc_c[1][2])[0]
        tst = _der_cocuklar(v, oktet[1], oktet[2])[0]
        gen = _der_cocuklar(v, tst[1], tst[2])[4]
        return _der_zaman(gen[0], v[gen[1]:gen[2]])
    except Exception:
        return None


def pkcs7_imzalayanlari_coz(v: bytes) -> list:
    """Veri içindeki tüm PKCS#7/CAdES SignedData yapılarını bulur; zaman damgası jetonlarını atlar.
    [{'ad', 'tc', 'saglayici', 'imza_zamani', 'zaman_damgasi', 'zaman_damgasi_saati'}] döndürür."""
    sonuc = []
    desen = b"\x06" + bytes([len(_OID_SIGNED_DATA)]) + _OID_SIGNED_DATA
    konum = v.find(desen)
    while konum != -1:
        try:
            i = konum + len(desen)
            etiket, b, e, _ = _der_oku(v, i)
            if etiket == 0xA0:
                sd = _der_cocuklar(v, b, e)[0]
                c = _der_cocuklar(v, sd[1], sd[2])
                enc = _der_cocuklar(v, c[2][1], c[2][2])
                if v[enc[0][1]:enc[0][2]] != _OID_TSTINFO:      # zaman damgası jetonu değil
                    sertifikalar, imzacilar = [], []
                    for et, cb, ce in c[3:]:
                        if et == 0xA0:
                            for kt, kb, ke in _der_cocuklar(v, cb, ce):
                                if kt == 0x30:
                                    try:
                                        sertifikalar.append(_der_sertifika(v, kb, ke))
                                    except Exception:
                                        pass
                        elif et == 0x31:
                            imzacilar = _der_cocuklar(v, cb, ce)
                    sonuc.extend(_imzacilari_eslestir(v, sertifikalar, imzacilar))
        except Exception:
            pass
        konum = v.find(desen, konum + 1)
    # Aynı kişi birden çok kez (ör. paralel imza kopyaları) çıkmasın
    tekil, gorulen = [], set()
    for s in sonuc:
        if s["ad"] and s["ad"] not in gorulen:
            gorulen.add(s["ad"])
            tekil.append(s)
    return tekil


def _imzacilari_eslestir(v, sertifikalar, imzacilar):
    sonuc = []
    for _, sb, se in imzacilar:
        si = _der_cocuklar(v, sb, se)
        sertifika = None
        sid = si[1]
        if sid[0] == 0x30:      # issuerAndSerialNumber
            ic = _der_cocuklar(v, sid[1], sid[2])
            yayinci, seri = v[ic[0][1]:ic[0][2]], v[ic[1][1]:ic[1][2]]
            sertifika = next((s for s in sertifikalar if s["yayinci_ham"] == yayinci and s["seri"] == seri), None)
        if sertifika is None:
            # Eşleşmezse: başka bir sertifikayı yayınlamamış (son kullanıcı) sertifikayı al
            adaylar =[s for s in sertifikalar if s["konu"].get("CN") not in {x["yayinci"].get("CN") for x in sertifikalar}]
            sertifika = (adaylar or sertifikalar or [None])[0]
        imza_zamani, damga, damga_saati = None, False, None
        for et, ab, ae in si[3:]:
            if et in (0xA0, 0xA1):
                for _, xb, xe in _der_cocuklar(v, ab, ae):
                    at = _der_cocuklar(v, xb, xe)
                    oid = v[at[0][1]:at[0][2]]
                    degerler = _der_cocuklar(v, at[1][1], at[1][2])
                    if oid == _OID_IMZA_ZAMANI and degerler:
                        imza_zamani = _der_zaman(degerler[0][0], v[degerler[0][1]:degerler[0][2]])
                    elif oid == _OID_ZAMAN_DAMGASI and degerler:
                        damga = True
                        damga_saati = _der_zaman_damgasi_saati(v, degerler[0][1], degerler[0][2])
        if sertifika:
            konu = sertifika["konu"]
            ad = konu.get("CN") or " ".join(x for x in (konu.get("ad"), konu.get("soyad")) if x)
            sonuc.append({"ad": ad, "tc": konu.get("seri"), "unvan": konu.get("unvan"),
                          "saglayici": sertifika["yayinci"].get("CN") or sertifika["yayinci"].get("O"),
                          "imza_zamani": imza_zamani, "zaman_damgasi": damga, "zaman_damgasi_saati": damga_saati})
    return sonuc


def _xades_sertifikalari(metin: bytes) -> list:
    """XML (XAdES) imzalardaki <X509Certificate> içinden son kullanıcı sertifikasını çıkarır."""
    sertifikalar = []
    for blob in re.findall(rb"<(?:\w+:)?X509Certificate>([^<]+)</", metin):
        try:
            der = base64.b64decode(re.sub(rb"\s+", b"", blob))
            _, b, e, _ = _der_oku(der, 0)
            sertifikalar.append(_der_sertifika(der, b, e))
        except Exception:
            pass
    cn_yayinci = {s["yayinci"].get("CN") for s in sertifikalar}
    sonuc = []
    for s in sertifikalar:
        if s["konu"].get("CN") not in cn_yayinci:
            sonuc.append({"ad": s["konu"].get("CN"), "tc": s["konu"].get("seri"), "unvan": s["konu"].get("unvan"),
                          "saglayici": s["yayinci"].get("CN") or s["yayinci"].get("O"),
                          "imza_zamani": None, "zaman_damgasi": False, "zaman_damgasi_saati": None})
    return sonuc


def imza_verisinden_imzalayanlar(ham: bytes) -> list:
    """Ham imza dosyası (DER, BER, Base64/PEM veya XML) → imzalayan listesi."""
    adaylar = [ham]
    if not ham[:1] == b"\x30":
        metin = re.sub(rb"-----[A-Z0-9 ]+-----", b"", ham)
        try:
            adaylar.append(base64.b64decode(re.sub(rb"\s+", b"", metin), validate=False))
        except Exception:
            pass
    for veri in adaylar:
        imzacilar = pkcs7_imzalayanlari_coz(veri)
        if imzacilar:
            return imzacilar
    return _xades_sertifikalari(ham)


def _imzalayan_metni(imzacilar):
    return "; ".join(i["ad"] for i in imzacilar if i.get("ad"))


def _imza_sonucuna_yaz(sonuc, imzacilar, varsayilan):
    sonuc["imzalayanlar"] = imzacilar
    sonuc["imzalayan"] = _imzalayan_metni(imzacilar) or varsayilan
    saglayicilar = sorted({i["saglayici"] for i in imzacilar if i.get("saglayici")})
    sonuc["saglayici"] = ", ".join(saglayicilar) or None
    zamanlar = [i["imza_zamani"] for i in imzacilar if i.get("imza_zamani")]
    if zamanlar and not sonuc.get("tarih"):
        sonuc["tarih"] = zamanlar[0]
    if any(i.get("zaman_damgasi") for i in imzacilar):
        sonuc["zaman_damgasi"] = True
        sonuc["zaman_damgasi_saati"] = next((i["zaman_damgasi_saati"] for i in imzacilar if i.get("zaman_damgasi_saati")), None)


def eimza_bilgisi_al(dosya_yolu: str) -> dict:
    uzanti = os.path.splitext(dosya_yolu)[1].lower()
    sonuc = {
        "imzali": False,
        "imzalayan": None,
        "imzalayanlar": [],
        "saglayici": None,
        "zaman_damgasi": False,
        "zaman_damgasi_saati": None,
        "tarih": None,
        "detay": "Dijital imza tespit edilemedi."
    }

    if not os.path.exists(dosya_yolu):
        return sonuc

    try:
        if uzanti == '.udf':
            with zipfile.ZipFile(dosya_yolu, 'r') as zf:
                isimler = zf.namelist()
                imza_dosyalari = [f for f in isimler if 'sign' in f.lower() or 'signature' in f.lower()]
                
                if imza_dosyalari:
                    sonuc["imzali"] = True
                    imzacilar = []
                    for ad in imza_dosyalari:
                        imzacilar += [i for i in imza_verisinden_imzalayanlar(zf.read(ad))
                                      if i["ad"] not in {x["ad"] for x in imzacilar}]
                    _imza_sonucuna_yaz(sonuc, imzacilar, "Nitelikli Elektronik İmza (UDF)")
                    if not sonuc["zaman_damgasi"]:
                        ham_veri = b"".join(zf.read(ad) for ad in imza_dosyalari)
                        sonuc["zaman_damgasi"] = _OID_ZAMAN_DAMGASI in ham_veri or b"TimeStamp" in ham_veri
                    sonuc["detay"] = f"Geçerli UYAP e-İmzası\nİmzalayan: {sonuc['imzalayan']}"
                    if sonuc["zaman_damgasi"]:
                        sonuc["detay"] += "\nZaman Damgası: Onaylı (TSA Mevcut)"
                else:
                    if 'content.xml' in isimler:
                        xml_raw = zf.read('content.xml')
                        if b"Signature" in xml_raw or b"ds:Signature" in xml_raw:
                            sonuc["imzali"] = True
                            _imza_sonucuna_yaz(sonuc, _xades_sertifikalari(xml_raw), "UDF XML İmzası")
                            sonuc["detay"] = "UDF XML Formatında İmzalı Evrak"

        elif uzanti == '.pdf':
            with open(dosya_yolu, 'rb') as f:
                pdf_bytes = f.read()

            if b"/ByteRange" in pdf_bytes and (b"/Contents" in pdf_bytes or b"/adbe.pkcs7" in pdf_bytes):
                sonuc["imzali"] = True

                # PAdES: her imza sözlüğünün /Contents <hex> alanı bir PKCS#7/CAdES yapısıdır
                imzacilar = []
                for hex_blok in re.findall(rb'/Contents\s*<([0-9A-Fa-f\s]{200,})>', pdf_bytes):
                    try:
                        der = bytes.fromhex(re.sub(rb"\s+", b"", hex_blok).decode("ascii"))
                    except ValueError:
                        continue
                    imzacilar += [i for i in pkcs7_imzalayanlari_coz(der) if i["ad"] not in {x["ad"] for x in imzacilar}]

                name_eslesme = re.search(rb'/Name\s*\(([^)]+)\)', pdf_bytes)
                if not name_eslesme:
                    name_eslesme = re.search(rb'/Name\s*<([^>]+)>', pdf_bytes)

                if imzacilar:
                    _imza_sonucuna_yaz(sonuc, imzacilar, "Elektronik İmza (PAdES)")
                elif name_eslesme:
                    isim_raw = name_eslesme.group(1)
                    try:
                        sonuc["imzalayan"] = isim_raw.decode('utf-8', errors='ignore')
                    except Exception:
                        sonuc["imzalayan"] = isim_raw.decode('iso-8859-9', errors='ignore')
                else:
                    sonuc["imzalayan"] = "Elektronik İmza (PAdES)"

                date_eslesme = re.search(rb'/M\s*\(D:([0-9]{14})', pdf_bytes)
                if date_eslesme and not sonuc["tarih"]:
                    d_str = date_eslesme.group(1).decode('ascii')
                    try:
                        dt = datetime.datetime.strptime(d_str, "%Y%m%d%H%M%S")
                        sonuc["tarih"] = dt.strftime("%d.%m.%Y %H:%M:%S")
                    except Exception:
                        pass

                if b"/ETSI.RFC3161" in pdf_bytes or b"TimeStamp" in pdf_bytes:
                    sonuc["zaman_damgasi"] = True

                sonuc["detay"] = f"Geçerli PAdES e-İmzası\nİmzalayan: {sonuc['imzalayan']}"
                if sonuc["tarih"]:
                    sonuc["detay"] += f"\nİmza Tarihi: {sonuc['tarih']}"
                if sonuc["zaman_damgasi"]:
                    sonuc["detay"] += "\nResmi Zaman Damgası: Mevcut"
    except Exception:
        pass

    return sonuc


# ============================================================================
# GELİŞMİŞ UDF MOTORU (WORD & PDF -> UDF VE WORD DÖNÜŞTÜRÜCÜSÜ)
# ============================================================================
_XML_GECERSIZ = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_NS_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_NS_V = "{urn:schemas-microsoft-com:vml}"
_NS_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class UdfBelge:
    YAZI_TIPI = "Times New Roman"

    def __init__(self):
        self.paragraflar = []

    def paragraf_ekle(self, ogeler, hiza=0):
        self.paragraflar.append((hiza, list(ogeler)))

    def tablo_ekle(self, satirlar, genislikler):
        self.paragraflar.append(("TABLO", satirlar, list(genislikler)))

    @staticmethod
    def _temiz(metin):
        metin = _XML_GECERSIZ.sub("", metin).replace("\r", "").replace("\n", " ")
        return metin.replace("]]>", "] ]>")

    def xml_uret(self):
        icerik = []
        elemanlar = []
        durum = {"konum": 0}
        yazi = self.YAZI_TIPI

        def paragraf_xml(hiza, ogeler):
            konum = durum["konum"]
            temiz_ogeler = []
            for o in ogeler:
                if o[0] == "metin":
                    t = self._temiz(o[1])
                    if t:
                        temiz_ogeler.append(("metin", t, o[2] if len(o) > 2 and o[2] else {}))
                else:
                    temiz_ogeler.append(o)

            satirlar = ['<paragraph Alignment="%d">' % int(hiza)]
            if not temiz_ogeler:
                icerik.append("\n")
                satirlar.append('<content family="%s" size="12" startOffset="%d" length="1" />' % (yazi, konum))
                konum += 1
            else:
                for i, o in enumerate(temiz_ogeler):
                    son = (i == len(temiz_ogeler) - 1)
                    if o[0] == "metin":
                        t = o[1] + ("\n" if son else "")
                        b = o[2]
                        run_yazi = b.get("family") or self.YAZI_TIPI
                        boyut = int(b.get("size") or 12)
                        ek = ""
                        if b.get("bold"):
                            ek += 'bold="true" '
                        if b.get("italic"):
                            ek += 'italic="true" '
                        if b.get("underline"):
                            ek += 'underline="true" '
                        icerik.append(t)
                        satirlar.append('<content %sfamily="%s" size="%d" startOffset="%d" length="%d" />'
                                        % (ek, run_yazi, boyut, konum, len(t)))
                        konum += len(t)
                    else:
                        icerik.append("\u00b8")
                        satirlar.append('<image imageData="%s" width="%.1f" height="%.1f" '
                                        'startOffset="%d" length="1" />' % (o[1], o[2], o[3], konum))
                        konum += 1
                        if son:
                            icerik.append("\n")
                            satirlar.append('<content family="%s" size="12" startOffset="%d" length="1" />'
                                            % (yazi, konum))
                            konum += 1
            satirlar.append("</paragraph>")
            durum["konum"] = konum
            return "".join(satirlar)

        for oge in self.paragraflar:
            if oge[0] == "TABLO":
                _, satirlar, genislikler = oge
                parcalar = ['<table tableName="Sabit" columnCount="%d" columnSpans="%s" border="borderCell">'
                            % (len(genislikler), ",".join("%d" % max(1, round(g)) for g in genislikler))]
                for n, satir in enumerate(satirlar, 1):
                    parcalar.append('<row rowName="row%d" rowType="dataRow">' % n)
                    for hucre in satir:
                        ek = ""
                        if hucre.get("rowspan", 1) > 1:
                            ek += ' rowSpan="%d"' % hucre["rowspan"]
                        if hucre.get("colspan", 1) > 1:
                            ek += ' colSpan="%d"' % hucre["colspan"]
                        parcalar.append("<cell%s>" % ek)
                        pars = hucre["paragraflar"] or [(0, [])]
                        for h, og in pars:
                            parcalar.append(paragraf_xml(h, og))
                        parcalar.append("</cell>")
                    parcalar.append("</row>")
                parcalar.append("</table>")
                elemanlar.append("".join(parcalar))
            else:
                elemanlar.append(paragraf_xml(oge[0], oge[1]))

        return (
            '<?xml version="1.0" encoding="UTF-8" ?>\n'
            '<template format_id="1.8" >\n'
            '<content><![CDATA[' + "".join(icerik) + ']]></content>\n'
            '<properties><pageFormat mediaSizeName="1" leftMargin="42.525000000000006" '
            'rightMargin="42.525000000000006" topMargin="42.525000000000006" '
            'bottomMargin="42.525000000000006" paperOrientation="1" headerFOffset="20.0" '
            'footerFOffset="20.0" /></properties>\n'
            '<elements resolver="hvl-default">\n' + "\n".join(elemanlar) + '\n</elements>\n'
            '<styles><style name="default" description="Geçerli" family="Dialog" size="12" '
            'bold="false" italic="false" foreground="-13421773" />'
            '<style name="hvl-default" family="Times New Roman" size="12" description="Gövde" /></styles>\n'
            '</template>'
        )

    def kaydet(self, yol):
        xml = self.xml_uret()
        ET.fromstring(xml.encode("utf-8"))
        with zipfile.ZipFile(yol, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("content.xml", xml.encode("utf-8"))


def _resmi_hazirla_udf(blob, azami_piksel=1600, jpeg_zorla=False):
    try:
        img = Image.open(io.BytesIO(blob))
        img.load()
    except Exception:
        return None
    fmt = (img.format or "").upper()
    try:
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            img = img.convert("RGBA")
            zemin = Image.new("RGB", img.size, "white")
            zemin.paste(img, mask=img.split()[3])
            img = zemin
        elif img.mode != "RGB":
            img = img.convert("RGB")
        w, h = img.size
        if w > azami_piksel:
            img = img.resize((azami_piksel, max(1, round(h * azami_piksel / w))), Image.LANCZOS)
            w, h = img.size
        cikti = io.BytesIO()
        if jpeg_zorla or fmt in ("JPEG", "MPO"):
            img.save(cikti, "JPEG", quality=88, optimize=True)
            uzanti = "jpg"
        else:
            img.save(cikti, "PNG", optimize=True)
            uzanti = "png"
        return cikti.getvalue(), uzanti, w, h
    except Exception:
        return None


def _resim_ogesi_udf(blob, genislik_pt, max_pt, resimler, ist, jpeg_zorla=False):
    hazir = _resmi_hazirla_udf(blob, jpeg_zorla=jpeg_zorla)
    if hazir is None:
        ist["atlanan"] += 1
        return None
    veri, uzanti, pw, ph = hazir
    gen = float(genislik_pt) if genislik_pt else float(pw)
    gen = max(16.0, min(gen, float(max_pt)))
    yuk = gen * ph / pw
    resimler.append((veri, uzanti))
    ist["resim"] += 1
    return ("resim", base64.b64encode(veri).decode("ascii"), gen, yuk)


def _docx_hiza(p):
    a = p.alignment
    if a is None:
        try:
            a = p.style.paragraph_format.alignment
        except Exception:
            a = None
    if a is None:
        return 0
    ad = str(getattr(a, "name", a)).upper()
    if "CENTER" in ad:
        return 1
    if "RIGHT" in ad:
        return 2
    if "JUSTIFY" in ad or "DISTRIBUTE" in ad:
        return 3
    return 0


def _docx_stil_degeri(p, ozellik):
    try:
        s = p.style
    except Exception:
        return None
    while s is not None:
        try:
            v = getattr(s.font, ozellik)
        except Exception:
            v = None
        if v is not None:
            return v
        s = s.base_style
    return None


def _docx_medya(el, d):
    kok = el
    if el.tag == _NS_MC + "AlternateContent":
        secimler = [c for c in el if c.tag == _NS_MC + "Choice"] or \
                   [c for c in el if c.tag == _NS_MC + "Fallback"]
        if not secimler:
            return []
        kok = secimler[0]
    genislik_pt = None
    ext = next(kok.iter(_NS_WP + "extent"), None)
    if ext is not None:
        try:
            genislik_pt = int(ext.get("cx")) / 12700.0
        except (TypeError, ValueError):
            genislik_pt = None
    bulunan = []
    for blip in kok.iter(_NS_A + "blip"):
        rid = blip.get(_NS_R + "embed")
        if rid and rid in d.part.related_parts:
            bulunan.append((d.part.related_parts[rid].blob, genislik_pt))
    for im in kok.iter(_NS_V + "imagedata"):
        rid = im.get(_NS_R + "id")
        if rid and rid in d.part.related_parts:
            bulunan.append((d.part.related_parts[rid].blob, None))
    return bulunan


def _docx_paragraf_isle(p, d, max_px, resimler, ist):
    hiza = _docx_hiza(p)
    stil_kalin = _docx_stil_degeri(p, "bold")
    stil_italik = _docx_stil_degeri(p, "italic")
    stil_punto = _docx_stil_degeri(p, "size")
    sonuc = []
    mevcut = []

    def kapat():
        sonuc.append((hiza, list(mevcut)))
        del mevcut[:]

    runlar = p._p.xpath("./w:r | ./w:hyperlink/w:r | ./w:ins/w:r | ./w:smartTag/w:r")
    for r in runlar:
        run = Run(r, p)
        f = run.font
        kalin = f.bold if f.bold is not None else stil_kalin
        italik = f.italic if f.italic is not None else stil_italik
        punto = f.size if f.size is not None else stil_punto
        try:
            boyut = int(round(punto.pt)) if punto is not None else 12
        except Exception:
            boyut = 12
        bicim = {"bold": bool(kalin), "italic": bool(italik),
                 "underline": bool(f.underline), "size": max(6, min(boyut, 72))}
        metin = []

        def metni_bosalt():
            if metin:
                mevcut.append(("metin", "".join(metin), bicim))
                del metin[:]

        for el in r.iterchildren():
            etiket = el.tag
            if etiket == qn("w:t"):
                metin.append(el.text or "")
            elif etiket == qn("w:tab"):
                metin.append("\t")
            elif etiket == qn("w:noBreakHyphen"):
                metin.append("-")
            elif etiket == qn("w:br"):
                if el.get(qn("w:type")) == "page":
                    continue
                metni_bosalt()
                kapat()
            elif etiket in (qn("w:drawing"), qn("w:pict"), qn("w:object"), _NS_MC + "AlternateContent"):
                metni_bosalt()
                for blob, gen_pt in _docx_medya(el, d):
                    oge = _resim_ogesi_udf(blob, gen_pt, max_px, resimler, ist)
                    if oge:
                        mevcut.append(oge)
        metni_bosalt()
    kapat()
    return sonuc


def _docx_hucre_span(tc):
    gs = tc.xpath("./w:tcPr/w:gridSpan")
    try:
        return max(1, int(gs[0].get(qn("w:val")))) if gs else 1
    except (TypeError, ValueError):
        return 1


def _docx_tablo_duz_isle(t, d, belge, max_px, resimler, ist):
    for satir in t.rows:
        gorulen = set()
        ogeler = []
        resimli = []
        for hucre in satir.cells:
            if hucre._tc in gorulen:
                continue
            gorulen.add(hucre._tc)
            if ogeler:
                ogeler.append(("metin", "\t", {}))
            ilk_paragraf = True
            for hp in hucre.paragraphs:
                for _, og in _docx_paragraf_isle(hp, d, max_px, resimler, ist):
                    metinler = [o for o in og if o[0] != "resim"]
                    resimli.extend(o for o in og if o[0] == "resim")
                    if metinler:
                        if not ilk_paragraf:
                            ogeler.append(("metin", " ", {}))
                        ogeler.extend(metinler)
                        ilk_paragraf = False
        belge.paragraf_ekle(ogeler, 0)
        for o in resimli:
            belge.paragraf_ekle([o], 1)


def _docx_tablo_isle(t, d, belge, max_px, resimler, ist):
    tbl = t._tbl
    trler = tbl.xpath("./w:tr")
    if not trler:
        return
    izgara = []
    for gc in tbl.xpath("./w:tblGrid/w:gridCol"):
        try:
            izgara.append(int(gc.get(qn("w:w"))) / 20.0)
        except (TypeError, ValueError):
            izgara.append(0.0)
    sutun_say = max([len(izgara)] + [
        sum(_docx_hucre_span(tc) for tc in tr.xpath("./w:tc")) for tr in trler])
    if len(izgara) != sutun_say or sum(izgara) <= 0:
        izgara = [510.0 / sutun_say] * sutun_say
    toplam = sum(izgara)
    if toplam > 510.0:
        izgara = [g * 510.0 / toplam for g in izgara]

    satirlar = []
    for tr in trler:
        satir = []
        sutun = 0
        for tc in tr.xpath("./w:tc"):
            span = _docx_hucre_span(tc)
            gen = sum(izgara[sutun:sutun + span]) or izgara[0]
            sutun += span
            pars = []
            for cocuk in tc.iterchildren():
                if cocuk.tag == qn("w:p"):
                    pars.extend(_docx_paragraf_isle(Paragraph(cocuk, d), d, max(30, int(gen) - 12), resimler, ist))
                elif cocuk.tag == qn("w:tbl"):
                    alt = UdfBelge()
                    _docx_tablo_duz_isle(Table(cocuk, d), d, alt, max(30, int(gen) - 12), resimler, ist)
                    pars.extend(alt.paragraflar)
            satir.append({"paragraflar": pars, "colspan": span})
        satirlar.append(satir)
    belge.tablo_ekle(satirlar, izgara)


def docx_to_udf_cevir(kaynak_docx_yol, max_px=510):
    d = docx.Document(kaynak_docx_yol)
    belge = UdfBelge()
    resimler = []
    ist = {"resim": 0, "atlanan": 0}
    for cocuk in d.element.body.iterchildren():
        if cocuk.tag == qn("w:p"):
            for hiza, ogeler in _docx_paragraf_isle(Paragraph(cocuk, d), d, max_px, resimler, ist):
                belge.paragraf_ekle(ogeler, hiza)
        elif cocuk.tag == qn("w:tbl"):
            _docx_tablo_isle(Table(cocuk, d), d, belge, max_px, resimler, ist)
    ist["paragraf"] = len(belge.paragraflar)
    return belge, ist, resimler


# ============================================================================
# PDF -> UDF & PDF -> WORD MOTORU (TOPLULUK KATKISI ANALİZ MOTORU)
# ============================================================================
def _alan(k):
    return max(0.0, k[2] - k[0]) * max(0.0, k[3] - k[1])


def _bolge_birlestir(bolgeler):
    bolgeler = [list(b) for b in bolgeler]
    degisti = True
    while degisti:
        degisti = False
        for i in range(len(bolgeler)):
            for j in range(i + 1, len(bolgeler)):
                a, b = bolgeler[i], bolgeler[j]
                if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                    bolgeler[i] = [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])]
                    del bolgeler[j]
                    degisti = True
                    break
            if degisti:
                break
    return [tuple(b) for b in bolgeler]


def _sekil_bolgeleri(sayfa, gen, yuk, metin_kutulari):
    DPI = 36
    pix = sayfa.get_pixmap(dpi=DPI)
    im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
    sx = im.width / gen
    sy = im.height / yuk
    cizim = ImageDraw.Draw(im)
    for x0, y0, x1, y1 in metin_kutulari:
        cizim.rectangle([x0 * sx - 2, y0 * sy - 2, x1 * sx + 2, y1 * sy + 2], fill=255)
    maske = im.point(lambda v: 255 if v < 225 else 0).filter(ImageFilter.MaxFilter(7))
    w, h = maske.size
    px = maske.load()
    gorulen = [[False] * w for _ in range(h)]
    bulunan = []
    for y in range(h):
        for x in range(w):
            if px[x, y] and not gorulen[y][x]:
                yigin = [(x, y)]
                gorulen[y][x] = True
                bx0 = bx1 = x
                by0 = by1 = y
                while yigin:
                    cx, cy = yigin.pop()
                    bx0 = min(bx0, cx); bx1 = max(bx1, cx)
                    by0 = min(by0, cy); by1 = max(by1, cy)
                    for nx, ny in ((cx + 1, cy), (cx - 1, cy), (cx, cy + 1), (cx, cy - 1)):
                        if 0 <= nx < w and 0 <= ny < h and px[nx, ny] and not gorulen[ny][nx]:
                            gorulen[ny][nx] = True
                            yigin.append((nx, ny))
                k = (bx0 / sx, by0 / sy, (bx1 + 1) / sx, (by1 + 1) / sy)
                bw, bh = k[2] - k[0], k[3] - k[1]
                if bw >= 30 and bh >= 30 and bw / bh < 8 and bh / bw < 8:
                    bulunan.append(k)
    sonuc = []
    for k in _bolge_birlestir(bulunan):
        sonuc.append((max(0.0, k[0] - 2), max(0.0, k[1] - 2), min(gen, k[2] + 2), min(yuk, k[3] + 2)))
    return sonuc


def _yazi_ailesi(font_adi):
    ad = (font_adi or "").lower()
    if "serif" in ad and "sans" not in ad:
        return "Times New Roman"
    if any(k in ad for k in ("sans", "arial", "helvet", "calibri", "verdana", "tahoma", "segoe", "gothic")):
        return "Arial"
    return "Times New Roman"


def _bloklari_birlestir(bloklar):
    sirali = sorted(bloklar, key=lambda b: (b["bbox"][1], b["bbox"][0]))
    sonuc = []
    for b in sirali:
        birlesti = False
        for a in reversed(sonuc[-8:]):
            ax0, ay0, ax1, ay1 = a["bbox"]
            bx0, by0, bx1, by1 = b["bbox"]
            satir_yuk = max(ay1 - ay0, 1.0) / max(len(a["lines"]), 1)
            bosluk = by0 - ay1
            ortak = min(ax1, bx1) - max(ax0, bx0)
            if -1.5 <= bosluk <= 0.6 * satir_yuk and ortak > 0.5 * min(ax1 - ax0, bx1 - bx0) and abs(ax0 - bx0) <= 20:
                a["lines"] = a["lines"] + b["lines"]
                a["bbox"] = (min(ax0, bx0), ay0, max(ax1, bx1), max(ay1, by1))
                birlesti = True
                break
        if not birlesti:
            sonuc.append({"type": 0, "bbox": tuple(b["bbox"]), "lines": list(b["lines"])})
    return sonuc


def _okuma_sirasi(ogeler, gen):
    orta = gen / 2.0

    def yayilir(o):
        return o["bbox"][0] < orta - 6 and o["bbox"][2] > orta + 6

    def anahtar(o):
        return (o["bbox"][1], o["bbox"][0])

    def kosu_coz(kosu):
        sol = [o for o in kosu if (o["bbox"][0] + o["bbox"][2]) / 2 < orta]
        sag = [o for o in kosu if (o["bbox"][0] + o["bbox"][2]) / 2 >= orta]
        gercek_sutun = (sol and sag
                        and max(o["bbox"][2] - o["bbox"][0] for o in sol) >= 0.25 * gen
                        and max(o["bbox"][2] - o["bbox"][0] for o in sag) >= 0.25 * gen)
        if not gercek_sutun:
            return sorted(kosu, key=anahtar)
        sol_ust = min(o["bbox"][1] for o in sol)
        sag_ust = min(o["bbox"][1] for o in sag)
        ust = [o for o in sag if o["bbox"][3] <= sol_ust + 1] + [o for o in sol if o["bbox"][3] <= sag_ust + 1]
        sol_kalan = [o for o in sol if o not in ust]
        sag_kalan = [o for o in sag if o not in ust]
        return sorted(ust, key=anahtar) + sorted(sol_kalan, key=anahtar) + sorted(sag_kalan, key=anahtar)

    sonuc, kosu = [], []
    for o in sorted(ogeler, key=anahtar):
        if yayilir(o):
            sonuc.extend(kosu_coz(kosu))
            kosu = []
            sonuc.append(o)
        else:
            kosu.append(o)
    sonuc.extend(kosu_coz(kosu))
    return sonuc


def _liste_gecisi(mevcut, sonraki_satir):
    if not mevcut:
        return False
    son = mevcut[-1][1].rstrip()
    ilk = ""
    for sp in sonraki_satir.get("spans", []):
        if sp.get("text", "").strip():
            ilk = sp["text"].lstrip()[:1]
            break
    if not son or not ilk or son[-1] in ".,;:-–—":
        return False
    return ilk.isupper() or ilk.isdigit() or ilk in "(“\"•"


def _pdf_tablolari(sayfa, govde, azami_gen=510.0):
    try:
        bulunan = sayfa.find_tables(vertical_strategy="lines_strict", horizontal_strategy="lines_strict").tables
        satir_listesi = []
        for b in sayfa.get_text("dict").get("blocks", []):
            if b.get("type") == 0:
                for ln in b.get("lines", []):
                    if any(sp.get("text", "").strip() for sp in ln.get("spans", [])):
                        satir_listesi.append(ln["bbox"])
    except Exception:
        return []

    def kume(degerler, tol=2.0):
        kenarlar = []
        for v in sorted(degerler):
            if kenarlar and v - kenarlar[-1] <= tol:
                continue
            kenarlar.append(v)
        return kenarlar

    def indeks(kenarlar, v, tol=2.0):
        return min(range(len(kenarlar)), key=lambda i: abs(kenarlar[i] - v))

    sonuc = []
    for t in bulunan:
        hucreler = [tuple(c) for c in (t.cells or []) if c]
        if len(hucreler) < 2:
            continue
        xk = kume([c[0] for c in hucreler] + [c[2] for c in hucreler])
        yk = kume([c[1] for c in hucreler] + [c[3] for c in hucreler])
        if len(xk) < 3:
            continue
        tx0, ty0, tx1, ty1 = min(xk), min(yk), max(xk), max(yk)
        icerde = [l for l in satir_listesi if tx0 - 1 <= (l[0] + l[2]) / 2 <= tx1 + 1 and ty0 - 1 <= (l[1] + l[3]) / 2 <= ty1 + 1]
        if not icerde:
            continue
        uymayan = 0
        for l in icerde:
            if not any(c[0] - 3 <= l[0] and l[2] <= c[2] + 3 and c[1] - 3 <= l[1] and l[3] <= c[3] + 3 for c in hucreler):
                uymayan += 1
        if uymayan > 0.15 * len(icerde):
            continue
        nsat, nsut = len(yk) - 1, len(xk) - 1
        yerlesim = {}
        kapli = [[False] * nsut for _ in range(nsat)]
        for c in hucreler:
            c0, c1 = indeks(xk, c[0]), indeks(xk, c[2])
            r0, r1 = indeks(yk, c[1]), indeks(yk, c[3])
            if c1 <= c0 or r1 <= r0 or kapli[r0][c0]:
                continue
            yerlesim[(r0, c0)] = (c, c1 - c0, r1 - r0)
            for r in range(r0, r1):
                for k in range(c0, c1):
                    kapli[r][k] = True

        def hucre_icerigi(rect, gen_h):
            r = pymupdf.Rect(*rect)
            r = pymupdf.Rect(r.x0 + 1, r.y0 + 1, r.x1 - 1, r.y1 - 1)
            try:
                sozluk = sayfa.get_text("dict", clip=r)
            except Exception:
                return []
            satirlar = []
            for b in sozluk.get("blocks", []):
                if b.get("type") != 0:
                    continue
                for ln in b.get("lines", []):
                    if any(sp.get("text", "").strip() for sp in ln.get("spans", [])):
                        satirlar.append(ln)
            satirlar.sort(key=lambda l: (round(l["bbox"][1] / 3.0), l["bbox"][0]))
            gruplar = []
            for ln in satirlar:
                if gruplar and abs(gruplar[-1][-1]["bbox"][1] - ln["bbox"][1]) < 2.5:
                    gruplar[-1].append(ln)
                else:
                    gruplar.append([ln])
            paragraflar, mevcut, onceki = [], [], None

            def bitir():
                if not mevcut:
                    return
                ogeler = list(mevcut)
                ogeler[0] = ("metin", ogeler[0][1].lstrip(), ogeler[0][2])
                ogeler[-1] = ("metin", ogeler[-1][1].rstrip(), ogeler[-1][2])
                x0 = min(l["bbox"][0] for g in mevcut_satirlar for l in g)
                x1 = max(l["bbox"][2] for g in mevcut_satirlar for l in g)
                orta = ((x0 + x1) / 2 - rect[0]) / max(gen_h, 1.0)
                hiza = 1 if abs(orta - 0.5) < 0.12 and (x0 - rect[0]) > 3 else 0
                if any(o[1].strip() for o in ogeler):
                    paragraflar.append((hiza, ogeler))
                del mevcut[:]
                del mevcut_satirlar[:]

            mevcut_satirlar = []
            for g in gruplar:
                y0 = min(l["bbox"][1] for l in g)
                y1 = max(l["bbox"][3] for l in g)
                if onceki is not None and mevcut and (y0 - onceki[1]) > 1.2 * max(onceki[1] - onceki[0], 6.0):
                    bitir()
                if mevcut:
                    mevcut.append(("metin", " ", mevcut[-1][2]))
                mevcut_satirlar.append(g)
                for ln in sorted(g, key=lambda l: l["bbox"][0]):
                    for sp in ln["spans"]:
                        metin = sp.get("text", "").replace("\u00ad", "")
                        if not metin:
                            continue
                        boyut = float(sp.get("size", 12))
                        ad = sp.get("font", "").lower()
                        bayrak = sp.get("flags", 0)
                        bicim = {
                            "bold": bool(bayrak & 16) or "bold" in ad,
                            "italic": bool(bayrak & 2) or "italic" in ad or "oblique" in ad,
                            "underline": False,
                            "size": max(6, min(int(round(boyut)), 72)),
                            "family": _yazi_ailesi(sp.get("font", "")),
                        }
                        if mevcut and mevcut[-1][2] == bicim:
                            mevcut[-1] = ("metin", mevcut[-1][1] + metin, bicim)
                        else:
                            mevcut.append(("metin", metin, bicim))
                onceki = (y0, y1)
            bitir()
            return paragraflar

        genislikler = [xk[i + 1] - xk[i] for i in range(nsut)]
        toplam = sum(genislikler)
        olcek = azami_gen / toplam if toplam > azami_gen else 1.0
        genislikler = [g * olcek for g in genislikler]

        satirlar = []
        for r in range(nsat):
            satir = []
            for k in range(nsut):
                if (r, k) in yerlesim:
                    rect, cs, rs = yerlesim[(r, k)]
                    satir.append({"paragraflar": hucre_icerigi(rect, rect[2] - rect[0]),
                                  "colspan": cs, "rowspan": rs})
                elif not kapli[r][k]:
                    satir.append({"paragraflar": [], "colspan": 1, "rowspan": 1})
            satirlar.append(satir)
        sonuc.append({"bbox": (xk[0], yk[0], xk[-1], yk[-1]), "satirlar": satirlar, "genislikler": genislikler})
    return sonuc


def _pdf_blok_paragraflari(blok, gen, govde_punto):
    satirlar = [s for s in blok.get("lines", []) if any(sp.get("text", "").strip() for sp in s.get("spans", []))]
    if not satirlar:
        return []
    sol = min(s["bbox"][0] for s in satirlar)
    sag = max(s["bbox"][2] for s in satirlar)
    blok_gen = max(sag - sol, 1.0)

    yasli = False
    if len(satirlar) >= 3:
        dayanan = sum(1 for s in satirlar[:-1] if s["bbox"][2] >= sag - 2.5)
        yasli = dayanan >= 0.7 * (len(satirlar) - 1) and blok_gen > 0.25 * gen

    paragraflar = []
    mevcut = []
    p_satirlari = []

    def bitir():
        if not mevcut:
            return
        x0 = min(s["bbox"][0] for s in p_satirlari)
        x1 = max(s["bbox"][2] for s in p_satirlari)
        y0 = p_satirlari[0]["bbox"][1]
        y1 = p_satirlari[-1]["bbox"][3]
        hiza = 0
        if yasli:
            hiza = 3
        elif len(p_satirlari) <= 3:
            merkez = (x0 + x1) / 2
            if abs(merkez - gen / 2) < 10 and (x1 - x0) < 0.85 * gen and x0 > 30:
                hiza = 1
            elif len(p_satirlari) == 1 and x1 >= gen - 100 and x0 > 0.55 * gen:
                hiza = 2
        ogeler = list(mevcut)
        if ogeler:
            ogeler[0] = ("metin", ogeler[0][1].lstrip(), ogeler[0][2])
            ogeler[-1] = ("metin", ogeler[-1][1].rstrip(), ogeler[-1][2])
        if any(o[1].strip() for o in ogeler):
            paragraflar.append((hiza, ogeler, y0, y1))
        del mevcut[:]
        del p_satirlari[:]

    for i, satir in enumerate(satirlar):
        if mevcut:
            son = mevcut[-1]
            ilk = ""
            for sp in satir["spans"]:
                if sp.get("text", "").strip():
                    ilk = sp["text"].lstrip()[:1]
                    break
            if son[1].endswith("-") and len(son[1]) > 1 and son[1][-2].isalnum() and ilk.islower():
                mevcut[-1] = ("metin", son[1][:-1], son[2])
            else:
                mevcut.append(("metin", " ", son[2]))
        p_satirlari.append(satir)
        for sp in satir["spans"]:
            metin = sp.get("text", "").replace("\u00ad", "")
            if not metin:
                continue
            boyut = float(sp.get("size", 12))
            ad = sp.get("font", "").lower()
            bayrak = sp.get("flags", 0)
            baslik = boyut >= max(13.0, 1.35 * govde_punto)
            bicim = {
                "bold": bool(bayrak & 16) or "bold" in ad or baslik,
                "italic": bool(bayrak & 2) or "italic" in ad or "oblique" in ad,
                "underline": False,
                "size": max(6, min(int(round(boyut)), 72)),
                "family": _yazi_ailesi(sp.get("font", "")),
            }
            if mevcut and mevcut[-1][2] == bicim:
                mevcut[-1] = ("metin", mevcut[-1][1] + metin, bicim)
            else:
                mevcut.append(("metin", metin, bicim))
        if i + 1 < len(satirlar):
            sonraki = satirlar[i + 1]
            tam_satir = satir["bbox"][2] >= sag - 0.20 * blok_gen
            girinti = sonraki["bbox"][0] > satir["bbox"][0] + 6
            if not tam_satir or girinti or (not yasli and _liste_gecisi(mevcut, sonraki)):
                bitir()
    bitir()
    return paragraflar


def _govde_punto(metin_bloklari):
    agirlik = {}
    for b in metin_bloklari:
        for s in b.get("lines", []):
            for sp in s.get("spans", []):
                n = len(sp.get("text", "").strip())
                if n:
                    k = round(float(sp.get("size", 12)))
                    agirlik[k] = agirlik.get(k, 0) + n
    if sum(agirlik.values()) < 30:
        return 10.0
    return float(max(agirlik, key=agirlik.get))


def pdf_den_udf(yol, belge, max_px, sayfa_resmi, resimler, sayfa_sinirlari=None, tablo_destegi=True):
    ist = {"resim": 0, "atlanan": 0}
    doc = pymupdf.open(yol)
    try:
        for sayfa in doc:
            gen = float(sayfa.rect.width)
            yuk = float(sayfa.rect.height)
            if sayfa_sinirlari is not None:
                sayfa_sinirlari.append(len(belge.paragraflar))
            ist.setdefault("sayfa_gen", gen)
            ist.setdefault("sayfa_yuk", yuk)
            if sayfa_resmi:
                pix = sayfa.get_pixmap(dpi=150)
                oge = _resim_ogesi_udf(pix.tobytes("png"), max_px, max_px, resimler, ist, jpeg_zorla=True)
                if oge:
                    belge.paragraf_ekle([oge], 1)
                continue

            bloklar = sayfa.get_text("dict").get("blocks", [])
            metin_bloklari = [b for b in bloklar if b.get("type") == 0 and any(sp.get("text", "").strip() for s in b.get("lines", []) for sp in s.get("spans", []))]
            resim_bloklari = [b for b in bloklar if b.get("type") == 1]
            govde = _govde_punto(metin_bloklari)

            ogeler = []
            arka_plan = any(_alan(b["bbox"]) >= 0.8 * gen * yuk for b in resim_bloklari)

            if arka_plan:
                kutular = [b["bbox"] for b in metin_bloklari]
                bolgeler = _sekil_bolgeleri(sayfa, gen, yuk, kutular)
                if not metin_bloklari or (bolgeler and max(_alan(k) for k in bolgeler) >= 0.7 * gen * yuk):
                    pix = sayfa.get_pixmap(dpi=150)
                    oge = _resim_ogesi_udf(pix.tobytes("png"), max_px, max_px, resimler, ist, jpeg_zorla=True)
                    if oge:
                        belge.paragraf_ekle([oge], 1)
                    continue
                for k in bolgeler:
                    try:
                        pix = sayfa.get_pixmap(dpi=150, clip=pymupdf.Rect(*k))
                        oge = _resim_ogesi_udf(pix.tobytes("png"), k[2] - k[0], max_px, resimler, ist, jpeg_zorla=True)
                    except Exception:
                        oge = None
                        ist["atlanan"] += 1
                    if oge:
                        ogeler.append({"tur": "resim", "bbox": k, "oge": oge})

                def sekil_icinde(b):
                    bx0, by0, bx1, by1 = b["bbox"]
                    alan_b = max(_alan(b["bbox"]), 1e-6)
                    for k in bolgeler:
                        kx = max(0.0, min(bx1, k[2]) - max(bx0, k[0]))
                        ky = max(0.0, min(by1, k[3]) - max(by0, k[1]))
                        if kx * ky >= 0.7 * alan_b:
                            return True
                    return False
                metin_bloklari = [b for b in metin_bloklari if not sekil_icinde(b)]
            else:
                for b in resim_bloklari:
                    if b.get("width", 0) < 10 or b.get("height", 0) < 10:
                        continue
                    x0, y0, x1, y1 = b["bbox"]
                    oge = _resim_ogesi_udf(b["image"], x1 - x0, max_px, resimler, ist)
                    if oge:
                        ogeler.append({"tur": "resim", "bbox": b["bbox"], "oge": oge})

            if not arka_plan and tablo_destegi:
                tablolar = _pdf_tablolari(sayfa, govde, max_px if max_px >= 510 else 510.0)
                if tablolar:
                    def satir_tabloda(ln):
                        x0, y0, x1, y1 = ln["bbox"]
                        mx, my = (x0 + x1) / 2, (y0 + y1) / 2
                        for tb in tablolar:
                            k = tb["bbox"]
                            if k[0] - 1 <= mx <= k[2] + 1 and k[1] - 1 <= my <= k[3] + 1:
                                return True
                        return False
                    yeni = []
                    for blk in metin_bloklari:
                        kalan = [ln for ln in blk.get("lines", []) if not satir_tabloda(ln)]
                        if not kalan:
                            continue
                        if len(kalan) != len(blk.get("lines", [])):
                            blk = {"type": 0, "lines": kalan,
                                   "bbox": (min(l["bbox"][0] for l in kalan), min(l["bbox"][1] for l in kalan),
                                            max(l["bbox"][2] for l in kalan), max(l["bbox"][3] for l in kalan))}
                        yeni.append(blk)
                    metin_bloklari = yeni
                    for tb in tablolar:
                        ogeler.append({"tur": "tablo", "bbox": tb["bbox"], "tablo": tb})

            metin_bloklari = _bloklari_birlestir(metin_bloklari)
            for b in metin_bloklari:
                ogeler.append({"tur": "metin", "bbox": b["bbox"], "blok": b})

            onceki_y1 = None
            onceki_yuk = None
            icerik_var = False
            for o in _okuma_sirasi(ogeler, gen):
                x0, y0, x1, y1 = o["bbox"]
                if o["tur"] == "tablo":
                    if onceki_y1 is not None and (y0 - onceki_y1) > 8:
                        belge.paragraf_ekle([], 0)
                    belge.tablo_ekle(o["tablo"]["satirlar"], o["tablo"]["genislikler"])
                    onceki_y1, onceki_yuk = y1, None
                    icerik_var = True
                    continue
                if o["tur"] == "resim":
                    merkez = (x0 + x1) / 2
                    if (x1 - x0) >= 0.6 * gen or abs(merkez - gen / 2) < 0.1 * gen:
                        hiza = 1
                    elif merkez < 0.4 * gen:
                        hiza = 0
                    else:
                        hiza = 2
                    belge.paragraf_ekle([o["oge"]], hiza)
                    onceki_y1, onceki_yuk = y1, None
                    icerik_var = True
                    continue
                paragraflar = _pdf_blok_paragraflari(o["blok"], gen, govde)
                for n, (hiza, og, py0, py1) in enumerate(paragraflar):
                    if onceki_y1 is not None and py0 > onceki_y1:
                        ref = onceki_yuk or (py1 - py0)
                        if (py0 - onceki_y1) > 1.4 * ref:
                            belge.paragraf_ekle([], 0)
                    belge.paragraf_ekle(og, hiza)
                    onceki_y1 = py1
                    onceki_yuk = max(py1 - py0, 8.0)
                    icerik_var = True
            if icerik_var:
                belge.paragraf_ekle([], 0)
    finally:
        doc.close()
    ist["paragraf"] = len(belge.paragraflar)
    return ist


def belge_docx_yaz(belge, yol, sayfa_gen=595.0, sayfa_yuk=842.0, sayfa_sinirlari=None, kenar=42.5):
    d = docx.Document()
    bolum = d.sections[0]
    bolum.page_width = Pt(sayfa_gen)
    bolum.page_height = Pt(sayfa_yuk)
    for ad in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(bolum, ad, Pt(kenar))
    normal = d.styles["Normal"]
    normal.font.name = UdfBelge.YAZI_TIPI
    normal.font.size = Pt(12)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(3)
    normal.paragraph_format.line_spacing = 1.0
    yazdirilabilir = sayfa_gen - 2 * kenar
    hizalar = {0: WD_ALIGN_PARAGRAPH.LEFT, 1: WD_ALIGN_PARAGRAPH.CENTER,
               2: WD_ALIGN_PARAGRAPH.RIGHT, 3: WD_ALIGN_PARAGRAPH.JUSTIFY}
    sinirlar = set(i for i in (sayfa_sinirlari or []) if i > 0)
    for i, (hiza, ogeler) in enumerate(o for o in belge.paragraflar if o[0] != "TABLO"):
        if (i + 1) in sinirlar and not ogeler:
            continue
        p = d.add_paragraph()
        if i in sinirlar:
            p.paragraph_format.page_break_before = True
        p.alignment = hizalar.get(hiza, WD_ALIGN_PARAGRAPH.LEFT)
        for o in ogeler:
            if o[0] == "metin":
                metin = _XML_GECERSIZ.sub("", o[1])
                if not metin:
                    continue
                r = p.add_run(metin)
                b = o[2] or {}
                r.bold = bool(b.get("bold"))
                r.italic = bool(b.get("italic"))
                r.underline = bool(b.get("underline"))
                r.font.name = b.get("family") or UdfBelge.YAZI_TIPI
                r.font.size = Pt(b.get("size") or 12)
            else:
                gen = min(float(o[2]), yazdirilabilir)
                p.add_run().add_picture(io.BytesIO(base64.b64decode(o[1])), width=Pt(gen))
    d.save(yol)


def pdf_den_docx(yol, hedef, sayfa_resmi=False, sayfa_sonlari=True):
    ac = pymupdf.open(yol)
    try:
        r = ac[0].rect
        sayfa_gen, sayfa_yuk = float(r.width), float(r.height)
    finally:
        ac.close()
    kenar = 42.5
    belge = UdfBelge()
    resimler = []
    sinirlar = []
    ist = pdf_den_udf(yol, belge, sayfa_gen - 2 * kenar, sayfa_resmi, resimler, sinirlar, tablo_destegi=False)
    belge_docx_yaz(belge, hedef, sayfa_gen, sayfa_yuk, sinirlar if sayfa_sonlari else None, kenar)
    return ist


# ============================================================================
# GÖRSEL SAYFA ÖNİZLEME & DÜZENLEME PENCERESİ (THUMBNAIL GRID)
# ============================================================================
class PdfGorselOnizlemePenceresi(ctk.CTkToplevel):

    def __init__(self, parent, dosya_yolu):
        super().__init__(parent)
        self.parent = parent
        self.dosya_yolu = dosya_yolu
        self.doc = pymupdf.open(dosya_yolu)
        self.toplam_sayfa = len(self.doc)

        self.sayfa_durumlari = []
        for i in range(self.toplam_sayfa):
            self.sayfa_durumlari.append({
                "orijinal_idx": i,
                "rotation": self.doc[i].rotation
            })

        self.secili_sayfalar = set()
        self.gorsel_referanslari = {}

        # ── Sürükle-bırak durum değişkenleri ──────────────────
        self._surukleme_aktif   = False   # Gerçek sürükleme başladı mı?
        self._surukle_kaynak    = None    # Sürüklenen kart indexi
        self._surukle_baslangic = None    # Fare başlangıç (x, y)
        self._ghost_pencere     = None    # Kayan yarı saydam pencere
        self._hedef_idx         = None    # Anlık bırakma hedefi
        self._kart_widgetlari   = []      # Her sıra no için kart widget referansı
        self.SUTUN_SAYISI       = 4

        self.title(f"Sayfa Önizleme & Yönetim: {os.path.basename(dosya_yolu)}")
        self.geometry("980x680")
        self.minsize(860, 560)
        self.configure(fg_color=("#F8FAFC", "#0F172A"))
        self.after(100, self.lift)

        self._arayuz_kur()
        self._sayfalari_ciz()

    def _arayuz_kur(self):
        ust_bar = ctk.CTkFrame(
            self,
            fg_color=("#FFFFFF", "#1E293B"),
            height=54,
            corner_radius=0,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        ust_bar.pack(fill="x", side="top")
        ust_bar.pack_propagate(False)

        sol_bilgi = ctk.CTkFrame(ust_bar, fg_color="transparent")
        sol_bilgi.pack(side="left", padx=16, pady=10)

        self.lbl_sayfa_sayisi = ctk.CTkLabel(
            sol_bilgi,
            text=f"Toplam: {len(self.sayfa_durumlari)} Sayfa",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.lbl_sayfa_sayisi.pack(side="left")

        self.lbl_secim = ctk.CTkLabel(
            sol_bilgi,
            text="• 0 sayfa seçildi",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#64748B", "#94A3B8")
        )
        self.lbl_secim.pack(side="left", padx=8)

        sag_aksiyon = ctk.CTkFrame(ust_bar, fg_color="transparent")
        sag_aksiyon.pack(side="right", padx=16)

        ctk.CTkButton(
            sag_aksiyon,
            text="⟳ Seçilenleri 90° Döndür",
            command=self._secilenleri_dondur,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            height=32,
            corner_radius=6
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            sag_aksiyon,
            text="🗑 Seçilenleri Sil",
            command=self._secilenleri_sil,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
            fg_color=("#FEE2E2", "#7F1D1D"),
            hover_color=("#FECACA", "#991B1B"),
            text_color=("#991B1B", "#FCA5A5"),
            height=32,
            corner_radius=6
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            sag_aksiyon,
            text="💾 Değişiklikleri Yeni PDF Olarak Kaydet",
            command=self._degisiklikleri_kaydet,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            height=32,
            corner_radius=6
        ).pack(side="left", padx=(8, 0))

        alt_bar = ctk.CTkFrame(
            self,
            fg_color=("#FFFFFF", "#1E293B"),
            height=32,
            corner_radius=0,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        alt_bar.pack(fill="x", side="bottom")
        alt_bar.pack_propagate(False)

        ctk.CTkLabel(
            alt_bar,
            text="💡 Tıklayarak seç • Sürükleyerek sıra değiştir • ⟳ Döndür • 🗑 Sil",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#64748B", "#94A3B8")
        ).pack(side="left", padx=16)

        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=("#F8FAFC", "#0F172A"),
            corner_radius=0
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=12, pady=10)

        for c in range(4):
            self.scroll_frame.columnconfigure(c, weight=1, uniform="grid_col")

    def _sayfa_thumbnail_al(self, orijinal_idx, rotation):
        sayfa = self.doc[orijinal_idx]
        pix = sayfa.get_pixmap(dpi=42, annots=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        net_derece = (rotation - sayfa.rotation) % 360
        if net_derece != 0:
            img = img.rotate(-net_derece, expand=True)

        img.thumbnail((150, 195), Image.Resampling.LANCZOS)
        return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

    def _sayfalari_ciz(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        self._kart_widgetlari = []
        self.gorsel_referanslari.clear()
        self.lbl_sayfa_sayisi.configure(text=f"Toplam: {len(self.sayfa_durumlari)} Sayfa")
        self.lbl_secim.configure(text=f"• {len(self.secili_sayfalar)} sayfa seçildi")

        if not self.sayfa_durumlari:
            ctk.CTkLabel(
                self.scroll_frame,
                text="Tüm sayfalar silindi!\nLütfen pencereyi kapatın veya işlemi iptal edin.",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13),
                text_color=("#94A3B8", "#64748B")
            ).pack(pady=100)
            return

        for sira_no, item in enumerate(self.sayfa_durumlari):
            row = sira_no // self.SUTUN_SAYISI
            col = sira_no % self.SUTUN_SAYISI

            is_selected  = sira_no in self.secili_sayfalar
            is_hedef     = sira_no == self._hedef_idx and self._surukleme_aktif

            kart = ctk.CTkFrame(
                self.scroll_frame,
                fg_color=("#FFFFFF", "#1E293B"),
                corner_radius=8,
                border_width=2 if (is_selected or is_hedef) else 1,
                border_color=(
                    ("#F59E0B", "#FBBF24") if is_hedef else
                    (("#3B82F6", "#60A5FA") if is_selected else ("#E2E8F0", "#334155"))
                )
            )
            kart.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            self._kart_widgetlari.append(kart)

            ust_satir = ctk.CTkFrame(kart, fg_color="transparent", height=24)
            ust_satir.pack(fill="x", padx=8, pady=(6, 2))

            chk_text  = "✓" if is_selected else "○"
            chk_color = "#3B82F6" if is_selected else ("#94A3B8", "#64748B")
            lbl_chk = ctk.CTkLabel(
                ust_satir, text=chk_text,
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                text_color=chk_color
            )
            lbl_chk.pack(side="left")

            lbl_no = ctk.CTkLabel(
                ust_satir, text=f"Sayfa {sira_no + 1}",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            )
            lbl_no.pack(side="right")

            ctk_img = self._sayfa_thumbnail_al(item["orijinal_idx"], item["rotation"])
            self.gorsel_referanslari[sira_no] = ctk_img

            lbl_img = ctk.CTkLabel(kart, text="", image=ctk_img, cursor="fleur")
            lbl_img.pack(padx=8, pady=4)

            # ── Event binding: seçim + sürükle-bırak ──────────
            for widget in (kart, lbl_img, lbl_no, lbl_chk, ust_satir):
                widget.bind("<ButtonPress-1>",   lambda e, i=sira_no: self._surukle_basla(e, i))
                widget.bind("<B1-Motion>",        lambda e, i=sira_no: self._surukle_hareket(e, i))
                widget.bind("<ButtonRelease-1>",  lambda e, i=sira_no: self._surukle_birak(e, i))

            # Alt butonlar (döndür / sil) - bunlara sürükleme binding'i YOK
            alt_satir = ctk.CTkFrame(kart, fg_color="transparent")
            alt_satir.pack(fill="x", padx=6, pady=(4, 6))

            ctk.CTkButton(
                alt_satir, text="⟳ 90°",
                command=lambda idx=sira_no: self._tekil_dondur(idx),
                width=50, height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("#F1F5F9", "#334155"),
                hover_color=("#E2E8F0", "#475569"),
                text_color=("#1E293B", "#F8FAFC"),
                corner_radius=5
            ).pack(side="left", expand=True, fill="x", padx=(0, 2))

            ctk.CTkButton(
                alt_satir, text="🗑 Sil",
                command=lambda idx=sira_no: self._tekil_sil(idx),
                width=50, height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("#FEE2E2", "#7F1D1D"),
                hover_color=("#FECACA", "#991B1B"),
                text_color=("#991B1B", "#FCA5A5"),
                corner_radius=5
            ).pack(side="right", expand=True, fill="x", padx=(2, 0))

    # ================================================================
    # SÜRÜKLE-BIRAK MOTORUR
    # ================================================================
    def _surukle_basla(self, event, idx):
        """Fare basıldığında başlangıç konumunu kaydet."""
        self._surukle_kaynak    = idx
        self._surukleme_aktif   = False
        self._surukle_baslangic = (event.x_root, event.y_root)
        self._ghost_pencere     = None

    def _surukle_hareket(self, event, idx):
        """Fare hareket ettikçe ghost pencereyi güncelle."""
        if self._surukle_kaynak is None:
            return

        dx = abs(event.x_root - self._surukle_baslangic[0])
        dy = abs(event.y_root - self._surukle_baslangic[1])

        # 5px eşiği: küçük hareket = henüz sürükleme değil
        if not self._surukleme_aktif:
            if dx < 5 and dy < 5:
                return
            self._surukleme_aktif = True
            self._ghost_olustur(self._surukle_kaynak)

        # Hedef indexi hesapla ve vurgula
        yeni_hedef = self._koordinattan_idx(event.x_root, event.y_root)
        if yeni_hedef != self._hedef_idx:
            self._hedef_idx = yeni_hedef
            self._sayfalari_ciz()

        # Ghost pencereyi fareyle birlikte sürükle
        if self._ghost_pencere and self._ghost_pencere.winfo_exists():
            self._ghost_pencere.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")

    def _surukle_birak(self, event, idx):
        """Fare bırakıldığında sayfayı taşı veya seçim yap."""
        # Ghost pencereyi kapat
        if self._ghost_pencere and self._ghost_pencere.winfo_exists():
            self._ghost_pencere.destroy()
        self._ghost_pencere = None

        kaynak = self._surukle_kaynak
        hedef  = self._hedef_idx

        self._surukleme_aktif = False
        self._hedef_idx       = None
        self._surukle_kaynak  = None

        if not self._surukleme_aktif and kaynak == idx:
            # Gerçek sürükleme olmadı → seçim toggle
            pass  # aşağıda işlenecek

        if kaynak is not None and hedef is not None and kaynak != hedef:
            # Sayfayı taşı
            self.secili_sayfalar.discard(kaynak)
            eleman = self.sayfa_durumlari.pop(kaynak)
            # kaynak gidince hedef kayabilir
            gercek_hedef = hedef if hedef <= kaynak else hedef
            # Eğer kaynak < hedef ise, pop sonrası hedef 1 azaldı
            if kaynak < hedef:
                gercek_hedef = hedef - 1
            else:
                gercek_hedef = hedef
            self.sayfa_durumlari.insert(gercek_hedef, eleman)
            self.secili_sayfalar.clear()
            self._sayfalari_ciz()
        else:
            # Seçim toggle (sürükleme yoktu)
            if kaynak is not None:
                if kaynak in self.secili_sayfalar:
                    self.secili_sayfalar.remove(kaynak)
                else:
                    self.secili_sayfalar.add(kaynak)
            self._sayfalari_ciz()

    def _ghost_olustur(self, idx):
        """Sürüklenen sayfanın thumbnail'ini taşıyan yarı saydam pencere."""
        if idx >= len(self.sayfa_durumlari):
            return
        try:
            item = self.sayfa_durumlari[idx]
            sayfa = self.doc[item["orijinal_idx"]]
            pix   = sayfa.get_pixmap(dpi=36, annots=False)
            img   = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            net_derece = (item["rotation"] - sayfa.rotation) % 360
            if net_derece:
                img = img.rotate(-net_derece, expand=True)
            img.thumbnail((100, 130), Image.Resampling.LANCZOS)

            # Yarı saydam çerçeve
            cerceve = Image.new("RGBA", (img.width + 8, img.height + 8), (30, 100, 220, 160))
            cerceve.paste(img, (4, 4))

            self._ghost_img_ref = ctk.CTkImage(
                light_image=Image.fromarray(
                    __import__('numpy', fromlist=['']).array(cerceve)[:, :, :3]
                    if False else img
                ),
                dark_image=img,
                size=img.size
            )

            ghost = tk.Toplevel(self)
            ghost.overrideredirect(True)
            ghost.attributes("-alpha", 0.70)
            ghost.attributes("-topmost", True)
            ghost.configure(bg="#1E40AF")
            ghost.geometry(f"108x138+0+0")

            lbl = tk.Label(ghost, image=self._ghost_img_ref._light_image,
                           bg="#1E40AF", bd=0)
            lbl.pack(padx=4, pady=4)

            self._ghost_pencere = ghost
        except Exception:
            self._ghost_pencere = None

    def _koordinattan_idx(self, x_root, y_root):
        """Ekran koordinatlarından hangi kart indexini işaret ettiğini bul."""
        for i, kart in enumerate(self._kart_widgetlari):
            try:
                if not kart.winfo_exists():
                    continue
                kx = kart.winfo_rootx()
                ky = kart.winfo_rooty()
                kw = kart.winfo_width()
                kh = kart.winfo_height()
                if kx <= x_root <= kx + kw and ky <= y_root <= ky + kh:
                    return i
            except Exception:
                continue
        # Grid dışındaysa en yakın sona yerleştir
        return len(self.sayfa_durumlari) - 1





    def _tekil_dondur(self, idx):
        self.sayfa_durumlari[idx]["rotation"] = (self.sayfa_durumlari[idx]["rotation"] + 90) % 360
        self._sayfalari_ciz()

    def _tekil_sil(self, idx):
        del self.sayfa_durumlari[idx]
        self.secili_sayfalar = {i if i < idx else i - 1 for i in self.secili_sayfalar if i != idx}
        self._sayfalari_ciz()

    def _secilenleri_dondur(self):
        if not self.secili_sayfalar:
            messagebox.showinfo("Seçim Yok", "Lütfen döndürmek istediğiniz sayfaları tıklayarak seçin.")
            return
        for idx in self.secili_sayfalar:
            self.sayfa_durumlari[idx]["rotation"] = (self.sayfa_durumlari[idx]["rotation"] + 90) % 360
        self._sayfalari_ciz()

    def _secilenleri_sil(self):
        if not self.secili_sayfalar:
            messagebox.showinfo("Seçim Yok", "Lütfen silmek istediğiniz sayfaları tıklayarak seçin.")
            return
        if len(self.secili_sayfalar) == len(self.sayfa_durumlari):
            messagebox.showwarning("Uyarı", "Belgedeki tüm sayfaları silemezsiniz.")
            return

        onay = messagebox.askyesno("Sayfa Silme Onayı", f"Seçilen {len(self.secili_sayfalar)} adet sayfa silinsin mi?")
        if not onay:
            return

        yeni_durumlar = []
        for i, item in enumerate(self.sayfa_durumlari):
            if i not in self.secili_sayfalar:
                yeni_durumlar.append(item)

        self.sayfa_durumlari = yeni_durumlar
        self.secili_sayfalar.clear()
        self._sayfalari_ciz()

    def _degisiklikleri_kaydet(self):
        if not self.sayfa_durumlari:
            messagebox.showerror("Hata", "Kaydedilecek sayfa bulunmuyor!")
            return

        hedef_yol = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Dosyası", "*.pdf")],
            title="Düzenlenmiş PDF'i Kaydet",
            initialfile="duzenlenmis_" + os.path.basename(self.dosya_yolu)
        )
        if not hedef_yol:
            return

        try:
            yeni_doc = pymupdf.open()
            for item in self.sayfa_durumlari:
                yeni_doc.insert_pdf(self.doc, from_page=item["orijinal_idx"], to_page=item["orijinal_idx"])
                son_eklenen = yeni_doc[-1]
                son_eklenen.set_rotation(item["rotation"])

            yeni_doc.save(hedef_yol, garbage=4, deflate=True)
            yeni_doc.close()

            messagebox.showinfo("Başarılı", f"Düzenlenmiş PDF başarıyla kaydedildi!\n\nDosya: {os.path.basename(hedef_yol)}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme sırasında hata oluştu:\n{str(e)}")

    def destroy(self):
        try:
            self.doc.close()
        except Exception:
            pass
        super().destroy()


# ============================================================================
# ANA UYGULAMA PENCERESİ
# ============================================================================
class ModernUyapAsistani(CTkDnD):
    def __init__(self):
        super().__init__()
        self.ayarlar = ayarlari_yukle()

        self.title("UYAP Evrak Asistanı")
        self.geometry("1140x720")
        self.minsize(1060, 660)
        self.configure(fg_color=("#F8FAFC", "#0F172A"))

        self.dosyalar = []
        self.dosya_imza_durumlari = {}
        if IS_MAC:
            # Tk 9 macOS'ta ::tk::unsupported::MacWindowStyle'ı kaldırdı; tkinter.simpledialog (askstring)
            # bunu çağırıp hata veriyor. Eksikse etkisiz bir karşılık tanımlanır, pencere normal açılır.
            self.tk.eval('namespace eval ::tk::unsupported {}\n'
                         'if {[info commands ::tk::unsupported::MacWindowStyle] eq ""} '
                         '{proc ::tk::unsupported::MacWindowStyle args {}}')
        self._uygulama_ikonu_yukle()
        self._arayuz_olustur()

        # Önceki sürümün AI modeli (Qwen2.5-1.5B) artık kullanılmıyor; diskte yer kaplamasın
        temizlenen_mb = ai_eski_modelleri_temizle()
        if temizlenen_mb:
            mesaj = f"AI modeli {AI_MODEL_ADI} olarak güncellendi; eski model silindi ({temizlenen_mb} MB boşaltıldı)."
            if not ai_model_indirildi() and self.ayarlar.get("ai_ozet_aktif"):
                mesaj += " Yeni modeli Ayarlar > Yapay Zeka'dan indirebilirsiniz."
            self.durum_etiketi.configure(text=mesaj)

        # Uygulama açıldıktan 3 saniye sonra arka planda güncelleme kontrolü yap
        if self.ayarlar.get("otomatik_guncelleme_kontrol", True):
            self.after(3000, self._guncelleme_kontrolu_baslat)

    def _uygulama_ikonu_yukle(self):
        if IS_MAC:
            # macOS .ico desteklemez; pencere/Dock simgesi PNG'den ayarlanır
            png = resource_path("logo.png")
            if os.path.exists(png):
                try:
                    self._ikon_resmi = tk.PhotoImage(file=png)
                    self.iconphoto(True, self._ikon_resmi)
                except Exception:
                    pass
            return
        ico = resource_path("logo.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(default=ico)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # GÜNCELLEME DENETLEYİCİSİ
    # ------------------------------------------------------------------
    def _guncelleme_kontrolu_baslat(self):
        """Arka plan thread'inde GitHub'ı sorgular; UI'yı bloklamaz."""
        def _kontrol():
            sonuc = guncelleme_kontrol_et()
            if sonuc:
                # Tkinter iş parçacığı güvenliği: after() ile ana thread'e geç
                self.after(0, lambda: self._guncelleme_bildirimi_goster(sonuc))

        t = threading.Thread(target=_kontrol, daemon=True)
        t.start()

    def _guncelleme_bildirimi_goster(self, bilgi: dict):
        """Yeni sürüm bulunduğunda özel bildirim penceresi açar."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Güncelleme Mevcut 🎉")
        dialog.geometry("480x320")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        # Başlık
        ctk.CTkLabel(
            dialog,
            text="🚀  Yeni Sürüm Kullanılabilir!",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=15, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(pady=(22, 4))

        ctk.CTkLabel(
            dialog,
            text=f"Mevcut sürüm: v{APP_VERSION}   →   Yeni sürüm: v{bilgi['surum']}",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#475569", "#94A3B8")
        ).pack(pady=(0, 12))

        # Sürüm notları
        kart = ctk.CTkFrame(
            dialog,
            fg_color=("#F1F5F9", "#1E293B"),
            corner_radius=8,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        kart.pack(fill="x", padx=24, pady=(0, 14))

        ctk.CTkLabel(
            kart,
            text="Sürüm Notları:",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"),
            text_color=("#334155", "#CBD5E1")
        ).pack(anchor="w", padx=12, pady=(8, 2))

        notlar_kutusu = ctk.CTkTextbox(
            kart,
            height=90,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10),
            fg_color="transparent",
            border_width=0
        )
        notlar_kutusu.pack(fill="x", padx=8, pady=(0, 8))
        notlar_kutusu.insert("0.0", bilgi["notlar"])
        notlar_kutusu.configure(state="disabled")

        # Butonlar
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=(0, 18))

        def github_ac():
            webbrowser.open(bilgi["url"])
            dialog.destroy()

        ctk.CTkButton(
            btn_frame,
            text="GitHub'da İncele & İndir",
            command=github_ac,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            btn_frame,
            text="Daha Sonra",
            command=dialog.destroy,
            fg_color=("#E2E8F0", "#334155"),
            hover_color=("#CBD5E1", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            height=36,
            corner_radius=8
        ).pack(side="left", fill="x", expand=True)

    def _arayuz_olustur(self):
        # 1. ÜST HEADER
        header = ctk.CTkFrame(
            self,
            fg_color=("#FFFFFF", "#1E293B"),
            height=64,
            corner_radius=0,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        logo_alani = ctk.CTkFrame(header, fg_color="transparent")
        logo_alani.pack(side="left", padx=24, pady=12)

        ctk.CTkLabel(
            logo_alani,
            text="⚖ UYAP Evrak Asistanı",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=17, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(side="left")

        ctk.CTkLabel(
            logo_alani,
            text="PRO",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"),
            fg_color=("#EEF2F6", "#334155"),
            text_color=("#475569", "#E2E8F0"),
            corner_radius=6,
            padx=7,
            pady=2
        ).pack(side="left", padx=10)

        sag_aksiyon_frame = ctk.CTkFrame(header, fg_color="transparent")
        sag_aksiyon_frame.pack(side="right", padx=24)

        mevcut_tema = self.ayarlar.get("tema", "Light")
        self.tema_switch = ctk.CTkSwitch(
            sag_aksiyon_frame,
            text="🌙" if mevcut_tema == "Dark" else "☀️",
            command=self._tema_degistir,
            font=ctk.CTkFont(size=14),
            text_color=("#0F172A", "#F8FAFC"),
            width=46,
            height=24,
            switch_width=38,
            switch_height=20,
            progress_color=("#334155", "#64748B")
        )
        if mevcut_tema == "Dark":
            self.tema_switch.select()
        else:
            self.tema_switch.deselect()

        self.tema_switch.pack(side="left", padx=(0, 14))

        ctk.CTkButton(
            sag_aksiyon_frame,
            text="Nasıl Kullanılır?",
            command=self.rehber_penceresi_ac,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            corner_radius=8,
            width=120,
            height=32
        ).pack(side="left")

        # 2. ALT DURUM ÇUBUĞU
        alt_bar = ctk.CTkFrame(
            self,
            fg_color=("#FFFFFF", "#1E293B"),
            height=38,
            corner_radius=0,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        alt_bar.pack(fill="x", side="bottom")
        alt_bar.pack_propagate(False)

        self.progress_bar = ctk.CTkProgressBar(alt_bar, width=140, height=8, mode="indeterminate", corner_radius=4)
        self.progress_bar.pack(side="right", padx=16, pady=14)
        self.progress_bar.set(0)

        ctk.CTkButton(
            alt_bar,
            text="⚙",
            command=self.ayarlar_penceresi_ac,
            font=ctk.CTkFont(size=15),
            fg_color="transparent",
            hover_color=("#E2E8F0", "#334155"),
            text_color=("#64748B", "#94A3B8"),
            width=32,
            height=26,
            corner_radius=6
        ).pack(side="right", padx=(0, 4))

        self.durum_etiketi = ctk.CTkLabel(
            alt_bar,
            text="Hazır. Evraklarınızı sol alana sürükleyebilir veya 'Dosya Ekle' ile seçebilirsiniz.",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#64748B", "#94A3B8")
        )
        self.durum_etiketi.pack(side="left", padx=18)

        # 3. ANA GÖVDE
        govde = ctk.CTkFrame(self, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=20, pady=16)

        # SOL KART
        sol_kart = ctk.CTkFrame(
            govde,
            fg_color=("#FFFFFF", "#1E293B"),
            corner_radius=12,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        sol_kart.pack(side="left", fill="both", expand=True, padx=(0, 10))

        sol_baslik_satir = ctk.CTkFrame(sol_kart, fg_color="transparent")
        sol_baslik_satir.pack(fill="x", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            sol_baslik_satir,
            text="İŞLENECEK DOSYALAR",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
            text_color=("#94A3B8", "#64748B")
        ).pack(side="left")

        self.dosya_sayisi_badge = ctk.CTkLabel(
            sol_baslik_satir,
            text="0 dosya",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#64748B", "#94A3B8")
        )
        self.dosya_sayisi_badge.pack(side="right")

        self.liste_frame = ctk.CTkScrollableFrame(
            sol_kart,
            fg_color=("#F8FAFC", "#0F172A"),
            corner_radius=8
        )
        self.liste_frame.pack(fill="both", expand=True, padx=16, pady=6)

        self.bos_bilgi_etiketi = ctk.CTkLabel(
            self.liste_frame,
            text="Belgeleri buraya sürükleyin\nveya aşağıdaki butondan ekleyin\n(PDF, Word, Görsel, UDF)\n\n💡 İpucu: PDF'e çift tıklayarak sayfalarını görsel önizleyebilirsiniz.",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#94A3B8", "#64748B"),
            justify="center"
        )
        self.bos_bilgi_etiketi.pack(pady=110)

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
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8,
            height=34
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            liste_alt_bar,
            text="▲",
            command=lambda: self._dosya_sirala(-1),
            width=36,
            height=34,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            corner_radius=8
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            liste_alt_bar,
            text="▼",
            command=lambda: self._dosya_sirala(1),
            width=36,
            height=34,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            corner_radius=8
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            liste_alt_bar,
            text="Kaldır",
            command=self.secileni_sil,
            width=65,
            height=34,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            corner_radius=8
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            liste_alt_bar,
            text="Temizle",
            command=self.listeyi_temizle,
            width=65,
            height=34,
            fg_color=("#FEE2E2", "#7F1D1D"),
            hover_color=("#FECACA", "#991B1B"),
            text_color=("#991B1B", "#FCA5A5"),
            corner_radius=8
        ).pack(side="right")

        # SAĞ KART
        sag_kart = ctk.CTkFrame(
            govde,
            fg_color=("#FFFFFF", "#1E293B"),
            width=480,
            corner_radius=12,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        sag_kart.pack(side="right", fill="both")
        sag_kart.pack_propagate(False)

        self.tabview = ctk.CTkTabview(
            sag_kart,
            fg_color="transparent",
            segmented_button_fg_color=("#E2E8F0", "#0F172A"),
            segmented_button_selected_color=("#1E293B", "#3B82F6"),
            segmented_button_selected_hover_color=("#0F172A", "#2563EB"),
            segmented_button_unselected_color=("#F8FAFC", "#1E293B"),
            segmented_button_unselected_hover_color=("#E2E8F0", "#334155"),
            corner_radius=8
        )
        self.tabview.pack(fill="both", expand=True, padx=14, pady=10)

        def _sekme_renklerini_guncelle(secilen_sekme):
            try:
                for ad, btn in self.tabview._segmented_button._buttons_dict.items():
                    if ad == secilen_sekme:
                        btn.configure(text_color="#FFFFFF")
                    else:
                        btn.configure(text_color=("#334155", "#94A3B8"))
            except Exception:
                pass

        self._sekme_renklerini_guncelle = _sekme_renklerini_guncelle

        def _sekme_sec(ad):
            if IS_MAC:
                # CTkTabview.set eski sekmeyi 100 ms sonra gizliyor; macOS'ta (Tk 9) üstteki sekme
                # gizlenince alttaki yeniden çizilmiyor ve eski içerik ekranda kalıyor. Önce gizle.
                for sekme_adi, cerceve in self.tabview._tab_dict.items():
                    if sekme_adi != ad:
                        cerceve.grid_forget()
            self.tabview.set(ad)

        self._sekme_sec = _sekme_sec

        try:
            self.tabview._segmented_button.configure(
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                height=38,
                command=lambda val: (self._sekme_renklerini_guncelle(val), self._sekme_sec(val))
            )
        except Exception:
            pass

        self.tab_birlestir = self.tabview.add("Birleştir")
        self.tab_duzenle = self.tabview.add("Sayfalar")
        self.tab_donusturucu = self.tabview.add("Dönüştür")
        self.tab_guvenlik = self.tabview.add("Araçlar")

        self._sekme_birlestir_olustur()
        self._sekme_duzenle_olustur()
        self._sekme_donusturucu_olustur()
        self._sekme_guvenlik_olustur()

        self._sekme_renklerini_guncelle(self.tabview.get())

    def _tema_degistir(self, kaynak=None):
        # Hangi switch tetiklediyse onun değeri esas alınır; diğeri ona eşlenir
        kaynak = kaynak or self.tema_switch
        yeni_dark = kaynak.get() == 1

        # Ayarlar penceresi kapandıysa switch'i yok edilmiştir; ona dokunulmamalı
        ayar_switch = getattr(self, "_tema_ayar_switch", None)
        try:
            if ayar_switch is not None and not ayar_switch.winfo_exists():
                ayar_switch = None
        except tk.TclError:
            ayar_switch = None

        for sw in (self.tema_switch, ayar_switch):
            if sw is not None and sw is not kaynak:
                sw.select() if yeni_dark else sw.deselect()

        if yeni_dark:
            ctk.set_appearance_mode("Dark")
            self.tema_switch.configure(text="🌙")
            self.ayarlar["tema"] = "Dark"
            ayar_kaydet("tema", "Dark")
        else:
            ctk.set_appearance_mode("Light")
            self.tema_switch.configure(text="☀️")
            self.ayarlar["tema"] = "Light"
            ayar_kaydet("tema", "Light")
        self._liste_arayuzunu_yenile()
        self._sekme_renklerini_guncelle(self.tabview.get())

    # ========================================================
    # AYARLAR PENCERESİ (Sağ alt ⚙ butonundan açılır)
    # ========================================================
    def ayarlar_penceresi_ac(self):
        """Sağ alt köşedeki ⚙ butonuna basıldığında ayarlar penceresini açar."""
        # Zaten açıksa öne getir, yenisini açma
        if hasattr(self, "_ayarlar_penceresi") and self._ayarlar_penceresi.winfo_exists():
            self._ayarlar_penceresi.lift()
            self._ayarlar_penceresi.focus()
            return

        pencere = ctk.CTkToplevel(self)
        pencere.title("Ayarlar")
        pencere.geometry("480x560")
        pencere.resizable(False, True)
        self._ayarlar_penceresi = pencere

        # Ana pencereye göre konumlandır (sağ alt)
        def _konumlandir():
            pw = self.winfo_x() + self.winfo_width()
            ph = self.winfo_y() + self.winfo_height()
            pencere.geometry(f"480x560+{pw - 500}+{ph - 590}")
        self.after(50, _konumlandir)

        # CTkToplevel render tamamlandıktan sonra içeriği inşa et
        def _icerik_olustur():
            if not pencere.winfo_exists():
                return

            scrollable = ctk.CTkScrollableFrame(
                pencere,
                fg_color=("#F1F5F9", "#0F172A")
            )
            scrollable.pack(fill="both", expand=True, padx=0, pady=0)

            def baslik(metin):
                ctk.CTkLabel(
                    scrollable, text=metin,
                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"),
                    text_color=("#94A3B8", "#64748B")
                ).pack(anchor="w", padx=16, pady=(14, 4))

            def kart():
                f = ctk.CTkFrame(
                    scrollable,
                    fg_color=("#FFFFFF", "#1E293B"),
                    corner_radius=10,
                    border_width=1,
                    border_color=("#E2E8F0", "#334155")
                )
                f.pack(fill="x", padx=12, pady=3)
                return f

            # ── UYGULAMA AYARLARI ─────────────────────────────────
            baslik("UYGULAMA AYARLARI")

            k1 = kart()
            ust_satir = ctk.CTkFrame(k1, fg_color="transparent")
            ust_satir.pack(fill="x", padx=14, pady=(12, 4))
            ctk.CTkLabel(
                ust_satir,
                text="🔄  Otomatik Güncelleme Denetimi",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            ).pack(side="left")

            self._guncelleme_switch_var = ctk.BooleanVar(
                value=bool(self.ayarlar.get("otomatik_guncelleme_kontrol", True))
            )

            def _guncelleme_toggle():
                yeni = self._guncelleme_switch_var.get()
                self.ayarlar["otomatik_guncelleme_kontrol"] = yeni
                ayar_kaydet("otomatik_guncelleme_kontrol", yeni)
                durum = "etkin" if yeni else "devre dışı"
                self.durum_etiketi.configure(
                    text=f"Otomatik güncelleme denetimi {durum} bırakıldı."
                )

            ctk.CTkSwitch(
                ust_satir,
                text="",
                variable=self._guncelleme_switch_var,
                command=_guncelleme_toggle,
                width=42,
                switch_width=36,
                switch_height=20,
                progress_color=("#334155", "#3B82F6")
            ).pack(side="right")

            ctk.CTkLabel(
                k1,
                text="Uygulama açılırken GitHub'dan yeni sürüm olup olmadığını\nkontrol eder. İnternet yoksa sessizce atlanır.",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                text_color=("#64748B", "#94A3B8"),
                justify="left"
            ).pack(anchor="w", padx=14, pady=(0, 12))

            # Manuel kontrol kartı
            k2 = kart()
            ctk.CTkLabel(
                k2,
                text="Manuel Güncelleme Kontrolü",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            ).pack(anchor="w", padx=14, pady=(12, 2))

            ctk.CTkLabel(
                k2,
                text=f"Şu anki sürüm: v{APP_VERSION}",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                text_color=("#64748B", "#94A3B8")
            ).pack(anchor="w", padx=14, pady=(0, 8))

            self._manuel_kontrol_btn = ctk.CTkButton(
                k2,
                text="🔍  Şimdi Güncelleme Ara",
                command=self._manuel_guncelleme_kontrol,
                fg_color=("#0F172A", "#3B82F6"),
                hover_color=("#1E293B", "#2563EB"),
                text_color="#FFFFFF",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                height=34,
                corner_radius=8
            )
            self._manuel_kontrol_btn.pack(fill="x", padx=14, pady=(0, 14))

            # ── YAPAY ZEKA ────────────────────────────────────────
            baslik("YAPAY ZEKA")
            self._ai_ayar_karti_olustur(kart())

            # ── GÖRÜNÜM AYARLARI ──────────────────────────────────
            baslik("GÖRÜNÜM AYARLARI")
            k3 = kart()
            tema_satir = ctk.CTkFrame(k3, fg_color="transparent")
            tema_satir.pack(fill="x", padx=14, pady=12)
            ctk.CTkLabel(
                tema_satir,
                text="🌗  Karanlık Tema",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            ).pack(side="left")
            ctk.CTkLabel(
                tema_satir,
                text="(Başlık switch'i ile de değiştirilebilir)",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10),
                text_color=("#94A3B8", "#64748B")
            ).pack(side="left", padx=8)

            self._tema_ayar_switch = ctk.CTkSwitch(
                tema_satir,
                text="",
                command=lambda: self._tema_degistir(self._tema_ayar_switch),
                width=42,
                switch_width=36,
                switch_height=20,
                progress_color=("#334155", "#3B82F6")
            )
            if self.ayarlar.get("tema", "Light") == "Dark":
                self._tema_ayar_switch.select()
            self._tema_ayar_switch.pack(side="right")

            # ── HAKKINDA ──────────────────────────────────────────
            baslik("HAKKINDA")
            k4 = kart()
            for satir_metin, satir_font_boyut in [
                (f"⚖  UYAP Evrak Asistanı  v{APP_VERSION}", 13),
                ("Geliştirici: Tolunay Akay", 11),
                ("tolunay.akay@galahukuk.com", 11),
                ("github.com/tolunayakay/uyap-evrak-asistani", 10),
            ]:
                ctk.CTkLabel(
                    k4,
                    text=satir_metin,
                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=satir_font_boyut),
                    text_color=("#334155", "#CBD5E1") if satir_font_boyut == 13
                    else ("#64748B", "#94A3B8")
                ).pack(anchor="w", padx=14, pady=(8 if satir_font_boyut == 13 else 1, 0))

            ctk.CTkButton(
                k4,
                text="GitHub'da Aç",
                command=lambda: webbrowser.open(
                    "https://github.com/tolunayakay/uyap-evrak-asistani"
                ),
                fg_color=("#E2E8F0", "#334155"),
                hover_color=("#CBD5E1", "#475569"),
                text_color=("#334155", "#F8FAFC"),
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                height=30,
                corner_radius=8
            ).pack(anchor="w", padx=14, pady=(6, 14))

            # ── EMEĞİ GEÇENLER ────────────────────────────────────
            baslik("EMEĞİ GEÇENLER (CREDITS)")
            k5 = kart()
            # Hakkında kartındaki satırlarla aynı yazı tipi, boyut ve renk
            emegi_gecenler = ("Zafer Kaya", "yanu34")
            for n, isim in enumerate(emegi_gecenler):
                ctk.CTkLabel(
                    k5,
                    text=f"- {isim}",
                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                    text_color=("#64748B", "#94A3B8")
                ).pack(anchor="w", padx=14, pady=(8 if n == 0 else 1, 10 if n == len(emegi_gecenler) - 1 else 0))

            pencere.lift()
            pencere.focus()

        # Pencere render olduktan sonra içeriği çiz
        pencere.after(200, _icerik_olustur)

    def _manuel_guncelleme_kontrol(self):
        """Kullanıcı 'Şimdi Ara' butonuna basınca arka planda kontrol yapar."""
        self._manuel_kontrol_btn.configure(
            text="⏳  Kontrol ediliyor...", state="disabled"
        )
        self.durum_etiketi.configure(text="GitHub güncelleme kontrolü yapılıyor...")

        def _kontrol():
            sonuc = guncelleme_kontrol_et()
            def _guncelle_ui():
                self._manuel_kontrol_btn.configure(
                    text="🔍  Şimdi Güncelleme Ara", state="normal"
                )
                if sonuc:
                    self._guncelleme_bildirimi_goster(sonuc)
                else:
                    self.durum_etiketi.configure(
                        text=f"✅  Zaten en güncel sürümdesiniz (v{APP_VERSION})."
                    )
            self.after(0, _guncelle_ui)

        threading.Thread(target=_kontrol, daemon=True).start()

    # ========================================================
    # AI EVRAK ÖZETİ: AYARLAR, MODEL İNDİRME, ÖZET PENCERESİ
    # ========================================================
    def _ai_ozet_kullanilabilir(self):
        return bool(self.ayarlar.get("ai_ozet_aktif", False))

    def _ai_ayar_karti_olustur(self, kart):
        """Ayarlar penceresindeki AI kartını (yeniden) çizer."""
        for w in kart.winfo_children():
            w.destroy()

        ust = ctk.CTkFrame(kart, fg_color="transparent")
        ust.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(
            ust, text="✨  AI Evrak Özeti",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(side="left")

        self._ai_switch_var = ctk.BooleanVar(value=self._ai_ozet_kullanilabilir())
        ctk.CTkSwitch(
            ust, text="", variable=self._ai_switch_var,
            command=lambda: self._ai_switch_degisti(kart),
            width=42, switch_width=36, switch_height=20,
            progress_color=("#334155", "#3B82F6")
        ).pack(side="right")

        ctk.CTkLabel(
            kart,
            text="Listedeki evrakların yanında 'AI Özet' butonu gösterir. Özet, internete\n"
                 f"veri göndermeden bilgisayarınızda{'' if IS_MAC else ' (CPU)'} çalışan {AI_MODEL_ADI} modeliyle üretilir.",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#64748B", "#94A3B8"),
            justify="left"
        ).pack(anchor="w", padx=14, pady=(0, 6))

        if not ai_kutuphane_mevcut():
            ctk.CTkLabel(
                kart, text="⚠ llama-cpp-python kütüphanesi bulunamadı; özellik kullanılamaz.",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
                text_color=("#B91C1C", "#F87171")
            ).pack(anchor="w", padx=14, pady=(0, 12))
            return

        self._ai_islem_birimi_satiri(kart)

        if ai_model_indirildi():
            boyut = os.path.getsize(ai_model_yolu()) / (1024 * 1024)
            ctk.CTkLabel(
                kart, text=f"✅ Model indirildi ({boyut:,.0f} MB)".replace(",", "."),
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
                text_color=("#166534", "#4ADE80")
            ).pack(anchor="w", padx=14, pady=(0, 6))
            ctk.CTkButton(
                kart, text="🗑  Modeli Sil",
                command=lambda: self._ai_model_sil(kart),
                fg_color=("#FEE2E2", "#7F1D1D"),
                hover_color=("#FECACA", "#991B1B"),
                text_color=("#991B1B", "#FEE2E2"),
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                height=32, corner_radius=8
            ).pack(fill="x", padx=14, pady=(0, 14))
        else:
            ctk.CTkLabel(
                kart, text=f"Model indirilmedi (yaklaşık {AI_MODEL_BOYUT_METNI}, tek seferlik).",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                text_color=("#64748B", "#94A3B8")
            ).pack(anchor="w", padx=14, pady=(0, 6))
            ctk.CTkButton(
                kart, text=f"⬇  Modeli İndir (~{AI_MODEL_BOYUT_METNI})",
                command=lambda: self._ai_model_indirme_penceresi(lambda: self._ai_kart_yenile(kart)),
                fg_color=("#0F172A", "#3B82F6"),
                hover_color=("#1E293B", "#2563EB"),
                text_color="#FFFFFF",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                height=32, corner_radius=8
            ).pack(fill="x", padx=14, pady=(0, 14))

    def _ai_islem_birimi_satiri(self, kart):
        """Otomatik / İşlemci / Ekran Kartı seçimi ve algılanan kartların bilgisi."""
        satir = ctk.CTkFrame(kart, fg_color="transparent")
        satir.pack(fill="x", padx=14, pady=(4, 2))
        ctk.CTkLabel(satir, text="İşlem birimi:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
                     text_color=("#0F172A", "#F8FAFC")).pack(side="left")

        motor_var = ai_motor_klasoru() is not None or ai_metal_destekli()
        ad_to_kod = {v: k for k, v in AI_ISLEM_BIRIMLERI.items()}
        kayitli = self.ayarlar.get("ai_islem_birimi", "otomatik")
        bilgi_lbl = ctk.CTkLabel(kart, text="", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                                 text_color=("#64748B", "#94A3B8"), justify="left", wraplength=420)

        def bilgi_yaz(gpular=None):
            try:
                if not bilgi_lbl.winfo_exists():
                    return
            except tk.TclError:
                return
            if not motor_var:
                bilgi_lbl.configure(text="Bu sürümde ekran kartı desteği yok (ai_motor bulunamadı); işlemci kullanılır."
                                    if not IS_MAC else "Kurulu AI kütüphanesi Metal desteği içermiyor; işlemci kullanılır.")
                return
            if IS_MAC:
                mod = self.ayarlar.get("ai_islem_birimi", "otomatik")
                secilen = "İşlemci (CPU)" if mod == "cpu" else "Apple ekran kartı (Metal)"
                bilgi_lbl.configure(text=f"Kullanılacak: {secilen}\n(Apple Silicon'da ekran kartı ile özetler belirgin şekilde daha hızlı üretilir.)")
                return
            if gpular is None:
                bilgi_lbl.configure(text="Ekran kartları algılanıyor...")
                return
            mod = self.ayarlar.get("ai_islem_birimi", "otomatik")
            if gpular:
                kartlar = " • ".join(f"{g['ad']} ({g['tur']}, {g['bellek_mb'] / 1024:.0f} GB)".replace(".", ",") for g in gpular)
                metin = f"Algılanan: {kartlar}"
            else:
                metin = "Ekran kartı algılanamadı."
            vk, secilen = ai_cihaz_sec(mod, gpular)
            metin += f"\nKullanılacak: {secilen}"
            if mod == "otomatik":
                metin += "\n(Otomatik: yeterli belleğe sahip harici kart varsa o, yoksa işlemci. Dahili kartlar işlemciden yavaş olduğu için atlanır.)"
            if _AI_MOTOR_DURUMU is not None and _AI_MOTOR_DURUMU.get("motor") == "vulkan" and vk != _AI_MOTOR_DURUMU.get("vk"):
                metin += "\n⚠ Bu değişiklik uygulama yeniden başlatıldığında geçerli olur."
            bilgi_lbl.configure(text=metin)

        def degisti(gorunen):
            kod = ad_to_kod.get(gorunen, "otomatik")
            self.ayarlar["ai_islem_birimi"] = kod
            ayar_kaydet("ai_islem_birimi", kod)
            if not getattr(self, "_ai_ozet_calisiyor", False):
                AiOzetMotoru.bosalt()   # bir sonraki özette yeni ayarla yüklensin
            bilgi_yaz(_AI_GPU_ONBELLEK if _AI_GPU_ONBELLEK is not None else None)
            self.durum_etiketi.configure(text=f"AI işlem birimi: {gorunen}")

        secici = ctk.CTkSegmentedButton(satir, values=list(AI_ISLEM_BIRIMLERI.values()), command=degisti,
                                        height=28, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"))
        secici.set(AI_ISLEM_BIRIMLERI.get(kayitli, "Otomatik"))
        secici.pack(side="right")
        if not motor_var:
            secici.configure(state="disabled")
        bilgi_lbl.pack(anchor="w", padx=14, pady=(0, 8))

        if IS_MAC or not motor_var or _AI_GPU_ONBELLEK is not None:
            bilgi_yaz(_AI_GPU_ONBELLEK)
        else:
            bilgi_yaz(None)
            threading.Thread(target=ai_gpu_listele, daemon=True).start()

            # İş parçacığı yalnızca önbelleği doldurur; arayüz ana iş parçacığından yoklanır
            def bekle():
                try:
                    if not bilgi_lbl.winfo_exists():
                        return
                except tk.TclError:
                    return
                if _AI_GPU_ONBELLEK is None:
                    bilgi_lbl.after(300, bekle)
                else:
                    bilgi_yaz(_AI_GPU_ONBELLEK)
            bilgi_lbl.after(300, bekle)

    def _ai_kart_yenile(self, kart):
        try:
            if kart.winfo_exists():
                self._ai_ayar_karti_olustur(kart)
        except tk.TclError:
            pass

    def _ai_ayar_yaz(self, deger):
        self.ayarlar["ai_ozet_aktif"] = deger
        ayar_kaydet("ai_ozet_aktif", deger)
        self._liste_arayuzunu_yenile()

    def _ai_switch_degisti(self, kart):
        acik = bool(self._ai_switch_var.get())
        if not acik:
            self._ai_ayar_yaz(False)
            self.durum_etiketi.configure(text="AI Evrak Özeti kapatıldı.")
            return

        if not ai_kutuphane_mevcut():
            self._ai_switch_var.set(False)
            messagebox.showerror("Kütüphane Yok", "llama-cpp-python kütüphanesi bulunamadı.\nKurulum: pip install llama-cpp-python")
            return

        self._ai_ayar_yaz(True)
        self.durum_etiketi.configure(text="AI Evrak Özeti açıldı.")
        if not ai_model_indirildi():
            if messagebox.askyesno(
                "Model Gerekli",
                f"AI özeti için yaklaşık {AI_MODEL_BOYUT_METNI}'lık {AI_MODEL_ADI} model dosyasının bir kez indirilmesi gerekiyor.\n\n"
                "Şimdi indirilsin mi?",
                parent=self._ayarlar_penceresi if hasattr(self, "_ayarlar_penceresi") else self
            ):
                self._ai_model_indirme_penceresi(lambda: self._ai_kart_yenile(kart))
            else:
                self._ai_switch_var.set(False)
                self._ai_ayar_yaz(False)

    def _ai_model_sil(self, kart):
        if getattr(self, "_ai_ozet_calisiyor", False):
            messagebox.showwarning("Meşgul", "Bir özet işlemi sürerken model silinemez.")
            return
        if not messagebox.askyesno(
            "Modeli Sil",
            "AI modeli bilgisayarınızdan silinecek ve AI Evrak Özeti kapatılacak.\n"
            "Tekrar kullanmak için modeli yeniden indirmeniz gerekir. Devam edilsin mi?",
            parent=self._ayarlar_penceresi if hasattr(self, "_ayarlar_penceresi") else self
        ):
            return
        AiOzetMotoru.bosalt()
        try:
            for yol in (ai_model_yolu(), ai_model_yolu() + ".part"):
                if os.path.exists(yol):
                    os.remove(yol)
        except OSError as e:
            messagebox.showerror("Hata", f"Model silinemedi:\n{e}")
            return
        self._ai_ayar_yaz(False)
        self._ai_kart_yenile(kart)
        self.durum_etiketi.configure(text="AI modeli silindi.")

    def _ai_model_indirme_penceresi(self, bitince=None):
        """Modeli arka planda indirir; yüzde ve MB ilerlemesini gösterir. Yarım kalan indirme devam ettirilir."""
        if getattr(self, "_ai_indirme_penceresi", None) is not None and self._ai_indirme_penceresi.winfo_exists():
            self._ai_indirme_penceresi.lift()
            return

        pencere = ctk.CTkToplevel(self)
        pencere.title("AI Modeli İndiriliyor")
        pencere.geometry("460x250")
        pencere.resizable(False, False)
        pencere.after(100, pencere.lift)
        self._ai_indirme_penceresi = pencere

        ctk.CTkLabel(pencere, text="⬇  AI Modeli İndiriliyor", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=15, weight="bold")).pack(anchor="w", padx=20, pady=(18, 2))
        ctk.CTkLabel(pencere, text=AI_MODEL_DOSYA_ADI, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#64748B", "#94A3B8")).pack(anchor="w", padx=20)

        yuzde_lbl = ctk.CTkLabel(pencere, text="%0", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=22, weight="bold"))
        yuzde_lbl.pack(anchor="w", padx=20, pady=(12, 0))
        bar = ctk.CTkProgressBar(pencere, height=14)
        bar.set(0)
        bar.pack(fill="x", padx=20, pady=(4, 6))
        detay_lbl = ctk.CTkLabel(pencere, text="Bağlanıyor...", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#475569", "#CBD5E1"))
        detay_lbl.pack(anchor="w", padx=20)

        iptal = threading.Event()
        hedef = ai_model_yolu()
        gecici = hedef + ".part"
        # İş parçacığı yalnızca bu sözlüğü günceller; arayüzü ana iş parçacığı okur
        durum = {"inen": 0, "toplam": 0, "hiz": 0.0, "bitti": False, "hata": None, "iptal": False}

        def mb(b):
            return f"{b / (1024 * 1024):,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")

        def guncelle(inen, toplam, hiz):
            durum.update(inen=inen, toplam=toplam, hiz=hiz)

        def bitti(hata=None, iptal_edildi=False):
            durum.update(hata=hata, iptal=iptal_edildi, bitti=True)

        def ekran_guncelle():
            if durum["bitti"]:
                self._ai_indirme_penceresi = None
                if pencere.winfo_exists():
                    pencere.destroy()
                if durum["iptal"]:
                    self.durum_etiketi.configure(text="Model indirme iptal edildi (kaldığı yerden devam ettirilebilir).")
                elif durum["hata"]:
                    messagebox.showerror("İndirme Hatası", f"Model indirilemedi:\n{durum['hata']}\n\nTekrar denediğinizde kaldığı yerden devam eder.")
                else:
                    self.durum_etiketi.configure(text="AI modeli indirildi. AI Özet kullanıma hazır.")
                    messagebox.showinfo("Tamamlandı", "AI modeli indirildi.\nListedeki evrakların yanındaki '✨ AI Özet' butonunu kullanabilirsiniz.")
                    self._liste_arayuzunu_yenile()
                if bitince:
                    bitince()
                return
            if not pencere.winfo_exists():
                return
            inen, toplam, hiz = durum["inen"], durum["toplam"], durum["hiz"]
            if toplam:
                oran = inen / toplam
                bar.set(oran)
                yuzde_lbl.configure(text=f"%{oran * 100:.0f}")
                kalan = ""
                if hiz > 0:
                    sn = (toplam - inen) / hiz
                    kalan = f"  •  kalan ~{int(sn // 60)} dk {int(sn % 60)} sn"
                detay_lbl.configure(text=f"{mb(inen)} / {mb(toplam)} MB  •  {mb(hiz)} MB/sn{kalan}")
            pencere.after(200, ekran_guncelle)

        def indir():
            import time
            try:
                mevcut = os.path.getsize(gecici) if os.path.exists(gecici) else 0
                basliklar = {"User-Agent": f"UyapEvrakAsistani/{APP_VERSION}"}
                if mevcut:
                    basliklar["Range"] = f"bytes={mevcut}-"
                istek = urllib.request.Request(AI_MODEL_URL, headers=basliklar)
                with urllib.request.urlopen(istek, timeout=30) as yanit:
                    if mevcut and yanit.status != 206:
                        mevcut = 0  # Sunucu devam etmeyi desteklemedi; baştan indir
                    kalan_boyut = int(yanit.headers.get("Content-Length", 0))
                    toplam = mevcut + kalan_boyut
                    inen = mevcut
                    son_ui, pencere_bas, pencere_inen = 0.0, time.time(), 0
                    hiz = 0.0
                    with open(gecici, "ab" if mevcut else "wb") as f:
                        while True:
                            if iptal.is_set():
                                bitti(iptal_edildi=True)
                                return
                            parca = yanit.read(1024 * 256)
                            if not parca:
                                break
                            f.write(parca)
                            inen += len(parca)
                            pencere_inen += len(parca)
                            simdi = time.time()
                            if simdi - pencere_bas >= 1.0:
                                hiz = pencere_inen / (simdi - pencere_bas)
                                pencere_bas, pencere_inen = simdi, 0
                            if simdi - son_ui >= 0.2:
                                son_ui = simdi
                                guncelle(inen, toplam, hiz)
                guncelle(inen, toplam, hiz)
                if toplam and inen != toplam:
                    raise IOError("Bağlantı kesildi, dosya eksik indi.")
                with open(gecici, "rb") as f:
                    if f.read(4) != b"GGUF":
                        os.remove(gecici)
                        raise IOError("İndirilen dosya geçerli bir GGUF modeli değil.")
                os.replace(gecici, hedef)
                bitti()
            except Exception as e:
                bitti(hata=str(e))

        def iptal_et():
            if messagebox.askyesno("İptal", "Model indirmesi durdurulsun mu?\n(Daha sonra kaldığı yerden devam ettirilebilir.)", parent=pencere):
                iptal.set()

        ctk.CTkButton(
            pencere, text="İptal", command=iptal_et,
            fg_color=("#E2E8F0", "#334155"), hover_color=("#CBD5E1", "#475569"),
            text_color=("#334155", "#F8FAFC"), height=32, corner_radius=8
        ).pack(pady=(14, 0))
        pencere.protocol("WM_DELETE_WINDOW", iptal_et)

        threading.Thread(target=indir, daemon=True).start()
        pencere.after(200, ekran_guncelle)

    def ai_ozet_penceresi(self, yol):
        if not ai_kutuphane_mevcut():
            messagebox.showerror("Kütüphane Yok", "llama-cpp-python kütüphanesi bulunamadı.\nKurulum: pip install llama-cpp-python")
            return
        if not ai_model_indirildi():
            if messagebox.askyesno("Model Gerekli", f"{AI_MODEL_ADI} AI modeli henüz indirilmemiş (~{AI_MODEL_BOYUT_METNI}).\nŞimdi indirilsin mi?"):
                self._ai_model_indirme_penceresi()
            return
        if getattr(self, "_ai_ozet_calisiyor", False):
            messagebox.showwarning("Meşgul", "Başka bir evrak özetleniyor. Lütfen bitmesini bekleyin.")
            return

        try:
            ham_metin = belge_metni_cikar(yol)
        except Exception as e:
            messagebox.showerror("Hata", f"Belge okunamadı:\n{e}")
            return
        if len(ham_metin.strip()) < 40:
            messagebox.showwarning(
                "Metin Bulunamadı",
                "Bu belgede okunabilir metin katmanı yok (taranmış görsel olabilir).\nAI özeti yalnızca metin içeren belgelerde çalışır."
            )
            return
        metin, kirpildi = ai_metni_hazirla(ham_metin)
        otomatik_tur = belge_turu_tespit(ham_metin)
        tur_adlari = {k: v["ad"] for k, v in AI_BELGE_TURLERI.items()}
        ad_to_tur = {v: k for k, v in tur_adlari.items()}

        ad = os.path.basename(yol)
        pencere = ctk.CTkToplevel(self)
        pencere.title(f"AI Özet: {ad}")
        pencere.geometry("680x600")
        pencere.minsize(520, 420)
        pencere.after(100, pencere.lift)

        ctk.CTkLabel(pencere, text=f"✨  {ad}", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=14, weight="bold")).pack(anchor="w", padx=16, pady=(14, 0))
        bilgi = f"{len(ham_metin):,} karakter".replace(",", ".")
        if kirpildi:
            bilgi += "  •  Belge uzun olduğu için giriş, gelişme ve sonuç/hüküm bölümleri özetlendi."
        ctk.CTkLabel(pencere, text=bilgi, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#64748B", "#94A3B8"), wraplength=640, justify="left").pack(anchor="w", padx=16)

        tur_satir = ctk.CTkFrame(pencere, fg_color="transparent")
        tur_satir.pack(fill="x", padx=16, pady=(8, 0))
        ctk.CTkLabel(tur_satir, text="Belge türü:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold")).pack(side="left")
        tur_menu = ctk.CTkOptionMenu(tur_satir, values=list(tur_adlari.values()), width=220, height=28,
                                     font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12))
        tur_menu.set(tur_adlari[otomatik_tur])
        tur_menu.pack(side="left", padx=8)
        tur_notu = ctk.CTkLabel(tur_satir, text="(otomatik algılandı)", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                                text_color=("#64748B", "#94A3B8"))
        tur_notu.pack(side="left")
        yeniden_btn = ctk.CTkButton(tur_satir, text="↻ Yeniden Özetle", width=130, height=28, state="disabled",
                                    fg_color=("#EDE9FE", "#4C1D95"), hover_color=("#DDD6FE", "#5B21B6"),
                                    text_color=("#5B21B6", "#EDE9FE"), corner_radius=6,
                                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"))
        yeniden_btn.pack(side="right")

        durum_lbl = ctk.CTkLabel(pencere, text="Model yükleniyor...", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"), text_color=("#1D4ED8", "#60A5FA"))
        durum_lbl.pack(anchor="w", padx=16, pady=(8, 2))
        bar = ctk.CTkProgressBar(pencere, mode="indeterminate", height=6)
        bar.pack(fill="x", padx=16)
        bar.start()

        kutu = ctk.CTkTextbox(pencere, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13), wrap="word")
        kutu.pack(fill="both", expand=True, padx=16, pady=(10, 6))

        ctk.CTkLabel(
            pencere,
            text="⚠ AI özeti hata içerebilir; hukuki işlemlerde daima asıl belgeyi esas alın.",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10), text_color=("#B45309", "#FBBF24")
        ).pack(anchor="w", padx=16)

        alt = ctk.CTkFrame(pencere, fg_color="transparent")
        alt.pack(fill="x", padx=16, pady=(6, 14))

        durdur = threading.Event()
        durum = {}
        import time

        def kopyala():
            pencere.clipboard_clear()
            pencere.clipboard_append(kutu.get("0.0", "end").strip())
            self.durum_etiketi.configure(text="AI özeti panoya kopyalandı.")

        kopyala_btn = ctk.CTkButton(alt, text="📋 Panoya Kopyala", command=kopyala, state="disabled",
                                    fg_color=("#0F172A", "#3B82F6"), hover_color=("#1E293B", "#2563EB"),
                                    text_color="#FFFFFF", corner_radius=8, height=32)
        kopyala_btn.pack(side="left")

        def kapat():
            durdur.set()
            pencere.destroy()

        durdur_btn = ctk.CTkButton(alt, text="⏹ Durdur", command=durdur.set,
                                   fg_color=("#E2E8F0", "#334155"), hover_color=("#CBD5E1", "#475569"),
                                   text_color=("#334155", "#F8FAFC"), corner_radius=8, height=32, width=100)
        durdur_btn.pack(side="right")
        pencere.protocol("WM_DELETE_WINDOW", kapat)

        def temizle(m):
            m = m.replace("**", "").replace("__", "")
            m = re.sub(r"(?m)^\s*#+\s*", "", m)
            m = re.sub(r"(?m)^\s*[-*]\s+", "• ", m)
            return m

        def ekran_guncelle():
            if not pencere.winfo_exists():
                return
            gecen = int(time.time() - durum["basla"])
            if durum["asama"] == "algiliyor":
                durum_lbl.configure(text=f"Ekran kartları algılanıyor...  ({gecen} sn)")
            elif durum["asama"] == "dogruluyor":
                durum_lbl.configure(text=f"Ekran kartı ilk kez doğrulanıyor (yalnızca bir kez yapılır)...  ({gecen} sn)")
            elif durum["asama"] == "yukleniyor":
                durum_lbl.configure(text=f"Model belleğe yükleniyor...  ({gecen} sn)")
            elif durum["asama"] == "okuyor":
                sure_notu = "Uzun belge; bilgisayar hızına göre birkaç dakika sürebilir." if durum["token"] > 4000 else "Kısa sürecek."
                durum_lbl.configure(text=f"Belge okunuyor (~{durum['token']:,} token)... {sure_notu}  ({gecen} sn)".replace(",", "."))
            elif durum["asama"] == "yaziyor":
                if bar.cget("mode") == "indeterminate":
                    bar.stop()
                    bar.configure(mode="determinate")
                    bar.set(1)
                durum_lbl.configure(text=f"Özet yazılıyor...  ({gecen} sn)")

            yeni = temizle(durum["metin"])
            if yeni != kutu.get("0.0", "end-1c"):
                kutu.delete("0.0", "end")
                kutu.insert("0.0", yeni)
                kutu.see("end")

            if durum["bitti"]:
                bar.stop()
                bar.configure(mode="determinate")
                bar.set(1)
                durdur_btn.configure(text="Kapat", command=kapat)
                kopyala_btn.configure(state="normal" if durum["metin"].strip() else "disabled")
                tur_menu.configure(state="normal")
                yeniden_btn.configure(state="normal")
                if durum["hata"]:
                    durum_lbl.configure(text="❌ Özet oluşturulamadı.", text_color=("#B91C1C", "#F87171"))
                    kutu.insert("end", f"\n\nHata: {durum['hata']}")
                elif durdur.is_set():
                    durum_lbl.configure(text=f"⏹ Durduruldu ({gecen} sn).", text_color=("#B45309", "#FBBF24"))
                else:
                    cihaz = f"  •  {durum['cihaz']}" if durum.get("cihaz") else ""
                    durum_lbl.configure(text=f"✅ Özet tamamlandı ({gecen} sn){cihaz}", text_color=("#166534", "#4ADE80"))
                self.durum_etiketi.configure(text=f"AI özeti hazır: {ad}")
                return
            pencere.after(150, ekran_guncelle)

        def calistir(tur):
            try:
                with AiOzetMotoru._kilit:
                    llm = AiOzetMotoru.yukle(durum_bildir=lambda a: durum.update(asama=a))
                    durum["cihaz"] = AiOzetMotoru.cihaz_aciklama
                    mesajlar, token = AiOzetMotoru.mesajlari_hazirla(llm, metin, tur)
                    durum["token"], durum["asama"] = token, "okuyor"
                    akis = llm.create_chat_completion(
                        messages=mesajlar,
                        max_tokens=AI_CIKTI_TOKEN,
                        temperature=0.1,
                        top_p=0.9,
                        # İsim/tutar/tarih aynen tekrarlanmalı; tekrar cezası bunları bozar
                        repeat_penalty=1.0,
                        stream=True,
                    )
                    for parca in akis:
                        if durdur.is_set():
                            break
                        icerik = parca["choices"][0].get("delta", {}).get("content")
                        if icerik:
                            durum["asama"] = "yaziyor"
                            durum["metin"] += icerik
                    try:
                        akis.close()
                    except Exception:
                        pass
            except Exception as e:
                durum["hata"] = str(e)
            finally:
                self._ai_ozet_calisiyor = False
                durum["bitti"] = True

        def baslat():
            if getattr(self, "_ai_ozet_calisiyor", False):
                messagebox.showwarning("Meşgul", "Başka bir evrak özetleniyor. Lütfen bitmesini bekleyin.", parent=pencere)
                return
            tur = ad_to_tur.get(tur_menu.get(), "genel")
            tur_notu.configure(text="(otomatik algılandı)" if tur == otomatik_tur else "(elle seçildi)")
            durdur.clear()
            durum.clear()
            durum.update(metin="", bitti=False, hata=None, asama="okuyor" if AiOzetMotoru.yuklu_mu() else "yukleniyor",
                         token=0, basla=time.time())
            kutu.delete("0.0", "end")
            durum_lbl.configure(text="Hazırlanıyor...", text_color=("#1D4ED8", "#60A5FA"))
            bar.configure(mode="indeterminate")
            bar.start()
            durdur_btn.configure(text="⏹ Durdur", command=durdur.set)
            kopyala_btn.configure(state="disabled")
            tur_menu.configure(state="disabled")
            yeniden_btn.configure(state="disabled")
            self._ai_ozet_calisiyor = True
            self.durum_etiketi.configure(text=f"AI özeti hazırlanıyor: {ad} ({tur_adlari[tur]})")
            threading.Thread(target=calistir, args=(tur,), daemon=True).start()
            pencere.after(150, ekran_guncelle)

        yeniden_btn.configure(command=baslat)
        baslat()

    # ========================================================
    # SEKME 1: BİRLEŞTİR
    # ========================================================

    def _sekme_birlestir_olustur(self):
        tab = self.tab_birlestir

        ctk.CTkLabel(tab, text="SIKIŞTIRMA MODU", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 4))
        self.mod_secimi = ctk.StringVar(value="akilli_vektor")

        ctk.CTkRadioButton(
            tab,
            text="Akıllı Mod (Metinleri Koru, Görselleri Sıkıştır)",
            variable=self.mod_secimi,
            value="akilli_vektor",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#334155", "#F8FAFC")
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            tab,
            text="Agresif Raster (150 DPI Taranmış Evrak Küçültme)",
            variable=self.mod_secimi,
            value="maksimum_raster",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#334155", "#F8FAFC")
        ).pack(anchor="w", pady=3)

        ctk.CTkLabel(tab, text="BATES NUMARALANDIRMA", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(14, 4))
        bates_frame = ctk.CTkFrame(
            tab,
            fg_color=("#F8FAFC", "#0F172A"),
            corner_radius=8,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        bates_frame.pack(fill="x", pady=2)

        self.bates_aktif = ctk.BooleanVar(value=False)
        self.bates_switch = ctk.CTkSwitch(
            bates_frame,
            text="Sayfa Altına Numara Bas",
            variable=self.bates_aktif,
            command=self._bates_toggle,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#334155", "#F8FAFC")
        )
        self.bates_switch.pack(side="left", padx=12, pady=10)

        self.bates_on_ek = ctk.CTkEntry(bates_frame, width=90, height=28, placeholder_text="EK-1", state="disabled")
        self.bates_on_ek.insert(0, "EK-1")
        self.bates_on_ek.pack(side="right", padx=12)

        ctk.CTkLabel(tab, text="EK LİSTESİ (FİHRİST)", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(14, 4))
        fihrist_frame = ctk.CTkFrame(
            tab,
            fg_color=("#F8FAFC", "#0F172A"),
            corner_radius=8,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        fihrist_frame.pack(fill="x", pady=2)

        self.fihrist_aktif = ctk.BooleanVar(value=bool(self.ayarlar.get("fihrist_aktif", False)))
        self.fihrist_switch = ctk.CTkSwitch(
            fihrist_frame,
            text="Ek Listesi Ekle",
            variable=self.fihrist_aktif,
            command=self._fihrist_toggle,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
            text_color=("#334155", "#F8FAFC")
        )
        self.fihrist_switch.pack(side="left", padx=12, pady=10)

        self.fihrist_baslik = ctk.CTkEntry(fihrist_frame, width=110, height=28, placeholder_text="EK LİSTESİ")
        kayitli_baslik = self.ayarlar.get("fihrist_baslik", "EK LİSTESİ")
        if kayitli_baslik == "DELİL / EK LİSTESİ":
            kayitli_baslik = "EK LİSTESİ"
        self.fihrist_baslik.insert(0, kayitli_baslik)
        self.fihrist_baslik.pack(side="right", padx=(6, 12))

        kayitli_konum = self.ayarlar.get("fihrist_konum", "Başa")
        self.fihrist_konum = ctk.CTkSegmentedButton(
            fihrist_frame,
            values=["Başa", "Sona"],
            height=28,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
            command=lambda v: (self.ayarlar.update({"fihrist_konum": v}), ayar_kaydet("fihrist_konum", v))
        )
        self.fihrist_konum.set(kayitli_konum if kayitli_konum in ("Başa", "Sona") else "Başa")
        self.fihrist_konum.pack(side="right")
        self._fihrist_toggle(kaydet=False)

        ctk.CTkLabel(tab, text="EYLEMLER", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(16, 6))

        ctk.CTkButton(
            tab,
            text="📑 Listedeki Belgeleri Birleştir (PDF Yap)",
            command=lambda: self._thread_calistir(self.tumunu_birlestir_pdf_yap),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
            height=40,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🗜 Seçili PDF'i Sıkıştır & Küçült",
            command=lambda: self._thread_calistir(self.pdf_tekil_sikistir),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="⚖ UYAP Akıllı MB Bölücüsü (Hızlı Parçala)",
            command=lambda: self._thread_calistir(self.pdf_akilli_boyut_bol),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 2: SAYFALAR
    # ========================================================
    def _sekme_duzenle_olustur(self):
        tab = self.tab_duzenle

        ctk.CTkLabel(tab, text="GÖRSEL SAYFA YÖNETİMİ", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🖼 Sayfaları Görsel Önizle & Yönet (Grid Panel)",
            command=self.pdf_gorsel_onizleme_ac,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkLabel(tab, text="HIZLI SAYFA AYIRMA VE DÖNDÜRME", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(12, 6))

        ctk.CTkButton(
            tab,
            text="🔄 Sayfa Döndür (90° / 180°)",
            command=self.pdf_dondur_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🧹 Boş Sayfaları Temizle (Akıllı Filtre)",
            command=lambda: self._thread_calistir(self.pdf_bos_sayfalari_temizle),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="✂ Sayfa Sayfa Böl (Her Sayfayı Tek Dosya Yap)",
            command=lambda: self._thread_calistir(self.pdf_sayfa_sayfa_bol),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🎯 Belirli Sayfa Aralığını Ayır",
            command=self.pdf_aralik_bol,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 3: DÖNÜŞTÜR (YENİ MODÜLLER EKLENDİ)
    # ========================================================
    def _sekme_donusturucu_olustur(self):
        tab = self.tab_donusturucu

        ctk.CTkLabel(tab, text="UYAP VE OFİS DÖNÜŞÜMLERİ", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 4))

        ctk.CTkButton(
            tab,
            text="📋 PDF ➔ UDF (.udf) (Şablon Editörüne Aktar)",
            command=self.pdf_to_udf_dialog,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=3)

        ctk.CTkButton(
            tab,
            text="📝 PDF ➔ Word (.docx)",
            command=self.pdf_to_docx_dialog,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=3)

        ctk.CTkButton(
            tab,
            text="📄 UDF (.udf) ➔ Word (.docx)",
            command=lambda: self._thread_calistir(self.udf_to_word),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=3)

        ctk.CTkButton(
            tab,
            text="📝 Word (.docx) ➔ UDF (.udf)",
            command=lambda: self._thread_calistir(self.word_to_udf),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=3)

        ctk.CTkButton(
            tab,
            text="📑 Word (.docx) ➔ PDF",
            command=lambda: self._thread_calistir(self.word_to_pdf),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=3)

        ctk.CTkButton(
            tab,
            text="🔍 Evraktaki Metinleri Görüntüle / Kopyala",
            command=self.metin_oku_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=34,
            corner_radius=8
        ).pack(fill="x", pady=(6, 3))

    # ========================================================
    # SEKME 4: ARAÇLAR (KVKK, KAŞE, FİLİGRAN, METAVERİ)
    # ========================================================
    def _sekme_guvenlik_olustur(self):
        tab = self.tab_guvenlik

        ctk.CTkLabel(tab, text="HUKUKİ İŞLEMLER VE GÜVENLİK", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🔏 e-İmza & Zaman Damgası Doğrula (Ön Kontrol)",
            command=self.eimza_dogrula_penceresi,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🛡 'Aslı Gibidir' Dijital Kaşesi Bas",
            command=self.pdf_kase_ekle_penceresi,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🕶 KVKK Akıllı Karartma (T.C., IBAN, Tel)",
            command=self.akilli_karartma_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🔏 Filigran Bas (Watermark)",
            command=self.pdf_filigran_ekle_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkButton(
            tab,
            text="🧹 Metaveri Kazı (PDF & Word Yazar/Revizyon Temizliği)",
            command=self.metaveri_kazima_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)


    def _bates_toggle(self):
        if self.bates_aktif.get():
            self.bates_on_ek.configure(state="normal")
        else:
            self.bates_on_ek.configure(state="disabled")

    def _fihrist_toggle(self, kaydet=True):
        aktif = bool(self.fihrist_aktif.get())
        self.fihrist_baslik.configure(state="normal" if aktif else "disabled")
        self.fihrist_konum.configure(state="normal" if aktif else "disabled")
        if kaydet:
            self.ayarlar["fihrist_aktif"] = aktif
            ayar_kaydet("fihrist_aktif", aktif)

    # ========================================================
    # LİSTE VE DOSYA YÖNETİMİ
    # ========================================================
    def _liste_arayuzunu_yenile(self):
        for widget in self.liste_frame.winfo_children():
            widget.destroy()

        if not self.dosyalar:
            self.bos_bilgi_etiketi = ctk.CTkLabel(
                self.liste_frame,
                text="Belgeleri buraya sürükleyin\nveya aşağıdaki butondan ekleyin\n(PDF, Word, Görsel, UDF)\n\n💡 İpucu: PDF'e çift tıklayarak sayfalarını görsel önizleyebilirsiniz.",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12),
                text_color=("#94A3B8", "#64748B"),
                justify="center"
            )
            self.bos_bilgi_etiketi.pack(pady=110)
            self.dosya_sayisi_badge.configure(text="0 dosya")
            self.secili_dosya_indeksi = None
            return

        self.secili_dosya_indeksi = getattr(self, "secili_dosya_indeksi", None)
        if self.secili_dosya_indeksi is not None and self.secili_dosya_indeksi >= len(self.dosyalar):
            self.secili_dosya_indeksi = len(self.dosyalar) - 1

        self.dosya_satir_widgetleri = []
        ai_aktif = self._ai_ozet_kullanilabilir()
        for i, yol in enumerate(self.dosyalar):
            uzanti = os.path.splitext(yol)[1].lower()
            boyut_mb = os.path.getsize(yol) / (1024 * 1024)
            simge_haritasi = {'.pdf': '📄', '.udf': '⚖', '.docx': '📝'}
            simge = simge_haritasi.get(uzanti, '🖼')
            ad = os.path.basename(yol)

            is_selected = (self.secili_dosya_indeksi == i)
            arkaplan = ("#E2E8F0", "#334155") if is_selected else ("#FFFFFF", "#1E293B")

            satir = ctk.CTkFrame(self.liste_frame, fg_color=arkaplan, corner_radius=8, height=36)
            satir.pack(fill="x", pady=2)
            satir.pack_propagate(False)

            def secim_yap(idx=i):
                self.secili_dosya_indeksi = idx
                self._liste_arayuzunu_yenile()

            def cift_tikla(yol_hedef=yol):
                if yol_hedef.lower().endswith(".pdf"):
                    PdfGorselOnizlemePenceresi(self, yol_hedef)

            satir.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))
            satir.bind("<Double-Button-1>", lambda e, y=yol: cift_tikla(y))

            sol_bilgi_box = ctk.CTkFrame(satir, fg_color="transparent")
            sol_bilgi_box.pack(side="left", padx=10, fill="y")
            sol_bilgi_box.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))
            sol_bilgi_box.bind("<Double-Button-1>", lambda e, y=yol: cift_tikla(y))

            lbl_ad = ctk.CTkLabel(
                sol_bilgi_box,
                text=f"{simge}  {ad}",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold" if is_selected else "normal"),
                text_color=("#0F172A", "#F8FAFC")
            )
            lbl_ad.pack(side="left")
            lbl_ad.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))
            lbl_ad.bind("<Double-Button-1>", lambda e, y=yol: cift_tikla(y))

            imza_bilgi = self.dosya_imza_durumlari.get(yol)
            if imza_bilgi and imza_bilgi.get("imzali"):
                rozet_metin = "✓ e-İmzalı"
                if imza_bilgi.get("zaman_damgasi"):
                    rozet_metin += " • ⏱ Z.Damgalı"
                lbl_imza = ctk.CTkLabel(
                    sol_bilgi_box,
                    text=rozet_metin,
                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=10, weight="bold"),
                    fg_color=("#DCFCE7", "#064E3B"),
                    text_color=("#166534", "#6EE7B7"),
                    corner_radius=4,
                    padx=6,
                    pady=1
                )
                lbl_imza.pack(side="left", padx=8)

            lbl_boyut = ctk.CTkLabel(
                satir,
                text=f"{boyut_mb:.2f} MB",
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
                text_color=("#64748B", "#94A3B8")
            )
            lbl_boyut.pack(side="right", padx=10)
            lbl_boyut.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))
            lbl_boyut.bind("<Double-Button-1>", lambda e, y=yol: cift_tikla(y))

            if ai_aktif and uzanti in ('.pdf', '.docx', '.udf'):
                ctk.CTkButton(
                    satir,
                    text="✨ AI Özet",
                    command=lambda y=yol: self.ai_ozet_penceresi(y),
                    width=74,
                    height=24,
                    fg_color=("#EDE9FE", "#4C1D95"),
                    hover_color=("#DDD6FE", "#5B21B6"),
                    text_color=("#5B21B6", "#EDE9FE"),
                    font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
                    corner_radius=6
                ).pack(side="right", padx=(0, 2))

            # Sağdaki öğeler (boyut, AI butonu) önce yerleşsin; uzun dosya adı onları ezmesin
            sol_bilgi_box.pack_forget()
            sol_bilgi_box.pack(side="left", padx=10, fill="y")

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
            self.dosya_imza_durumlari[dosya_yolu] = eimza_bilgisi_al(dosya_yolu)
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
                silinecek = self.dosyalar[self.secili_dosya_indeksi]
                if silinecek in self.dosya_imza_durumlari:
                    del self.dosya_imza_durumlari[silinecek]
                del self.dosyalar[self.secili_dosya_indeksi]
                self.secili_dosya_indeksi = None
                self._liste_arayuzunu_yenile()
                self.durum_etiketi.configure(text="Seçilen belge kaldırıldı.")

    def listeyi_temizle(self):
        self.dosyalar.clear()
        self.dosya_imza_durumlari.clear()
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

    def _ek_listesi_sayfasi_ekle(self, doc, ekler, baslik, sona=False):
        """
        Birleştirilmiş belgenin başına (veya sonuna) Ek Listesi tablosu ekler.
        ekler: [(dosya_yolu, baslangic_sayfa_idx, sayfa_sayisi), ...]
        Sayfa aralıkları ek listesi hariç (ilk ek sayfası = S. 1) gösterilir.
        """
        if not ekler:
            return
        icerik_sayfa_sayisi = len(doc)

        if TURKCE_FONT_YOLU:
            kalin_yol = os.path.join(os.path.dirname(TURKCE_FONT_YOLU), "arialbd.ttf")
            if not os.path.exists(kalin_yol):
                kalin_yol = os.path.join(os.path.dirname(TURKCE_FONT_YOLU), "Arial Bold.ttf")   # macOS
            if not os.path.exists(kalin_yol):
                kalin_yol = TURKCE_FONT_YOLU
            font_n = pymupdf.Font(fontfile=TURKCE_FONT_YOLU)
            font_b = pymupdf.Font(fontfile=kalin_yol)
            param_n = {"fontfile": TURKCE_FONT_YOLU, "fontname": "fihrist_n"}
            param_b = {"fontfile": kalin_yol, "fontname": "fihrist_b"}
        else:
            font_n = font_b = pymupdf.Font("helv")
            param_n = param_b = {"fontname": "helv"}

        sayfa_w, sayfa_h = 595.0, 842.0
        margin = 42.0
        tablo_ust = 124.0
        satir_h = 22.0
        alt_bosluk = 70.0
        lacivert = (0.06, 0.09, 0.16)
        gri = (0.35, 0.40, 0.47)
        cizgi = (0.80, 0.84, 0.89)

        # Sütunlar: (başlık, genişlik, hizalama)
        tablo_w = sayfa_w - 2 * margin
        sutunlar = [("Ek No", 58.0, "sol"), ("Belge Adı", 0.0, "sol"),
                    ("Sayfa Aralığı", 92.0, "orta"), ("Sayfa", 52.0, "orta")]
        sabit = sum(s[1] for s in sutunlar)
        sutunlar[1] = (sutunlar[1][0], tablo_w - sabit, "sol")

        satir_kapasitesi = int((sayfa_h - tablo_ust - satir_h - alt_bosluk) // satir_h)
        liste_sayfa_sayisi = (len(ekler) + satir_kapasitesi - 1) // satir_kapasitesi
        toplam_ek_sayfa = sum(e[2] for e in ekler)
        # Liste başa eklenirse ek sayfaları liste sayfası kadar kayar
        liste_ilk_idx = icerik_sayfa_sayisi if sona else 0
        icerik_ofset = 0 if sona else liste_sayfa_sayisi
        bugun = datetime.datetime.now().strftime("%d.%m.%Y")

        def sigdir(metin, font, fs, genislik):
            if font.text_length(metin, fontsize=fs) <= genislik:
                return metin
            while metin and font.text_length(metin + "…", fontsize=fs) > genislik:
                metin = metin[:-1]
            return metin.rstrip() + "…"

        def hucre_yaz(sayfa, x, w, y, metin, hizalama, font, param, fs, renk):
            metin = sigdir(metin, font, fs, w - 12)
            uzunluk = font.text_length(metin, fontsize=fs)
            if hizalama == "orta":
                tx = x + (w - uzunluk) / 2.0
            elif hizalama == "sag":
                tx = x + w - 6 - uzunluk
            else:
                tx = x + 6
            sayfa.insert_text(pymupdf.Point(tx, y + satir_h / 2.0 + fs * 0.36), metin, fontsize=fs, color=renk, **param)

        def ortali_yaz(sayfa, y, metin, font, param, fs, renk):
            metin = sigdir(metin, font, fs, tablo_w)
            x = (sayfa_w - font.text_length(metin, fontsize=fs)) / 2.0
            sayfa.insert_text(pymupdf.Point(x, y), metin, fontsize=fs, color=renk, **param)

        for s_idx in range(liste_sayfa_sayisi):
            sayfa = doc.new_page(pno=liste_ilk_idx + s_idx, width=sayfa_w, height=sayfa_h)

            ust_baslik = baslik if s_idx == 0 else f"{baslik} (devam)"
            ortali_yaz(sayfa, 72, ust_baslik, font_b, param_b, 16, lacivert)
            ortali_yaz(sayfa, 92, f"Düzenleme Tarihi: {bugun}   •   Toplam {len(ekler)} ek, {toplam_ek_sayfa} sayfa",
                       font_n, param_n, 9, gri)
            sayfa.draw_line(pymupdf.Point(margin, 104), pymupdf.Point(sayfa_w - margin, 104), color=cizgi, width=0.8)

            # Başlık satırı
            y = tablo_ust
            sayfa.draw_rect(pymupdf.Rect(margin, y, sayfa_w - margin, y + satir_h), color=None, fill=lacivert)
            x = margin
            for ad, w, hiz in sutunlar:
                hucre_yaz(sayfa, x, w, y, ad, hiz, font_b, param_b, 9, (1, 1, 1))
                x += w
            y += satir_h

            parca = ekler[s_idx * satir_kapasitesi:(s_idx + 1) * satir_kapasitesi]
            for j, (dosya, bas_idx, adet) in enumerate(parca):
                ek_no = s_idx * satir_kapasitesi + j + 1
                satir_rect = pymupdf.Rect(margin, y, sayfa_w - margin, y + satir_h)
                if j % 2 == 1:
                    sayfa.draw_rect(satir_rect, color=None, fill=(0.95, 0.96, 0.98))
                sayfa.draw_line(satir_rect.bl, satir_rect.br, color=cizgi, width=0.5)

                ad = os.path.splitext(os.path.basename(dosya))[0].replace("_", " ").strip()
                ilk, son = bas_idx + 1, bas_idx + adet
                aralik = f"S. {ilk}" if adet == 1 else f"S. {ilk}-{son}"

                degerler = [f"EK-{ek_no}", ad, aralik, str(adet)]
                x = margin
                for k, ((_, w, hiz), deger) in enumerate(zip(sutunlar, degerler)):
                    font, param = (font_b, param_b) if k == 0 else (font_n, param_n)
                    hucre_yaz(sayfa, x, w, y, deger, hiz, font, param, 9, lacivert)
                    x += w

                # Satıra tıklayınca ilgili ekin ilk sayfasına git
                sayfa.insert_link({
                    "kind": pymupdf.LINK_GOTO,
                    "from": satir_rect,
                    "page": icerik_ofset + bas_idx,
                    "to": pymupdf.Point(0, 0),
                })
                y += satir_h

            sayfa.draw_rect(pymupdf.Rect(margin, tablo_ust, sayfa_w - margin, y), color=cizgi, width=0.8)

            ortali_yaz(sayfa, sayfa_h - 40,
                       "Sayfa aralıkları bu liste hariç, ilk ek sayfası S. 1 kabul edilerek gösterilmiştir. Satırlara tıklayarak ilgili eke gidebilirsiniz.",
                       font_n, param_n, 7.5, gri)
            if liste_sayfa_sayisi > 1:
                ortali_yaz(sayfa, sayfa_h - 26, f"Liste Sayfası {s_idx + 1} / {liste_sayfa_sayisi}", font_n, param_n, 7.5, gri)

        # PDF okuyucularda sayfa numarası: liste i, ii... ; ekler 1, 2, 3...
        etiket_liste = {"startpage": liste_ilk_idx, "prefix": "", "style": "r", "firstpagenum": 1}
        etiket_icerik = {"startpage": icerik_ofset, "prefix": "", "style": "D", "firstpagenum": 1}
        try:
            doc.set_page_labels([etiket_icerik, etiket_liste] if sona else [etiket_liste, etiket_icerik])
        except Exception:
            pass

        # Yer imleri (Bookmarks)
        toc = []
        for n, (dosya, bas_idx, _) in enumerate(ekler, start=1):
            ad = os.path.splitext(os.path.basename(dosya))[0].replace("_", " ").strip()
            toc.append([1, f"EK-{n}: {ad}", icerik_ofset + bas_idx + 1])
        liste_yer_imi = [1, baslik, liste_ilk_idx + 1]
        doc.set_toc(toc + [liste_yer_imi] if sona else [liste_yer_imi] + toc)

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
            ekler = []

            for dosya in pdf_ve_resimler:
                baslangic = len(yeni_doc)
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

                eklenen = len(yeni_doc) - baslangic
                if eklenen > 0:
                    ekler.append((dosya, baslangic, eklenen))

            # Bates numarası yalnızca ek sayfalarına basılır; ek listesi numarasız kalır
            self._bates_numaralari_bas(yeni_doc)

            if self.fihrist_aktif.get():
                baslik = " ".join(self.fihrist_baslik.get().split()) or "EK LİSTESİ"
                self.ayarlar["fihrist_baslik"] = baslik
                ayar_kaydet("fihrist_baslik", baslik)
                self.durum_etiketi.configure(text="Ek listesi oluşturuluyor...")
                self._ek_listesi_sayfasi_ekle(yeni_doc, ekler, baslik, sona=self.fihrist_konum.get() == "Sona")

            # Gömülen Arial fontunu yalnızca kullanılan karakterlere indir (MB tasarrufu)
            try:
                yeni_doc.subset_fonts()
            except Exception:
                pass

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
    def pdf_gorsel_onizleme_ac(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Görsel Önizleme paneli yalnızca PDF dosyaları içindir.")
            return
        PdfGorselOnizlemePenceresi(self, kaynak_yol)

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

        ctk.CTkLabel(dialog, text=f"Toplam Sayfa: {toplam_sayfa}", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold")).pack(pady=(14, 6))

        derece_var = ctk.IntVar(value=90)
        radio_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        radio_bar.pack(pady=4)

        ctk.CTkRadioButton(radio_bar, text="90° Sağa", variable=derece_var, value=90).pack(side="left", padx=8)
        ctk.CTkRadioButton(radio_bar, text="90° Sola", variable=derece_var, value=270).pack(side="left", padx=8)
        ctk.CTkRadioButton(radio_bar, text="180°", variable=derece_var, value=180).pack(side="left", padx=8)

        ctk.CTkLabel(dialog, text="Döndürülecek Sayfalar (Örn: 1-3, 5 | Boşsa tümü):", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11)).pack(pady=(10, 2))
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

        ctk.CTkButton(
            dialog,
            text="Uygula ve Kaydet",
            command=dondur_kaydet,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=16)

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
    # 3. DÖNÜŞTÜR İŞLEMLERİ (PDF ➔ UDF & PDF ➔ WORD DAHİL)
    # ========================================================
    def pdf_to_udf_dialog(self):
        kaynak_yol = None
        if hasattr(self, "secili_dosya_indeksi") and self.secili_dosya_indeksi is not None and self.dosyalar:
            aday = self.dosyalar[self.secili_dosya_indeksi]
            if aday.lower().endswith('.pdf'):
                kaynak_yol = aday

        if not kaynak_yol:
            kaynak_yol = filedialog.askopenfilename(
                title="Dönüştürülecek PDF Dosyasını Seçin",
                filetypes=[("PDF Dosyaları", "*.pdf")]
            )

        if not kaynak_yol:
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("PDF ➔ UDF Dönüştürücü")
        dialog.geometry("450x330")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text="📋 UYAP Şablon Editörüne Aktarma", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=14, weight="bold")).pack(pady=(16, 4))
        ctk.CTkLabel(dialog, text=os.path.basename(kaynak_yol), font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#64748B", "#94A3B8")).pack(pady=(0, 10))

        mod_var = ctk.StringVar(value="metin")
        ctk.CTkRadioButton(dialog, text="Metin + Resimler (Düzenlenebilir Metin - Önerilen)", variable=mod_var, value="metin").pack(anchor="w", padx=24, pady=3)
        ctk.CTkRadioButton(dialog, text="Her Sayfayı Resim Olarak Ekle (Taranmış Evrak Modu)", variable=mod_var, value="sayfa").pack(anchor="w", padx=24, pady=3)

        yedek_resim_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="Resimleri yedek olarak ayrıca klasöre kaydet", variable=yedek_resim_var).pack(anchor="w", padx=24, pady=(10, 4))

        oto_ac_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="Oluşturulunca dosyayı UYAP Editörde aç", variable=oto_ac_var).pack(anchor="w", padx=24, pady=4)

        def baslat():
            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".udf",
                filetypes=[("UYAP Dokümanı", "*.udf")],
                title="UDF Olarak Kaydet",
                initialfile=os.path.splitext(os.path.basename(kaynak_yol))[0] + ".udf"
            )
            if not hedef_yol:
                return

            dialog.destroy()

            def gorev():
                self.durum_etiketi.configure(text="PDF, UDF formatına dönüştürülüyor...")
                try:
                    belge = UdfBelge()
                    resimler = []
                    ist = pdf_den_udf(kaynak_yol, belge, max_px=510, sayfa_resmi=(mod_var.get() == "sayfa"), resimler=resimler)
                    belge.kaydet(hedef_yol)

                    klasor_notu = ""
                    if yedek_resim_var.get() and resimler:
                        klasor = os.path.splitext(hedef_yol)[0] + "_resimler"
                        os.makedirs(klasor, exist_ok=True)
                        for n, (veri, uz) in enumerate(resimler, 1):
                            with open(os.path.join(klasor, f"resim_{n:03d}.{uz}"), "wb") as f:
                                f.write(veri)
                        klasor_notu = f"\n\nYedek görseller klasöre çıkarıldı:\n{klasor}"

                    self.durum_etiketi.configure(text=f"PDF ➔ UDF tamamlandı: {ist['paragraf']} paragraf, {ist['resim']} resim.")
                    mesaj = f"PDF başarıyla UDF biçimine dönüştürüldü!\n\nAktarılan Paragraf: {ist['paragraf']}\nAktarılan Resim: {ist['resim']}\nDosya: {os.path.basename(hedef_yol)}{klasor_notu}"
                    if ist.get("atlanan", 0) > 0:
                        mesaj += f"\n\nNot: {ist['atlanan']} adet vektörel görsel editörden elle eklenmelidir."

                    messagebox.showinfo("Başarılı", mesaj)

                    if oto_ac_var.get():
                        try:
                            dosyayi_ac(hedef_yol)
                        except Exception:
                            pass
                except Exception as e:
                    messagebox.showerror("Hata", f"PDF ➔ UDF dönüştürme hatası:\n{str(e)}")

            self._thread_calistir(gorev)

        ctk.CTkButton(
            dialog,
            text="📋 UDF'e Dönüştür ve Kaydet",
            command=baslat,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(pady=(16, 0), padx=24, fill="x")

    def pdf_to_docx_dialog(self):
        kaynak_yol = None
        if hasattr(self, "secili_dosya_indeksi") and self.secili_dosya_indeksi is not None and self.dosyalar:
            aday = self.dosyalar[self.secili_dosya_indeksi]
            if aday.lower().endswith('.pdf'):
                kaynak_yol = aday

        if not kaynak_yol:
            kaynak_yol = filedialog.askopenfilename(
                title="Word'e Dönüştürülecek PDF Seçin",
                filetypes=[("PDF Dosyaları", "*.pdf")]
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

        def gorev():
            self.durum_etiketi.configure(text="PDF Word belgesine aktarılıyor...")
            try:
                ist = pdf_den_docx(kaynak_yol, hedef_yol, sayfa_resmi=False, sayfa_sonlari=True)
                self.durum_etiketi.configure(text=f"PDF ➔ Word tamamlandı: {ist['paragraf']} paragraf.")
                messagebox.showinfo("Başarılı", f"PDF başarıyla Word formatına dönüştürüldü!\n\nAktarılan Paragraf: {ist['paragraf']}\nAktarılan Görsel: {ist['resim']}\nDosya: {os.path.basename(hedef_yol)}")
                try:
                    dosyayi_ac(hedef_yol)
                except Exception:
                    pass
            except Exception as e:
                messagebox.showerror("Hata", f"PDF ➔ Word dönüştürme hatası:\n{str(e)}")

        self._thread_calistir(gorev)

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

    def word_to_udf(self):
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
            defaultextension=".udf",
            filetypes=[("UYAP Dokümanı", "*.udf")],
            title="UDF (.udf) Olarak Kaydet",
            initialfile=os.path.splitext(os.path.basename(kaynak_yol))[0] + ".udf"
        )
        if not hedef_yol:
            return

        try:
            self.durum_etiketi.configure(text="Word belgesi UDF formatına dönüştürülüyor...")
            belge, ist, resimler = docx_to_udf_cevir(kaynak_yol, max_px=510)
            belge.kaydet(hedef_yol)

            self.durum_etiketi.configure(text=f"Word ➔ UDF tamamlandı: {ist['paragraf']} paragraf, {ist['resim']} resim.")
            mesaj = f"Word belgesi UDF formatına dönüştürüldü!\n\nAktarılan Paragraf: {ist['paragraf']}\nAktarılan Resim: {ist['resim']}\nDosya: {os.path.basename(hedef_yol)}"
            if ist.get("atlanan", 0) > 0:
                mesaj += f"\n\nNot: {ist['atlanan']} adet vektörel/desteklenmeyen görsel editörden elle eklenmelidir."
            messagebox.showinfo("Başarılı", mesaj)
        except Exception as e:
            messagebox.showerror("Hata", f"Word ➔ UDF dönüştürme hatası:\n{str(e)}")

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

        if pythoncom is not None:
            pythoncom.CoInitialize()

        try:
            self.durum_etiketi.configure(text="Word belgesi PDF'e dönüştürülüyor...")
            abs_kaynak = os.path.abspath(kaynak_yol)
            abs_hedef = os.path.abspath(hedef_yol)

            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(abs_kaynak)
                doc.SaveAs(abs_hedef, FileFormat=17)
                doc.Close()
                word.Quit()
            except Exception:
                if docx_to_pdf_convert is not None:
                    eski_stderr = sys.stderr
                    if sys.stderr is None:
                        sys.stderr = io.StringIO()
                    try:
                        docx_to_pdf_convert(abs_kaynak, abs_hedef)
                    except SystemExit:
                        # docx2pdf macOS'ta hata olunca sys.exit(1) çağırıyor; uygulama kapanmasın
                        raise RuntimeError("Microsoft Word belgeyi PDF'e dönüştüremedi.")
                    finally:
                        sys.stderr = eski_stderr
                    if IS_MAC and not os.path.exists(abs_hedef):
                        raise RuntimeError("Microsoft Word PDF dosyasını oluşturmadı.")
                else:
                    raise RuntimeError("Microsoft Word veya docx2pdf motoruna erişilemedi.")

            self.durum_etiketi.configure(text="Word ➔ PDF dönüşümü tamamlandı.")
            messagebox.showinfo("Başarılı", f"Word belgesi başarıyla PDF'e dönüştürüldü!\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror(
                "Dönüşüm Hatası",
                f"Word ➔ PDF dönüştürme hatası:\n{str(e)}\n\n"
                + ("Not: Mac'inizde Microsoft Word yüklü olmalı ve ilk kullanımda çıkan \"Word'ü denetleme\" "
                   "izni verilmelidir (Sistem Ayarları › Gizlilik ve Güvenlik › Otomasyon).\n\n"
                   "Word \"Dosyaya Erişim İzni Ver\" penceresi açarsa \"Seçin...\" ile belgenin bulunduğu "
                   "klasörü seçip \"Erişim İzni Ver\"e basın ve dönüştürmeyi tekrar deneyin."
                   if IS_MAC else
                   "Not: Bilgisayarınızda Microsoft Word yüklü olmalı ve arka planda açık olan Word pencereleri kapatılmalıdır.")
            )
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

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

            txt = ctk.CTkTextbox(pencere, font=ctk.CTkFont(family=SABIT_GENISLIK_FONTU, size=12))
            txt.pack(fill="both", expand=True, padx=16, pady=(16, 10))
            txt.insert("0.0", metin_icerik)

            def kopyala():
                pencere.clipboard_clear()
                pencere.clipboard_append(txt.get("0.0", "end"))
                messagebox.showinfo("Kopyalandı", "Metin panoya kopyalandı!")

            ctk.CTkButton(
                pencere,
                text="📋 Panoya Kopyala",
                command=kopyala,
                fg_color=("#0F172A", "#3B82F6"),
                hover_color=("#1E293B", "#2563EB"),
                text_color="#FFFFFF",
                corner_radius=8
            ).pack(pady=(0, 16))
            self.durum_etiketi.configure(text="Metin hazır.")
        except Exception as e:
            messagebox.showerror("Hata", f"Metin okunamadı:\n{str(e)}")

    # ========================================================
    # 4. ARAÇLAR: KVKK & KAŞE & GÜVENLİK İŞLEMLERİ
    # ========================================================
    def eimza_dogrula_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol:
            return

        uzanti = os.path.splitext(kaynak_yol)[1].lower()
        if uzanti not in ('.udf', '.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "e-İmza kontrolü yalnızca UDF ve PDF dosyalarında uygulanabilir.")
            return

        bilgi = eimza_bilgisi_al(kaynak_yol)

        dialog = ctk.CTkToplevel(self)
        dialog.title("e-İmza & Güvenlik Analizi")
        dialog.geometry("480x380")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(
            dialog,
            text="🔏 Belge Güvenlik ve İmza Doğrulama",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=16, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(pady=(20, 2))

        ctk.CTkLabel(
            dialog,
            text=os.path.basename(kaynak_yol),
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#64748B", "#94A3B8")
        ).pack(pady=(0, 14))

        kart = ctk.CTkFrame(
            dialog,
            fg_color=("#F8FAFC", "#1E293B"),
            corner_radius=10,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        kart.pack(fill="x", padx=24, pady=4)

        durum_baslik = "✅ Nitelikli Elektronik İmza Bulundu" if bilgi["imzali"] else "❌ İmzalanmamış Belge (İmza Yok)"
        durum_renk = ("#166534", "#4ADE80") if bilgi["imzali"] else ("#991B1B", "#F87171")

        ctk.CTkLabel(
            kart,
            text=durum_baslik,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold"),
            text_color=durum_renk
        ).pack(anchor="w", padx=16, pady=(14, 6))

        def satir_ekle(etiket, deger):
            f = ctk.CTkFrame(kart, fg_color="transparent")
            f.pack(fill="x", padx=16, pady=3)
            ctk.CTkLabel(f, text=etiket, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"), text_color=("#475569", "#CBD5E1"), width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(f, text=deger or "-", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#0F172A", "#F8FAFC"),
                         anchor="w", justify="left", wraplength=300).pack(side="left", fill="x", expand=True)

        def kisi_metni(i):
            metin = i["ad"]
            if i.get("unvan"):
                metin += f" ({i['unvan']})"
            tc = re.sub(r"\D", "", i.get("tc") or "")
            if len(tc) == 11:   # Sertifikadaki seri numarası T.C. kimlik no'dur; KVKK gereği maskelenir
                metin += f"  •  T.C. {tc[:3]}******{tc[-2:]}"
            return metin

        imzacilar = bilgi.get("imzalayanlar") or []
        if imzacilar:
            for n, i in enumerate(imzacilar):
                satir_ekle("İmzalayan Kişi:" if len(imzacilar) == 1 else f"{n + 1}. İmzalayan:", kisi_metni(i))
        else:
            satir_ekle("İmzalayan Kişi:", bilgi["imzalayan"])
        if bilgi.get("saglayici"):
            satir_ekle("Sertifika Sağlayıcı:", bilgi["saglayici"])
        satir_ekle("İmza Formatı:", "CAdES / UYAP UDF" if uzanti == '.udf' else "PAdES / PDF")
        if bilgi.get("tarih"):
            satir_ekle("İmza Zamanı:", bilgi["tarih"])
        if bilgi["zaman_damgasi"]:
            damga = "Mevcut (TSA Onaylı)" + (f" — {bilgi['zaman_damgasi_saati']}" if bilgi.get("zaman_damgasi_saati") else "")
        else:
            damga = "Bulunmuyor (Yerel Saat)"
        satir_ekle("Zaman Damgası:", damga)
        ek_satir = max(0, len(imzacilar) - 1) + (1 if bilgi.get("saglayici") else 0)
        dialog.geometry(f"480x{380 + 26 * ek_satir}")

        ctk.CTkButton(
            dialog,
            text="Kapat",
            command=dialog.destroy,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"),
            height=34,
            corner_radius=8
        ).pack(pady=(20, 0), padx=24, fill="x")

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

        ctk.CTkLabel(dialog, text="Kaşe Başlığı:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(12, 2))
        baslik_entry = ctk.CTkEntry(dialog)
        baslik_entry.insert(0, self.ayarlar.get("kase_baslik", "ASLI GİBİDİR"))
        baslik_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Avukat İsim / Sicil No / Unvan:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        avukat_entry = ctk.CTkEntry(dialog, placeholder_text="Av. Ad Soyad - Baro Sicil No")
        kayitli_avukat = self.ayarlar.get("avukat_bilgisi", "")
        if kayitli_avukat:
            avukat_entry.insert(0, kayitli_avukat)
        avukat_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Tarih / Ek Açıklama:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        ek_entry = ctk.CTkEntry(dialog)
        ek_entry.insert(0, f"Tarih: {bugun}")
        ek_entry.pack(fill="x", padx=20)

        kayitli_konum_gorunum = self.ayarlar.get("kase_konum", "Sağ Alt")
        konum_haritasi = {"Sağ Alt": "sag_alt", "Sol Alt": "sol_alt", "Alt Orta": "alt_orta", "Sağ Üst": "sag_ust"}
        konum_var = ctk.StringVar(value=konum_haritasi.get(kayitli_konum_gorunum, "sag_alt"))

        ctk.CTkLabel(dialog, text="Konum:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        konum_opt = ctk.CTkOptionMenu(
            dialog,
            values=["Sağ Alt", "Sol Alt", "Alt Orta", "Sağ Üst"],
            command=lambda v: (konum_var.set(konum_haritasi.get(v, "sag_alt")), self.ayarlar.update({"kase_konum": v}), ayar_kaydet("kase_konum", v))
        )
        konum_opt.set(kayitli_konum_gorunum)
        konum_opt.pack(fill="x", padx=20)

        kayitli_renk = self.ayarlar.get("kase_renk", "kirmizi")
        renk_var = ctk.StringVar(value=kayitli_renk)
        renk_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        renk_frame.pack(fill="x", padx=20, pady=8)

        def renk_degisti():
            secilen = renk_var.get()
            self.ayarlar["kase_renk"] = secilen
            ayar_kaydet("kase_renk", secilen)

        ctk.CTkRadioButton(renk_frame, text="Kırmızı", variable=renk_var, value="kirmizi", command=renk_degisti).pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(renk_frame, text="Lacivert", variable=renk_var, value="lacivert", command=renk_degisti).pack(side="left")

        aralik_entry = ctk.CTkEntry(dialog, placeholder_text="Sayfa No (Boşsa Tümü | Örn: 1, 3-5)")
        aralik_entry.pack(fill="x", padx=20, pady=8)

        def uygula():
            try:
                secilenler = sayfa_araligi_ayristir(aralik_entry.get().strip(), toplam_sayfa)
            except ValueError as ve:
                messagebox.showerror("Hata", str(ve))
                return

            yeni_avukat = avukat_entry.get().strip()
            yeni_baslik = baslik_entry.get().strip()
            if yeni_avukat:
                self.ayarlar["avukat_bilgisi"] = yeni_avukat
                ayar_kaydet("avukat_bilgisi", yeni_avukat)
            if yeni_baslik:
                self.ayarlar["kase_baslik"] = yeni_baslik
                ayar_kaydet("kase_baslik", yeni_baslik)

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

        ctk.CTkButton(
            dialog,
            text="🛡 Kaşeyi Bas ve Kaydet",
            command=uygula,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=16)

    def pdf_filigran_ekle_penceresi(self):
        kaynak_yol = self._secili_dosya_al()
        if not kaynak_yol or not kaynak_yol.lower().endswith('.pdf'):
            messagebox.showwarning("Geçersiz Dosya", "Filigran Basma yalnızca PDF dosyaları içindir.")
            return

        with pymupdf.open(kaynak_yol) as doc_temp:
            toplam_sayfa = len(doc_temp)

        renk_haritasi = {
            "Gri": (0.45, 0.45, 0.45),
            "Kırmızı": (0.80, 0.05, 0.05),
            "Lacivert": (0.0, 0.18, 0.38),
            "Siyah": (0.0, 0.0, 0.0),
        }

        dialog = ctk.CTkToplevel(self)
        dialog.title("Filigran Bas (Watermark)")
        dialog.geometry("460x560")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text="Hazır Filigranlar:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(12, 4))

        metin_entry = ctk.CTkEntry(dialog, placeholder_text="Filigran metnini yazın (Örn: TASLAKTIR)")

        def hazir_sec(metin):
            metin_entry.delete(0, "end")
            metin_entry.insert(0, metin)

        hazir_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        hazir_frame.pack(fill="x", padx=20)
        hazir_frame.grid_columnconfigure((0, 1), weight=1)
        hazir_metinler = ["GİZLİDİR", "TASLAKTIR", "DELİL NİTELİĞİNDEDİR"]
        for i, etiket in enumerate(hazir_metinler):
            ctk.CTkButton(
                hazir_frame,
                text=etiket,
                command=lambda m=etiket: hazir_sec(m),
                fg_color=("#F1F5F9", "#334155"),
                hover_color=("#E2E8F0", "#475569"),
                text_color=("#1E293B", "#F8FAFC"),
                font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"),
                height=30,
                corner_radius=6
            ).grid(row=i // 2, column=i % 2, columnspan=2 if i == len(hazir_metinler) - 1 and i % 2 == 0 else 1, sticky="ew", padx=3, pady=3)

        ctk.CTkLabel(dialog, text="Filigran Metni (elle düzenlenebilir):", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(10, 2))
        kayitli_metin = self.ayarlar.get("filigran_metin", "TASLAKTIR")
        if kayitli_metin:
            metin_entry.insert(0, kayitli_metin)
        metin_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Renk:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(10, 2))
        renk_opt = ctk.CTkOptionMenu(dialog, values=list(renk_haritasi.keys()))
        kayitli_renk = self.ayarlar.get("filigran_renk", "Gri")
        renk_opt.set(kayitli_renk if kayitli_renk in renk_haritasi else "Gri")
        renk_opt.pack(fill="x", padx=20)

        try:
            kayitli_opaklik = float(self.ayarlar.get("filigran_opaklik", 0.25))
        except (TypeError, ValueError):
            kayitli_opaklik = 0.25
        kayitli_opaklik = min(0.8, max(0.05, kayitli_opaklik))

        opaklik_etiket = ctk.CTkLabel(dialog, text="", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold"))
        opaklik_etiket.pack(anchor="w", padx=20, pady=(10, 2))

        def opaklik_guncelle(v):
            opaklik_etiket.configure(text=f"Saydamlık (Opaklık): %{int(round(float(v) * 100))}")

        opaklik_slider = ctk.CTkSlider(dialog, from_=0.05, to=0.8, number_of_steps=15, command=opaklik_guncelle)
        opaklik_slider.set(kayitli_opaklik)
        opaklik_slider.pack(fill="x", padx=20)
        opaklik_guncelle(kayitli_opaklik)

        arkada_var = ctk.BooleanVar(value=bool(self.ayarlar.get("filigran_arkada", True)))
        ctk.CTkCheckBox(
            dialog,
            text="Metnin arkasına yerleştir (taranmış sayfalarda otomatik öne alınır)",
            variable=arkada_var,
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11)
        ).pack(anchor="w", padx=20, pady=(12, 4))

        aralik_entry = ctk.CTkEntry(dialog, placeholder_text="Sayfa No (Boşsa Tümü | Örn: 1, 3-5)")
        aralik_entry.pack(fill="x", padx=20, pady=8)

        def uygula():
            metin = " ".join(metin_entry.get().split())
            if not metin:
                messagebox.showwarning("Eksik Bilgi", "Lütfen filigran metnini girin veya hazır bir seçenek seçin.", parent=dialog)
                return

            try:
                secilenler = sayfa_araligi_ayristir(aralik_entry.get().strip(), toplam_sayfa)
            except ValueError as ve:
                messagebox.showerror("Hata", str(ve), parent=dialog)
                return

            renk_adi = renk_opt.get()
            opaklik = round(float(opaklik_slider.get()), 2)
            arkada = bool(arkada_var.get())
            for anahtar, deger in (("filigran_metin", metin), ("filigran_renk", renk_adi),
                                   ("filigran_opaklik", opaklik), ("filigran_arkada", arkada)):
                self.ayarlar[anahtar] = deger
                ayar_kaydet(anahtar, deger)

            hedef_yol = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Dosyası", "*.pdf")],
                title="Filigranlı PDF'i Kaydet",
                initialfile="filigranli_" + os.path.basename(kaynak_yol)
            )
            if not hedef_yol or not self._ayni_dosya_kontrolu(kaynak_yol, hedef_yol):
                return

            try:
                renk = renk_haritasi.get(renk_adi, renk_haritasi["Gri"])
                if TURKCE_FONT_YOLU:
                    font_obj = pymupdf.Font(fontfile=TURKCE_FONT_YOLU)
                    font_param = {"fontfile": TURKCE_FONT_YOLU, "fontname": "filigran_tr"}
                else:
                    font_obj = pymupdf.Font("helv")
                    font_param = {"fontname": "helv"}
                birim_uzunluk = font_obj.text_length(metin, fontsize=1) or 1.0
                aci = 45

                with pymupdf.open(kaynak_yol) as doc:
                    for idx in secilenler:
                        sayfa = doc[idx]
                        # Görünen (döndürülmüş) sayfa ölçüleri
                        w, h = sayfa.rect.width, sayfa.rect.height

                        # 45° çaprazda sayfaya sığacak en büyük yazı boyutu
                        hedef_uzunluk = min(w, h) * 1.41421356 * 0.80
                        f_size = max(10.0, min(120.0, hedef_uzunluk / birim_uzunluk))
                        uzunluk = birim_uzunluk * f_size

                        # Sayfa döndürülmüşse merkez ve açıyı ham koordinatlara çevir
                        merkez = pymupdf.Point(w / 2.0, h / 2.0) * sayfa.derotation_matrix
                        baslangic = pymupdf.Point(merkez.x - uzunluk / 2.0, merkez.y + f_size * 0.35)

                        # Taranmış (metinsiz) sayfada arkaya basılırsa görsel tarafından örtülür
                        sayfa_arkada = arkada and bool(sayfa.get_text().strip())

                        sayfa.insert_text(
                            baslangic,
                            metin,
                            fontsize=f_size,
                            color=renk,
                            fill_opacity=opaklik,
                            stroke_opacity=opaklik,
                            morph=(merkez, pymupdf.Matrix(aci + sayfa.rotation)),
                            overlay=not sayfa_arkada,
                            **font_param
                        )

                    doc.save(hedef_yol, garbage=4, deflate=True)

                dialog.destroy()
                self.durum_etiketi.configure(text=f"Filigran seçilen {len(secilenler)} sayfaya uygulandı.")
                messagebox.showinfo("Başarılı", f"Filigran basıldı!\nDosya: {os.path.basename(hedef_yol)}")
            except Exception as e:
                messagebox.showerror("Hata", f"Filigran basma hatası:\n{str(e)}")

        ctk.CTkButton(
            dialog,
            text="🔏 Filigranı Bas ve Kaydet",
            command=uygula,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=14)

    def metaveri_kazima_penceresi(self):
        uygun_uzantilar = ('.pdf', '.docx')
        tum_uygunlar = [d for d in self.dosyalar if d.lower().endswith(uygun_uzantilar)]
        if not tum_uygunlar:
            messagebox.showwarning("Belge Yok", "Metaveri kazıma için listeye PDF veya Word (.docx) belgesi ekleyin.")
            return

        secili = None
        if getattr(self, "secili_dosya_indeksi", None) is not None and self.secili_dosya_indeksi < len(self.dosyalar):
            aday = self.dosyalar[self.secili_dosya_indeksi]
            if aday.lower().endswith(uygun_uzantilar):
                secili = aday

        dialog = ctk.CTkToplevel(self)
        dialog.title("Metaveri Kazıma (Metadata Sanitizer)")
        dialog.geometry("480x430")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(dialog, text="🧹 Hassas Arka Plan Bilgilerini Temizle", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=15, weight="bold")).pack(anchor="w", padx=20, pady=(16, 4))
        ctk.CTkLabel(
            dialog,
            text="Yazar adı, son düzenleyen kullanıcı/bilgisayar, şirket, oluşturma tarihleri,\n"
                 "düzenleme süresi, şablon yolu, revizyon geçmişi ve değişiklik/yorum\n"
                 "yazarları silinir. Belgenin içeriği değişmez.",
            font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11),
            text_color=("#475569", "#CBD5E1"),
            justify="left"
        ).pack(anchor="w", padx=20)

        ctk.CTkLabel(dialog, text="Kapsam:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(14, 2))
        kapsam_var = ctk.StringVar(value="secili" if secili else "tumu")
        secili_radio = ctk.CTkRadioButton(
            dialog,
            text=f"Yalnızca seçili belge ({os.path.basename(secili)})" if secili else "Yalnızca seçili belge (PDF/Word seçili değil)",
            variable=kapsam_var, value="secili"
        )
        secili_radio.pack(anchor="w", padx=20, pady=3)
        if not secili:
            secili_radio.configure(state="disabled")
        ctk.CTkRadioButton(
            dialog,
            text=f"Listedeki tüm PDF ve Word belgeleri ({len(tum_uygunlar)} adet)",
            variable=kapsam_var, value="tumu"
        ).pack(anchor="w", padx=20, pady=3)

        ctk.CTkLabel(dialog, text="Kayıt:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11, weight="bold")).pack(anchor="w", padx=20, pady=(12, 2))
        kayit_var = ctk.StringVar(value="kopya")
        ctk.CTkRadioButton(dialog, text="Temiz kopya oluştur (temiz_ önekiyle, klasör seçilir)", variable=kayit_var, value="kopya").pack(anchor="w", padx=20, pady=3)
        ctk.CTkRadioButton(dialog, text="Orijinal dosyaların üzerine yaz", variable=kayit_var, value="uzerine").pack(anchor="w", padx=20, pady=3)

        def uygula():
            hedefler = [secili] if kapsam_var.get() == "secili" and secili else list(tum_uygunlar)

            imzalilar = [os.path.basename(d) for d in hedefler
                         if d.lower().endswith('.pdf') and eimza_bilgisi_al(d).get("imzali")]
            if imzalilar and not messagebox.askyesno(
                "e-İmzalı Belge",
                "Şu belgeler elektronik imzalı görünüyor:\n\n• " + "\n• ".join(imzalilar[:8]) +
                "\n\nMetaveri kazıma bu belgelerdeki imzayı geçersiz kılar. Devam edilsin mi?",
                parent=dialog
            ):
                return

            if kayit_var.get() == "kopya":
                klasor = filedialog.askdirectory(
                    title="Temiz kopyaların kaydedileceği klasörü seçin",
                    initialdir=os.path.dirname(hedefler[0]),
                    parent=dialog
                )
                if not klasor:
                    return
            else:
                klasor = None
                if not messagebox.askyesno(
                    "Onay",
                    f"{len(hedefler)} belgenin orijinali temizlenmiş haliyle değiştirilecek.\nBu işlem geri alınamaz. Devam edilsin mi?",
                    parent=dialog
                ):
                    return

            rapor, basarili, hatali = [], 0, 0
            for i, kaynak in enumerate(hedefler, start=1):
                ad = os.path.basename(kaynak)
                self.durum_etiketi.configure(text=f"Metaveri kazınıyor ({i}/{len(hedefler)}): {ad}")
                self.update_idletasks()

                hedef = os.path.join(klasor, "temiz_" + ad) if klasor else kaynak + ".kazima_tmp"
                try:
                    if os.path.abspath(hedef) == os.path.abspath(kaynak):
                        raise ValueError("Hedef dosya kaynakla aynı.")
                    if kaynak.lower().endswith('.pdf'):
                        bulunanlar = pdf_metaveri_kazi(kaynak, hedef)
                    else:
                        bulunanlar = docx_metaveri_kazi(kaynak, hedef)
                    if not klasor:
                        os.replace(hedef, kaynak)
                    basarili += 1
                    rapor.append(f"✅ {ad}")
                    rapor += [f"     • {b}" for b in bulunanlar] or ["     • Hassas metaveri bulunamadı (yine de sıfırlandı)."]
                except Exception as e:
                    hatali += 1
                    rapor.append(f"❌ {ad}: {e}")
                    if not klasor and os.path.exists(hedef):
                        try:
                            os.remove(hedef)
                        except OSError:
                            pass

            dialog.destroy()
            self.durum_etiketi.configure(text=f"Metaveri kazıma tamamlandı: {basarili} başarılı, {hatali} hatalı.")

            sonuc = ctk.CTkToplevel(self)
            sonuc.title("Metaveri Kazıma Raporu")
            sonuc.geometry("600x440")
            sonuc.after(100, sonuc.lift)
            ozet = f"{basarili} belge temizlendi" + (f", {hatali} belgede hata oluştu" if hatali else "") + "."
            if klasor:
                ozet += f"\nKayıt klasörü: {klasor}"
            ctk.CTkLabel(sonuc, text=ozet, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"), justify="left").pack(anchor="w", padx=16, pady=(14, 6))
            ctk.CTkLabel(sonuc, text="Silinen bilgiler:", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#64748B", "#94A3B8")).pack(anchor="w", padx=16)
            kutu = ctk.CTkTextbox(sonuc, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12))
            kutu.pack(fill="both", expand=True, padx=16, pady=(4, 8))
            kutu.insert("0.0", "\n".join(rapor))
            kutu.configure(state="disabled")
            ctk.CTkButton(sonuc, text="Kapat", command=sonuc.destroy, fg_color=("#0F172A", "#3B82F6"),
                          hover_color=("#1E293B", "#2563EB"), text_color="#FFFFFF", corner_radius=8).pack(pady=(0, 14))

        ctk.CTkButton(
            dialog,
            text="🧹 Metaverileri Kazı",
            command=uygula,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=18)

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

        ctk.CTkLabel(dialog, text="Otomatik Karartma Seçenekleri", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=13, weight="bold")).pack(anchor="w", padx=20, pady=(14, 6))

        var_tc = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="Resmi T.C. Kimlik Algoritmasıyla Doğrulanan Numaralar", variable=var_tc).pack(anchor="w", padx=20, pady=4)

        var_iban = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dialog, text="IBAN Numaraları (TR...)", variable=var_iban).pack(anchor="w", padx=20, pady=4)

        var_tel = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(dialog, text="Telefon Numaraları (05xx / +90)", variable=var_tel).pack(anchor="w", padx=20, pady=4)

        ctk.CTkLabel(dialog, text="Özel Maskelenecek Kelimeler (Virgülle ayırın):", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11)).pack(anchor="w", padx=20, pady=(12, 2))
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

        ctk.CTkButton(
            dialog,
            text="🛡 Kalıcı Karart ve Kaydet",
            command=uygula,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=18)

    # Kullanım rehberi içeriği: (simge, bölüm başlığı, kısa açıklama, [(madde adı, açıklama), ...])
    REHBER_BOLUMLERI = [
        ("📑", "Birleştir", "Belgeleri tek ve UYAP'a uygun bir PDF'te toplayın", [
            ("Birleştir (PDF Yap)", "Belgeleri sol alana sürükleyin, sıralarını ▲ ▼ ile ayarlayın ve "
                                    "'Listedeki Belgeleri Birleştir' butonuna basın."),
            ("Görseller", "Dekont, tutanak gibi fotoğraflar yönü bozulmadan A4 sayfaya yerleştirilir."),
            ("Sayfa numarası", "'Sayfa Altına Numara Bas' açıksa her sayfanın altına "
                               "'EK-1 | Sayfa 1/5' biçiminde numara eklenir."),
            ("Ek Listesi", "Birleştirilen belgenin başına veya sonuna her ekin adını ve sayfa aralığını "
                           "(örn. EK-1: Fatura S. 1-3) gösteren bir liste eklenir. Satıra tıklayınca o eke gidilir."),
            ("Sıkıştır & Küçült", "Akıllı mod metinleri koruyarak, Agresif mod taranmış evrakları "
                                  "daha çok küçültür."),
            ("UYAP MB Bölücü", "Büyük dosyaları UYAP sınırına (9,5 MB) göre, sayfaları bölmeden parçalara ayırır."),
        ]),
        ("🖼", "Sayfalar", "PDF sayfalarını görün, düzenleyin ve ayırın", [
            ("Görsel Önizleme", "Sayfaları küçük resimler hâlinde görün; sürükleyerek sıralayın, "
                                "döndürün veya silin. Listedeki PDF'e çift tıklamak da bu paneli açar."),
            ("Sayfa Döndür", "Seçtiğiniz sayfaları 90° veya 180° döndürür."),
            ("Boş Sayfaları Temizle", "Çift taraflı taramalardaki boş sayfaları bulur, onayınızla siler."),
            ("Sayfa Böl / Aralık Ayır", "Her sayfayı ayrı dosya yapar ya da istediğiniz sayfa aralığını "
                                        "(örn. 3-7) ayrı bir PDF olarak kaydeder."),
        ]),
        ("🔄", "Dönüştür", "UDF, Word ve PDF arasında geçiş yapın", [
            ("PDF ➔ UDF", "PDF'i metin, tablo ve görselleriyle UYAP Şablon Editörü'nde açılabilen UDF'e çevirir."),
            ("PDF ➔ Word", "Sütun yapısını, tabloları ve başlıkları koruyarak düzenlenebilir Word belgesi üretir."),
            ("UDF ➔ Word / Word ➔ UDF", "UYAP evraklarını Word'de düzenleyin, dilekçelerinizi resim ve "
                                        "tablolarıyla UDF'e taşıyın."),
            ("Word ➔ PDF", "Word belgesini PDF'e çevirir (bilgisayarda Microsoft Word kurulu olmalıdır)."),
            ("Metinleri Görüntüle", "Evraktaki metni okuyup tek tıkla panoya kopyalayın."),
        ]),
        ("🛡", "Araçlar", "İmza kontrolü, kaşe, karartma ve gizlilik", [
            ("e-İmza Doğrula", "UDF veya PDF'i imzalayan kişilerin adını, unvanını, sertifika sağlayıcısını "
                               "ve zaman damgası saatini gösterir."),
            ("'Aslı Gibidir' Kaşesi", "Kaşe, sayfanın yönüne göre seçtiğiniz köşeye yerleşir; tercihleriniz hatırlanır."),
            ("KVKK Karartma", "Doğrulanmış T.C. kimlik numaralarını, IBAN'ları ve telefonları kalıcı olarak "
                              "karartır; belgenin gizli bilgilerini de temizler."),
            ("Filigran Bas", "Sayfalara 45° açıyla yarı saydam 'GİZLİDİR', 'TASLAKTIR' gibi bir yazı ekler."),
            ("Metaveri Kazı", "PDF ve Word belgelerindeki yazar adı, bilgisayar adı ve revizyon geçmişi "
                              "gibi gizli bilgileri siler. Belgenin içeriği değişmez."),
        ]),
        ("✨", "AI Evrak Özeti", "Evrakları bilgisayarınızda, internete göndermeden özetleyin", [
            ("Açmak için", "Ayarlar (⚙) > Yapay Zeka'dan özelliği açın. Model (3,1 GB) yalnızca bir kez indirilir."),
            ("Kullanım", "Listedeki PDF, Word ve UDF evraklarının yanındaki '✨ AI Özet' butonuna basın."),
            ("Belge türü", "Sözleşme, dilekçe, karar, ihtarname gibi türler otomatik algılanır ve özet ona göre "
                           "yazılır. Yanlışsa özet penceresinden türü değiştirip 'Yeniden Özetle'ye basın."),
            ("Hız", "Ayarlar'daki 'İşlem birimi' Otomatik'te bırakılırsa uygun bir ekran kartı varsa "
                    "o kullanılır (3-5 kat daha hızlı), yoksa işlemci."),
            ("Unutmayın", "AI özeti hata içerebilir; hukuki işlemlerde daima asıl belgeyi esas alın."),
        ]),
    ]

    def rehber_penceresi_ac(self):
        if getattr(self, "_rehber_penceresi", None) is not None and self._rehber_penceresi.winfo_exists():
            self._rehber_penceresi.lift()
            self._rehber_penceresi.focus()
            return

        pencere = ctk.CTkToplevel(self)
        pencere.title("Nasıl Kullanılır?")
        pencere.geometry("700x640")
        pencere.minsize(560, 420)
        pencere.configure(fg_color=("#F1F5F9", "#0F172A"))
        pencere.after(100, pencere.lift)
        self._rehber_penceresi = pencere

        ust = ctk.CTkFrame(pencere, fg_color="transparent")
        ust.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkLabel(ust, text="Nasıl Kullanılır?", font=ctk.CTkFont(family=ARAYUZ_FONTU, size=20, weight="bold"),
                     text_color=("#0F172A", "#F8FAFC")).pack(anchor="w")
        ctk.CTkLabel(ust, text="UYAP Evrak Asistanı'nın özellikleri",
                     font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12), text_color=("#64748B", "#94A3B8")).pack(anchor="w")

        govde = ctk.CTkScrollableFrame(pencere, fg_color="transparent")
        govde.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        aciklama_etiketleri = []
        for no, (simge, baslik, alt_baslik, maddeler) in enumerate(self.REHBER_BOLUMLERI, start=1):
            kart = ctk.CTkFrame(govde, fg_color=("#FFFFFF", "#1E293B"), corner_radius=12,
                                border_width=1, border_color=("#E2E8F0", "#334155"))
            kart.pack(fill="x", padx=8, pady=6)

            bas = ctk.CTkFrame(kart, fg_color="transparent")
            bas.pack(fill="x", padx=18, pady=(14, 2))
            ctk.CTkLabel(bas, text=str(no), width=26, height=26, corner_radius=13,
                         fg_color=("#0F172A", "#3B82F6"), text_color="#FFFFFF",
                         font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold")).pack(side="left")
            ctk.CTkLabel(bas, text=baslik, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=16, weight="bold"),
                         text_color=("#0F172A", "#F8FAFC")).pack(side="left", padx=(10, 0))
            ctk.CTkLabel(kart, text=alt_baslik, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12), height=18,
                         text_color=("#64748B", "#94A3B8")).pack(anchor="w", padx=(54, 18), pady=(0, 10))

            for m, (ad, aciklama) in enumerate(maddeler):
                satir = ctk.CTkFrame(kart, fg_color="transparent")
                satir.pack(fill="x", padx=(54, 18), pady=(0, 16 if m == len(maddeler) - 1 else 10))
                ctk.CTkLabel(satir, text=ad, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12, weight="bold"), height=18,
                             text_color=("#1E293B", "#E2E8F0"), anchor="w").pack(anchor="w")
                lbl = ctk.CTkLabel(satir, text=aciklama, font=ctk.CTkFont(family=ARAYUZ_FONTU, size=12), height=18,
                                   text_color=("#475569", "#94A3B8"), anchor="w", justify="left", wraplength=500)
                lbl.pack(anchor="w", pady=(1, 0))
                aciklama_etiketleri.append(lbl)

        ctk.CTkLabel(govde, text="💡 İpucu: Bir PDF'e çift tıklayarak sayfalarını hızlıca önizleyebilirsiniz.",
                     font=ctk.CTkFont(family=ARAYUZ_FONTU, size=11), text_color=("#64748B", "#94A3B8")
                     ).pack(anchor="w", padx=12, pady=(6, 12))

        # Pencere boyutu değişince açıklamalar yeni genişliğe göre satır kırsın.
        # event.width ekran pikseli, wraplength ise ölçeklenmemiş birimdir (ör. %125 ekran ölçeği).
        def genislige_uydur(event):
            olcek = govde._get_widget_scaling() or 1.0
            yeni = max(280, int(event.width / olcek) - 110)
            for lbl in aciklama_etiketleri:
                if lbl.cget("wraplength") != yeni:
                    lbl.configure(wraplength=yeni)
        # add="+": CTkScrollableFrame'in kaydırma bölgesini güncelleyen kendi <Configure> bağlantısı ezilmemeli
        govde.bind("<Configure>", genislige_uydur, add="+")

def ai_oz_test(cikti_yolu: str):
    """Paketlenmiş EXE'de AI altyapısını doğrular (penceresiz EXE konsola yazamadığı için sonucu dosyaya yazar).
    Kullanım: "UYAP Evrak Asistanı.exe" --ai-test sonuc.json"""
    import time
    sonuc = {"exe_mi": bool(getattr(sys, "frozen", False)), "ayarlar_karti_kutuphaneyi_goruyor": ai_kutuphane_mevcut(),
             "ai_motor_klasoru": ai_motor_klasoru(), "islem_birimi_ayari": ayarlari_yukle().get("ai_islem_birimi", "otomatik"),
             "kutuphane": None, "dll_klasoru": None,
             "model_indirildi": ai_model_indirildi(), "model_yuklendi": False, "cevap": None, "hata": None}
    try:
        if sonuc["model_indirildi"]:
            # Motor, llama_cpp içe aktarılmadan önce hazırlanmalı (Vulkan DLL'leri bu sırayla yüklenir)
            t0 = time.time()
            llm = AiOzetMotoru.yukle()
        import llama_cpp
        sonuc["kutuphane"] = llama_cpp.__version__
        motor = _AI_MOTOR_DURUMU and _AI_MOTOR_DURUMU.get("motor") == "vulkan" and ai_motor_klasoru()
        lib = motor or os.path.join(os.path.dirname(llama_cpp.__file__), "lib")
        sonuc["dll_klasoru"] = sorted(f for f in os.listdir(lib) if f.endswith((".dll", ".dylib"))) if os.path.isdir(lib) else None
        if sonuc["model_indirildi"]:
            sonuc["model_yuklendi"] = True
            r = llm.create_chat_completion(messages=[{"role": "user", "content": "Türkiye'nin başkenti neresidir? Tek kelimeyle cevap ver."}],
                                           max_tokens=8, temperature=0)
            sonuc["cevap"] = r["choices"][0]["message"]["content"].strip()
            sonuc["sure_sn"] = round(time.time() - t0, 1)
            sonuc["kullanilan_cihaz"] = AiOzetMotoru.cihaz_aciklama
            sonuc["motor"] = _AI_MOTOR_DURUMU
            sonuc["algilanan_gpular"] = ai_gpu_listele()
    except Exception as e:
        sonuc["hata"] = f"{type(e).__name__}: {e}"
    with open(cikti_yolu, "w", encoding="utf-8") as f:
        json.dump(sonuc, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    if IS_MAC:
        # Paketlenmiş .app'te multiprocessing yardımcı süreçleri (ör. docx2pdf/tqdm kilidinin başlattığı
        # resource_tracker) uygulamanın kendisiyle başlatılır; bu çağrı olmadan ikinci bir pencere açılır.
        import multiprocessing
        multiprocessing.freeze_support()
    if len(sys.argv) >= 3 and sys.argv[1] == "--ai-test":
        ai_oz_test(sys.argv[2])
        if IS_MAC:
            AiOzetMotoru.bosalt()   # Metal: model bırakılmadan çıkılırsa llama.cpp çöküyor (aşağıya bkz.)
        sys.exit(0)
    # AI ekran kartı algılama / doğrulama alt süreçleri (penceresiz çalışır)
    if len(sys.argv) >= 3 and sys.argv[1] == "--ai-cihazlar":
        _ai_cihazlari_yaz(sys.argv[2])
        sys.exit(0)
    if len(sys.argv) >= 4 and sys.argv[1] == "--ai-gpu-dogrula":
        _ai_gpu_dogrula_yaz(sys.argv[2], sys.argv[3])
        sys.exit(0)
    app = ModernUyapAsistani()
    app.mainloop()
    if IS_MAC:
        # macOS (Metal): model bellekteyken süreç kapanırsa llama.cpp çıkışta çöküyor (GGML_ASSERT).
        # Önce model bırakılır; özet sürerken model kullanımda olduğundan süreç doğrudan sonlandırılır.
        if getattr(app, "_ai_ozet_calisiyor", False):
            os._exit(0)
        AiOzetMotoru.bosalt()