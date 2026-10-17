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
# GİZLİ AYAR YÖNETİCİSİ (APPDATA / CONFIG DİZİNİNDE SAKLAR)
# ============================================================================
def get_ayar_dosya_yolu() -> str:
    if os.name == 'nt':
        taban_dizin = os.environ.get('APPDATA', os.path.expanduser('~'))
        klasor = os.path.join(taban_dizin, 'UyapAsistani')
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
        "kase_renk": "kirmizi"
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


# ============================================================================
# e-İMZA & ZAMAN DAMGASI DOĞRULAMA / TESPİT MOTORU
# ============================================================================
def eimza_bilgisi_al(dosya_yolu: str) -> dict:
    uzanti = os.path.splitext(dosya_yolu)[1].lower()
    sonuc = {
        "imzali": False,
        "imzalayan": None,
        "zaman_damgasi": False,
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
                    ham_veri = zf.read(imza_dosyalari[0])
                    
                    if b"1.2.840.113549.1.9.16.1.4" in ham_veri or b"TimeStamp" in ham_veri or b"Zaman Damgasi" in ham_veri:
                        sonuc["zaman_damgasi"] = True
                    
                    cn_eslesme = re.findall(rb"CN=([^\x00-\x1f,]+)", ham_veri)
                    if not cn_eslesme:
                        cn_eslesme = re.findall(rb"(?:Av\.|Hakim|Savci|Katip)[\s\w]+", ham_veri)

                    if cn_eslesme:
                        try:
                            sonuc["imzalayan"] = cn_eslesme[0].decode('utf-8', errors='ignore').strip()
                        except Exception:
                            sonuc["imzalayan"] = cn_eslesme[0].decode('iso-8859-9', errors='ignore').strip()
                    else:
                        sonuc["imzalayan"] = "Nitelikli Elektronik İmza (UDF)"

                    sonuc["detay"] = f"Geçerli UYAP e-İmzası\nİmzalayan: {sonuc['imzalayan']}"
                    if sonuc["zaman_damgasi"]:
                        sonuc["detay"] += "\nZaman Damgası: Onaylı (TSA Mevcut)"
                else:
                    if 'content.xml' in isimler:
                        xml_raw = zf.read('content.xml')
                        if b"Signature" in xml_raw or b"ds:Signature" in xml_raw:
                            sonuc["imzali"] = True
                            sonuc["imzalayan"] = "UDF XML İmzası"
                            sonuc["detay"] = "UDF XML Formatında İmzalı Evrak"

        elif uzanti == '.pdf':
            with open(dosya_yolu, 'rb') as f:
                pdf_bytes = f.read()

            if b"/ByteRange" in pdf_bytes and (b"/Contents" in pdf_bytes or b"/adbe.pkcs7" in pdf_bytes):
                sonuc["imzali"] = True
                
                name_eslesme = re.search(rb'/Name\s*\(([^)]+)\)', pdf_bytes)
                if not name_eslesme:
                    name_eslesme = re.search(rb'/Name\s*<([^>]+)>', pdf_bytes)
                
                if name_eslesme:
                    isim_raw = name_eslesme.group(1)
                    try:
                        sonuc["imzalayan"] = isim_raw.decode('utf-8', errors='ignore')
                    except Exception:
                        sonuc["imzalayan"] = isim_raw.decode('iso-8859-9', errors='ignore')
                else:
                    sonuc["imzalayan"] = "Elektronik İmza (PAdES)"

                date_eslesme = re.search(rb'/M\s*\(D:([0-9]{14})', pdf_bytes)
                if date_eslesme:
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
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.lbl_sayfa_sayisi.pack(side="left")

        self.lbl_secim = ctk.CTkLabel(
            sol_bilgi,
            text="• 0 sayfa seçildi",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#64748B", "#94A3B8")
        )
        self.lbl_secim.pack(side="left", padx=8)

        sag_aksiyon = ctk.CTkFrame(ust_bar, fg_color="transparent")
        sag_aksiyon.pack(side="right", padx=16)

        ctk.CTkButton(
            sag_aksiyon,
            text="⟳ Seçilenleri 90° Döndür",
            command=self._secilenleri_dondur,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            text="💡 İpucu: Sayfa kartına tıklayarak seçebilir, altındaki butonlarla tek tek döndürebilir veya silebilirsiniz.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
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

        self.gorsel_referanslari.clear()
        self.lbl_sayfa_sayisi.configure(text=f"Toplam: {len(self.sayfa_durumlari)} Sayfa")
        self.lbl_secim.configure(text=f"• {len(self.secili_sayfalar)} sayfa seçildi")

        if not self.sayfa_durumlari:
            ctk.CTkLabel(
                self.scroll_frame,
                text="Tüm sayfalar silindi!\nLütfen pencereyi kapatın veya işlemi iptal edin.",
                font=ctk.CTkFont(family="Segoe UI", size=13),
                text_color=("#94A3B8", "#64748B")
            ).pack(pady=100)
            return

        sutun_sayisi = 4
        for sira_no, item in enumerate(self.sayfa_durumlari):
            row = sira_no // sutun_sayisi
            col = sira_no % sutun_sayisi

            is_selected = sira_no in self.secili_sayfalar

            kart = ctk.CTkFrame(
                self.scroll_frame,
                fg_color=("#FFFFFF", "#1E293B"),
                corner_radius=8,
                border_width=2 if is_selected else 1,
                border_color=("#3B82F6", "#60A5FA") if is_selected else ("#E2E8F0", "#334155")
            )
            kart.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

            ust_satir = ctk.CTkFrame(kart, fg_color="transparent", height=24)
            ust_satir.pack(fill="x", padx=8, pady=(6, 2))

            chk_text = "✓" if is_selected else "○"
            chk_color = "#3B82F6" if is_selected else ("#94A3B8", "#64748B")
            lbl_chk = ctk.CTkLabel(
                ust_satir,
                text=chk_text,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color=chk_color
            )
            lbl_chk.pack(side="left")

            lbl_no = ctk.CTkLabel(
                ust_satir,
                text=f"Sayfa {sira_no + 1}",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color=("#0F172A", "#F8FAFC")
            )
            lbl_no.pack(side="right")

            ctk_img = self._sayfa_thumbnail_al(item["orijinal_idx"], item["rotation"])
            self.gorsel_referanslari[sira_no] = ctk_img

            lbl_img = ctk.CTkLabel(kart, text="", image=ctk_img, cursor="hand2")
            lbl_img.pack(padx=8, pady=4)

            def toggle_secim(idx=sira_no):
                if idx in self.secili_sayfalar:
                    self.secili_sayfalar.remove(idx)
                else:
                    self.secili_sayfalar.add(idx)
                self._sayfalari_ciz()

            kart.bind("<Button-1>", lambda e, idx=sira_no: toggle_secim(idx))
            lbl_img.bind("<Button-1>", lambda e, idx=sira_no: toggle_secim(idx))
            lbl_no.bind("<Button-1>", lambda e, idx=sira_no: toggle_secim(idx))
            lbl_chk.bind("<Button-1>", lambda e, idx=sira_no: toggle_secim(idx))

            alt_satir = ctk.CTkFrame(kart, fg_color="transparent")
            alt_satir.pack(fill="x", padx=6, pady=(4, 6))

            ctk.CTkButton(
                alt_satir,
                text="⟳ 90°",
                command=lambda idx=sira_no: self._tekil_dondur(idx),
                width=50,
                height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("#F1F5F9", "#334155"),
                hover_color=("#E2E8F0", "#475569"),
                text_color=("#1E293B", "#F8FAFC"),
                corner_radius=5
            ).pack(side="left", expand=True, fill="x", padx=(0, 2))

            ctk.CTkButton(
                alt_satir,
                text="🗑 Sil",
                command=lambda idx=sira_no: self._tekil_sil(idx),
                width=50,
                height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("#FEE2E2", "#7F1D1D"),
                hover_color=("#FECACA", "#991B1B"),
                text_color=("#991B1B", "#FCA5A5"),
                corner_radius=5
            ).pack(side="right", expand=True, fill="x", padx=(2, 0))

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
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(side="left")

        ctk.CTkLabel(
            logo_alani,
            text="PRO",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
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
            text="Hakkında",
            command=self.hakkinda_penceresi_ac,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#334155", "#F8FAFC"),
            corner_radius=8,
            width=90,
            height=32
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            sag_aksiyon_frame,
            text="Nasıl Kullanılır?",
            command=self.rehber_penceresi_ac,
            font=ctk.CTkFont(family="Segoe UI", size=12),
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

        self.durum_etiketi = ctk.CTkLabel(
            alt_bar,
            text="Hazır. Evraklarınızı sol alana sürükleyebilir veya 'Dosya Ekle' ile seçebilirsiniz.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
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
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#94A3B8", "#64748B")
        ).pack(side="left")

        self.dosya_sayisi_badge = ctk.CTkLabel(
            sol_baslik_satir,
            text="0 dosya",
            font=ctk.CTkFont(family="Segoe UI", size=11),
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
            font=ctk.CTkFont(family="Segoe UI", size=12),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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

        self._sekme_renklerini_guncelle(self.tabview.get())

    def _tema_degistir(self):
        if self.tema_switch.get() == 1:
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
    # SEKME 1: BİRLEŞTİR
    # ========================================================
    def _sekme_birlestir_olustur(self):
        tab = self.tab_birlestir

        ctk.CTkLabel(tab, text="SIKIŞTIRMA MODU", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 4))
        self.mod_secimi = ctk.StringVar(value="akilli_vektor")

        ctk.CTkRadioButton(
            tab,
            text="Akıllı Mod (Metinleri Koru, Görselleri Sıkıştır)",
            variable=self.mod_secimi,
            value="akilli_vektor",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#334155", "#F8FAFC")
        ).pack(anchor="w", pady=3)

        ctk.CTkRadioButton(
            tab,
            text="Agresif Raster (150 DPI Taranmış Evrak Küçültme)",
            variable=self.mod_secimi,
            value="maksimum_raster",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#334155", "#F8FAFC")
        ).pack(anchor="w", pady=3)

        ctk.CTkLabel(tab, text="BATES NUMARALANDIRMA", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(14, 4))
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
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#334155", "#F8FAFC")
        )
        self.bates_switch.pack(side="left", padx=12, pady=10)

        self.bates_on_ek = ctk.CTkEntry(bates_frame, width=90, height=28, placeholder_text="EK-1", state="disabled")
        self.bates_on_ek.insert(0, "EK-1")
        self.bates_on_ek.pack(side="right", padx=12)

        ctk.CTkLabel(tab, text="EYLEMLER", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(16, 6))

        ctk.CTkButton(
            tab,
            text="📑 Listedeki Belgeleri Birleştir (PDF Yap)",
            command=lambda: self._thread_calistir(self.tumunu_birlestir_pdf_yap),
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 2: SAYFALAR
    # ========================================================
    def _sekme_duzenle_olustur(self):
        tab = self.tab_duzenle

        ctk.CTkLabel(tab, text="GÖRSEL SAYFA YÖNETİMİ", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🖼 Sayfaları Görsel Önizle & Yönet (Grid Panel)",
            command=self.pdf_gorsel_onizleme_ac,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=38,
            corner_radius=8
        ).pack(fill="x", pady=4)

        ctk.CTkLabel(tab, text="HIZLI SAYFA AYIRMA VE DÖNDÜRME", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(12, 6))

        ctk.CTkButton(
            tab,
            text="🔄 Sayfa Döndür (90° / 180°)",
            command=self.pdf_dondur_penceresi,
            fg_color=("#F1F5F9", "#334155"),
            hover_color=("#E2E8F0", "#475569"),
            text_color=("#1E293B", "#F8FAFC"),
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", pady=4)

    # ========================================================
    # SEKME 3: DÖNÜŞTÜR (YENİ MODÜLLER EKLENDİ)
    # ========================================================
    def _sekme_donusturucu_olustur(self):
        tab = self.tab_donusturucu

        ctk.CTkLabel(tab, text="UYAP VE OFİS DÖNÜŞÜMLERİ", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 4))

        ctk.CTkButton(
            tab,
            text="📋 PDF ➔ UDF (.udf) (Şablon Editörüne Aktar)",
            command=self.pdf_to_udf_dialog,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=34,
            corner_radius=8
        ).pack(fill="x", pady=(6, 3))

    # ========================================================
    # SEKME 4: KVKK & KAŞE
    # ========================================================
    def _sekme_guvenlik_olustur(self):
        tab = self.tab_guvenlik

        ctk.CTkLabel(tab, text="HUKUKİ İŞLEMLER VE GÜVENLİK", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=("#94A3B8", "#64748B")).pack(anchor="w", pady=(4, 6))

        ctk.CTkButton(
            tab,
            text="🔏 e-İmza & Zaman Damgası Doğrula (Ön Kontrol)",
            command=self.eimza_dogrula_penceresi,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
                text="Belgeleri buraya sürükleyin\nveya aşağıdaki butondan ekleyin\n(PDF, Word, Görsel, UDF)\n\n💡 İpucu: PDF'e çift tıklayarak sayfalarını görsel önizleyebilirsiniz.",
                font=ctk.CTkFont(family="Segoe UI", size=12),
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
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold" if is_selected else "normal"),
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
                    font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
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
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=("#64748B", "#94A3B8")
            )
            lbl_boyut.pack(side="right", padx=10)
            lbl_boyut.bind("<Button-1>", lambda e, idx=i: secim_yap(idx))
            lbl_boyut.bind("<Double-Button-1>", lambda e, y=yol: cift_tikla(y))

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

        ctk.CTkLabel(dialog, text="📋 UYAP Şablon Editörüne Aktarma", font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")).pack(pady=(16, 4))
        ctk.CTkLabel(dialog, text=os.path.basename(kaynak_yol), font=ctk.CTkFont(family="Segoe UI", size=11), text_color=("#64748B", "#94A3B8")).pack(pady=(0, 10))

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
                            os.startfile(hedef_yol)
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
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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
                    os.startfile(hedef_yol)
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
                    finally:
                        sys.stderr = eski_stderr
                else:
                    raise RuntimeError("Microsoft Word veya docx2pdf motoruna erişilemedi.")

            self.durum_etiketi.configure(text="Word ➔ PDF dönüşümü tamamlandı.")
            messagebox.showinfo("Başarılı", f"Word belgesi başarıyla PDF'e dönüştürüldü!\nDosya: {os.path.basename(hedef_yol)}")
        except Exception as e:
            messagebox.showerror(
                "Dönüşüm Hatası",
                f"Word ➔ PDF dönüştürme hatası:\n{str(e)}\n\n"
                "Not: Bilgisayarınızda Microsoft Word yüklü olmalı ve arka planda açık olan Word pencereleri kapatılmalıdır."
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

            txt = ctk.CTkTextbox(pencere, font=ctk.CTkFont(family="Consolas", size=12))
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
    # 4. KVKK & KAŞE & GÜVENLİK İŞLEMLERİ
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
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(pady=(20, 2))

        ctk.CTkLabel(
            dialog,
            text=os.path.basename(kaynak_yol),
            font=ctk.CTkFont(family="Segoe UI", size=11),
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
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=durum_renk
        ).pack(anchor="w", padx=16, pady=(14, 6))

        def satir_ekle(etiket, deger):
            f = ctk.CTkFrame(kart, fg_color="transparent")
            f.pack(fill="x", padx=16, pady=3)
            ctk.CTkLabel(f, text=etiket, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=("#475569", "#CBD5E1"), width=120, anchor="w").pack(side="left")
            ctk.CTkLabel(f, text=deger or "-", font=ctk.CTkFont(family="Segoe UI", size=11), text_color=("#0F172A", "#F8FAFC"), anchor="w").pack(side="left", fill="x", expand=True)

        satir_ekle("İmzalayan Kişi:", bilgi["imzalayan"])
        satir_ekle("İmza Formatı:", "CAdES / UYAP UDF" if uzanti == '.udf' else "PAdES / PDF")
        if bilgi.get("tarih"):
            satir_ekle("İmza Zamanı:", bilgi["tarih"])
        satir_ekle("Zaman Damgası:", "Mevcut (TSA Onaylı)" if bilgi["zaman_damgasi"] else "Bulunmuyor (Yerel Saat)")

        ctk.CTkButton(
            dialog,
            text="Kapat",
            command=dialog.destroy,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
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

        ctk.CTkLabel(dialog, text="Kaşe Başlığı:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(12, 2))
        baslik_entry = ctk.CTkEntry(dialog)
        baslik_entry.insert(0, self.ayarlar.get("kase_baslik", "ASLI GİBİDİR"))
        baslik_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Avukat İsim / Sicil No / Unvan:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        avukat_entry = ctk.CTkEntry(dialog, placeholder_text="Av. Ad Soyad - Baro Sicil No")
        kayitli_avukat = self.ayarlar.get("avukat_bilgisi", "")
        if kayitli_avukat:
            avukat_entry.insert(0, kayitli_avukat)
        avukat_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(dialog, text="Tarih / Ek Açıklama:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
        ek_entry = ctk.CTkEntry(dialog)
        ek_entry.insert(0, f"Tarih: {bugun}")
        ek_entry.pack(fill="x", padx=20)

        kayitli_konum_gorunum = self.ayarlar.get("kase_konum", "Sağ Alt")
        konum_haritasi = {"Sağ Alt": "sag_alt", "Sol Alt": "sol_alt", "Alt Orta": "alt_orta", "Sağ Üst": "sag_ust"}
        konum_var = ctk.StringVar(value=konum_haritasi.get(kayitli_konum_gorunum, "sag_alt"))

        ctk.CTkLabel(dialog, text="Konum:", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")).pack(anchor="w", padx=20, pady=(8, 2))
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

        ctk.CTkButton(
            dialog,
            text="🛡 Kalıcı Karart ve Kaydet",
            command=uygula,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            corner_radius=8
        ).pack(pady=18)

    def rehber_penceresi_ac(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Kullanım Rehberi")
        pencere.geometry("680x560")
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
• 'Sayfaları Görsel Önizle & Yönet' paneli veya listedeki PDF'e çift tıklayarak sayfaları küçük resimler halinde görebilir, tek tek veya toplu olarak döndürüp silebilirsiniz.
• Boş sayfaları temizleme butonu çift taraflı taramalardaki boş sayfaları tespit eder ve onayınızı alarak temizler.
• UYAP Akıllı MB Bölücüsü büyük dosyaları UYAP limitine (9.5 MB) göre sayfa bütünlüğünü bozmadan ayırır.

3. Dönüştür:
• PDF ➔ UDF (.udf): PDF evraklarını düzenlenebilir metin katmanları, akıllı tablo yapıları ve gömülü görselleriyle UYAP Şablon Editörünün açabileceği UDF formatına dönüştürür.
• PDF ➔ Word (.docx): PDF belgelerini sütun yapısını, tabloları ve başlıkları koruyarak düzenlenebilir Word evrakına çevirir.
• UDF (.udf) ➔ Word: UYAP UDF dosyalarını tek tıkla Microsoft Word (.docx) formatına aktarır.
• Word (.docx) ➔ UDF: Hazırladığınız dilekçeleri resim ve tablolarıyla UDF editörüne taşır.

4. KVKK, Kaşe & Güvenlik:
• 'e-İmza & Zaman Damgası Doğrula': Seçili UDF veya PDF'in nitelikli elektronik imza sahibini ve TÜBİTAK/TSA zaman damgası durumunu tek tıkla analiz eder.
• 'Aslı Gibidir' kaşesi belgenin yönüne duyarlı şekilde otomatik yerleşir. Tercihleriniz hatırlanır.
• KVKK Karartma modülü yalnızca geçerli T.C. Kimlik numaralarını algoritmayla doğrulayarak ve IBAN'ları tespit ederek maskeler; ardından metaverileri PDF'ten tamamen kazır."""

        txt.insert("0.0", rehber.strip())
        txt.configure(state="disabled")

    def hakkinda_penceresi_ac(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Hakkında")
        dialog.geometry("460x390")
        dialog.resizable(False, False)
        dialog.after(100, dialog.lift)

        ctk.CTkLabel(
            dialog,
            text="⚖ UYAP Evrak Asistanı",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(pady=(20, 2))

        ctk.CTkLabel(
            dialog,
            text="Sürüm 1.1.0 • UYAP Kullanıcıları İçin Tasarlandı",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#64748B", "#94A3B8")
        ).pack(pady=(0, 16))

        kart = ctk.CTkFrame(
            dialog,
            fg_color=("#F8FAFC", "#1E293B"),
            corner_radius=10,
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        kart.pack(fill="x", padx=24, pady=4)

        ctk.CTkLabel(
            kart,
            text="Proje Geliştiricisi:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#334155", "#E2E8F0")
        ).pack(anchor="w", padx=14, pady=(10, 1))

        ctk.CTkLabel(
            kart,
            text="Tolunay Akay",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(anchor="w", padx=14, pady=(0, 10))

        ctk.CTkLabel(
            kart,
            text="Katkıda Bulunanlar (Credits):",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#334155", "#E2E8F0")
        ).pack(anchor="w", padx=14, pady=(4, 1))

        ctk.CTkLabel(
            kart,
            text="• Zafer Kaya",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#64748B", "#94A3B8")
        ).pack(anchor="w", padx=14, pady=(0, 12))

        def mail_at():
            webbrowser.open("mailto:tolunay.akay@galahukuk.com")

        ctk.CTkButton(
            dialog,
            text="✉ Geliştirici ile İletişim / Geri Bildirim",
            command=mail_at,
            fg_color=("#0F172A", "#3B82F6"),
            hover_color=("#1E293B", "#2563EB"),
            text_color="#FFFFFF",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=36,
            corner_radius=8
        ).pack(fill="x", padx=24, pady=(18, 8))

        ctk.CTkLabel(
            dialog,
            text="Hata bildirimi ve yeni özellik önerileri için dilediğiniz zaman ulaşabilirsiniz.",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=("#94A3B8", "#64748B")
        ).pack()


if __name__ == "__main__":
    app = ModernUyapAsistani()
    app.mainloop()