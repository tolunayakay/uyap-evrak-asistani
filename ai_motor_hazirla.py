# -*- coding: utf-8 -*-
"""
AI motoru (GPU / Vulkan) DLL'lerini hazırlar.

llama-cpp-python'un hazır paketi yalnızca işlemcide (CPU) çalışır. Bu betik, resmi
llama.cpp sürümünün Vulkan paketini indirir, SHA256 ile doğrular ve gereken DLL'leri
proje içindeki 'ai_motor' klasörüne çıkarır. Uygulama bu klasörü bulursa ekran kartını
(NVIDIA / AMD / Intel) kullanabilir; bulamazsa eskisi gibi yalnızca işlemcide çalışır.

Kullanım (EXE derlemeden önce bir kez):
    python ai_motor_hazirla.py

Not: Sürüm, llama-cpp-python 0.3.36 ile birlikte test edildi. Sürümü değiştirirseniz
uygulamanın '--ai-test' öz testiyle yeniden doğrulayın.
"""
import hashlib
import io
import os
import sys
import urllib.request
import zipfile

LLAMA_CPP_SURUM = "b11734"
PAKET_URL = (f"https://github.com/ggml-org/llama.cpp/releases/download/{LLAMA_CPP_SURUM}/"
             f"llama-{LLAMA_CPP_SURUM}-bin-win-vulkan-x64.zip")
PAKET_SHA256 = "b6e3f7aec77d0634f094bf0468b723509d440e17b0b2f25880b011b791e5d831"

HEDEF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ai_motor")


def gerekli_mi(ad: str) -> bool:
    ad = ad.lower()
    return ad in ("llama.dll", "ggml.dll", "ggml-base.dll", "ggml-vulkan.dll", "libomp.dll") or \
        (ad.startswith("ggml-cpu-") and ad.endswith(".dll"))


def main():
    print(f"İndiriliyor: {PAKET_URL}")
    with urllib.request.urlopen(urllib.request.Request(PAKET_URL, headers={"User-Agent": "UyapEvrakAsistani"}), timeout=120) as r:
        veri = r.read()
    ozet = hashlib.sha256(veri).hexdigest()
    if ozet != PAKET_SHA256:
        sys.exit(f"HATA: SHA256 uyuşmuyor!\n  beklenen: {PAKET_SHA256}\n  gelen:    {ozet}")
    print("SHA256 doğrulandı.")

    os.makedirs(HEDEF, exist_ok=True)
    for eski in os.listdir(HEDEF):
        if eski.lower().endswith(".dll"):
            os.remove(os.path.join(HEDEF, eski))

    toplam = 0
    with zipfile.ZipFile(io.BytesIO(veri)) as z:
        for bilgi in z.infolist():
            ad = os.path.basename(bilgi.filename)
            if bilgi.is_dir() or not gerekli_mi(ad):
                continue
            with open(os.path.join(HEDEF, ad), "wb") as f:
                f.write(z.read(bilgi))
            toplam += bilgi.file_size
            print(f"  + {ad} ({bilgi.file_size / 1048576:.1f} MB)")
    with open(os.path.join(HEDEF, "SURUM.txt"), "w", encoding="utf-8") as f:
        f.write(f"llama.cpp {LLAMA_CPP_SURUM} (Vulkan)\n{PAKET_URL}\nSHA256 {PAKET_SHA256}\n")
    print(f"Tamamlandı: {HEDEF} ({toplam / 1048576:.1f} MB)")


if __name__ == "__main__":
    main()
