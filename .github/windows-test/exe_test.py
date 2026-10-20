# -*- coding: utf-8 -*-
"""
Derlenen Windows .exe'yi dener: AI öz testi (--ai-test) ve pencerenin açılıp açık kalması.
Kullanım: python exe_test.py <exe yolu> <çıktı klasörü>
"""
import json
import os
import subprocess
import sys
import time

exe, cikti = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
os.makedirs(cikti, exist_ok=True)
sonuc = {"exe": os.path.basename(exe), "boyut_mb": round(os.path.getsize(exe) / 1048576, 1)}

ai_json = os.path.join(cikti, "exe_ai_test.json")
t0 = time.time()
r = subprocess.run([exe, "--ai-test", ai_json], timeout=900)
sonuc["ai_test_cikis_kodu"] = r.returncode
try:
    with open(ai_json, encoding="utf-8") as f:
        ai = json.load(f)
    sonuc["ai_test"] = {k: ai.get(k) for k in ("exe_mi", "kutuphane", "ai_motor_klasoru", "model_yuklendi", "cevap",
                                                "kullanilan_cihaz", "hata")}
    sonuc["ai_test"]["sure_sn"] = round(time.time() - t0, 1)
except Exception as e:
    sonuc["ai_test"] = {"hata": f"sonuç okunamadı: {e}"}

p = subprocess.Popen([exe])
time.sleep(20)
sonuc["acildi"] = p.poll() is None
try:
    from PIL import ImageGrab
    ImageGrab.grab().save(os.path.join(cikti, "ekran_exe.png"))
except Exception as e:
    sonuc["ekran"] = str(e)
p.kill()
subprocess.run(["taskkill", "/F", "/IM", os.path.basename(exe)], capture_output=True)

with open(os.path.join(cikti, "exe_sonuc.json"), "w", encoding="utf-8") as f:
    json.dump(sonuc, f, ensure_ascii=False, indent=1)
print(json.dumps(sonuc, ensure_ascii=False, indent=1))
