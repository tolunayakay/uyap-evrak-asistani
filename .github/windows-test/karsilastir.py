# -*- coding: utf-8 -*-
"""
main ile macOS dalının Windows test sonuçlarını karşılaştırır, GitHub özet sayfasına tablo yazar.
Dalda main'de geçip burada kalan bir test varsa veya .exe / AI kontrolü başarısızsa hata koduyla çıkar.
Kullanım: python karsilastir.py <main klasörü> <dal klasörü> <exe_sonuc.json>
"""
import json
import os
import sys

ana_kl, dal_kl, exe_json = sys.argv[1:4]


def oku(yol, varsayilan=None):
    try:
        with open(yol, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return varsayilan


ana = {s["test"]: s for s in oku(os.path.join(ana_kl, "sonuclar.json"), [])}
dal = {s["test"]: s for s in oku(os.path.join(dal_kl, "sonuclar.json"), [])}
ai_ana = oku(os.path.join(ana_kl, "ai.json"), {})
ai_dal = oku(os.path.join(dal_kl, "ai.json"), {})
exe = oku(exe_json, {})

satirlar = ["## Windows doğrulama: main ↔ feat/macos-uyumlulugu", "",
            "| Test | main | bu dal | Not (bu dal) |", "|---|---|---|---|"]
gerileme = []
for ad in list(dict.fromkeys(list(ana) + list(dal))):
    a, d = ana.get(ad, {}).get("durum", "—"), dal.get(ad, {}).get("durum", "—")
    if a == "GEÇTİ" and d != "GEÇTİ":
        gerileme.append(ad)
    isaret = lambda x: {"GEÇTİ": "✅", "KALDI": "❌"}.get(x, x)
    satirlar.append(f"| {ad} | {isaret(a)} | {isaret(d)} | {dal.get(ad, {}).get('not', '').replace('|', '/')[:110]} |")

satirlar += ["", "### AI Evrak Özeti (işlemci)", "", "| | main | bu dal |", "|---|---|---|"]
for k in ("tur", "cihaz", "sure_sn", "eksik", "hata"):
    satirlar.append(f"| {k} | {ai_ana.get(k, '—')} | {ai_dal.get(k, '—')} |")
ai_tamam = ai_dal.get("ozet") and not ai_dal.get("eksik") and not ai_dal.get("hata")

satirlar += ["", "### Bu daldan derlenen .exe", ""]
for k, v in exe.items():
    satirlar.append(f"- **{k}:** {v}")
exe_tamam = exe.get("acildi") and exe.get("ai_test", {}).get("cevap")

gecen = sum(s["durum"] == "GEÇTİ" for s in dal.values())
satirlar += ["", f"**Sonuç:** bu dalda {gecen}/{len(dal)} test geçti; main'e göre gerileme: "
             f"{', '.join(gerileme) if gerileme else 'yok'}; AI: {'tamam' if ai_tamam else 'SORUNLU'}; "
             f".exe: {'tamam' if exe_tamam else 'SORUNLU'}"]
metin = "\n".join(satirlar)
print(metin)
if os.environ.get("GITHUB_STEP_SUMMARY"):
    with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
        f.write(metin + "\n")
sys.exit(0 if (not gerileme and gecen == len(dal) and ai_tamam and exe_tamam) else 1)
