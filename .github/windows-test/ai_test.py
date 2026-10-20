# -*- coding: utf-8 -*-
"""
Windows'ta AI Evrak Özeti'ni (uygulamanın kendi motoruyla, işlemcide) dener.
Kullanım: python ai_test.py <uyap_asistan.py'nin bulunduğu klasör> <sonuç.json>
"""
import json
import os
import sys
import time

KAYNAK = os.path.abspath(sys.argv[1])
SONUC_YOLU = os.path.abspath(sys.argv[2])
os.chdir(KAYNAK)
sys.path.insert(0, KAYNAK)
sys.argv = ["uyap_asistan.py"]
import uyap_asistan as u

METIN = ("T.C. İSTANBUL 3. ASLİYE HUKUK MAHKEMESİ\nESAS NO: 2026/123\nKARAR NO: 2026/456\n"
         "DAVACI: Ahmet Yılmaz\nVEKİLİ: Av. Mehmet Demir\nDAVALI: Ayşe Şahin\nDAVA: Kira bedelinin tespiti\n\n"
         "Davacı vekili dava dilekçesinde; müvekkiline ait dairenin davalıya aylık 8.000 TL bedelle kiraya "
         "verildiğini belirterek kira bedelinin aylık 30.000 TL olarak tespitini talep etmiştir. Mahkemece "
         "yaptırılan bilirkişi incelemesinde emsal kira bedelinin aylık 25.000 TL olduğu tespit edilmiştir.\n\n"
         "HÜKÜM: Davanın KISMEN KABULÜ ile aylık kira bedelinin 25.000 TL olarak TESPİTİNE karar verildi.")

sonuc = {"model_indirildi": u.ai_model_indirildi()}
t0 = time.time()
try:
    llm = u.AiOzetMotoru.yukle()
    tur = u.belge_turu_tespit(METIN)
    metin, _ = u.ai_metni_hazirla(METIN)
    mesajlar, _ = u.AiOzetMotoru.mesajlari_hazirla(llm, metin, tur)
    r = llm.create_chat_completion(messages=mesajlar, max_tokens=u.AI_CIKTI_TOKEN, temperature=0.1, top_p=0.9,
                                   repeat_penalty=1.0)
    ozet = r["choices"][0]["message"]["content"]
    sonuc.update(tur=tur, cihaz=u.AiOzetMotoru.cihaz_aciklama, sure_sn=round(time.time() - t0, 1), ozet=ozet,
                 eksik=[k for k in ("Ahmet Yılmaz", "Ayşe Şahin", "25.000", "2026/123") if k not in ozet])
    u.AiOzetMotoru.bosalt()
except Exception as e:
    sonuc["hata"] = f"{type(e).__name__}: {e}"
with open(SONUC_YOLU, "w", encoding="utf-8") as f:
    json.dump(sonuc, f, ensure_ascii=False, indent=1)
print(json.dumps({k: v for k, v in sonuc.items() if k != "ozet"}, ensure_ascii=False))
