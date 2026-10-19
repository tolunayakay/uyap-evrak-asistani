# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('logo.ico', '.')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('llama_cpp')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# AI motoru (Vulkan / ekran kartı desteği): 'python ai_motor_hazirla.py' ile üretilir.
# Klasör olduğu gibi kopyalanır (uygulama DLL'leri bu klasörden yükler).
import os, glob
if glob.glob(os.path.join('ai_motor', 'ggml-vulkan.dll')):
    datas += [(f, 'ai_motor') for f in glob.glob(os.path.join('ai_motor', '*.dll'))]
else:
    print("UYARI: ai_motor klasörü yok; EXE yalnızca işlemcide (CPU) çalışacak. "
          "Ekran kartı desteği için önce 'python ai_motor_hazirla.py' çalıştırın.")


# Windows dosya özellikleri (Özellikler > Ayrıntılar) için sürüm bilgisi.
# Sürüm, uyap_asistan.py içindeki APP_VERSION'dan okunur; tek kaynaktan yönetilir.
import re
from PyInstaller.utils.win32.versioninfo import (VSVersionInfo, FixedFileInfo, StringFileInfo, StringTable,
                                                 StringStruct, VarFileInfo, VarStruct)
with open('uyap_asistan.py', encoding='utf-8') as _f:
    APP_VERSION = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', _f.read(), re.M).group(1)
_surum = tuple(int(x) for x in APP_VERSION.split('.')) + (0,) * (4 - len(APP_VERSION.split('.')))
surum_bilgisi = VSVersionInfo(
    ffi=FixedFileInfo(filevers=_surum, prodvers=_surum),
    kids=[
        StringFileInfo([StringTable('041F04B0', [   # Türkçe, Unicode
            StringStruct('ProductName', 'UYAP Evrak Asistanı'),
            StringStruct('FileDescription', 'UYAP Evrak Asistanı'),
            StringStruct('CompanyName', 'Tolunay Akay'),
            StringStruct('LegalCopyright', 'Tolunay Akay'),
            StringStruct('ProductVersion', APP_VERSION),
            StringStruct('FileVersion', APP_VERSION),
            StringStruct('OriginalFilename', 'UYAP Evrak Asistanı.exe'),
        ])]),
        VarFileInfo([VarStruct('Translation', [0x041F, 1200])]),
    ],
)
print(f"Sürüm: {APP_VERSION}")


a = Analysis(
    ['uyap_asistan.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='UYAP Evrak Asistanı',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo.ico'],
    version=surum_bilgisi,
)
