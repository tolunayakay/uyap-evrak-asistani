# -*- mode: python ; coding: utf-8 -*-
# macOS .app paketi. Windows EXE için "UYAP Evrak Asistanı.spec" kullanılır.
# Derleme:  python -m PyInstaller --noconfirm "UYAP Evrak Asistanı (macOS).spec"
import re
from PyInstaller.utils.hooks import collect_all

datas = [('logo.png', '.')]
binaries = []
hiddenimports = []
for paket in ('llama_cpp', 'customtkinter', 'tkinterdnd2'):
    tmp_ret = collect_all(paket)
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Sürüm, uyap_asistan.py içindeki APP_VERSION'dan okunur; tek kaynaktan yönetilir.
with open('uyap_asistan.py', encoding='utf-8') as _f:
    APP_VERSION = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', _f.read(), re.M).group(1)
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
    [],
    exclude_binaries=True,
    name='UYAP Evrak Asistanı',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='UYAP Evrak Asistanı',
)
app = BUNDLE(
    coll,
    name='UYAP Evrak Asistanı.app',
    icon='logo.icns',
    bundle_identifier='com.tolunayakay.uyapevrakasistani',
    version=APP_VERSION,
    info_plist={
        'CFBundleDisplayName': 'UYAP Evrak Asistanı',
        'CFBundleShortVersionString': APP_VERSION,
        'CFBundleVersion': APP_VERSION,
        'LSMinimumSystemVersion': '12.0',
        'NSHighResolutionCapable': True,
        'NSHumanReadableCopyright': 'Tolunay Akay',
        # Word ➔ PDF dönüştürme, Microsoft Word'ü Apple Events ile kullanır
        'NSAppleEventsUsageDescription': 'Word belgelerini PDF\'e dönüştürmek için Microsoft Word kullanılır.',
    },
)

# PyInstaller, ana programın SDK sürümünü Python'unkine eşitliyor. Python.org Python'u eski bir SDK
# (ör. 12.1) ile derlendiğinde macOS, Metal'i eski uyumluluk modunda çalıştırıyor ve AI özeti yeni
# Mac'lerde (M4 vb.) "nil Metal pipeline" hatasıyla çöküyor. SDK bilgisi derleyen makinenin SDK'sına
# yükseltilir (en düşük macOS sürümü değişmez) ve paket yeniden imzalanır.
import os, subprocess
_ana_program = os.path.join(DISTPATH, 'UYAP Evrak Asistanı.app', 'Contents', 'MacOS', 'UYAP Evrak Asistanı')
_yapi = subprocess.run(['otool', '-l', _ana_program], capture_output=True, text=True).stdout
_minos = re.search(r'minos (\S+)', _yapi).group(1)
_sdk_eski = re.search(r'sdk (\S+)', _yapi).group(1)
_sdk_yeni = subprocess.run(['xcrun', '--show-sdk-version'], capture_output=True, text=True).stdout.strip()
if int(_sdk_eski.split('.')[0]) < 14:
    subprocess.run(['xcrun', 'vtool', '-set-build-version', 'macos', _minos, _sdk_yeni, '-replace',
                    '-output', _ana_program + '.tmp', _ana_program], check=True)
    os.replace(_ana_program + '.tmp', _ana_program)
    os.chmod(_ana_program, 0o755)
    subprocess.run(['codesign', '--force', '--deep', '--sign', '-',
                    os.path.join(DISTPATH, 'UYAP Evrak Asistanı.app')], check=True)
    print(f"Ana program SDK sürümü {_sdk_eski} → {_sdk_yeni} (en düşük macOS {_minos})")
