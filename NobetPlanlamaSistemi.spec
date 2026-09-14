# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['nobet_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # httpx/anyio/certifi import grafinden otomatik bulunur; elle eklemeye
    # gerek yok. sniffio kurulu degil ve senkron httpx istemcisi onu
    # kullanmiyor (dev ortaminda da yok, sorunsuz calisiyor).
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # openpyxl bunlari ISTEGE BAGLI olarak import ediyor (numpy: ekstra sayi
    # tipleri, PIL: Excel'e resim gomme). Ikisini de kullanmiyoruz - biz sadece
    # metin/sayi yaziyoruz. Ancak PyInstaller bunlari bulup binlerce modulu
    # tarayip pakete koyuyordu; derlemeyi asiri agirlastiran (ve PC'yi
    # cokerten) asil sebep buydu. Disarida birakmak exe'yi de kucultuyor.
    # numpy/PIL yokken openpyxl try/except ile sorunsuz calisiyor (test edildi).
    excludes=[
        'numpy', 'PIL',
        'setuptools', 'pkg_resources', 'distutils',  # sadece kurulum araclari
        'pytest', 'unittest', 'doctest', 'pydoc',    # test/dokuman araclari
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# Acilis (splash) ekrani: onefile exe kendini gecici klasore acarken (birkac
# saniye) bu gorsel gorunur, boylece 'dosya olmus mu' hissi olmaz. Uygulama
# hazir olunca kod icinden pyi_splash.close() ile kapatilir.
splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(18, 240),
    text_size=9,
    text_color='#C8D6EB',
    text_default='Yükleniyor…',
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    splash,
    splash.binaries,
    [],
    name='NobetPlanlamaSistemi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX zaten sistemde kurulu degil (upx=True bos calisiyordu). Karisiklik
    # olmasin diye acikca kapatildi.
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
