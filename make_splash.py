"""Acilis (splash) gorselini uretir: splash.png. Bir kez calistirilir; ciktisi
spec dosyasinda kullanilir. Onefile exe acilirken (kendini gecici klasore
acarken) bu gorsel gosterilir, boylece 'dosya olmus mu' hissi olmaz."""
from PIL import Image, ImageDraw, ImageFont

W, H = 460, 260
UST = (48, 84, 150)      # koyu mavi (uygulamanin baslik rengi 305496)
ALT = (32, 58, 108)

img = Image.new("RGB", (W, H), UST)
d = ImageDraw.Draw(img)

# Basit dikey degrade
for y in range(H):
    t = y / H
    r = int(UST[0] * (1 - t) + ALT[0] * t)
    g = int(UST[1] * (1 - t) + ALT[1] * t)
    b = int(UST[2] * (1 - t) + ALT[2] * t)
    d.line([(0, y), (W, y)], fill=(r, g, b))


def font(size, bold=True):
    adaylar = ["segoeuib.ttf" if bold else "segoeui.ttf", "arialbd.ttf" if bold else "arial.ttf"]
    for ad in adaylar:
        try:
            return ImageFont.truetype(ad, size)
        except OSError:
            continue
    return ImageFont.load_default()


def ortala(metin, f, y, renk):
    kutu = d.textbbox((0, 0), metin, font=f)
    w = kutu[2] - kutu[0]
    d.text(((W - w) / 2, y), metin, font=f, fill=renk)


# Takvim benzeri basit bir ikon
ix, iy, isz = W // 2 - 26, 40, 52
d.rounded_rectangle([ix, iy, ix + isz, iy + isz], radius=7, fill=(255, 255, 255))
d.rectangle([ix, iy, ix + isz, iy + 15], fill=(230, 133, 112))  # ust serit (tatil rengi)
d.line([ix + 14, iy - 6, ix + 14, iy + 8], fill=(255, 255, 255), width=4)
d.line([ix + isz - 14, iy - 6, ix + isz - 14, iy + 8], fill=(255, 255, 255), width=4)
for gx in range(3):
    for gy in range(2):
        px = ix + 10 + gx * 16
        py = iy + 24 + gy * 12
        d.rectangle([px, py, px + 9, py + 7], fill=(196, 218, 238))

ortala("NÖBET PLANLAMA SİSTEMİ", font(26), 112, (255, 255, 255))
ortala("Vardiya ve Nöbet Yönetimi", font(13, bold=False), 150, (200, 214, 235))

img.save("splash.png")
print("splash.png yazildi:", img.size)
