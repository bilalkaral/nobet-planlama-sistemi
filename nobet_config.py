"""Supabase baglanti ayarlari.

Kendi Supabase projenin bilgilerini asagiya yaz (Project Settings > API).
Kurulum adimlari: KURULUM.md

anon anahtarini burada tutmak GUVENLIDIR. Bu anahtar zaten istemcide
calismak uzere tasarlanmistir; tek basina hicbir sey yapamaz, her istek
veritabanindaki guvenlik kurallarindan (RLS) gecer. Kullanici giris
yapmadan bu anahtarla hicbir veri okunamaz/yazilamaz.

service_role anahtarini ASLA buraya yazma. O anahtar tum RLS kurallarini
atlar; exe'ye gomulurse herkes tum veriye sinirsiz erisir ve rol sistemi
tamamen anlamsiz hale gelir.
"""
from __future__ import annotations

# Ornek: "https://abcdefghijklmnop.supabase.co"
SUPABASE_URL = "https://PROJE-REF.supabase.co"

# Project Settings > API > "anon public" anahtari (eyJhbGciOi... ile baslar)
SUPABASE_ANON_KEY = "BURAYA-ANON-KEY"

# Giris ekraninda kullanici sadece "ayse" yazar; arkada bu alan adiyla
# birlestirip "ayse@nobet.local" e-postasina cevrilir. Supabase panelinde
# kullanici olustururken de ayni alan adi kullanilmali.
EMAIL_DOMAIN = "nobet.local"

# L.G. devir kurali.
#   True  : Sadece EKSIK mesai devreder. Onceki ayi arti (fazla mesai) ile
#           kapatan kisinin L.G.'si 0 baslar - fazla mesai ayrica odendigi
#           icin sonraki aya tasinmaz.
#   False : Hem eksik hem fazla devreder (matematiksel bakiye).
SADECE_EKSIK_DEVRET = True

# Degisiklikleri kac saniyede bir kontrol edelim. Websocket yerine yoklama
# kullaniyoruz (hastane aglari websocket'i sik sik kapatir). Her yoklamada
# tum tablo degil, sadece son bakistan sonra degisen satirlar cekilir.
YOKLAMA_SANIYE = 5
