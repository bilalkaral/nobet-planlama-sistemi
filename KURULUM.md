# Nöbet Sistemi — Veritabanı Kurulumu

Bu adımları **bir kere** yapman yeterli. Yaklaşık 20 dakika sürer.
DB bilmene gerek yok; hazır SQL'i kopyala-yapıştır yapacaksın.

---

## 1) Supabase projesi oluştur

1. https://supabase.com → **Start your project** → ücretsiz hesap (kredi kartı istemez).
2. **New project**:
   - **Name:** `nobet-sistemi`
   - **Database Password:** güçlü bir şifre belirle → **bir yere kaydet**, sonra lazım olabilir.
   - **Region:** `Central EU (Frankfurt)` — Türkiye'ye en yakını, en hızlısı bu.
3. **Create new project** → hazırlanması 1-2 dakika sürer.

---

## 2) Tabloları kur

1. Sol menü → **SQL Editor** → **New query**.
2. Bu klasördeki **`schema.sql`** dosyasını aç, **tamamını** kopyala, editöre yapıştır.
3. **Run** (veya Ctrl+Enter).
4. Altta **"Success. No rows returned"** görmelisin.

> Yanlışlıkla iki kere çalıştırırsan bir şey bozulmaz, dosya buna göre yazıldı.

---

## 3) Dışarıdan kayıt olmayı kapat

Bu **önemli** — yoksa internetten herkes hesap açıp çizelgeyi görebilir.

Sol menü → **Authentication** → **Sign In / Providers** → **Email** →
**"Allow new users to sign up"** seçeneğini **KAPAT** → **Save**.

Böylece sadece senin elle oluşturduğun hesaplar girebilir.

---

## 4) Kullanıcıları oluştur

Sol menü → **Authentication** → **Users** → **Add user** →
**Create new user**. Her kişi için sadece üç şey:

| Alan | Ne yazacaksın |
|---|---|
| **Email** | `kullaniciadi@nobet.local` (örn. `ayse@nobet.local`) |
| **Password** | O kişiye vereceğin şifre |
| **Auto Confirm User** | ✅ **İşaretle** — yoksa giriş yapamaz |

> **User Metadata alanını boş bırak.** Rolleri sonraki adımda SQL ile vereceğiz.

**Kimlerin hesabı olacak?** Sadece **giriş yapacak** kişilerin: başhemşire
ve her bölümün sorumlusu. Çizelgedeki personelin hesabı olması gerekmez —
onlar `personeller` tablosuna gider.

> İsteğe bağlı: personel kendi nöbetini görebilsin istersen onlara da
> `personel` rolüyle (salt okuma) hesap açabilirsin.

---

## 5) Rolleri ata

Yeni açılan her hesap varsayılan olarak `personel` (salt okuma) rolündedir.
Hesapları açtıktan sonra **SQL Editor** → **New query** → aşağıdakini kendi
kullanıcı adlarına ve bölümlerine göre düzenleyip → **Run**:

```sql
-- Bölümleri oluştur
insert into public.bolumler (ad) values
  ('Bölüm A'),
  ('Bölüm B')
on conflict (ad) do nothing;

-- Başhemşire
update public.profiller
   set rol = 'bashemsire', bolum_id = null
 where kullanici_adi = 'bashemsire';

-- Sorumlular: (kullanıcı adı, bölüm adı)
update public.profiller p
   set rol = 'sorumlu', bolum_id = b.id
  from (values
    ('bolum_a', 'Bölüm A'),
    ('bolum_b', 'Bölüm B')
  ) as m(kadi, bolum_ad)
  join public.bolumler b on b.ad = m.bolum_ad
 where p.kullanici_adi = m.kadi;

-- Kontrol
select p.kullanici_adi, p.rol, coalesce(b.ad, '(yok)') as bolum
  from public.profiller p
  left join public.bolumler b on b.id = p.bolum_id
 order by p.rol, p.kullanici_adi;
```

Kontrol tablosunda `rol` = `personel` kalan sorumlu varsa e-postası yanlış
yazılmıştır; panelden düzeltip sorguyu tekrar çalıştır.

> **Bölüm adının bir harfi bile farklı** olursa (`3. Kat` / `3. KAT`)
> veritabanı bunları iki ayrı bölüm sayar; sorumlu kendi personelini göremez.

Sonraki bölümler ve **personel** uygulamanın içinden, başhemşire hesabıyla
**Personeller** sekmesinden eklenir.

> Roller sadece şu üçü olabilir: **`bashemsire`**, **`sorumlu`**, **`personel`**.
> **En az bir kişi `bashemsire` olmalı** — yoksa sistemi kimse yönetemez.
> Giriş ekranında `@nobet.local` yazılmaz; sadece `ayse` yazılır,
> uygulama arkada e-postaya çevirir.

---

## 6) Bağlantı bilgilerini uygulamaya yaz

Sol menü → **Project Settings** (⚙️) → **API**. Şu ikisini kopyala:

- **Project URL** → örn. `https://xxxxxxxx.supabase.co`
- **anon public** anahtarı → `eyJhbGciOi...` diye başlayan uzun yazı

Bunları **`nobet_config.py`** içindeki `SUPABASE_URL` ve `SUPABASE_ANON_KEY`
alanlarına yaz.

> `anon` anahtarı uygulamada durabilir — istemcide çalışması için
> tasarlanmıştır, asıl güvenlik veritabanı kurallarıyla (RLS) sağlanır.
> **`service_role`** anahtarını ise **asla** hiçbir yere yazma —
> o anahtar tüm güvenlik kurallarını atlar.

---

## Roller ne yapabilir?

| | Başhemşire | Sorumlu | Personel |
|---|---|---|---|
| Çizelgeyi görmek | ✅ (tümü) | ✅ (kendi bölümü) | ✅ |
| Kendi bölümünün nöbetini değiştirmek | ✅ | ✅ | ❌ |
| Başka bölümün nöbetini değiştirmek | ✅ | ❌ | ❌ |
| Personel / bölüm ekleme | ✅ | ❌ | ❌ |
| Kod saatlerini değiştirmek | ✅ | ❌ | ❌ |
| Ay ayarları (hedef saat, resmi tatil) | ✅ | ❌ | ❌ |
| Geçmiş ayları düzenlemek | ✅ | 🔒 başhemşire onayıyla | ❌ |

Bu kurallar **veritabanında** (RLS) zorunlu tutuluyor. Yani biri exe'yi
kurcalasa bile başka bölümün nöbetini değiştiremez — sunucu reddeder.
(Geçmiş ay kilidi uygulama tarafında bir iş akışı kontrolüdür.)
