# Nöbet Planlama Sistemi

Hemşire/personel nöbet çizelgesi hazırlamak için çok kullanıcılı masaüstü
uygulaması. Tkinter + [tksheet](https://github.com/ragardner/tksheet)
arayüzü, veritabanı olarak [Supabase](https://supabase.com) (PostgreSQL + RLS).

## Özellikler

- **Nöbet çizelgesi** — kişi × gün tablosu, kod yazınca otomatik renklenir
  ve saat hesaplanır (Toplam, Fark, L.G./L.Ç.)
- **Rol bazlı yetki** — `bashemsire` (her şey) › `sorumlu` (kendi bölümü) ›
  `personel` (salt okuma); kurallar veritabanında RLS ile zorunlu
- **Canlı senkron** — birden fazla bilgisayar aynı anda çalışabilir,
  değişiklikler birkaç saniyede diğerlerine yansır
- **Arka planda kayıt** — yavaş internette arayüz donmaz
- **L.G. devri** — önceki ayın eksik mesaisi yeni aya taşınır
- **Saat aralığı girişi** — `Ç 10.00-13.30` (saatlik çalışma),
  `İ 10.00-13.30` (saatlik izin)
- **Özel kodlar** — `@D` (farklı bölümde çalışma), `İN=6` (kod=saat)
- **Hafta sonu / resmi tatil** renklendirmesi, resmi tatilde çalışma raporu
- **Personel raporu** — çoklu kişi seçimi, Excel'e aktarma
- **Geçmiş ay kilidi** — sorumlu, geçmiş ayı başhemşire onayıyla açabilir
- İşlem kayıtları (log) veritabanında tetikleyicilerle tutulur

## Dosyalar

| Dosya | Açıklama |
|---|---|
| `nobet_app.py` | Arayüz (giriş ekranı, sekmeler) — **başlangıç noktası** |
| `nobet_core.py` | Saat hesabı, kod ayrıştırma, Excel çıktısı |
| `nobet_db.py` | httpx tabanlı Supabase istemcisi (Auth + PostgREST) |
| `nobet_sync.py` | Veritabanı ↔ uygulama veri dönüşümü, senkron |
| `nobet_config.py` | **Supabase bağlantı ayarları** (doldurulmalı) |
| `schema.sql` | Tablolar, RLS kuralları, tetikleyiciler, varsayılan kodlar |
| `NobetPlanlamaSistemi.spec` | PyInstaller exe tanımı |
| `make_splash.py` | Açılış görseli (`splash.png`) üretir |
| `dagitim/kurulum.bat` | Exe'yi hedef bilgisayara kurar, kısayol açar |

## Kurulum

### 1. Veritabanı

Adım adım rehber: **[KURULUM.md](KURULUM.md)**

Kısaca: Supabase projesi aç → `schema.sql` çalıştır → dışarıdan kaydı kapat →
kullanıcıları oluştur → rolleri ata → personeli uygulamadan ekle.

### 2. Bağlantı ayarları

`nobet_config.py` içine kendi projenin bilgilerini yaz:

```python
SUPABASE_URL = "https://PROJE-REF.supabase.co"
SUPABASE_ANON_KEY = "eyJhbGciOi..."
```

> ⚠️ Buraya **yalnızca `anon` anahtarını** yaz, `service_role` anahtarını asla.

### 3. Çalıştırma

Python 3.11 önerilir.

```bash
pip install -r requirements.txt
python nobet_app.py
```

Giriş ekranında kullanıcı adı yazılır (`bashemsire`), uygulama arkada
`bashemsire@nobet.local` e-postasına çevirir.

### 4. Exe derleme (isteğe bağlı)

```bash
pip install pyinstaller pillow
python make_splash.py
python -m PyInstaller NobetPlanlamaSistemi.spec --noconfirm
```

Çıktı: `dist/NobetPlanlamaSistemi.exe` (tek dosya, ~16 MB). Hedef bilgisayara
`dagitim/kurulum.bat` ile birlikte gönderilir.

> Derleme ağırdır; bilgisayar zorlanıyorsa işlemi düşük öncelikte çalıştırın.
> `numpy`/`PIL` spec'te bilerek hariç tutulmuştur — openpyxl bunları isteğe
> bağlı kullanır, dahil edilirse derleme çok şişer.

## Kişisel veri uyarısı

Bu depo **boş şablondur**; gerçek personel adları, çizelgeler veya
bağlantı anahtarları içermez. Kendi verilerinizi (`.xlsx` dosyaları,
personel listeleri vb.) herkese açık bir depoya göndermeyin.
