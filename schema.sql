-- ============================================================
--  NÖBET PLANLAMA SİSTEMİ  —  Supabase Veritabanı Kurulumu
-- ============================================================
--  Supabase panelinde:  SQL Editor -> New query  açıp bu dosyanın
--  TAMAMINI yapıştır ve "Run" bas. Bir kere çalıştırman yeterli.
--  Tekrar çalıştırırsan bir şey bozulmaz (her şey "if not exists").
--
--  ROLLER
--    bashemsire : Herkesi görür, herkesi düzenler. Personel ekler/çıkarır,
--                 kod tanımlarını ve ay ayarlarını değiştirir, logları okur.
--    sorumlu    : SADECE kendi bölümündeki personeli düzenler.
--                 Diğer bölümleri görür ama değiştiremez.
--    personel   : Her şeyi görür, hiçbir şeyi değiştiremez (salt okuma).
--
--  Roller sadece arayüzde değil, veritabanı kurallarıyla (RLS) korunur.
--  Yani biri exe'yi kurcalasa bile başka bölümün nöbetini değiştiremez.
-- ============================================================


-- ============================================================
--  1) TABLOLAR
-- ============================================================

-- ---------- Bölümler ----------
create table if not exists public.bolumler (
  id         uuid primary key default gen_random_uuid(),
  ad         text unique not null,
  created_at timestamptz not null default now()
);

-- ---------- Kullanıcı profilleri (rol bilgisi burada) ----------
create table if not exists public.profiller (
  id            uuid primary key references auth.users on delete cascade,
  kullanici_adi text unique not null,
  ad_soyad      text,
  rol           text not null default 'personel'
                check (rol in ('bashemsire', 'sorumlu', 'personel')),
  -- Sorumlu rolü için: hangi bölümden sorumlu olduğu.
  bolum_id      uuid references public.bolumler(id) on delete set null,
  created_at    timestamptz not null default now()
);

-- ---------- Personeller (nöbete girenler) ----------
-- NOT: Bu tablo "kullanıcı" değildir. Nöbet çizelgesindeki kişilerdir.
-- Bir personelin sisteme girme hesabı olmak zorunda değil.
create table if not exists public.personeller (
  id         uuid primary key default gen_random_uuid(),
  sicil_no   text,
  ad_soyad   text not null,
  bolum_id   uuid references public.bolumler(id) on delete set null,
  unvan      text,
  aktif      boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists personeller_bolum_idx on public.personeller (bolum_id);

-- ---------- Kod tanımları (D=10 saat, N=14 saat ...) ----------
create table if not exists public.kod_tanimlari (
  kod        text primary key,
  aciklama   text,
  saat       numeric not null default 0,
  sira       int not null default 0,
  updated_at timestamptz not null default now()
);

-- ---------- Aylar (her ay bir satır) ----------
create table if not exists public.aylar (
  id             uuid primary key default gen_random_uuid(),
  yil            int not null,
  ay_no          int not null check (ay_no between 1 and 12),
  hedef_saat     numeric not null default 210,
  resmi_tatiller int[] not null default '{}',
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now(),
  unique (yil, ay_no)
);

-- ---------- Ay + Personel satırı (çizelgedeki bir satır) ----------
-- L.G. (geçen aydan devreden) burada tutulur.
create table if not exists public.ay_personel (
  id          uuid primary key default gen_random_uuid(),
  ay_id       uuid not null references public.aylar(id) on delete cascade,
  personel_id uuid not null references public.personeller(id) on delete cascade,
  lg          numeric not null default 0,
  updated_at  timestamptz not null default now(),
  unique (ay_id, personel_id)
);

create index if not exists ay_personel_ay_idx on public.ay_personel (ay_id);

-- ---------- Gün hücreleri (çizelgenin asıl verisi) ----------
-- Her hücre AYRI bir satır. Böylece iki sorumlu aynı anda farklı kişileri
-- düzenlerken birbirinin verisini ezmez.
--
-- ÖNEMLİ: Bir hücre temizlenirken satır SİLİNMEZ, kod = '' yapılır.
-- Sebebi: istemci "updated_at > son_bakılan" diye sadece değişenleri çekiyor;
-- satır silinseydi bu sorgu silinmeyi göremez, hücre diğer PC'lerde dolu kalırdı.
create table if not exists public.nobet_gunleri (
  id             uuid primary key default gen_random_uuid(),
  ay_personel_id uuid not null references public.ay_personel(id) on delete cascade,
  gun            int not null check (gun between 1 and 31),
  kod            text not null default '',
  aciklama       text,
  updated_at     timestamptz not null default now(),
  updated_by     uuid references public.profiller(id) on delete set null,
  unique (ay_personel_id, gun)
);

-- Yoklama sorgusunun (updated_at > X) hızlı olması için:
create index if not exists nobet_gunleri_updated_idx on public.nobet_gunleri (updated_at desc);
create index if not exists nobet_gunleri_ap_idx on public.nobet_gunleri (ay_personel_id);

-- ---------- İşlem kayıtları (log) ----------
create table if not exists public.islem_kayitlari (
  id            bigserial primary key,
  kullanici_id  uuid references public.profiller(id) on delete set null,
  kullanici_adi text,
  tablo         text not null,
  islem         text not null,          -- INSERT / UPDATE / DELETE
  kayit_id      uuid,
  aciklama      text,                   -- insan okuyabilsin diye özet
  eski_deger    jsonb,
  yeni_deger    jsonb,
  created_at    timestamptz not null default now()
);

create index if not exists islem_kayitlari_created_idx on public.islem_kayitlari (created_at desc);


-- ============================================================
--  2) ROL / YETKİ FONKSİYONLARI
-- ============================================================
--  Hepsi "security definer": RLS'i atlayarak profiller tablosuna bakarlar.
--  Böylece "profiller'i okumak için profiller'e bakmak" sonsuz döngüsü olmaz.

create or replace function public.aktif_rol()
returns text language sql security definer stable
set search_path = public as $$
  select rol from public.profiller where id = auth.uid();
$$;

create or replace function public.is_bashemsire()
returns boolean language sql security definer stable
set search_path = public as $$
  select coalesce(public.aktif_rol() = 'bashemsire', false);
$$;

create or replace function public.benim_bolumum()
returns uuid language sql security definer stable
set search_path = public as $$
  select bolum_id from public.profiller where id = auth.uid();
$$;

-- Bir bölüm üzerinde yazma yetkim var mı?
--   başhemşire -> her bölümde
--   sorumlu    -> sadece kendi bölümünde
--   personel   -> hiçbir yerde
create or replace function public.bolum_yetkim_var(p_bolum_id uuid)
returns boolean language sql security definer stable
set search_path = public as $$
  select public.is_bashemsire()
      or (public.aktif_rol() = 'sorumlu'
          and p_bolum_id is not null
          and p_bolum_id = public.benim_bolumum());
$$;

-- Bir personel üzerinde yazma yetkim var mı?
create or replace function public.personel_yetkim_var(p_personel_id uuid)
returns boolean language sql security definer stable
set search_path = public as $$
  select public.is_bashemsire()
      or exists (
        select 1 from public.personeller p
        where p.id = p_personel_id
          and public.bolum_yetkim_var(p.bolum_id)
      );
$$;

-- Bir gün hücresi üzerinde yazma yetkim var mı?
create or replace function public.gun_yetkim_var(p_ay_personel_id uuid)
returns boolean language sql security definer stable
set search_path = public as $$
  select public.is_bashemsire()
      or exists (
        select 1 from public.ay_personel ap
        where ap.id = p_ay_personel_id
          and public.personel_yetkim_var(ap.personel_id)
      );
$$;


-- ============================================================
--  3) TETİKLEYİCİLER (otomatik çalışan işler)
-- ============================================================

-- ---------- Yeni kullanıcı -> otomatik profil ----------
-- Supabase panelinden kullanıcı eklerken User Metadata'ya şunu yazarsın:
--   { "kullanici_adi": "ayse", "ad_soyad": "Ayşe Yılmaz",
--     "rol": "sorumlu", "bolum": "Acil" }
-- "bolum" ADIYLA yazılır; aşağıdaki fonksiyon id'sini kendi bulur.
create or replace function public.yeni_kullanici_profili()
returns trigger language plpgsql security definer
set search_path = public as $$
declare
  v_bolum_id uuid;
  v_bolum_ad text;
begin
  v_bolum_ad := new.raw_user_meta_data ->> 'bolum';
  if v_bolum_ad is not null and v_bolum_ad <> '' then
    -- Bölüm yoksa oluştur ki kullanıcı eklerken sıra takılmasın.
    insert into public.bolumler (ad) values (v_bolum_ad)
      on conflict (ad) do nothing;
    select id into v_bolum_id from public.bolumler where ad = v_bolum_ad;
  end if;

  insert into public.profiller (id, kullanici_adi, ad_soyad, rol, bolum_id)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'kullanici_adi', split_part(new.email, '@', 1)),
    new.raw_user_meta_data ->> 'ad_soyad',
    coalesce(new.raw_user_meta_data ->> 'rol', 'personel'),
    v_bolum_id
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.yeni_kullanici_profili();

-- ---------- updated_at otomatik güncelle ----------
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists personeller_touch    on public.personeller;
drop trigger if exists kod_tanimlari_touch  on public.kod_tanimlari;
drop trigger if exists aylar_touch          on public.aylar;
drop trigger if exists ay_personel_touch    on public.ay_personel;

create trigger personeller_touch   before update on public.personeller
  for each row execute function public.touch_updated_at();
create trigger kod_tanimlari_touch before update on public.kod_tanimlari
  for each row execute function public.touch_updated_at();
create trigger aylar_touch         before update on public.aylar
  for each row execute function public.touch_updated_at();
create trigger ay_personel_touch   before update on public.ay_personel
  for each row execute function public.touch_updated_at();

-- ---------- Nöbet hücresi: "kim, ne zaman" damgasını SUNUCU koyar ----------
-- updated_by istemciden alınsaydı, biri exe'yi kurcalayıp değişikliği
-- başkasının üstüne yazabilirdi. Log kaydı bu yüzden buradan damgalanır.
create or replace function public.nobet_gunleri_damgala()
returns trigger language plpgsql security definer
set search_path = public as $$
begin
  new.updated_at = now();
  new.updated_by = auth.uid();
  return new;
end;
$$;

drop trigger if exists nobet_gunleri_touch on public.nobet_gunleri;
create trigger nobet_gunleri_touch
  before insert or update on public.nobet_gunleri
  for each row execute function public.nobet_gunleri_damgala();

-- ---------- Nöbet hücresi değişikliğini logla (okunur özetle) ----------
create or replace function public.logla_nobet_gunu()
returns trigger language plpgsql security definer
set search_path = public as $$
declare
  v_ad     text;
  v_yil    int;
  v_ay     int;
  v_eski   text;
  v_yeni   text;
  v_kadi   text;
  v_ap_id  uuid;
  v_id     uuid;
  v_gun    int;
  v_eski_j jsonb;
  v_yeni_j jsonb;
begin
  -- DELETE'te NEW, INSERT'te OLD boştur; alanlarına doğrudan dokunulmaz.
  if tg_op = 'DELETE' then
    v_ap_id := old.ay_personel_id; v_id := old.id; v_gun := old.gun;
    v_eski  := nullif(coalesce(old.kod, ''), ''); v_yeni := null;
    v_eski_j := to_jsonb(old); v_yeni_j := null;
  elsif tg_op = 'INSERT' then
    v_ap_id := new.ay_personel_id; v_id := new.id; v_gun := new.gun;
    v_eski  := null; v_yeni := nullif(coalesce(new.kod, ''), '');
    v_eski_j := null; v_yeni_j := to_jsonb(new);
  else
    v_ap_id := new.ay_personel_id; v_id := new.id; v_gun := new.gun;
    v_eski  := nullif(coalesce(old.kod, ''), '');
    v_yeni  := nullif(coalesce(new.kod, ''), '');
    v_eski_j := to_jsonb(old); v_yeni_j := to_jsonb(new);
    -- Gerçekten bir değişiklik yoksa log şişirme.
    if v_eski is not distinct from v_yeni
       and old.aciklama is not distinct from new.aciklama then
      return new;
    end if;
  end if;

  select p.ad_soyad, a.yil, a.ay_no
    into v_ad, v_yil, v_ay
    from public.ay_personel ap
    join public.personeller p on p.id = ap.personel_id
    join public.aylar a       on a.id = ap.ay_id
   where ap.id = v_ap_id;

  select kullanici_adi into v_kadi from public.profiller where id = auth.uid();

  insert into public.islem_kayitlari
    (kullanici_id, kullanici_adi, tablo, islem, kayit_id, aciklama, eski_deger, yeni_deger)
  values (
    auth.uid(), coalesce(v_kadi, '?'), 'nobet_gunleri', tg_op, v_id,
    format('%s — %s/%s %s. gün: %s -> %s',
           coalesce(v_ad, '?'), v_ay, v_yil, v_gun,
           coalesce(v_eski, 'boş'), coalesce(v_yeni, 'boş')),
    v_eski_j, v_yeni_j
  );
  return coalesce(new, old);
end;
$$;

drop trigger if exists nobet_gunleri_log on public.nobet_gunleri;
create trigger nobet_gunleri_log
  after insert or update or delete on public.nobet_gunleri
  for each row execute function public.logla_nobet_gunu();

-- ---------- Personel ekleme/çıkarmayı logla ----------
create or replace function public.logla_personel()
returns trigger language plpgsql security definer
set search_path = public as $$
declare
  v_kadi   text;
  v_ozet   text;
  v_id     uuid;
  v_eski_j jsonb;
  v_yeni_j jsonb;
begin
  select kullanici_adi into v_kadi from public.profiller where id = auth.uid();

  -- DELETE'te NEW, INSERT'te OLD boştur; alanlarına doğrudan dokunulmaz.
  if tg_op = 'DELETE' then
    v_id := old.id; v_ozet := format('Personel silindi: %s', old.ad_soyad);
    v_eski_j := to_jsonb(old); v_yeni_j := null;
  elsif tg_op = 'INSERT' then
    v_id := new.id; v_ozet := format('Personel eklendi: %s', new.ad_soyad);
    v_eski_j := null; v_yeni_j := to_jsonb(new);
  else
    v_id := new.id; v_ozet := format('Personel güncellendi: %s', new.ad_soyad);
    v_eski_j := to_jsonb(old); v_yeni_j := to_jsonb(new);
  end if;

  insert into public.islem_kayitlari
    (kullanici_id, kullanici_adi, tablo, islem, kayit_id, aciklama, eski_deger, yeni_deger)
  values (
    auth.uid(), coalesce(v_kadi, '?'), 'personeller', tg_op, v_id, v_ozet,
    v_eski_j, v_yeni_j
  );
  return coalesce(new, old);
end;
$$;

drop trigger if exists personeller_log on public.personeller;
create trigger personeller_log
  after insert or update or delete on public.personeller
  for each row execute function public.logla_personel();


-- ============================================================
--  4) GÜVENLİK (Row Level Security)
-- ============================================================

alter table public.bolumler        enable row level security;
alter table public.profiller       enable row level security;
alter table public.personeller     enable row level security;
alter table public.kod_tanimlari   enable row level security;
alter table public.aylar           enable row level security;
alter table public.ay_personel     enable row level security;
alter table public.nobet_gunleri   enable row level security;
alter table public.islem_kayitlari enable row level security;

-- ---------- BÖLÜMLER: herkes okur, sadece başhemşire yönetir ----------
drop policy if exists "bolumler_select" on public.bolumler;
create policy "bolumler_select" on public.bolumler
  for select to authenticated using (true);

drop policy if exists "bolumler_write" on public.bolumler;
create policy "bolumler_write" on public.bolumler
  for all to authenticated
  using (public.is_bashemsire()) with check (public.is_bashemsire());

-- ---------- PROFİLLER: herkes okur, sadece başhemşire rol atar ----------
drop policy if exists "profiller_select" on public.profiller;
create policy "profiller_select" on public.profiller
  for select to authenticated using (true);

drop policy if exists "profiller_write" on public.profiller;
create policy "profiller_write" on public.profiller
  for all to authenticated
  using (public.is_bashemsire()) with check (public.is_bashemsire());

-- ---------- PERSONELLER ----------
-- Herkes tüm personeli GÖRÜR (başka bölümün çizelgesi de görünsün diye).
drop policy if exists "personeller_select" on public.personeller;
create policy "personeller_select" on public.personeller
  for select to authenticated using (true);

-- Ekleme: başhemşire her bölüme, sorumlu sadece kendi bölümüne.
drop policy if exists "personeller_insert" on public.personeller;
create policy "personeller_insert" on public.personeller
  for insert to authenticated
  with check (public.bolum_yetkim_var(bolum_id));

-- Güncelleme: hem mevcut hem yeni bölümde yetki şart.
-- (Yoksa sorumlu, personeli başka bölüme "kaçırabilirdi".)
drop policy if exists "personeller_update" on public.personeller;
create policy "personeller_update" on public.personeller
  for update to authenticated
  using (public.bolum_yetkim_var(bolum_id))
  with check (public.bolum_yetkim_var(bolum_id));

drop policy if exists "personeller_delete" on public.personeller;
create policy "personeller_delete" on public.personeller
  for delete to authenticated
  using (public.bolum_yetkim_var(bolum_id));

-- ---------- KOD TANIMLARI: herkes okur, sadece başhemşire değiştirir ----------
-- Saat karşılıkları herkesin hesabını etkilediği için sorumluya açık değil.
drop policy if exists "kod_tanimlari_select" on public.kod_tanimlari;
create policy "kod_tanimlari_select" on public.kod_tanimlari
  for select to authenticated using (true);

drop policy if exists "kod_tanimlari_write" on public.kod_tanimlari;
create policy "kod_tanimlari_write" on public.kod_tanimlari
  for all to authenticated
  using (public.is_bashemsire()) with check (public.is_bashemsire());

-- ---------- AYLAR: herkes okur, sadece başhemşire açar/ayarlar ----------
-- Hedef saat ve resmi tatiller tüm bölümleri ilgilendirir.
drop policy if exists "aylar_select" on public.aylar;
create policy "aylar_select" on public.aylar
  for select to authenticated using (true);

drop policy if exists "aylar_write" on public.aylar;
create policy "aylar_write" on public.aylar
  for all to authenticated
  using (public.is_bashemsire()) with check (public.is_bashemsire());

-- ---------- AY_PERSONEL: herkes okur, bölüm yetkisi olan yazar ----------
drop policy if exists "ay_personel_select" on public.ay_personel;
create policy "ay_personel_select" on public.ay_personel
  for select to authenticated using (true);

drop policy if exists "ay_personel_write" on public.ay_personel;
create policy "ay_personel_write" on public.ay_personel
  for all to authenticated
  using (public.personel_yetkim_var(personel_id))
  with check (public.personel_yetkim_var(personel_id));

-- ---------- NÖBET GÜNLERİ: herkes okur, bölüm yetkisi olan yazar ----------
drop policy if exists "nobet_gunleri_select" on public.nobet_gunleri;
create policy "nobet_gunleri_select" on public.nobet_gunleri
  for select to authenticated using (true);

drop policy if exists "nobet_gunleri_write" on public.nobet_gunleri;
create policy "nobet_gunleri_write" on public.nobet_gunleri
  for all to authenticated
  using (public.gun_yetkim_var(ay_personel_id))
  with check (public.gun_yetkim_var(ay_personel_id));

-- ---------- LOGLAR: sadece başhemşire okur ----------
-- Yazma yetkisi kimseye verilmez; kayıtları tetikleyiciler (security definer)
-- yazar, o yüzden insert politikası gerekmez.
drop policy if exists "islem_kayitlari_select" on public.islem_kayitlari;
create policy "islem_kayitlari_select" on public.islem_kayitlari
  for select to authenticated using (public.is_bashemsire());


-- ============================================================
--  5) BAŞLANGIÇ VERİSİ — Kod Tanımları
-- ============================================================
--  Uygulamadaki varsayılan kod listesi. Saatleri sonradan başhemşire
--  arayüzden değiştirebilir; burada sadece ilk doldurma yapılır.

insert into public.kod_tanimlari (kod, aciklama, saat, sira) values
  ('D',     'Gündüz',                     10,   1),
  ('N',     'Gece',                       14,   2),
  ('24',    '24 Saat Nöbet',              24,   3),
  ('Yİ',    'Yıllık İzin',                10,   4),
  ('R',     'Rapor',                      10,   5),
  ('RHS',   'Resmi Hafta Tatili',          0,   6),
  ('GÇ',    'Gebe Çalışan',              7.5,   7),
  ('Sİ',    'SüT İzni Kullanan',          10,   8),
  ('Eİ',    'Evlilik İzni',               10,   9),
  ('RA',    'Radyasyonlu Alan Çalışanı',   7,  10),
  ('DSİ',   'Doğum Sonrası İzin',         10,  11),
  ('DÖİ',   'Doğum Öncesi İzin',          10,  12),
  ('E',     'Eğitim İzni',                10,  13),
  ('İİ',    'İdari İzin',                 10,  14),
  ('N+1',   'Gece+1 Saat',                15,  15),
  ('N+2',   'Gece+2 Saat',                16,  16),
  ('İST',   'İstifa',                      0,  17),
  ('N+1,5', 'Gece+1,5',                 15.5,  18),
  ('N+3,5', 'Gece+3,5',                 17.5,  19),
  ('N+4',   'Gece+4 Saat',                18,  20),
  ('D+1',   'Gündüz+1 Saat',              11,  21),
  ('D+2',   'Gündüz+2 Saat',              12,  22),
  ('D+3',   'Gündüz+3 Saat',              13,  23),
  ('D+4',   'Gündüz+4 Saat',              14,  24),
  ('24+1',  '24+1 Saat',                  25,  25),
  ('Şİ',    'Şua İzni',                   10,  26),
  ('RAÖ',   'Radyasyonlu Alan 14:30-21:30', 7, 27),
  ('RAG',   'Radyasyonlu Alan Gece 21:30-07:30', 10, 28),
  ('İN',    'İcap Nöbeti (hücrede İN=SAAT ile özel saat girilir)', 4, 29)
on conflict (kod) do nothing;


-- ============================================================
--  KURULUM BİTTİ
-- ============================================================
--  Sırada: Authentication -> Users -> Add user ile kullanıcıları oluştur.
--  Detaylar için KURULUM.md dosyasına bak.
-- ============================================================
