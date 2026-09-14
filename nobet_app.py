"""Nöbet Planlama Sistemi - masaüstü uygulama (Tkinter + tksheet).

Çalıştırmak için: python nobet_app.py
"""
from __future__ import annotations

import calendar
import datetime
import os
import queue
import tempfile
import threading
import webbrowser
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox, simpledialog, ttk

from tksheet import Sheet

from nobet_core import (
    AY_ADLARI,
    CALISMA_SAAT_COLOR,
    CODE_COLORS,
    DAY_COUNT,
    FARKLI_BOLUM_COLOR,
    HOLIDAY_HEADER_COLOR,
    IZIN_SAAT_COLOR,
    WEEKEND_HEADER_COLOR,
    NobetData,
    cell_hours,
    export_personel_raporu_excel,
    is_calisma_kodu,
    normalize_code,
    parse_cell_code,
    parse_saat_araligi,
    tr_upper,
)
from nobet_db import AuthError, DbError, SupabaseClient
from nobet_config import YOKLAMA_SANIYE
import nobet_sync

APP_TITLE = "Nöbet Planlama Sistemi"

ROL_ETIKET = {
    "bashemsire": "Başhemşire",
    "sorumlu": "Sorumlu",
    "personel": "Personel (salt okuma)",
}

FARK_POS_BG = "9DC0F1"
FARK_POS_FG = "1F4E78"
FARK_NEG_BG = "F5A3A3"
FARK_NEG_FG = "9C0006"
DEFAULT_BG = "FFFFFF"
DEFAULT_FG = "000000"
# Tablo daima 31 gun sutunu tasir; ayin cekmedigi gunler (orn. Subat 30-31)
# bu renkle kapatilir ve salt-okunur yapilir. Diger gri tonlarindan (RHS,
# IST, hafta sonu) ayirt edilebilmesi icin belirgin sekilde koyu.
DISABLED_DAY_BG = "9E9E9E"
# Yetki disindaki (baska bolume ait) satirlarin ad/gorev/bolum sutunlari bu
# renkle isaretlenir; boylece sorumlu hangi satirlara dokunamayacagini gorur.
# Gun hucrelerine uygulanmaz, cunku oradaki kod renkleri ustune yaziliyor.
LOCKED_ROW_BG = "D8D8D8"
LOCKED_ROW_FG = "6B6B6B"

# tksheet varsayılan sağ-tık menüsü İngilizce geliyor; tüm sayfalarda
# aynı Türkçe etiketleri kullanmak için tek yerden yönetiyoruz.
TURKISH_SHEET_LABELS = dict(
    edit_cell_label="Hücreyi Düzenle",
    cut_label="Kes",
    copy_label="Kopyala",
    copy_plain_label="Metni Kopyala",
    paste_label="Yapıştır",
    delete_label="Sil",
    clear_contents_label="İçeriği Temizle",
    delete_rows_label="Satırları Sil",
    insert_rows_above_label="Üste Satır Ekle",
    insert_rows_below_label="Alta Satır Ekle",
    insert_row_label="Satır Ekle",
    delete_columns_label="Sütunları Sil",
    insert_columns_left_label="Sola Sütun Ekle",
    insert_columns_right_label="Sağa Sütun Ekle",
    insert_column_label="Sütun Ekle",
    select_all_label="Tümünü Seç",
    undo_label="Geri Al",
    redo_label="Yinele",
    sort_cells_label="Artan Sırala",
    sort_cells_reverse_label="Azalan Sırala",
    sort_row_label="Değerleri Artan Sırala",
    sort_row_reverse_label="Değerleri Azalan Sırala",
    sort_column_label="Değerleri Artan Sırala",
    sort_column_reverse_label="Değerleri Azalan Sırala",
    sort_rows_label="Satırları Artan Sırala",
    sort_rows_reverse_label="Satırları Azalan Sırala",
    sort_columns_label="Sütunları Artan Sırala",
    sort_columns_reverse_label="Sütunları Azalan Sırala",
)


def setup_sheet(sheet: Sheet) -> None:
    """Her tablo icin ortak kurulum: Turkce menu etiketleri + guvenli olmayan
    siralamalarin kapatilmasi. Tek noktadan yapiliyor ki yeni bir tablo
    eklendiginde korumasiz kalmasin."""
    sheet.set_options(redraw=False, **TURKISH_SHEET_LABELS)
    disable_unsafe_sorting(sheet)


def disable_unsafe_sorting(sheet: Sheet) -> None:
    """tksheet'in yerlesik siralamalarini kapatir.

    Bu siralamalar SATIR BUTUNLUGUNU BOZUYOR: 'Sütunları Sırala' yalnizca o
    sutunun degerlerini sirallayip diger sutunlari yerinde birakiyor, boylece
    isimler birimlerden/nobet kodlarindan kayiyor. Cizelgede bu, karismis
    kodlarin YANLIS kisilere ait olarak veritabanina yazilmasina yol acar.
    Guvenli siralama (tum satiri birlikte tasiyan) ayrica saglanir."""
    for isim in ("sort_cells", "sort_rows", "sort_columns"):
        try:
            sheet.disable_bindings(isim)
        except Exception:
            pass


def fmt_num(x) -> str:
    try:
        x = float(x)
    except (TypeError, ValueError):
        return str(x)
    if x == int(x):
        return str(int(x))
    return str(round(x, 2))


def dosya_zaman_damgasi() -> str:
    """Excel dosya adlarinin sonuna eklenen guncel tarih-saat.
    Ornek: 18.07.2026_17.45 (Windows dosya adinda gecerli karakterler)."""
    return datetime.datetime.now().strftime("%d.%m.%Y_%H.%M")


def compute_usage_rows(codes, days_values):
    """Verilen gun hucresi metinlerinden (bir kisinin ya da tum cizelgenin)
    kod bazinda kullanim adedi ve gercek saat katkisini hesaplar. Katki
    hesabinda hucreye ozel girilen saat (orn. 'İN=6') Kod_Tanimlari'ndaki
    sabit saatin onune gecer, boylece rapor gercek girilen degerleri yansitir.
    """
    codes_map = {c["kod"]: float(c.get("saat") or 0) for c in codes}
    counts: dict = {}
    hours_by_code: dict = {}
    # Saat araligi girisleri (Ç/İ 10.00-13.30) kod tablosunda olmadigi icin
    # ayri toplanip listenin sonuna kendi satirlari olarak eklenir.
    sa_counts = {"Ç": 0, "İ": 0}
    sa_hours = {"Ç": 0.0, "İ": 0.0}
    for raw in days_values:
        sa = parse_saat_araligi(raw)
        if sa is not None:
            tip, sure, _ = sa
            sa_counts[tip] += 1
            sa_hours[tip] += sure if tip == "Ç" else -sure
            continue
        kod, override, _ = parse_cell_code(raw)
        if not kod:
            continue
        counts[kod] = counts.get(kod, 0) + 1
        saat = override if override is not None else codes_map.get(kod, 0.0)
        hours_by_code[kod] = hours_by_code.get(kod, 0.0) + saat
    rows = []
    for c in codes:
        kod = c["kod"]
        adet = counts.get(kod, 0)
        katki = hours_by_code.get(kod, 0.0)
        rows.append([kod, c.get("aciklama", ""), fmt_num(c.get("saat") or 0), adet, fmt_num(katki)])
    if sa_counts["Ç"]:
        rows.append(["Ç (saatlik)", "Saatlik Çalışma (Ç ss.dd-ss.dd)", "-", sa_counts["Ç"], fmt_num(sa_hours["Ç"])])
    if sa_counts["İ"]:
        rows.append(["İ (saatlik)", "Saatlik İzin (İ ss.dd-ss.dd)", "-", sa_counts["İ"], fmt_num(sa_hours["İ"])])
    return rows, sum(hours_by_code.values()) + sa_hours["Ç"] + sa_hours["İ"]


class AyAyarlariDialog(tk.Toplevel):
    """Bir ayin hedef saatini ve resmi tatil gunlerini duzenler. Eski 'Ayarlar'
    sekmesinin ay bazina inmis hali - deger DB'deki aylar tablosuna yazilir."""

    def __init__(self, master, yil, ay, hedef, tatiller, gun_sayisi):
        super().__init__(master)
        self.title(f"Ay Ayarları — {AY_ADLARI[ay - 1]} {yil}")
        self.resizable(False, False)
        self.configure(padx=20, pady=16)
        self.sonuc = None
        self._gun_sayisi = gun_sayisi

        ttk.Label(self, text="Aylık Hedef Saat:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=5)
        self.hedef_var = tk.StringVar(value=str(hedef))
        ttk.Entry(self, textvariable=self.hedef_var, width=14).grid(row=0, column=1, sticky="w", pady=5)

        ttk.Label(self, text=f"Resmi Tatil Günleri\n(1-{gun_sayisi}, virgülle):", justify="right").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=5
        )
        self.tatil_var = tk.StringVar(value=", ".join(str(d) for d in (tatiller or [])))
        ttk.Entry(self, textvariable=self.tatil_var, width=24).grid(row=1, column=1, sticky="w", pady=5)

        self.hata = ttk.Label(self, text="", foreground=ENTRY_BORDER_ERROR)
        self.hata.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="İptal", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Kaydet", command=self._kaydet).pack(side="right")

        self.bind("<Return>", lambda e: self._kaydet())
        self.update_idletasks()
        self.geometry(f"+{master.winfo_rootx() + 140}+{master.winfo_rooty() + 120}")
        self.grab_set()

    def _kaydet(self):
        try:
            hedef = float(self.hedef_var.get().replace(",", "."))
            if hedef < 0:
                raise ValueError
        except ValueError:
            self.hata.config(text="Hedef saat geçerli bir sayı olmalı.")
            return
        tatiller = []
        for parca in self.tatil_var.get().replace(";", ",").split(","):
            parca = parca.strip()
            if parca.isdigit():
                g = int(parca)
                if 1 <= g <= self._gun_sayisi:
                    tatiller.append(g)
        self.sonuc = (hedef, sorted(set(tatiller)))
        self.grab_release()
        self.destroy()


class KodTanimlariTab(ttk.Frame):
    HEADERS = ["Kod", "Açıklama", "Saat"]

    def __init__(self, master, data: NobetData, db=None, on_saved=None):
        super().__init__(master)
        self.data = data
        self.db = db
        self.on_saved = on_saved or (lambda: None)
        # Kod saatleri herkesin hesabini etkiledigi icin sadece basemsire
        # degistirebilir.
        self.yetkili = db is None or db.is_bashemsire

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=6, pady=6)
        self.btn_add = ttk.Button(toolbar, text="+ Satır Ekle", command=self.add_row)
        self.btn_add.pack(side="left", padx=2)
        self.btn_del = ttk.Button(toolbar, text="Seçili Satırı Sil", command=self.delete_row)
        self.btn_del.pack(side="left", padx=2)
        self.btn_save = ttk.Button(toolbar, text="Kaydet", command=self._save_to_db)
        self.btn_save.pack(side="left", padx=8)
        if not self.yetkili:
            for b in (self.btn_add, self.btn_del, self.btn_save):
                b.state(["disabled"])
            ttk.Label(
                toolbar, text="🔒 Kod saatlerini yalnızca başhemşire değiştirebilir.",
                foreground="#B00020",
            ).pack(side="left", padx=12)
        else:
            ttk.Label(
                toolbar,
                text=("Her kodun kaç saat sayılacağını belirler. Değişiklikten sonra "
                      "'Kaydet' — herkesin hesabı bu saatlere göre yapılır. "
                      "Ayrıca çizelgede saatlik giriş yapılabilir: 'Ç 10.00-13.30' "
                      "çalışma saati EKLER, 'İ 10.00-13.30' izin saati DÜŞER."),
                foreground="#555", wraplength=560, justify="left",
            ).pack(side="left", padx=12)

        self.sheet = Sheet(self, headers=self.HEADERS, data=self._rows_from_data())
        self.sheet.enable_bindings()
        setup_sheet(self.sheet)
        self.sheet.pack(fill="both", expand=True, padx=6, pady=6)
        self.sheet.column_width(column=0, width=90)
        self.sheet.column_width(column=1, width=320)
        self.sheet.column_width(column=2, width=90)
        if not self.yetkili:
            self.sheet.readonly_columns(columns=[0, 1, 2], readonly=True)

    def _rows_from_data(self):
        return [[c["kod"], c["aciklama"], c["saat"]] for c in self.data.codes]

    def refresh(self):
        self.sheet.set_sheet_data(self._rows_from_data(), reset_col_positions=False)

    def add_row(self):
        self.sheet.insert_row(["", "", 0])

    def delete_row(self):
        sel = self.sheet.get_currently_selected()
        if sel:
            self.sheet.delete_row(sel.row)

    def _save_to_db(self):
        """Kod tablosunu DB'ye yazar (upsert). Silinen kodlar da DB'den kaldirilir."""
        codes = self.get_codes()
        if not codes:
            messagebox.showinfo(APP_TITLE, "En az bir kod olmalı.")
            return
        if self.db is None:
            self.data.codes = codes
            self.on_saved()
            return
        try:
            mevcut = {r["kod"] for r in self.db.select("kod_tanimlari", params={"select": "kod"})}
            yeni = {c["kod"] for c in codes}
            payload = [
                {"kod": c["kod"], "aciklama": c.get("aciklama") or "",
                 "saat": c.get("saat") or 0, "sira": i + 1}
                for i, c in enumerate(codes)
            ]
            self.db.insert("kod_tanimlari", payload, upsert=True, on_conflict="kod")
            for silinen in mevcut - yeni:
                self.db.delete("kod_tanimlari", {"kod": f"eq.{silinen}"})
        except DbError as exc:
            messagebox.showerror(APP_TITLE, f"Kaydedilemedi:\n{exc}")
            return
        messagebox.showinfo(APP_TITLE, "Kod tanımları kaydedildi.")
        self.on_saved()

    def get_codes(self):
        codes = []
        for row in self.sheet.get_sheet_data():
            if not row:
                continue
            kod = str(row[0]).strip() if row[0] is not None else ""
            if not kod:
                continue
            aciklama = row[1] if len(row) > 1 and row[1] is not None else ""
            try:
                saat = float(row[2]) if len(row) > 2 and row[2] not in (None, "") else 0.0
            except (TypeError, ValueError):
                saat = 0.0
            codes.append({"kod": kod, "aciklama": aciklama, "saat": saat})
        return codes


class PersonelDialog(tk.Toplevel):
    """Personel ekleme/duzenleme penceresi. Bolum bir listeden secilir; boylece
    elle yazimdan dogan 'ayni bolum iki kez' hatasi olusmaz."""

    def __init__(self, master, bolum_adlari, kayit=None, on_add_bolum=None):
        super().__init__(master)
        self.title("Personel Düzenle" if kayit else "Yeni Personel")
        self.resizable(False, False)
        self.configure(padx=20, pady=16)
        self.sonuc = None
        # Yeni bolum olusturma geri cagrisi (yalnizca basemsirede dolu gelir).
        self.on_add_bolum = on_add_bolum

        kayit = kayit or {}
        self.vars = {
            "sicil_no": tk.StringVar(value=str(kayit.get("sicil_no", "") or "")),
            "ad_soyad": tk.StringVar(value=str(kayit.get("ad_soyad", "") or "")),
            "unvan": tk.StringVar(value=str(kayit.get("unvan", "") or "")),
            "birim": tk.StringVar(value=str(kayit.get("birim", "") or "")),
        }
        self.aktif_var = tk.BooleanVar(value=str(kayit.get("aktif", "Aktif")) != "Pasif")

        satirlar = [("ad_soyad", "Ad Soyad *"), ("unvan", "Ünvan / Görev"), ("sicil_no", "Sicil No")]
        for i, (key, etiket) in enumerate(satirlar):
            ttk.Label(self, text=etiket + ":").grid(row=i, column=0, sticky="e", padx=(0, 8), pady=4)
            ttk.Entry(self, textvariable=self.vars[key], width=30).grid(row=i, column=1, sticky="w", pady=4)

        ttk.Label(self, text="Bölüm:").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        bolum_satiri = ttk.Frame(self)
        bolum_satiri.grid(row=3, column=1, sticky="w", pady=4)
        # Liste salt-okunur: elle yazim 'ayni bolum iki kez' hatasina yol acar
        # (bir zamanlar 'Bölüm A ' ile 'Bölüm A' ayri bolum olmustu).
        # Yeni bolum ancak asagidaki + dugmesiyle, kontrollu sekilde eklenir.
        self.bolum_cb = ttk.Combobox(
            bolum_satiri, textvariable=self.vars["birim"], values=[""] + list(bolum_adlari),
            state="readonly", width=26,
        )
        self.bolum_cb.pack(side="left")
        if self.on_add_bolum is not None:
            ttk.Button(bolum_satiri, text="＋", width=3, command=self._yeni_bolum).pack(
                side="left", padx=(4, 0)
            )

        ttk.Checkbutton(self, text="Aktif (nöbet çizelgesinde görünsün)", variable=self.aktif_var).grid(
            row=4, column=1, sticky="w", pady=(6, 0)
        )
        self.hata = ttk.Label(self, text="", foreground=ENTRY_BORDER_ERROR)
        self.hata.grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        btns = ttk.Frame(self)
        btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="İptal", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Kaydet", command=self._kaydet).pack(side="right")

        self.bind("<Return>", lambda e: self._kaydet())
        self.update_idletasks()
        x = master.winfo_rootx() + 120
        y = master.winfo_rooty() + 100
        self.geometry(f"+{x}+{y}")
        self.grab_set()

    def _yeni_bolum(self):
        """Yeni bolum olusturur ve listede secili hale getirir."""
        ad = simpledialog.askstring(
            "Yeni Bölüm", "Bölüm adı:", parent=self
        )
        if ad is None:
            return
        ad = ad.strip()
        if not ad:
            self.hata.config(text="Bölüm adı boş olamaz.")
            return
        try:
            eklendi = self.on_add_bolum(ad)
        except Exception as exc:
            self.hata.config(text=str(exc))
            return
        if not eklendi:
            return
        # Listeyi tazele ve yeni bolumu sec.
        mevcut = list(self.bolum_cb["values"])
        if eklendi not in mevcut:
            mevcut = [""] + sorted(x for x in mevcut + [eklendi] if x)
            self.bolum_cb["values"] = mevcut
        self.vars["birim"].set(eklendi)
        self.hata.config(text="")

    def _kaydet(self):
        ad = self.vars["ad_soyad"].get().strip()
        if not ad:
            self.hata.config(text="Ad Soyad boş olamaz.")
            return
        self.sonuc = {
            "sicil_no": self.vars["sicil_no"].get().strip(),
            "ad_soyad": ad,
            "unvan": self.vars["unvan"].get().strip(),
            "birim": self.vars["birim"].get().strip(),
            "aktif": "Aktif" if self.aktif_var.get() else "Pasif",
        }
        self.grab_release()
        self.destroy()


class PersonellerTab(ttk.Frame):
    # Ad Soyad en basta: listede aranan ilk bilgi o. (Sicil No cogu kayitta
    # bos oldugu icin bastayken sadece yer kapliyordu.)
    HEADERS = ["Ad Soyad", "Sicil No", "Birim", "Ünvan", "Aktif/Pasif"]
    AD_COL = 0  # Ad Soyad sutunu - secili kisiyi bulmak icin kullanilir

    def __init__(self, master, data: NobetData, on_export_to_schedule, db=None,
                 get_bolum_map=None, on_changed=None):
        super().__init__(master)
        self.data = data
        self.db = db
        self.get_bolum_map = get_bolum_map or (lambda: {})
        self.on_changed = on_changed or (lambda: None)
        self.on_export_to_schedule = on_export_to_schedule

        # Personel ekleme/cikarma sadece basemsirede. Sorumlu/personel yalniz gorur.
        self.yetkili = db is None or db.is_bashemsire

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=6, pady=6)
        self.btn_add = ttk.Button(toolbar, text="+ Personel Ekle", command=self.add_person)
        self.btn_add.pack(side="left", padx=2)
        self.btn_edit = ttk.Button(toolbar, text="Düzenle", command=self.edit_person)
        self.btn_edit.pack(side="left", padx=2)
        self.btn_del = ttk.Button(toolbar, text="Sil", command=self.delete_person)
        self.btn_del.pack(side="left", padx=2)
        if not self.yetkili:
            for b in (self.btn_add, self.btn_edit, self.btn_del):
                b.state(["disabled"])
            ttk.Label(
                toolbar, text="🔒 Personel ekleme/çıkarma yetkisi yalnızca başhemşiredir.",
                foreground="#B00020",
            ).pack(side="left", padx=12)
        else:
            ttk.Label(
                toolbar, text="Eklediğiniz personel Nöbet Çizelgesi'ne otomatik düşer.",
                foreground="#555",
            ).pack(side="left", padx=12)

        self.sheet = Sheet(self, headers=self.HEADERS, data=self._rows_from_data())
        self.sheet.enable_bindings()
        setup_sheet(self.sheet)
        self.sheet.pack(fill="both", expand=True, padx=6, pady=6)
        for col, width in enumerate((220, 80, 200, 160, 100)):
            self.sheet.column_width(column=col, width=width)
        # Tablo goruntu amaclidir; degisiklikler diyalog uzerinden yapilir ki
        # "kaydettim mi?" belirsizligi olmasin.
        self.sheet.readonly_columns(columns=list(range(len(self.HEADERS))), readonly=True)
        # Guvenli siralama: basliga tiklayinca TUM SATIR birlikte tasinir.
        # (tksheet'in kendi siralamasi kapali - satir butunlugunu bozuyordu.)
        self.sheet.extra_bindings("column_select", func=self._on_header_click)
        self._sort_key = None
        self._sort_reverse = False

    _SORT_ALANLARI = {0: "ad_soyad", 1: "sicil_no", 2: "birim", 3: "unvan", 4: "aktif"}

    def _on_header_click(self, event):
        """Sutun basligina tiklayinca o sutuna gore tum listeyi sirallar."""
        try:
            c = event["selected"].column
        except Exception:
            try:
                c = event.selected.column
            except Exception:
                return
        alan = self._SORT_ALANLARI.get(c)
        if alan is None:
            return
        # Ayni sutuna tekrar tiklanirsa yon degisir.
        if self._sort_key == alan:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_key, self._sort_reverse = alan, False
        self.data.personnel.sort(
            key=lambda p: tr_upper(str(p.get(alan, "") or "")), reverse=self._sort_reverse
        )
        self.refresh()

    def _rows_from_data(self):
        return [
            [p.get("ad_soyad", ""), p.get("sicil_no", ""), p.get("birim", ""), p.get("unvan", ""), p.get("aktif", "Aktif")]
            for p in self.data.personnel
        ]

    def refresh(self):
        self.sheet.set_sheet_data(self._rows_from_data(), reset_col_positions=False)
        ok = " ▼" if self._sort_reverse else " ▲"
        basliklar = [
            h + (ok if self._SORT_ALANLARI.get(i) == self._sort_key else "")
            for i, h in enumerate(self.HEADERS)
        ]
        self.sheet.headers(basliklar)

    # -- yardimcilar ----------------------------------------------------
    def _bolum_adlari(self):
        return sorted(self.get_bolum_map().values())

    def _bolum_id_of(self, ad):
        for bid, b_ad in self.get_bolum_map().items():
            if b_ad == ad:
                return bid
        return None

    def _yeni_bolum_olustur(self, ad):
        """Yeni bolumu veritabanina ekler; eklenen adi doner.

        Ayni bolumun iki kez olusmasi (orn. 'bölüm a' / 'Bölüm A')
        buradaki kontrolle engellenir - boyle bir kayma olursa sorumlu kendi
        personelini goremez hale gelir ve sebebi kolay anlasilmaz."""
        ad = " ".join(str(ad or "").split())  # bas/son bosluk + ic tekrarlari temizle
        if not ad:
            raise ValueError("Bölüm adı boş olamaz.")
        mevcut = self.get_bolum_map()
        for var_ad in mevcut.values():
            if tr_upper(var_ad) == tr_upper(ad):
                # Zaten var: yenisini yaratma, mevcut olani sec.
                return var_ad
        if self.db is None:
            return ad
        try:
            rows = self.db.insert("bolumler", [{"ad": ad}])
        except DbError as exc:
            raise ValueError(f"Bölüm eklenemedi: {exc}")
        if rows:
            # Canli haritayi guncelle ki diger ekranlar da yeni bolumu gorsun.
            mevcut[rows[0]["id"]] = rows[0].get("ad", ad)
        return ad

    def _selected_person(self):
        """Ekranda secili satirin personel kaydini doner.

        ONEMLI: Satir numarasi ile self.data.personnel[r] YAPILMAZ. Kullanici
        tabloyu siralayinca tksheet tablo verisini fiziksel olarak yeniden
        diziyor; o zaman ekrandaki 5. satir ile listedeki 5. kayit ayni kisi
        olmuyor ve yanlis kisi duzenlenip SILINEBILIYOR. Bu yuzden ekranda
        yazan Ad Soyad okunup kayit ona gore bulunur."""
        sel = self.sheet.get_currently_selected()
        if not sel:
            return None
        r = sel.row
        if r is None or not (0 <= r < self.sheet.get_total_rows()):
            return None
        ad = str(self.sheet.get_cell_data(r, self.AD_COL) or "").strip()
        if not ad:
            return None
        for p in self.data.personnel:
            if str(p.get("ad_soyad", "")).strip() == ad:
                return p
        return None

    # -- islemler -------------------------------------------------------
    def add_person(self):
        dlg = PersonelDialog(
            self, self._bolum_adlari(),
            on_add_bolum=self._yeni_bolum_olustur if self.yetkili else None,
        )
        self.wait_window(dlg)
        if not dlg.sonuc:
            return
        if self.db is None:
            self.data.personnel.append(dlg.sonuc)
            self.refresh()
            return
        try:
            nobet_sync.add_personel(
                self.db, dlg.sonuc["ad_soyad"], dlg.sonuc["unvan"],
                self._bolum_id_of(dlg.sonuc["birim"]), dlg.sonuc["sicil_no"],
            )
        except DbError as exc:
            messagebox.showerror(APP_TITLE, f"Personel eklenemedi:\n{exc}")
            return
        messagebox.showinfo(APP_TITLE, f"{dlg.sonuc['ad_soyad']} eklendi ve çizelgeye aktarıldı.")
        self.on_changed()

    def edit_person(self):
        kisi = self._selected_person()
        if not kisi:
            messagebox.showinfo(APP_TITLE, "Önce listeden bir personel seçin.")
            return
        dlg = PersonelDialog(
            self, self._bolum_adlari(), kayit=kisi,
            on_add_bolum=self._yeni_bolum_olustur if self.yetkili else None,
        )
        self.wait_window(dlg)
        if not dlg.sonuc:
            return
        if self.db is None or not kisi.get("_id"):
            kisi.update(dlg.sonuc)
            self.refresh()
            return
        try:
            self.db.update("personeller", {"id": f"eq.{kisi['_id']}"}, {
                "ad_soyad": dlg.sonuc["ad_soyad"],
                "unvan": dlg.sonuc["unvan"] or None,
                "sicil_no": dlg.sonuc["sicil_no"] or None,
                "bolum_id": self._bolum_id_of(dlg.sonuc["birim"]),
                "aktif": dlg.sonuc["aktif"] == "Aktif",
            })
        except DbError as exc:
            messagebox.showerror(APP_TITLE, f"Güncellenemedi:\n{exc}")
            return
        self.on_changed()

    def delete_person(self):
        kisi = self._selected_person()
        if not kisi:
            messagebox.showinfo(APP_TITLE, "Önce listeden bir personel seçin.")
            return
        ad = kisi.get("ad_soyad", "")
        if not messagebox.askyesno(
            APP_TITLE,
            f"{ad} silinsin mi?\n\nBu kişinin TÜM aylardaki nöbet kayıtları da silinir.\n"
            f"Sadece çizelgeden çıkarmak istiyorsanız 'Düzenle' ile Aktif işaretini kaldırın.",
        ):
            return
        if self.db is None or not kisi.get("_id"):
            self.data.personnel.remove(kisi)
            self.refresh()
            return
        try:
            nobet_sync.delete_personel(self.db, kisi["_id"])
        except DbError as exc:
            messagebox.showerror(APP_TITLE, f"Silinemedi:\n{exc}")
            return
        self.on_changed()

    def get_personnel(self):
        """Excel aktarimi icin - artik kaynak DB'den yuklenen listedir."""
        return list(self.data.personnel)

    def get_personnel_lookup(self):
        """Ad Soyad (büyük harf, kırpılmış) -> {gorev, bolum} sözlüğü."""
        lookup = {}
        for p in self.get_personnel():
            key = str(p.get("ad_soyad", "")).strip().upper()
            if key:
                lookup[key] = {"gorev": p.get("unvan", ""), "bolum": p.get("birim", "")}
        return lookup

    def _export_to_schedule(self):
        self.on_export_to_schedule(self.get_personnel())


class BashemsireOnayDialog(tk.Toplevel):
    """Gecmis ayi duzenlemeye acmak icin basemsire kimligini dogrular.

    Dogrulama AYRI bir istemciyle yapilir; boylece ekranda acik olan sorumlu
    oturumu bozulmaz. Sifre hicbir yere kaydedilmez, yalnizca dogrulama icin
    kullanilip atilir."""

    def __init__(self, master, ay_adi, yil):
        super().__init__(master)
        self.title("Başhemşire Onayı")
        self.resizable(False, False)
        self.configure(padx=24, pady=18)
        self.onaylandi = False

        ttk.Label(self, text="🔒 Geçmiş Ay Kilidi", font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            self,
            text=(f"{ay_adi} {yil} kapanmış bir aydır ve düzenlemeye kapalıdır.\n"
                  "Açmak için başhemşire kullanıcı adı ve şifresi gerekir."),
            foreground="#555", justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 14))

        ttk.Label(self, text="Kullanıcı adı:").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        self.user_var = tk.StringVar(value="bashemsire")
        self.user_entry = ttk.Entry(self, textvariable=self.user_var, width=26)
        self.user_entry.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(self, text="Şifre:").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        self.pass_var = tk.StringVar()
        self.pass_entry = ttk.Entry(self, textvariable=self.pass_var, width=26, show="•")
        self.pass_entry.grid(row=3, column=1, sticky="w", pady=4)

        self.hata = ttk.Label(self, text="", foreground=ENTRY_BORDER_ERROR, wraplength=300, justify="left")
        self.hata.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        btns = ttk.Frame(self)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="İptal", command=self.destroy).pack(side="right", padx=4)
        self.btn_ac = ttk.Button(btns, text="Kilidi Aç", command=self._dogrula)
        self.btn_ac.pack(side="right")

        self.bind("<Return>", lambda e: self._dogrula())
        self.update_idletasks()
        self.geometry(f"+{master.winfo_rootx() + 140}+{master.winfo_rooty() + 120}")
        self.pass_entry.focus_set()
        self.grab_set()

    def _dogrula(self):
        self.hata.config(text="")
        self.btn_ac.config(state="disabled", text="Kontrol ediliyor…")
        self.update_idletasks()
        gecici = SupabaseClient()
        try:
            gecici.sign_in(self.user_var.get(), self.pass_var.get())
            if not gecici.is_bashemsire:
                self.hata.config(text="Bu kullanıcı başhemşire değil; kilidi açamaz.")
                return
            self.onaylandi = True
            self.grab_release()
            self.destroy()
        except AuthError as exc:
            self.hata.config(text=str(exc))
        except Exception as exc:
            self.hata.config(text=f"Doğrulanamadı:\n{exc}")
        finally:
            gecici.close()
            if not self.onaylandi:
                try:
                    self.btn_ac.config(state="normal", text="Kilidi Aç")
                    self.pass_var.set("")
                    self.pass_entry.focus_set()
                except Exception:
                    pass


class PersonelSecDialog(tk.Toplevel):
    """Rapor icin coklu personel secme penceresi.

    Cizelgede Ctrl+tik ile secmek de mumkun, ama 148 satirda dogru satirlari
    bulmak zor; bu pencere arama kutusu ve toplu secim ile bunu kolaylastirir.
    """

    def __init__(self, master, kisiler):
        """kisiler: [(sheet_row_index, ad_soyad, gorev, bolum), ...]"""
        super().__init__(master)
        self.title("Rapor için Personel Seç")
        self.resizable(True, True)
        self.geometry("520x520")
        self.configure(padx=14, pady=12)
        self.kisiler = list(kisiler)
        self.sonuc = None

        ust = ttk.Frame(self)
        ust.pack(fill="x")
        ttk.Label(ust, text="Ara:").pack(side="left")
        self.arama = tk.StringVar()
        e = ttk.Entry(ust, textvariable=self.arama)
        e.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.arama.trace_add("write", lambda *a: self._doldur())

        ttk.Label(
            self, text="Ctrl ile tek tek, Shift ile aralık seçebilirsiniz.",
            foreground="#555",
        ).pack(anchor="w", pady=(6, 2))

        cerceve = ttk.Frame(self)
        cerceve.pack(fill="both", expand=True)
        self.liste = tk.Listbox(cerceve, selectmode="extended", activestyle="none")
        sb = ttk.Scrollbar(cerceve, orient="vertical", command=self.liste.yview)
        self.liste.configure(yscrollcommand=sb.set)
        self.liste.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.liste.bind("<Double-Button-1>", lambda e: self._tamam())

        self.sayac = ttk.Label(self, text="", foreground="#555")
        self.sayac.pack(anchor="w", pady=(6, 0))
        self.liste.bind("<<ListboxSelect>>", lambda e: self._sayaci_guncelle())

        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Tümünü Seç", command=self._tumunu_sec).pack(side="left")
        ttk.Button(btns, text="Temizle", command=lambda: self.liste.selection_clear(0, "end")).pack(
            side="left", padx=6
        )
        ttk.Button(btns, text="İptal", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Rapor Oluştur", command=self._tamam).pack(side="right", padx=6)

        self._doldur()
        e.focus_set()
        self.grab_set()

    def _doldur(self):
        filtre = (self.arama.get() or "").strip().upper()
        self.gosterilen = [
            k for k in self.kisiler
            if not filtre or filtre in f"{k[1]} {k[2]} {k[3]}".upper()
        ]
        self.liste.delete(0, "end")
        for _idx, ad, gorev, bolum in self.gosterilen:
            etiket = ad if not bolum else f"{ad}   —   {bolum}"
            self.liste.insert("end", etiket)
        self._sayaci_guncelle()

    def _tumunu_sec(self):
        self.liste.selection_set(0, "end")
        self._sayaci_guncelle()

    def _sayaci_guncelle(self):
        self.sayac.config(
            text=f"{len(self.liste.curselection())} kişi seçili  /  {len(self.gosterilen)} gösteriliyor"
        )

    def _tamam(self):
        secili = self.liste.curselection()
        if not secili:
            messagebox.showinfo(APP_TITLE, "En az bir personel seçin.", parent=self)
            return
        self.sonuc = [self.gosterilen[i][0] for i in secili]
        self.grab_release()
        self.destroy()


class NobetCizelgesiTab(ttk.Frame):
    STATIC_HEADERS = ["Ad Soyad", "Görev", "Bölüm"]
    LG_IDX = 3
    DAY_START_IDX = 4
    DAY_END_IDX = DAY_START_IDX + DAY_COUNT - 1
    TOPLAM_IDX = DAY_END_IDX + 1
    FARK_IDX = TOPLAM_IDX + 1
    LC_IDX = FARK_IDX + 1

    def __init__(self, master, data: NobetData, get_codes_map, get_hedef_saat,
                 get_personnel_lookup=None, on_change_month=None, on_data_synced=None,
                 on_generate_report=None, db=None, on_cell_saved=None,
                 on_edit_month_settings=None):
        super().__init__(master)
        self.data = data
        self.on_edit_month_settings = on_edit_month_settings
        self.db = db
        self.get_codes_map = get_codes_map
        self.get_hedef_saat = get_hedef_saat
        self.get_personnel_lookup = get_personnel_lookup or (lambda: {})
        self.on_change_month = on_change_month
        self.on_data_synced = on_data_synced or (lambda: None)
        self.on_generate_report = on_generate_report
        self.on_cell_saved = on_cell_saved or (lambda: None)
        # ad_soyad -> o kisiye ait schedule satiri (gizli _personel_id/_ap_id/
        # _bolum_id icerir). Yazma sirasinda hedef DB kaydini bulmak icin.
        self._meta_by_name = {}
        # Genel/ozel liste: sorumlu varsayilan olarak sadece kendi bolumunu
        # gorur (daha az satir = daha rahat ve hizli). Genel listeye gecebilir.
        # Sorumlu DAIMA yalnizca kendi bolumunu gorur (secenek yok). Basemsire
        # ise bolum/meslek ile suzebilir; bos = tumu.
        self._filter_bolum = tk.StringVar(value="")
        self._filter_meslek = tk.StringVar(value="")
        # Basemsire sifresiyle duzenlemeye acilan gecmis aylar: {(yil, ay)}.
        # Yalnizca bu oturum icin gecerli; program kapaninca kilit geri gelir.
        self._acilan_aylar = set()
        self._gecmis_kilit_aktif = False
        self._display_schedule = []  # o an ekranda gorunen satirlar (filtreli)

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=6, pady=4)
        ttk.Button(toolbar, text="◀ Önceki Ay", command=lambda: self._change_month(-1)).pack(side="left", padx=2)
        self.month_label = ttk.Label(toolbar, text="", font=("Segoe UI", 10, "bold"))
        self.month_label.pack(side="left", padx=6)
        ttk.Button(toolbar, text="Sonraki Ay ▶", command=lambda: self._change_month(1)).pack(side="left", padx=2)
        # Tam yenileme (canli senkron gun hucrelerini otomatik gunceller; bu
        # buton personel/ayar gibi diger degisiklikleri de bastan yukler).
        ttk.Button(toolbar, text="🔄 Yenile", command=lambda: self._change_month(0)).pack(side="left", padx=2)
        # Gecmis ay kilidi - yalnizca sorumlu/personel icin anlamli.
        self.btn_gecmis_kilit = ttk.Button(
            toolbar, text="🔒 Geçmiş Ay", command=self._toggle_gecmis_ay_kilidi
        )
        if self.db is not None and not self.db.is_bashemsire:
            self.btn_gecmis_kilit.pack(side="left", padx=(8, 2))
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
        # Ay Ayarlari (hedef saat + resmi tatil) yalnizca basemsirede. Eski
        # Ayarlar sekmesinin yerini tutar.
        if self.on_edit_month_settings is not None and (self.db is None or self.db.is_bashemsire):
            ttk.Button(toolbar, text="⚙ Ay Ayarları", command=self.on_edit_month_settings).pack(side="left", padx=2)
        if self.db is None:
            # Yerel (DB'siz) modda eski satir ekle/sil butonlari kalsin.
            self.btn_add_row = ttk.Button(toolbar, text="+ Personel Satırı Ekle", command=self.add_row)
            self.btn_add_row.pack(side="left", padx=2)
            self.btn_del_row = ttk.Button(toolbar, text="Seçili Satırı Sil", command=self.delete_row)
            self.btn_del_row.pack(side="left", padx=2)
        ttk.Button(toolbar, text="Tümünü Yeniden Hesapla", command=self.recalc_all).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Yazdır", command=self.print_preview).pack(side="left", padx=6)
        # Basemsire icin bolum + meslek filtresi. Sorumluya filtre gerekmez;
        # o zaten yalnizca kendi bolumunu gorur (asagida _visible_schedule).
        if self.db is not None and self.db.is_bashemsire:
            ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)
            ttk.Label(toolbar, text="Bölüm:").pack(side="left", padx=(2, 2))
            self.bolum_filter = ttk.Combobox(
                toolbar, textvariable=self._filter_bolum, state="readonly", width=22,
            )
            self.bolum_filter.pack(side="left", padx=2)
            self.bolum_filter.bind("<<ComboboxSelected>>", lambda e: self._on_filter_toggle())
            ttk.Label(toolbar, text="Meslek:").pack(side="left", padx=(8, 2))
            self.meslek_filter = ttk.Combobox(
                toolbar, textvariable=self._filter_meslek, state="readonly", width=18,
            )
            self.meslek_filter.pack(side="left", padx=2)
            self.meslek_filter.bind("<<ComboboxSelected>>", lambda e: self._on_filter_toggle())
            ttk.Button(toolbar, text="Temizle", width=8, command=self._clear_filters).pack(
                side="left", padx=(6, 2)
            )
        # Kayitlar arka planda gonderildigi icin kullaniciya durumu gosteriyoruz.
        self.save_status = tk.StringVar(value="")
        if self.db is not None:
            ttk.Label(toolbar, textvariable=self.save_status, foreground="#555").pack(
                side="right", padx=8
            )

        legend = ttk.Frame(self)
        legend.pack(fill="x", padx=6, pady=(0, 4))
        legend_items = [(kod, renk) for kod, renk in CODE_COLORS.items()]
        legend_items += [
            ("Ç Saatlik Çalışma", CALISMA_SAAT_COLOR),
            ("İ Saatlik İzin", IZIN_SAAT_COLOR),
            ("Farklı Bölüm (@KOD)", FARKLI_BOLUM_COLOR),
            ("Hafta Sonu", WEEKEND_HEADER_COLOR),
            ("Resmi Tatil", HOLIDAY_HEADER_COLOR),
            ("Ayda Yok", DISABLED_DAY_BG),
        ]
        legend_cols = 14
        for i, (kod, renk) in enumerate(legend_items):
            tk.Label(legend, text=f"  {kod}  ", bg="#" + renk, relief="groove", borderwidth=1).grid(
                row=i // legend_cols, column=i % legend_cols, padx=2, pady=1, sticky="w"
            )

        # -- seçili personelin detay paneli --
        detail = ttk.LabelFrame(self, text="Seçili Personel Detayı")
        detail.pack(fill="x", padx=6, pady=(0, 4))
        self.detail_vars = {k: tk.StringVar(value="-") for k in
                             ("ad", "bolum", "lg", "toplam", "hedef", "fark", "lc")}
        detail_fields = [
            ("ad", "Ad Soyad"), ("bolum", "Bölüm"), ("lg", "L.G."),
            ("toplam", "Toplam Saat"), ("hedef", "Hedef Saat"), ("fark", "Fark"), ("lc", "L.Ç."),
        ]
        for i, (key, label) in enumerate(detail_fields):
            ttk.Label(detail, text=label + ":").grid(row=0, column=2 * i, sticky="w", padx=(8, 2), pady=4)
            ttk.Label(detail, textvariable=self.detail_vars[key], font=("Segoe UI", 9, "bold")).grid(
                row=0, column=2 * i + 1, sticky="w", padx=(0, 8)
            )
        ttk.Button(detail, text="👥 Personel Seç", command=self._pick_people_for_report).grid(
            row=0, column=2 * len(detail_fields) + 1, sticky="w", padx=(4, 8)
        )
        ttk.Button(detail, text="📄 Rapor Oluştur", command=self._generate_report).grid(
            row=0, column=2 * len(detail_fields), sticky="w", padx=(12, 8)
        )
        # Yetki uyarisi: kilitli bir satir secilince burada belirir. Her tiklamada
        # acilir pencere gostermek rahatsiz edici olurdu, bu yuzden satir ici uyari.
        self.lock_hint = tk.StringVar(value="")
        self.lock_hint_label = ttk.Label(
            detail, textvariable=self.lock_hint, foreground="#B00020", font=("Segoe UI", 9, "bold")
        )
        self.lock_hint_label.grid(
            row=1, column=0, columnspan=2 * len(detail_fields) + 1, sticky="w", padx=8, pady=(0, 4)
        )
        self._locked_rows = set()

        headers = self.STATIC_HEADERS + ["L.G."] + [str(d) for d in range(1, DAY_COUNT + 1)] + ["Toplam Saat", "Fark", "L.Ç."]
        self.sheet = Sheet(
            self, headers=headers, data=self._rows_from_data(),
            note_corners=True, tooltip_width=170, tooltip_height=60, tooltip_hover_delay=400,
            table_grid_fg="#9A9A9A",
        )
        self.sheet.enable_bindings()
        # Coklu kisi raporu icin Ctrl/Shift ile satir secimi acikca acilir.
        try:
            self.sheet.enable_bindings("ctrl_select", "row_select", "ctrl_click_select")
        except Exception:
            pass
        setup_sheet(self.sheet)

        zoom_bar = ttk.Frame(self)
        zoom_bar.pack(fill="x", side="bottom", padx=6, pady=(0, 4))
        ttk.Button(zoom_bar, text="＋", width=3, command=self._zoom_in).pack(side="right", padx=2)
        self.zoom_label = ttk.Label(zoom_bar, text="100%", width=5, anchor="center")
        self.zoom_label.pack(side="right")
        ttk.Button(zoom_bar, text="－", width=3, command=self._zoom_out).pack(side="right", padx=2)
        ttk.Label(zoom_bar, text="Tablo Boyutu:").pack(side="right", padx=(0, 6))
        self.zoom_pct = 100

        self.sheet.pack(fill="both", expand=True, padx=6, pady=4)

        self.sheet.column_width(column=0, width=160)
        self.sheet.column_width(column=1, width=95)
        self.sheet.column_width(column=2, width=120)
        self.sheet.column_width(column=self.LG_IDX, width=48)
        for c in range(self.DAY_START_IDX, self.DAY_END_IDX + 1):
            self.sheet.column_width(column=c, width=26)
        for c in (self.TOPLAM_IDX, self.FARK_IDX, self.LC_IDX):
            self.sheet.column_width(column=c, width=64)

        self.sheet.readonly_columns(columns=[self.TOPLAM_IDX, self.FARK_IDX, self.LC_IDX], readonly=True)
        # DB'ye bagliyken ad/gorev/bolum buradan degil, Personeller sekmesinden
        # yonetilir; yanlislikla degistirilmesin diye salt-okunur. L.G. de
        # elle girilmez - onceki ayin L.C.'sinden otomatik devreder (turetilir).
        if self.db is not None:
            self.sheet.readonly_columns(columns=[0, 1, 2, self.LG_IDX], readonly=True)
        # DIKKAT: sadece "end_edit_cell" dinlemek yetmiyor - Delete tusu, sag-tik
        # "Icerigi Temizle", kes ve yapistir bu olayi tetiklemiyor. O yuzden
        # tum degisiklikleri kapsayan "all_modified_events" kullaniyoruz; aksi
        # halde silinen hucrenin rengi ekranda kaliyor ve DB'ye yazilmiyordu.
        self.sheet.extra_bindings("all_modified_events", func=self._on_modified)
        self.sheet.extra_bindings("cell_select", func=self._on_select)
        self.sheet.popup_menu_add_command(
            "Açıklama Ekle / Düzenle", self._edit_note_for_selected,
            table_menu=True, index_menu=False, header_menu=False, empty_space_menu=False,
        )

        self.after(50, self._full_refresh)

    # -- veri yardımcıları ------------------------------------------------
    def _visible_schedule(self):
        """Ekranda gosterilecek satirlar. Ayni satir nesneleri (kopya degil)
        donduruur - yazma/_meta tutarli kalir.

        Sorumlu: DAIMA sadece kendi bolumu (baska bolumu hic gormez).
        Basemsire: bolum ve/veya meslek filtresi; ikisi de bossa tum liste.
        """
        rows = list(self.data.schedule)
        if self.db is None:
            return rows
        if self.db.is_sorumlu:
            bid = self.db.bolum_id
            return [r for r in rows if r.get("_bolum_id") == bid]
        if self.db.is_bashemsire:
            bolum = (self._filter_bolum.get() or "").strip()
            meslek = (self._filter_meslek.get() or "").strip()
            if bolum == "(tümü)":
                bolum = ""
            if meslek == "(tümü)":
                meslek = ""
            if bolum:
                rows = [r for r in rows if str(r.get("bolum", "")).strip() == bolum]
            if meslek:
                rows = [r for r in rows if str(r.get("gorev", "")).strip() == meslek]
        return rows

    def _refresh_filter_options(self):
        """Bolum/meslek acilir listelerini mevcut veriden doldurur."""
        if not hasattr(self, "bolum_filter"):
            return
        bolumler = sorted({str(r.get("bolum", "")).strip() for r in self.data.schedule if str(r.get("bolum", "")).strip()})
        meslekler = sorted({str(r.get("gorev", "")).strip() for r in self.data.schedule if str(r.get("gorev", "")).strip()})
        self.bolum_filter["values"] = ["(tümü)"] + bolumler
        self.meslek_filter["values"] = ["(tümü)"] + meslekler

    def _clear_filters(self):
        self._filter_bolum.set("")
        self._filter_meslek.set("")
        self._on_filter_toggle()

    def _rows_from_data(self):
        rows = []
        for row in self._display_schedule:
            days = [row.get("days", {}).get(d, "") for d in range(1, DAY_COUNT + 1)]
            rows.append(
                [row.get("ad_soyad", ""), row.get("gorev", ""), row.get("bolum", ""), fmt_num(row.get("lg", 0))]
                + days + ["", "", ""]
            )
        return rows

    def refresh(self):
        # _meta tum listeden kurulur (isimle arama her zaman calissin); goruntu
        # ise filtreli olabilir.
        self._meta_by_name = {
            str(row.get("ad_soyad", "")).strip(): row for row in self.data.schedule
        }
        self._refresh_filter_options()
        self._display_schedule = self._visible_schedule()
        self.sheet.set_sheet_data(self._rows_from_data(), reset_col_positions=False)
        self.after(50, self._full_refresh)

    def _on_filter_toggle(self):
        """Genel/ozel liste anahtari degisince gorunumu yeniden kur."""
        self._display_schedule = self._visible_schedule()
        self.sheet.set_sheet_data(self._rows_from_data(), reset_col_positions=False)
        self.after(50, self._full_refresh)

    def _full_refresh(self):
        self._update_month_label()
        self._color_day_headers()
        self._apply_day_locks()
        self._apply_permission_locks()
        self.recalc_all()

    def _update_gecmis_ay_butonu(self):
        """Gecmis ay dugmesinin gorunumunu duruma gore ayarlar."""
        btn = getattr(self, "btn_gecmis_kilit", None)
        if btn is None:
            return
        if self.db is None or self.db.is_bashemsire:
            return
        if not self._is_gecmis_ay():
            # Gecerli/gelecek ay - kilit kavrami yok, dugmeyi pasiflestir.
            btn.config(text="🔒 Geçmiş Ay", state="disabled")
        elif self._gecmis_kilit_aktif:
            btn.config(text="🔒 Kilidi Aç", state="normal")
        else:
            btn.config(text="🔓 Açık — Kilitle", state="normal")

    def _is_gecmis_ay(self) -> bool:
        """Acik ay, icinde bulundugumuz aydan onceki bir ay mi?"""
        try:
            yil = int(self.data.settings.get("yil"))
            ay = int(self.data.settings.get("ay_no"))
        except (TypeError, ValueError):
            return False
        bugun = datetime.date.today()
        return (yil, ay) < (bugun.year, bugun.month)

    def _gecmis_ay_kilitli(self) -> bool:
        """Gecmis ay sorumluya kapali mi? Basemsire sifresiyle acilmissa degil.

        Amac: kapanmis aylarin verisi sonradan degistirilmesin. Basemsire
        kendi oturumunda zaten serbesttir; sorumlunun bilgisayarinda duzeltme
        gerekiyorsa basemsire sifresini girip o ayi acar."""
        if self.db is None or self.db.is_bashemsire:
            return False
        if not self._is_gecmis_ay():
            return False
        anahtar = (self.data.settings.get("yil"), self.data.settings.get("ay_no"))
        return anahtar not in self._acilan_aylar

    def _toggle_gecmis_ay_kilidi(self):
        """Gecmis ay kilidini basemsire sifresiyle acar / tekrar kilitler."""
        anahtar = (self.data.settings.get("yil"), self.data.settings.get("ay_no"))
        if anahtar in self._acilan_aylar:
            self._acilan_aylar.discard(anahtar)
            self._full_refresh()
            return
        dlg = BashemsireOnayDialog(self, self.data.settings.get("ay_adi", ""),
                                   self.data.settings.get("yil", ""))
        self.wait_window(dlg)
        if dlg.onaylandi:
            self._acilan_aylar.add(anahtar)
            self._full_refresh()
            messagebox.showinfo(
                APP_TITLE,
                f"{self.data.settings.get('ay_adi','')} {self.data.settings.get('yil','')} "
                "düzenlemeye açıldı.\n\nYapılan değişiklikler yine sizin adınıza kaydedilir.",
            )

    def _apply_permission_locks(self):
        """Role gore satir kilidi: personel her seyi salt-okur; sorumlu sadece
        kendi bolumunu; basemsire her yeri duzenler. Asil zorlama DB'de (RLS);
        bu, kullaniciya yetkisi olmayan hucreyi hic actirmamak icin.

        Ayrica GECMIS AYLAR sorumluya kapalidir (basemsire sifresiyle acilir).

        NOT: Ayin cekmedigi gunler icin sutun kilidi (_apply_day_locks) ayrica
        uygulanir; ikisi birlikte calisir."""
        n = self.sheet.get_total_rows()
        self._gecmis_kilit_aktif = self._gecmis_ay_kilitli()
        self._update_gecmis_ay_butonu()
        if n == 0:
            return
        # Once tum satir kilitlerini kaldir (ay/rol degisince kalinti kalmasin).
        self.sheet.readonly_rows(rows=list(range(n)), readonly=False)
        if self._gecmis_kilit_aktif:
            locked = list(range(n))  # gecmis ay - hepsi kapali
        elif self.db is None or self.db.is_bashemsire:
            locked = []
        elif self.db.rol == "personel":
            locked = list(range(n))
        else:  # sorumlu - gorunen (filtreli) satirlar uzerinden kilitle
            locked = [
                i for i, row in enumerate(self._display_schedule)
                if not self.db.can_edit_bolum(row.get("_bolum_id"))
            ]
        # Gorsel isaret: kilitli satirlarin ad/gorev/bolum sutunlari gri olur.
        # (Gun sutunlarina uygulamiyoruz; recalc_all oradaki kod renklerini
        # zaten ustune yaziyor.)
        self._locked_rows = set(locked)
        acik_cells = [(r, c) for r in range(n) if r not in self._locked_rows for c in (0, 1, 2)]
        kilit_cells = [(r, c) for r in self._locked_rows for c in (0, 1, 2)]
        if acik_cells:
            self.sheet.highlight_cells(
                cells=acik_cells, bg="#" + DEFAULT_BG, fg="#" + DEFAULT_FG, redraw=False
            )
        if kilit_cells:
            self.sheet.highlight_cells(
                cells=kilit_cells, bg="#" + LOCKED_ROW_BG, fg="#" + LOCKED_ROW_FG, redraw=False
            )
        if locked:
            self.sheet.readonly_rows(rows=locked, readonly=True)
        self._update_lock_hint()

    def _update_month_label(self):
        self.month_label.config(
            text=f"{self.data.settings.get('ay_adi','')} {self.data.settings.get('yil','')}"
        )

    def _color_day_headers(self):
        gun_sayisi = self.data.days_in_month()
        for i in range(DAY_COUNT):
            day = i + 1
            c = self.DAY_START_IDX + i
            if day > gun_sayisi:
                bg = DISABLED_DAY_BG
            elif self.data.is_holiday_day(day):
                bg = HOLIDAY_HEADER_COLOR
            elif self.data.is_weekend_day(day):
                bg = WEEKEND_HEADER_COLOR
            else:
                bg = None
            if bg:
                self.sheet.highlight_cells(column=c, bg="#" + bg, canvas="header", redraw=False)
            else:
                self.sheet.dehighlight_cells(column=c, canvas="header", redraw=False)
        self.sheet.refresh()

    def _day_bg_map(self):
        """Gun no -> BOS hucrenin arka plan rengi.

        Hafta sonu ve resmi tatil sutunlari bastan asagi renklenir ki goz
        bunlari kolayca secsin. DOLU hucreler kendi kod rengini korur
        (D/N gibi vardiya bilgisi renkten okunuyor, onu ezmiyoruz)."""
        gun_sayisi = self.data.days_in_month()
        m = {}
        for gun in range(1, DAY_COUNT + 1):
            if gun > gun_sayisi:
                m[gun] = DISABLED_DAY_BG
            elif self.data.is_holiday_day(gun):
                m[gun] = HOLIDAY_HEADER_COLOR
            elif self.data.is_weekend_day(gun):
                m[gun] = WEEKEND_HEADER_COLOR
            else:
                m[gun] = DEFAULT_BG
        return m

    def _apply_day_locks(self):
        """Ayin cekmedigi gun sutunlarini salt-okunur yapar; ay degistiginde
        onceki aydan kalan kilitleri de kaldirir."""
        gun_sayisi = self.data.days_in_month()
        valid = [self.DAY_START_IDX + i for i in range(gun_sayisi)]
        invalid = [self.DAY_START_IDX + i for i in range(gun_sayisi, DAY_COUNT)]
        if valid:
            self.sheet.readonly_columns(columns=valid, readonly=False)
        if invalid:
            self.sheet.readonly_columns(columns=invalid, readonly=True)

    def add_row(self, ad_soyad="", gorev="", bolum="", lg=0.0, redraw: bool = True):
        self.sheet.insert_row(
            [ad_soyad, gorev, bolum, fmt_num(lg)] + [""] * DAY_COUNT + ["0", "0", fmt_num(lg)], redraw=redraw
        )
        if redraw:
            self._recalc_row(self.sheet.get_total_rows() - 1)

    def delete_row(self):
        sel = self.sheet.get_currently_selected()
        if sel:
            self.sheet.delete_row(sel.row)

    def add_personnel_rows(self, personnel_rows):
        existing = {str(r[0]).strip() for r in self.sheet.get_sheet_data() if r and r[0]}
        added = 0
        for p in personnel_rows:
            ad = str(p.get("ad_soyad", "")).strip()
            if not ad or ad in existing:
                continue
            self.add_row(ad_soyad=ad, gorev=p.get("unvan", ""), bolum=p.get("birim", ""), redraw=False)
            existing.add(ad)
            added += 1
        self.recalc_all()
        return added

    def get_schedule(self):
        note_options = self.sheet.get_cell_options(key="note")
        gun_sayisi = self.data.days_in_month()
        rows = []
        for r_idx, r in enumerate(self.sheet.get_sheet_data()):
            if not r or not str(r[0]).strip():
                continue
            days = {}
            notes = {}
            for i, d in enumerate(range(1, gun_sayisi + 1)):
                v = r[self.DAY_START_IDX + i]
                if v not in (None, ""):
                    days[d] = str(v).strip()
                opt = note_options.get((r_idx, self.DAY_START_IDX + i))
                if opt and opt.get("note"):
                    notes[d] = opt["note"]
            try:
                lg = float(r[self.LG_IDX]) if r[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            rows.append({"ad_soyad": r[0], "gorev": r[1], "bolum": r[2], "lg": lg, "days": days, "notes": notes})
        return rows

    # -- hesaplama & renklendirme ---------------------------------------
    def _color_for_cell(self, raw: str):
        sa = parse_saat_araligi(raw)
        if sa is not None:
            # Saatlik giris: kullanici rengiden taniyabilsin diye ozel renkler.
            return CALISMA_SAAT_COLOR if sa[0] == "Ç" else IZIN_SAAT_COLOR
        kod, _, farkli_bolum = parse_cell_code(raw)
        if not kod:
            return DEFAULT_BG
        if farkli_bolum:
            return FARKLI_BOLUM_COLOR
        return CODE_COLORS.get(kod, DEFAULT_BG)

    def _recalc_row(self, r: int, codes_map=None, hedef=None, redraw: bool = True):
        if codes_map is None:
            codes_map = self.get_codes_map()
        if hedef is None:
            hedef = self.get_hedef_saat()
        total = 0.0
        row_data = self.sheet.get_row_data(r)
        gun_sayisi = self.data.days_in_month()
        gun_bg = self._day_bg_map()
        for i in range(DAY_COUNT):
            c = self.DAY_START_IDX + i
            if i >= gun_sayisi:
                self.sheet.highlight_cells(row=r, column=c, bg="#" + DISABLED_DAY_BG, redraw=False)
                continue
            raw = str(row_data[c] or "").strip()
            # Bos hucre: hafta sonu/tatil rengi. Dolu hucre: kendi kod rengi.
            renk = self._color_for_cell(raw) if raw else gun_bg[i + 1]
            self.sheet.highlight_cells(row=r, column=c, bg="#" + renk, redraw=False)
            saat = cell_hours(raw, codes_map)
            if saat is not None:
                total += saat
        try:
            lg = float(row_data[self.LG_IDX]) if row_data[self.LG_IDX] not in (None, "") else 0.0
        except (TypeError, ValueError):
            lg = 0.0
        fark = total - hedef
        lc = lg + fark
        self.sheet.set_cell_data(r, self.TOPLAM_IDX, fmt_num(total), redraw=False)
        self.sheet.set_cell_data(r, self.FARK_IDX, fmt_num(fark), redraw=False)
        self.sheet.set_cell_data(r, self.LC_IDX, fmt_num(lc), redraw=False)
        if fark < 0:
            bg, fg = FARK_NEG_BG, FARK_NEG_FG
        elif fark > 0:
            bg, fg = FARK_POS_BG, FARK_POS_FG
        else:
            bg, fg = DEFAULT_BG, DEFAULT_FG
        self.sheet.highlight_cells(row=r, column=self.TOPLAM_IDX, bg="#" + DEFAULT_BG, fg="#" + DEFAULT_FG, redraw=False)
        self.sheet.highlight_cells(row=r, column=self.FARK_IDX, bg="#" + bg, fg="#" + fg, redraw=False)
        self.sheet.highlight_cells(row=r, column=self.LC_IDX, bg="#" + bg, fg="#" + fg, redraw=False)
        if redraw:
            self.sheet.refresh()
            self._update_detail_panel_for_row(r)

    def recalc_all(self):
        """Tum satirlari topluca yeniden hesaplar. Optimizasyon: hucre hucre
        highlight_cells cagirmak yerine (148 satir x 31 gun ~ binlerce cagri)
        renge gore gruplayip her renk icin TEK cagri yapar. Kasmayi ciddi
        oranda azaltir."""
        codes_map = self.get_codes_map()
        hedef = self.get_hedef_saat()
        n = self.sheet.get_total_rows()
        bg_groups = defaultdict(list)       # gun hucresi rengi -> [(r, c)]
        fark_groups = defaultdict(list)     # (bg, fg) -> [(r, c)] (fark + L.C.)
        gun_bg = self._day_bg_map()         # bos hucreler icin hafta sonu/tatil rengi
        for r in range(n):
            row_data = self.sheet.get_row_data(r)
            total = 0.0
            for i in range(DAY_COUNT):
                c = self.DAY_START_IDX + i
                raw = str(row_data[c] or "").strip()
                renk = self._color_for_cell(raw) if raw else gun_bg[i + 1]
                bg_groups[renk].append((r, c))
                saat = cell_hours(raw, codes_map)
                if saat is not None:
                    total += saat
            try:
                lg = float(row_data[self.LG_IDX]) if row_data[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            fark = total - hedef
            lc = lg + fark
            self.sheet.set_cell_data(r, self.TOPLAM_IDX, fmt_num(total), redraw=False)
            self.sheet.set_cell_data(r, self.FARK_IDX, fmt_num(fark), redraw=False)
            self.sheet.set_cell_data(r, self.LC_IDX, fmt_num(lc), redraw=False)
            if fark < 0:
                bg, fg = FARK_NEG_BG, FARK_NEG_FG
            elif fark > 0:
                bg, fg = FARK_POS_BG, FARK_POS_FG
            else:
                bg, fg = DEFAULT_BG, DEFAULT_FG
            fark_groups[(DEFAULT_BG, DEFAULT_FG)].append((r, self.TOPLAM_IDX))
            fark_groups[(bg, fg)].append((r, self.FARK_IDX))
            fark_groups[(bg, fg)].append((r, self.LC_IDX))
        for color, cells in bg_groups.items():
            self.sheet.highlight_cells(cells=cells, bg="#" + color, redraw=False)
        for (bg, fg), cells in fark_groups.items():
            self.sheet.highlight_cells(cells=cells, bg="#" + bg, fg="#" + fg, redraw=False)
        self.sheet.refresh()

    def _changed_coords(self, event):
        """Bir degisiklik olayindan etkilenen (satir, sutun) ciftlerini cikarir.

        tksheet surumune/olay turune gore veri farkli yerde durabiliyor:
        toplu islemlerde event.cells.table ({(r,c): eski_deger}), tek hucre
        duzenlemesinde event.row/column. Ikisini de savunmaci okuyoruz."""
        coords = set()
        cells = getattr(event, "cells", None)
        table = getattr(cells, "table", None) if cells is not None else None
        if table is None and isinstance(event, dict):
            table = (event.get("cells") or {}).get("table")
        if table:
            for key in table:
                try:
                    r, c = int(key[0]), int(key[1])
                    coords.add((r, c))
                except (TypeError, ValueError, IndexError):
                    continue
        if not coords:
            r = c = None
            try:
                r, c = event["row"], event["column"]
            except Exception:
                r, c = getattr(event, "row", None), getattr(event, "column", None)
            if r is not None and c is not None:
                try:
                    coords.add((int(r), int(c)))
                except (TypeError, ValueError):
                    pass
        return coords

    def _on_modified(self, event):
        """Hucre yazma, SILME, kesme ve yapistirma dahil tum degisikliklerde
        calisir: kodu standartlastirir, satiri yeniden renklendirir ve DB'ye
        yazar."""
        if getattr(self, "_suppress_events", False):
            return
        n = self.sheet.get_total_rows()
        coords = {(r, c) for (r, c) in self._changed_coords(event) if 0 <= r < n}
        if not coords:
            return

        # Normalizasyon set_cell_data cagirdigi icin olay yeniden tetiklenir;
        # sonsuz donguyu onlemek uzere bayrakla bastiriyoruz.
        self._suppress_events = True
        try:
            for (r, c) in coords:
                if c == 0:
                    self._autofill_gorev_bolum(r)
                elif self.DAY_START_IDX <= c <= self.DAY_END_IDX:
                    raw = str(self.sheet.get_cell_data(r, c) or "")
                    norm = normalize_code(raw)
                    if norm != raw:
                        self.sheet.set_cell_data(r, c, norm, redraw=False)
            # Satiri bastan hesapla: bosalan hucrenin rengi de boylece temizlenir.
            for r in sorted({r for (r, _) in coords}):
                self._recalc_row(r, redraw=False)
            self.sheet.refresh()
        finally:
            self._suppress_events = False

        if self.db is not None:
            for (r, c) in sorted(coords):
                if c == self.LG_IDX or self.DAY_START_IDX <= c <= self.DAY_END_IDX:
                    self._save_cell_to_db(r, c)

    def _save_cell_to_db(self, r: int, c: int):
        """Duzenlenen gun hucresini ya da L.G.'yi DB'ye yazar. Yetki/aginda
        hata olursa satiri DB'deki son haline geri yukler."""
        ad = str(self.sheet.get_cell_data(r, 0) or "").strip()
        meta = self._meta_by_name.get(ad)
        if not meta or not meta.get("_personel_id"):
            return
        # Gecmis ay kilidi: satir kilidi yapistirma gibi yollarla asilabilecegi
        # icin yazma anında da kontrol ediyoruz.
        if self._gecmis_ay_kilitli():
            messagebox.showwarning(
                APP_TITLE,
                "Bu ay kapanmıştır; değişiklik kaydedilmedi.\n\n"
                "Düzeltme gerekiyorsa '🔒 Kilidi Aç' ile başhemşire onayı alın.",
            )
            self._reload_row_from_meta(r, meta)
            return
        if not self.db.can_edit_bolum(meta.get("_bolum_id")):
            messagebox.showwarning(APP_TITLE, "Bu personel sizin yetki alanınızda değil; değişiklik kaydedilmedi.")
            self._reload_row_from_meta(r, meta)
            return

        # Ekrandaki degerleri BURADA (ana is parcaciginda) okuyoruz; ag islemi
        # arka planda yapilacak, o sirada hucre degismis olabilir.
        if c == self.LG_IDX:
            try:
                deger = float(self.sheet.get_cell_data(r, self.LG_IDX) or 0)
            except (TypeError, ValueError):
                deger = 0.0
            job = {
                "tur": "lg", "r": r, "meta": meta, "lg": deger,
                "yil": int(self.data.settings["yil"]), "ay": int(self.data.settings["ay_no"]),
            }
        else:
            gun = c - self.DAY_START_IDX + 1
            kod = str(self.sheet.get_cell_data(r, c) or "").strip()
            opt = self.sheet.get_cell_options(key="note").get((r, c))
            job = {
                "tur": "gun", "r": r, "meta": meta, "gun": gun, "kod": kod,
                "aciklama": (opt.get("note") if opt else None) or None,
                "yil": int(self.data.settings["yil"]), "ay": int(self.data.settings["ay_no"]),
            }
        self._enqueue_write(job)

    # -- arka plan yazma ------------------------------------------------
    def _ensure_writer(self):
        """Yazma islerini sirayla yapan arka plan is parcacigini baslatir.
        Yavas internette her kaydin ag cevabini beklemek arayuzu donduruyordu;
        artik yazma arka planda, kullanici kesintisiz yazmaya devam eder.

        Sonuclar Tkinter'a arka plandan DOKUNULARAK degil, bir sonuc kuyruguna
        birakilip ana is parcaciginda yoklanarak isleniyor (Tkinter thread-safe
        degildir; arka plandan cagirmak nadiren cokmeye yol acar)."""
        if getattr(self, "_write_q", None) is None:
            self._write_q = queue.Queue()
            self._result_q = queue.Queue()
            self._write_thread = threading.Thread(
                target=self._writer_loop, name="nobet-writer", daemon=True
            )
            self._write_thread.start()
            self.after(150, self._poll_write_results)

    def _enqueue_write(self, job: dict):
        self._ensure_writer()
        self._pending = getattr(self, "_pending", 0) + 1
        self._update_save_status()
        self._write_q.put(job)

    def _writer_loop(self):
        """Arka plan is parcacigi: isleri sirayla yapar, sonucu kuyruga birakir."""
        while True:
            job = self._write_q.get()
            if job is None:
                break
            try:
                self._perform_write(job)
                self._result_q.put((job, None))
            except Exception as exc:  # ag/yetki hatasi
                self._result_q.put((job, exc))
            finally:
                self._write_q.task_done()

    def _poll_write_results(self):
        """Ana is parcaciginda periyodik calisir; tamamlanan yazmalari isler."""
        try:
            while True:
                job, exc = self._result_q.get_nowait()
                self._write_done(None if exc is None else (job, exc))
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            try:
                self.after(150, self._poll_write_results)
            except Exception:
                pass  # pencere kapanmis olabilir

    def _perform_write(self, job: dict):
        """Arka planda calisir. Tkinter'a DOKUNMAZ (thread-safe degil)."""
        meta = job["meta"]
        ay_id = self.data.settings.get("_ay_id")
        if not ay_id:
            ay_id = nobet_sync.ensure_ay(self.db, job.get("yil"), job.get("ay"))
            self.data.settings["_ay_id"] = ay_id
        ap_id = meta.get("_ap_id")
        if not ap_id:
            ap_id = nobet_sync.ensure_ay_personel(self.db, ay_id, meta["_personel_id"])
            meta["_ap_id"] = ap_id

        if job["tur"] == "lg":
            nobet_sync.save_lg(self.db, ap_id, job["lg"])
            meta["lg"] = job["lg"]
        else:
            nobet_sync.save_cell(self.db, ap_id, job["gun"], job["kod"], job["aciklama"])
            if job["kod"]:
                meta.setdefault("days", {})[job["gun"]] = job["kod"]
            else:
                meta.get("days", {}).pop(job["gun"], None)

    def _write_done(self, hata):
        """Ana is parcaciginda calisir: sayaci dusur, hata varsa bildir."""
        self._pending = max(0, getattr(self, "_pending", 1) - 1)
        self._update_save_status()
        if hata is None:
            self.on_cell_saved()
            return
        job, exc = hata
        messagebox.showerror(APP_TITLE, f"Kaydedilemedi:\n{exc}")
        r, meta = job.get("r"), job.get("meta")
        if r is not None and 0 <= r < self.sheet.get_total_rows():
            self._reload_row_from_meta(r, meta)

    def _update_save_status(self):
        """Bekleyen kayit sayisini gosterir - kullanici yazdiginin gittigini bilsin."""
        if not hasattr(self, "save_status"):
            return
        n = getattr(self, "_pending", 0)
        self.save_status.set(f"⏳ {n} kayıt gönderiliyor…" if n else "✓ Tüm değişiklikler kaydedildi")

    def apply_remote_changes(self, changes):
        """Baska bir PC'de yapilan degisiklikleri (canli senkron) ekrana ve
        yerel modele isler. Ana is parcaciginda cagrilir. Kendi yazdiklarimiz
        da bu listede olabilir (ayni degeri tekrar yazmak zararsiz)."""
        if not changes:
            return
        by_ap = {r.get("_ap_id"): r for r in self.data.schedule if r.get("_ap_id")}
        disp_index = {
            r.get("_ap_id"): i for i, r in enumerate(self._display_schedule) if r.get("_ap_id")
        }
        etkilenen = set()
        # set_cell_data _on_modified'i tetiklemesin (yoksa uzak degisikligi
        # kendi degisikligimiz sanip tekrar DB'ye yazardik = dongu).
        self._suppress_events = True
        try:
            for ch in changes:
                ap = ch.get("ay_personel_id")
                row = by_ap.get(ap)
                if not row:
                    continue  # baska aya/gorunmeyen kisiye ait - atla
                gun = ch.get("gun")
                if not isinstance(gun, int) or not (1 <= gun <= DAY_COUNT):
                    continue
                kod = (ch.get("kod") or "").strip()
                acik = ch.get("aciklama") or None
                if kod:
                    row.setdefault("days", {})[gun] = kod
                else:
                    row.get("days", {}).pop(gun, None)
                if acik:
                    row.setdefault("notes", {})[gun] = acik
                else:
                    row.get("notes", {}).pop(gun, None)
                ri = disp_index.get(ap)
                if ri is not None:
                    c = self.DAY_START_IDX + gun - 1
                    self.sheet.set_cell_data(ri, c, kod, redraw=False)
                    self.sheet.note(ri, c, note=acik)
                    etkilenen.add(ri)
            for ri in etkilenen:
                self._recalc_row(ri, redraw=False)
            if etkilenen:
                self.sheet.refresh()
        finally:
            self._suppress_events = False

    def _reload_row_from_meta(self, r: int, meta: dict):
        """Bir satiri DB'deki son bilinen haline (meta) geri yazar - basarisiz
        kaydin ekranda kalmamasi icin."""
        self.sheet.set_cell_data(r, self.LG_IDX, fmt_num(meta.get("lg", 0)), redraw=False)
        for i in range(DAY_COUNT):
            gun = i + 1
            self.sheet.set_cell_data(r, self.DAY_START_IDX + i, meta.get("days", {}).get(gun, ""), redraw=False)
        self._recalc_row(r)

    def _autofill_gorev_bolum(self, r: int):
        ad = str(self.sheet.get_cell_data(r, 0) or "").strip()
        if not ad:
            return
        info = self.get_personnel_lookup().get(ad.upper())
        if not info:
            return
        if not str(self.sheet.get_cell_data(r, 1) or "").strip():
            self.sheet.set_cell_data(r, 1, info.get("gorev", ""), redraw=False)
        if not str(self.sheet.get_cell_data(r, 2) or "").strip():
            self.sheet.set_cell_data(r, 2, info.get("bolum", ""), redraw=False)

    def _on_select(self, event):
        try:
            r = event["selected"].row
        except Exception:
            r = None
        if r is not None and 0 <= r < self.sheet.get_total_rows():
            self._update_detail_panel_for_row(r)
            self._update_lock_hint(r)

    def _update_lock_hint(self, r=None):
        """Secili satir yetki disindaysa kullaniciya nedenini yazar. Salt-okunur
        olmak tek basina yeterli degil; kullanici neden yazamadigini bilmeli."""
        if not hasattr(self, "lock_hint"):
            return
        if self.db is None or self.db.is_bashemsire:
            self.lock_hint.set("")
            return
        if getattr(self, "_gecmis_kilit_aktif", False):
            self.lock_hint.set(
                "🔒 Geçmiş ay — kapanmış aylar değiştirilemez. "
                "Düzeltme gerekiyorsa '🔒 Kilidi Aç' ile başhemşire onayı alın."
            )
            return
        if self.db.rol == "personel":
            self.lock_hint.set("🔒 Salt okuma yetkiniz var — çizelgede değişiklik yapamazsınız.")
            return
        # sorumlu
        if r is None or r not in getattr(self, "_locked_rows", set()):
            if getattr(self, "_locked_rows", set()):
                self.lock_hint.set(
                    "🔒 Gri satırlar başka bölümlere ait — görebilirsiniz, değiştiremezsiniz."
                )
            else:
                self.lock_hint.set("")
            return
        ad = str(self.sheet.get_cell_data(r, 0) or "").strip()
        bolum = str(self.sheet.get_cell_data(r, 2) or "").strip()
        self.lock_hint.set(
            f"🔒 {ad} — '{bolum}' bölümünde. Bu satıra dokunma izniniz yok, yalnızca kendi bölümünüzü düzenleyebilirsiniz."
        )

    def _update_detail_panel_for_row(self, r: int):
        row_data = self.sheet.get_row_data(r)
        if not row_data or not str(row_data[0]).strip():
            for v in self.detail_vars.values():
                v.set("-")
            return
        self.detail_vars["ad"].set(str(row_data[0]))
        self.detail_vars["bolum"].set(str(row_data[2]))
        self.detail_vars["lg"].set(str(row_data[self.LG_IDX]))
        self.detail_vars["toplam"].set(str(row_data[self.TOPLAM_IDX]))
        self.detail_vars["hedef"].set(fmt_num(self.get_hedef_saat()))
        self.detail_vars["fark"].set(str(row_data[self.FARK_IDX]))
        self.detail_vars["lc"].set(str(row_data[self.LC_IDX]))

    def _selected_report_rows(self):
        """Rapor icin secili satir indekslerini doner. Birden fazla satir
        secilebilir (satir basligindan surukle, Ctrl+tik ya da hucrelerden
        coklu secim); hicbiri yoksa imlecin bulundugu satir kullanilir."""
        secili = set()
        try:
            secili |= set(self.sheet.get_selected_rows())
        except Exception:
            pass
        try:
            # Hucre bazli secimde de satirlari topla (orn. birkac satirin
            # hucrelerini surukleyerek secmek).
            for (rr, _cc) in self.sheet.get_selected_cells():
                secili.add(rr)
        except Exception:
            pass
        if not secili:
            sel = self.sheet.get_currently_selected()
            if sel and sel.row is not None:
                secili.add(sel.row)
        n = self.sheet.get_total_rows()
        return sorted(r for r in secili if isinstance(r, int) and 0 <= r < n)

    def _pick_people_for_report(self):
        """'Personel Seç' penceresini acar; secilenler icin rapor olusturur."""
        n = self.sheet.get_total_rows()
        kisiler = []
        for r in range(n):
            row = self.sheet.get_row_data(r)
            ad = str(row[0] or "").strip() if row else ""
            if ad:
                kisiler.append((r, ad, str(row[1] or "").strip(), str(row[2] or "").strip()))
        if not kisiler:
            messagebox.showinfo(APP_TITLE, "Listede personel yok.")
            return
        dlg = PersonelSecDialog(self, kisiler)
        self.wait_window(dlg)
        if dlg.sonuc:
            self._generate_report(satirlar=dlg.sonuc)

    def _generate_report(self, satirlar=None):
        if satirlar is None:
            satirlar = self._selected_report_rows()
        if not satirlar:
            messagebox.showinfo(
                APP_TITLE,
                "Önce rapor almak istediğiniz personeli seçin.\n\n"
                "Birden fazla kişi için '👥 Personel Seç' düğmesini kullanın; "
                "ya da çizelgede Ctrl ile tıklayarak çoklu seçim yapın.",
            )
            return
        gun_sayisi = self.data.days_in_month()
        note_options = self.sheet.get_cell_options(key="note")
        kisiler = []
        for r in satirlar:
            row_data = self.sheet.get_row_data(r)
            if not row_data or not str(row_data[0]).strip():
                continue
            days = {}
            notes = {}
            for i in range(gun_sayisi):
                v = row_data[self.DAY_START_IDX + i]
                if v not in (None, ""):
                    days[i + 1] = str(v).strip()
                opt = note_options.get((r, self.DAY_START_IDX + i))
                if opt and opt.get("note"):
                    notes[i + 1] = opt["note"]
            try:
                lg = float(row_data[self.LG_IDX]) if row_data[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            kisiler.append({
                "ad_soyad": row_data[0], "gorev": row_data[1], "bolum": row_data[2],
                "lg": lg, "days": days, "notes": notes,
            })
        if not kisiler:
            messagebox.showinfo(APP_TITLE, "Seçili satırlarda personel yok.")
            return
        if self.on_generate_report:
            self.on_generate_report(kisiler)

    # -- ay geçişi ----------------------------------------------------------
    def _change_month(self, delta: int):
        if self.on_change_month:
            self.on_data_synced()
            self.on_change_month(delta)

    # -- açıklama (not) -------------------------------------------------------
    def _edit_note_for_selected(self):
        sel = self.sheet.get_currently_selected()
        if not sel:
            return
        r, c = sel.row, sel.column
        if not (self.DAY_START_IDX <= c <= self.DAY_END_IDX):
            messagebox.showinfo(APP_TITLE, "Açıklama sadece gün hücrelerine eklenebilir.")
            return
        if c - self.DAY_START_IDX >= self.data.days_in_month():
            messagebox.showinfo(APP_TITLE, "Bu gün seçili ayda bulunmuyor.")
            return
        # Yetki: baska bolumun hucresine aciklama eklenemesin (satir kilidi
        # sag-tik menusunu engellemez, bu yuzden burada da kontrol sart).
        if self.db is not None:
            ad = str(self.sheet.get_cell_data(r, 0) or "").strip()
            meta = self._meta_by_name.get(ad)
            if meta and not self.db.can_edit_bolum(meta.get("_bolum_id")):
                messagebox.showwarning(APP_TITLE, "Bu personel sizin yetki alanınızda değil.")
                return
        mevcut = ""
        opt = self.sheet.get_cell_options(key="note").get((r, c))
        if opt:
            mevcut = opt.get("note") or ""
        yeni = simpledialog.askstring(
            "Açıklama Ekle / Düzenle", "Bu hücre için açıklama:", initialvalue=mevcut, parent=self
        )
        if yeni is None:
            return
        if yeni.strip():
            self.sheet.note(r, c, note=yeni.strip())
        else:
            self.sheet.note(r, c, note=None)
        self.sheet.refresh()
        if self.db is not None:
            self._save_cell_to_db(r, c)

    # -- yazdırma -------------------------------------------------------------
    def print_preview(self):
        self.data.schedule = self.get_schedule()
        html = self.data.render_print_html()
        tmp_dir = tempfile.gettempdir()
        path = os.path.join(tmp_dir, "nobet_cizelgesi_yazdir.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        webbrowser.open("file:///" + path.replace("\\", "/"))

    # -- tablo yakınlaştırma ---------------------------------------------
    def _zoom_in(self):
        if self.zoom_pct >= 200:
            return
        try:
            self.sheet.zoom_in()
        except Exception:
            pass  # tksheet, tablo bos iken (0 satir) zoom sirasinda hata verebiliyor
        self.zoom_pct += 10
        self.zoom_label.config(text=f"{self.zoom_pct}%")

    def _zoom_out(self):
        if self.zoom_pct <= 50:
            return
        try:
            self.sheet.zoom_out()
        except Exception:
            pass  # tksheet, tablo bos iken (0 satir) zoom sirasinda hata verebiliyor
        self.zoom_pct -= 10
        self.zoom_label.config(text=f"{self.zoom_pct}%")


class PersonelRaporuTab(ttk.Frame):
    """Tek bir personelin nobet gunlerini ve vardiya dagilimini gosterir.
    Nobet Cizelgesi'ndeki 'Rapor Olustur' tusuyla doldurulur; buradaki
    degisiklikler 'Kaydet' ile ana cizelgeye geri yazilir."""

    AD_IDX = 0
    LG_IDX = 1
    DAY_START_IDX = 2
    DAY_END_IDX = DAY_START_IDX + DAY_COUNT - 1
    TOPLAM_IDX = DAY_END_IDX + 1
    FARK_IDX = TOPLAM_IDX + 1
    LC_IDX = FARK_IDX + 1

    def __init__(self, master, data: NobetData, get_codes, get_hedef_saat, on_save=None):
        super().__init__(master)
        self.data = data
        self.get_codes = get_codes
        self.get_hedef_saat = get_hedef_saat
        self.on_save = on_save
        self.people = []  # raporda gosterilen personel satirlari

        header = ttk.Frame(self)
        header.pack(fill="x", padx=6, pady=6)
        self.title_label = ttk.Label(header, text="Henüz personel seçilmedi.", font=("Segoe UI", 12, "bold"))
        self.title_label.pack(side="left")
        ttk.Button(header, text="Kaydet", command=self._save).pack(side="right", padx=4)
        ttk.Button(header, text="📥 Excel'e Aktar", command=self._export_excel).pack(side="right", padx=4)

        info = ttk.Frame(self)
        info.pack(fill="x", padx=6, pady=(0, 6))
        self.info_vars = {k: tk.StringVar(value="-") for k in
                           ("bolum", "gorev", "lg", "toplam", "hedef", "fark", "lc")}
        info_fields = [
            ("bolum", "Bölüm"), ("gorev", "Görev"), ("lg", "L.G."),
            ("toplam", "Toplam Saat"), ("hedef", "Hedef Saat"), ("fark", "Fark"), ("lc", "L.Ç."),
        ]
        for i, (key, label) in enumerate(info_fields):
            ttk.Label(info, text=label + ":").grid(row=0, column=2 * i, sticky="w", padx=(0 if i == 0 else 8, 2))
            ttk.Label(info, textvariable=self.info_vars[key], font=("Segoe UI", 9, "bold")).grid(
                row=0, column=2 * i + 1, sticky="w"
            )

        gun_headers = (
            ["Ad Soyad", "L.G."] + [str(d) for d in range(1, DAY_COUNT + 1)]
            + ["Toplam Saat", "Fark", "L.Ç."]
        )
        self.gun_sheet = Sheet(
            self, headers=gun_headers, data=[],
            note_corners=True, tooltip_width=170, tooltip_height=60, tooltip_hover_delay=400,
            table_grid_fg="#9A9A9A",
        )
        self.gun_sheet.enable_bindings()
        setup_sheet(self.gun_sheet)
        self.gun_sheet.pack(fill="x", padx=6, pady=4)
        self.gun_sheet.column_width(column=self.AD_IDX, width=180)
        self.gun_sheet.readonly_columns(columns=[self.AD_IDX], readonly=True)
        self.gun_sheet.column_width(column=self.LG_IDX, width=48)
        for c in range(self.DAY_START_IDX, self.DAY_END_IDX + 1):
            self.gun_sheet.column_width(column=c, width=26)
        for c in (self.TOPLAM_IDX, self.FARK_IDX, self.LC_IDX):
            self.gun_sheet.column_width(column=c, width=64)
        self.gun_sheet.readonly_columns(columns=[self.TOPLAM_IDX, self.FARK_IDX, self.LC_IDX], readonly=True)
        self.gun_sheet.extra_bindings("end_edit_cell", func=lambda e: self._recalc())
        self._apply_day_locks()
        self.gun_sheet.popup_menu_add_command(
            "Açıklama Ekle / Düzenle", self._edit_note,
            table_menu=True, index_menu=False, header_menu=False, empty_space_menu=False,
        )

        alt = ttk.Frame(self)
        alt.pack(fill="both", expand=True, padx=6, pady=(6, 6))

        sol = ttk.Frame(alt)
        sol.pack(side="left", fill="both", expand=True)
        ttk.Label(sol, text="Vardiya Dağılımı", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(4, 0))
        self.report_sheet = Sheet(sol, headers=["Kod", "Açıklama", "Saat", "Gün Sayısı", "Toplam Saat Katkısı"])
        self.report_sheet.enable_bindings()
        setup_sheet(self.report_sheet)
        self.report_sheet.pack(fill="both", expand=True, pady=(2, 0))
        self.report_sheet.readonly_columns(columns=[0, 1, 2, 3, 4], readonly=True)

        # Madde 6: resmi tatilde fiilen calisilan gunler ayri baslik altinda.
        sag = ttk.Frame(alt)
        sag.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.tatil_baslik = ttk.Label(
            sag, text="Resmi Tatilde Çalışma", font=("Segoe UI", 10, "bold"), foreground="#B23A20"
        )
        self.tatil_baslik.pack(anchor="w", pady=(4, 0))
        self.tatil_sheet = Sheet(sag, headers=["Ad Soyad", "Gün", "Kod", "Saat"])
        self.tatil_sheet.enable_bindings()
        setup_sheet(self.tatil_sheet)
        self.tatil_sheet.pack(fill="both", expand=True, pady=(2, 0))
        self.tatil_sheet.readonly_columns(columns=[0, 1, 2, 3], readonly=True)

    def _apply_day_locks(self):
        """Ayin cekmedigi gun sutunlarini salt-okunur yapar; Nobet Cizelgesi
        sekmesindeki davranisin aynisi."""
        gun_sayisi = self.data.days_in_month()
        valid = [self.DAY_START_IDX + i for i in range(gun_sayisi)]
        invalid = [self.DAY_START_IDX + i for i in range(gun_sayisi, DAY_COUNT)]
        if valid:
            self.gun_sheet.readonly_columns(columns=valid, readonly=False)
        if invalid:
            self.gun_sheet.readonly_columns(columns=invalid, readonly=True)

    def _color_day_headers(self):
        gun_sayisi = self.data.days_in_month()
        for i in range(DAY_COUNT):
            day = i + 1
            c = self.DAY_START_IDX + i
            if day > gun_sayisi:
                bg = DISABLED_DAY_BG
            elif self.data.is_holiday_day(day):
                bg = HOLIDAY_HEADER_COLOR
            elif self.data.is_weekend_day(day):
                bg = WEEKEND_HEADER_COLOR
            else:
                bg = None
            if bg:
                self.gun_sheet.highlight_cells(column=c, bg="#" + bg, canvas="header", redraw=False)
            else:
                self.gun_sheet.dehighlight_cells(column=c, canvas="header", redraw=False)
        self.gun_sheet.refresh()

    def clear(self):
        self.people = []
        self.title_label.config(text="Henüz personel seçilmedi.")
        for v in self.info_vars.values():
            v.set("-")
        self.gun_sheet.set_sheet_data(
            [], reset_highlights=True, reset_col_positions=False,
        )
        self.report_sheet.set_sheet_data([], reset_col_positions=False)
        self.tatil_sheet.set_sheet_data([], reset_col_positions=False)

    def show_person(self, row_dict):
        """Tek kisi - coklu gosterimin ozel hali."""
        self.show_people([row_dict])

    def show_people(self, row_dicts):
        """Bir ya da daha fazla personeli raporda gosterir."""
        self.people = list(row_dicts or [])
        if not self.people:
            self.clear()
            return
        if len(self.people) == 1:
            p = self.people[0]
            self.title_label.config(text=f"Personel Raporu: {p.get('ad_soyad','')}")
            self.info_vars["bolum"].set(p.get("bolum", ""))
            self.info_vars["gorev"].set(p.get("gorev", ""))
        else:
            self.title_label.config(text=f"Personel Raporu — {len(self.people)} kişi")
            bolumler = {str(p.get("bolum", "")).strip() for p in self.people if p.get("bolum")}
            self.info_vars["bolum"].set(
                next(iter(bolumler)) if len(bolumler) == 1 else f"{len(bolumler)} bölüm"
            )
            self.info_vars["gorev"].set("—")

        data = []
        for p in self.people:
            days = [p.get("days", {}).get(d, "") for d in range(1, DAY_COUNT + 1)]
            data.append([p.get("ad_soyad", ""), fmt_num(p.get("lg", 0))] + days + ["", "", ""])
        self.gun_sheet.set_sheet_data(data, reset_col_positions=False)
        for ri, p in enumerate(self.people):
            for d, text in p.get("notes", {}).items():
                self.gun_sheet.note(ri, self.DAY_START_IDX + d - 1, note=text)
        self._apply_day_locks()
        self._color_day_headers()
        self._recalc()

    def _recalc(self):
        codes = self.get_codes()
        codes_map = {c["kod"]: float(c.get("saat") or 0) for c in codes}
        hedef = self.get_hedef_saat()
        gun_sayisi = self.data.days_in_month()
        n = self.gun_sheet.get_total_rows()

        tum_gun_degerleri = []   # vardiya dagilimi (tum secililer birlikte)
        tatil_satirlari = []     # madde 6: resmi tatilde fiilen calisma
        toplam_hepsi = lg_hepsi = 0.0

        for r in range(n):
            row = self.gun_sheet.get_row_data(r)
            ad = str(row[self.AD_IDX] or "").strip()
            total = 0.0
            for i in range(DAY_COUNT):
                c = self.DAY_START_IDX + i
                if i >= gun_sayisi:
                    self.gun_sheet.highlight_cells(row=r, column=c, bg="#" + DISABLED_DAY_BG, redraw=False)
                    continue
                gun = i + 1
                raw = str(row[c] or "").strip()
                tum_gun_degerleri.append(raw)
                kod, override, farkli_bolum = parse_cell_code(raw)
                sa = parse_saat_araligi(raw)
                if sa is not None:
                    color = CALISMA_SAAT_COLOR if sa[0] == "Ç" else IZIN_SAAT_COLOR
                elif not kod:
                    # Bos hucre: hafta sonu/resmi tatil sutunu renklensin.
                    if self.data.is_holiday_day(gun):
                        color = HOLIDAY_HEADER_COLOR
                    elif self.data.is_weekend_day(gun):
                        color = WEEKEND_HEADER_COLOR
                    else:
                        color = DEFAULT_BG
                elif farkli_bolum:
                    color = FARKLI_BOLUM_COLOR
                else:
                    color = CODE_COLORS.get(kod, DEFAULT_BG)
                self.gun_sheet.highlight_cells(row=r, column=c, bg="#" + color, redraw=False)
                katki = cell_hours(raw, codes_map)
                if katki is not None:
                    total += katki
                # Resmi tatilde FIILEN calisma (izin/rapor kodlari ve 'İ ...'
                # saatlik izin sayilmaz; 'Ç ...' saatlik calisma sayilir)
                if self.data.is_holiday_day(gun) and is_calisma_kodu(raw):
                    gosterilecek_saat = sa[1] if sa is not None else (katki or 0.0)
                    tatil_satirlari.append([ad, str(gun), raw, fmt_num(gosterilecek_saat)])
            try:
                lg = float(row[self.LG_IDX]) if row[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            fark = total - hedef
            lc = lg + fark
            toplam_hepsi += total
            lg_hepsi += lg
            self.gun_sheet.set_cell_data(r, self.TOPLAM_IDX, fmt_num(total), redraw=False)
            self.gun_sheet.set_cell_data(r, self.FARK_IDX, fmt_num(fark), redraw=False)
            self.gun_sheet.set_cell_data(r, self.LC_IDX, fmt_num(lc), redraw=False)
            if fark < 0:
                bg, fg = FARK_NEG_BG, FARK_NEG_FG
            elif fark > 0:
                bg, fg = FARK_POS_BG, FARK_POS_FG
            else:
                bg, fg = DEFAULT_BG, DEFAULT_FG
            self.gun_sheet.highlight_cells(row=r, column=self.TOPLAM_IDX, bg="#" + DEFAULT_BG, fg="#" + DEFAULT_FG, redraw=False)
            self.gun_sheet.highlight_cells(row=r, column=self.FARK_IDX, bg="#" + bg, fg="#" + fg, redraw=False)
            self.gun_sheet.highlight_cells(row=r, column=self.LC_IDX, bg="#" + bg, fg="#" + fg, redraw=False)
        self.gun_sheet.refresh()

        cok = n > 1
        fark_hepsi = toplam_hepsi - (hedef * n if n else 0)
        self.info_vars["lg"].set(fmt_num(lg_hepsi) if not cok else f"Σ {fmt_num(lg_hepsi)}")
        self.info_vars["toplam"].set(fmt_num(toplam_hepsi) if not cok else f"Σ {fmt_num(toplam_hepsi)}")
        self.info_vars["hedef"].set(fmt_num(hedef) if not cok else f"{fmt_num(hedef)} × {n}")
        self.info_vars["fark"].set(fmt_num(fark_hepsi) if not cok else f"Σ {fmt_num(fark_hepsi)}")
        self.info_vars["lc"].set(fmt_num(lg_hepsi + fark_hepsi) if not cok else f"Σ {fmt_num(lg_hepsi + fark_hepsi)}")

        rows, _ = compute_usage_rows(codes, tum_gun_degerleri)
        rows = [row for row in rows if row[3] > 0]
        self.report_sheet.set_sheet_data(rows, reset_col_positions=False)
        for col, width in enumerate((80, 280, 70, 110, 150)):
            self.report_sheet.column_width(column=col, width=width)
        self._last_usage_rows = rows  # Excel aktarimi icin sakla

        # -- madde 6: resmi tatilde calisma ayri baslik --
        tatil_satirlari.sort(key=lambda x: (x[0], int(x[1])))
        self._last_tatil_rows = tatil_satirlari  # Excel aktarimi icin sakla
        self.tatil_sheet.set_sheet_data(tatil_satirlari, reset_col_positions=False)
        for col, width in enumerate((220, 60, 90, 90)):
            self.tatil_sheet.column_width(column=col, width=width)
        gun_adet = len(tatil_satirlari)
        saat_top = sum(float(x[3]) for x in tatil_satirlari) if tatil_satirlari else 0.0
        self.tatil_baslik.config(
            text=f"Resmi Tatilde Çalışma — {gun_adet} gün, toplam {fmt_num(saat_top)} saat"
            if gun_adet else "Resmi Tatilde Çalışma — yok"
        )

    def _export_excel(self):
        """Rapordaki 3 bolumu (gun cizelgesi, vardiya dagilimi, resmi tatilde
        calisma) tek bir Excel dosyasina aktarir. Ekrandaki GUNCEL hali
        (kullanicinin yaptigi degisiklikler dahil) disari verilir."""
        n = self.gun_sheet.get_total_rows()
        if not n:
            messagebox.showinfo(APP_TITLE, "Önce Nöbet Çizelgesi'nden personel seçip rapor oluşturun.")
            return
        gun_sayisi = self.data.days_in_month()
        note_options = self.gun_sheet.get_cell_options(key="note")
        # Ekrandaki guncel degerleri oku (kaydedilmemis duzenlemeler de gelsin)
        people = []
        for r in range(n):
            row = self.gun_sheet.get_row_data(r)
            ad = str(row[self.AD_IDX] or "").strip()
            if not ad:
                continue
            kaynak = self.people[r] if r < len(self.people) else {}
            days, notes = {}, {}
            for i in range(gun_sayisi):
                v = row[self.DAY_START_IDX + i]
                if v not in (None, ""):
                    days[i + 1] = str(v).strip()
                opt = note_options.get((r, self.DAY_START_IDX + i))
                if opt and opt.get("note"):
                    notes[i + 1] = opt["note"]
            try:
                lg = float(row[self.LG_IDX]) if row[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            people.append({
                "ad_soyad": ad,
                "gorev": kaynak.get("gorev", ""),
                "bolum": kaynak.get("bolum", ""),
                "lg": lg, "days": days, "notes": notes,
            })

        path = filedialog.asksaveasfilename(
            title="Personel Raporunu Excel'e Aktar",
            defaultextension=".xlsx",
            filetypes=[("Excel dosyaları", "*.xlsx")],
            initialfile=f"Personel_Raporu_{dosya_zaman_damgasi()}.xlsx",
        )
        if not path:
            return
        try:
            export_personel_raporu_excel(
                path,
                people=people,
                codes=self.get_codes(),
                hedef=self.get_hedef_saat(),
                ay_adi=self.data.settings.get("ay_adi", ""),
                yil=self.data.settings.get("yil", ""),
                gun_sayisi=gun_sayisi,
                tatil_gunleri=[g for g in range(1, gun_sayisi + 1) if self.data.is_holiday_day(g)],
                hafta_sonu_gunleri=[g for g in range(1, gun_sayisi + 1) if self.data.is_weekend_day(g)],
                usage_rows=getattr(self, "_last_usage_rows", []),
                tatil_rows=getattr(self, "_last_tatil_rows", []),
            )
        except PermissionError:
            messagebox.showerror(
                APP_TITLE,
                "Dosya oluşturulamadı. Aynı adlı dosya Excel'de açıksa kapatıp tekrar deneyin.",
            )
            return
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Excel'e aktarılamadı:\n{exc}")
            return
        messagebox.showinfo(APP_TITLE, f"Rapor Excel'e aktarıldı:\n{os.path.basename(path)}")

    def _edit_note(self):
        sel = self.gun_sheet.get_currently_selected()
        if not sel:
            return
        r, c = sel.row, sel.column
        if not (self.DAY_START_IDX <= c <= self.DAY_END_IDX):
            messagebox.showinfo(APP_TITLE, "Açıklama sadece gün hücrelerine eklenebilir.")
            return
        if c - self.DAY_START_IDX >= self.data.days_in_month():
            messagebox.showinfo(APP_TITLE, "Bu gün seçili ayda bulunmuyor.")
            return
        mevcut = ""
        opt = self.gun_sheet.get_cell_options(key="note").get((r, c))
        if opt:
            mevcut = opt.get("note") or ""
        yeni = simpledialog.askstring(
            "Açıklama Ekle / Düzenle", "Bu hücre için açıklama:", initialvalue=mevcut, parent=self
        )
        if yeni is None:
            return
        if yeni.strip():
            self.gun_sheet.note(r, c, note=yeni.strip())
        else:
            self.gun_sheet.note(r, c, note=None)
        self.gun_sheet.refresh()

    def _save(self):
        """Raporda gorunen TUM personelin degisikliklerini cizelgeye yazar."""
        n = self.gun_sheet.get_total_rows()
        if not n:
            messagebox.showinfo(APP_TITLE, "Kaydedilecek bir personel raporu yok.")
            return
        gun_sayisi = self.data.days_in_month()
        note_options = self.gun_sheet.get_cell_options(key="note")
        kaydedilen = 0
        for r in range(n):
            row = self.gun_sheet.get_row_data(r)
            ad = str(row[self.AD_IDX] or "").strip()
            if not ad:
                continue
            days = {}
            notes = {}
            for i in range(gun_sayisi):
                v = row[self.DAY_START_IDX + i]
                if v not in (None, ""):
                    days[i + 1] = str(v).strip()
                opt = note_options.get((r, self.DAY_START_IDX + i))
                if opt and opt.get("note"):
                    notes[i + 1] = opt["note"]
            try:
                lg = float(row[self.LG_IDX]) if row[self.LG_IDX] not in (None, "") else 0.0
            except (TypeError, ValueError):
                lg = 0.0
            if self.on_save:
                self.on_save(ad, days, notes, lg)
                kaydedilen += 1
        if kaydedilen:
            messagebox.showinfo(APP_TITLE, f"{kaydedilen} personelin değişiklikleri çizelgeye kaydedildi.")


class RaporlarTab(ttk.Frame):
    def __init__(self, master, data: NobetData, get_codes, get_schedule, get_hedef):
        super().__init__(master)
        self.data = data
        self.get_codes = get_codes
        self.get_schedule = get_schedule
        self.get_hedef = get_hedef

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=6, pady=6)
        ttk.Button(toolbar, text="Raporu Yenile", command=self.refresh).pack(side="left")
        self.summary_label = ttk.Label(toolbar, text="", font=("Segoe UI", 10, "bold"))
        self.summary_label.pack(side="left", padx=16)

        self.sheet = Sheet(self, headers=["Kod", "Açıklama", "Saat", "Kullanım Adedi", "Toplam Saat Katkısı"])
        self.sheet.enable_bindings()
        setup_sheet(self.sheet)
        self.sheet.pack(fill="both", expand=True, padx=6, pady=6)
        self.sheet.readonly_columns(columns=[0, 1, 2, 3, 4], readonly=True)

    def refresh(self):
        codes = self.get_codes()
        schedule = self.get_schedule()
        all_days = (raw for row in schedule for raw in row.get("days", {}).values())
        rows, toplam_saat = compute_usage_rows(codes, all_days)
        self.sheet.set_sheet_data(rows, reset_col_positions=False)
        for col, width in enumerate((80, 280, 70, 110, 150)):
            self.sheet.column_width(column=col, width=width)

        toplam_personel = len(schedule)
        hedef = self.get_hedef()
        self.summary_label.config(
            text=f"Personel: {toplam_personel}   |   Hedef Saat: {fmt_num(hedef)}   |   Genel Toplam Saat: {fmt_num(toplam_saat)}"
        )


ENTRY_BORDER_OK = "#B0B0B0"
ENTRY_BORDER_ERROR = "#E74C3C"
ENTRY_BG_ERROR = "#FDEDEC"


class AyarlarTab(ttk.Frame):
    def __init__(self, master, data: NobetData, on_apply):
        super().__init__(master)
        self.data = data
        self.on_apply = on_apply

        box = ttk.Frame(self)
        box.pack(anchor="nw", padx=20, pady=20)

        ttk.Label(box, text="Nöbet Planlama Sistemi - Ayarlar", font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 14)
        )

        self.vars = {
            "yil": tk.StringVar(value=str(data.settings.get("yil", ""))),
            "ay_no": tk.StringVar(value=str(data.settings.get("ay_no", ""))),
            "hedef_saat": tk.StringVar(value=str(data.settings.get("hedef_saat", ""))),
            "resmi_tatiller": tk.StringVar(
                value=", ".join(str(d) for d in (data.settings.get("resmi_tatiller") or []))
            ),
        }
        labels = [
            ("yil", "Yıl"),
            ("ay_no", "Ay No (1-12)"),
            ("hedef_saat", "Aylık Hedef Saat"),
            ("resmi_tatiller", "Resmi Tatil Günleri (örn: 1, 15, 23)"),
        ]
        self.entries = {}
        self.error_labels = {}
        for i, (key, label) in enumerate(labels, start=1):
            ttk.Label(box, text=label + ":").grid(row=i, column=0, sticky="w", pady=4, padx=(0, 10))
            entry = tk.Entry(
                box, textvariable=self.vars[key], width=30,
                highlightthickness=2, highlightbackground=ENTRY_BORDER_OK, highlightcolor=ENTRY_BORDER_OK,
                relief="flat",
            )
            entry.grid(row=i, column=1, sticky="w", pady=4, ipady=2)
            self.entries[key] = entry
            err = ttk.Label(box, text="", foreground=ENTRY_BORDER_ERROR)
            err.grid(row=i, column=2, sticky="w", padx=(8, 0))
            self.error_labels[key] = err
            if key == "ay_no":
                self.ay_adi_label = ttk.Label(box, text="", foreground="#555")
                self.ay_adi_label.grid(row=i, column=3, sticky="w", padx=(10, 0))

        for key in ("yil", "ay_no", "hedef_saat", "resmi_tatiller"):
            self.vars[key].trace_add("write", self._validate_all)
        self._validate_all()

        ttk.Button(box, text="Kaydet / Uygula", command=self._apply).grid(
            row=len(labels) + 1, column=0, columnspan=2, sticky="w", pady=(16, 0)
        )
        self.info = ttk.Label(box, text="", foreground="green")
        self.info.grid(row=len(labels) + 2, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _ay_adi_for(self, ay_no_text) -> str:
        try:
            n = int(ay_no_text)
        except (TypeError, ValueError):
            return "?"
        if 1 <= n <= 12:
            return AY_ADLARI[n - 1]
        return "?"

    def _set_field_ok(self, key: str, ok: bool, error_text: str = "Geçersiz"):
        entry = self.entries[key]
        color = ENTRY_BORDER_OK if ok else ENTRY_BORDER_ERROR
        entry.configure(
            highlightbackground=color, highlightcolor=color,
            bg="white" if ok else ENTRY_BG_ERROR,
        )
        self.error_labels[key].config(text="" if ok else error_text)

    def _parse_holiday_list(self, text):
        days = []
        for part in text.replace(";", ",").split(","):
            part = part.strip()
            if part.isdigit():
                days.append(int(part))
        return sorted(set(days))

    def _validate_all(self, *args) -> bool:
        valid = True

        try:
            yil = int(self.vars["yil"].get())
            yil_ok = 1900 <= yil <= 2100
        except ValueError:
            yil, yil_ok = None, False
        self._set_field_ok("yil", yil_ok)
        valid = valid and yil_ok

        try:
            ay_no = int(self.vars["ay_no"].get())
            ay_ok = 1 <= ay_no <= 12
        except ValueError:
            ay_no, ay_ok = None, False
        self._set_field_ok("ay_no", ay_ok)
        valid = valid and ay_ok
        self.ay_adi_label.config(text=self._ay_adi_for(self.vars["ay_no"].get()))

        try:
            hedef = float(self.vars["hedef_saat"].get())
            hedef_ok = hedef >= 0
        except ValueError:
            hedef, hedef_ok = None, False
        self._set_field_ok("hedef_saat", hedef_ok)
        valid = valid and hedef_ok

        resmi_tatiller = self._parse_holiday_list(self.vars["resmi_tatiller"].get())
        tatiller_ok = True
        if yil_ok and ay_ok:
            gun_sayisi = calendar.monthrange(yil, ay_no)[1]
            if any(d > gun_sayisi for d in resmi_tatiller):
                tatiller_ok = False
        self._set_field_ok("resmi_tatiller", tatiller_ok)
        valid = valid and tatiller_ok

        return valid

    def _apply(self):
        if not self._validate_all():
            self.info.config(text="Geçersiz alanlar var, önce kırmızı işaretli olanları düzeltin.", foreground=ENTRY_BORDER_ERROR)
            return
        yil = int(self.vars["yil"].get())
        ay_no = int(self.vars["ay_no"].get())
        hedef = float(self.vars["hedef_saat"].get())
        resmi_tatiller = self._parse_holiday_list(self.vars["resmi_tatiller"].get())
        self.data.settings["yil"] = yil
        self.data.settings["ay_no"] = ay_no
        self.data.settings["ay_adi"] = AY_ADLARI[ay_no - 1]
        self.data.settings["hedef_saat"] = hedef
        self.data.settings["resmi_tatiller"] = resmi_tatiller
        self.info.config(text="Ayarlar kaydedildi.", foreground="green")
        self.on_apply()

    def refresh(self):
        self.vars["yil"].set(str(self.data.settings.get("yil", "")))
        self.vars["ay_no"].set(str(self.data.settings.get("ay_no", "")))
        self.vars["hedef_saat"].set(str(self.data.settings.get("hedef_saat", "")))
        self.vars["resmi_tatiller"].set(
            ", ".join(str(d) for d in (self.data.settings.get("resmi_tatiller") or []))
        )
        self._validate_all()


class LoginDialog(tk.Toplevel):
    """Uygulamadan once acilan giris penceresi. Basarili girişte db.logged_in
    True olur ve pencere kapanir; kullanici pencereyi kaparsa giris yapilmamis
    sayilir."""

    def __init__(self, master, db: SupabaseClient):
        super().__init__(master)
        self.db = db
        self.title(APP_TITLE + " — Giriş")
        self.resizable(False, False)
        self.configure(padx=28, pady=24)

        ttk.Label(self, text="Nöbet Planlama Sistemi", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 4)
        )
        ttk.Label(self, text="Devam etmek için giriş yapın.", foreground="#555").grid(
            row=1, column=0, columnspan=2, pady=(0, 16)
        )

        ttk.Label(self, text="Kullanıcı adı:").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=5)
        self.user_var = tk.StringVar()
        self.user_entry = ttk.Entry(self, textvariable=self.user_var, width=26)
        self.user_entry.grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(self, text="Şifre:").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=5)
        self.pass_var = tk.StringVar()
        self.pass_entry = ttk.Entry(self, textvariable=self.pass_var, width=26, show="•")
        self.pass_entry.grid(row=3, column=1, sticky="w", pady=5)

        self.error_label = ttk.Label(self, text="", foreground=ENTRY_BORDER_ERROR, wraplength=280, justify="left")
        self.error_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.login_btn = ttk.Button(self, text="Giriş Yap", command=self._try_login)
        self.login_btn.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(14, 0))

        self.bind("<Return>", lambda e: self._try_login())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        # Pencereyi ekranin ortasina al ve odagi ver.
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{x}+{y}")
        self.user_entry.focus_set()
        self.grab_set()

    def _set_busy(self, busy: bool):
        self.login_btn.config(state="disabled" if busy else "normal", text="Bağlanıyor…" if busy else "Giriş Yap")
        self.update_idletasks()

    def _try_login(self):
        self.error_label.config(text="")
        self._set_busy(True)
        try:
            self.db.sign_in(self.user_var.get(), self.pass_var.get())
        except AuthError as exc:
            self._set_busy(False)
            self.error_label.config(text=str(exc))
            self.pass_entry.focus_set()
            return
        except Exception as exc:  # beklenmeyen hata da kullaniciyi kilitlemesin
            self._set_busy(False)
            self.error_label.config(text=f"Beklenmeyen hata:\n{exc}")
            return
        self.grab_release()
        self.destroy()

    def _cancel(self):
        self.grab_release()
        self.destroy()


class NobetApp(tk.Tk):
    def __init__(self, db: SupabaseClient):
        super().__init__()
        # Tcl/Tk, Windows ekran olcegini (orn. %125-150) algilayip TUM piksel
        # degerlerini (pencere boyutu, sutun genislikleri, yazi tipleri)
        # kendi icinde otomatik buyutuyor - bu da her seyin gerekenden cok
        # daha buyuk gorunmesine sebep oluyor. Bunu 1:1 piksel eslesmesine
        # sabitliyoruz ki kodda verdigimiz olculer ekranda birebir cikssin.
        try:
            self.tk.call("tk", "scaling", 1.0)
        except Exception:
            pass
        self.db = db
        self.relogin_requested = False
        self.title(APP_TITLE)
        self.geometry("1400x800")
        self.data = NobetData()

        self._build_menu()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        # Ayarlar sekmesi kaldirildi: yil/ay artik prev/next ile DB'den geliyor;
        # hedef saat + resmi tatiller ise cizelgedeki "Ay Ayarlari" (basemsire)
        # butonundan yonetiliyor.
        self.kod_tab = KodTanimlariTab(self.notebook, self.data, db=self.db, on_saved=self._on_codes_saved)
        self.personel_tab = PersonellerTab(
            self.notebook, self.data,
            on_export_to_schedule=self._export_personnel_to_schedule,
            db=self.db,
            get_bolum_map=lambda: getattr(self, "_bolum_map", {}),
            on_changed=self._on_personnel_changed,
        )
        self.cizelge_tab = NobetCizelgesiTab(
            self.notebook, self.data,
            get_codes_map=self._current_codes_map,
            get_hedef_saat=self._current_hedef_saat,
            get_personnel_lookup=self.personel_tab.get_personnel_lookup,
            on_change_month=self._change_month,
            on_data_synced=self._sync_schedule_from_cizelge,
            on_generate_report=self._open_personel_report,
            db=self.db,
            on_cell_saved=self._invalidate_lg_cache,
            on_edit_month_settings=self._edit_month_settings,
        )
        self.personel_rapor_tab = PersonelRaporuTab(
            self.notebook, self.data,
            get_codes=self.kod_tab.get_codes,
            get_hedef_saat=self._current_hedef_saat,
            on_save=self._save_personel_report,
        )
        self.rapor_tab = RaporlarTab(
            self.notebook,
            self.data,
            get_codes=self.kod_tab.get_codes,
            get_schedule=self.cizelge_tab.get_schedule,
            get_hedef=self._current_hedef_saat,
        )

        self.notebook.add(self.cizelge_tab, text="Nöbet Çizelgesi")
        self.notebook.add(self.personel_tab, text="Personeller")
        self.notebook.add(self.kod_tab, text="Kod Tanımları")
        self.notebook.add(self.personel_rapor_tab, text="Personel Raporu")
        self.notebook.add(self.rapor_tab, text="Raporlar")

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        status_bar = ttk.Frame(self)
        status_bar.pack(fill="x", side="bottom")
        self.status = tk.StringVar(value="Hazır.")
        ttk.Label(status_bar, textvariable=self.status, relief="sunken", anchor="w").pack(
            side="left", fill="x", expand=True
        )
        self.login_status = tk.StringVar(value="")
        ttk.Label(status_bar, textvariable=self.login_status, relief="sunken", anchor="e", padding=(8, 0)).pack(
            side="right"
        )
        self._update_login_status()

        # Pencere ciziltikten hemen sonra veriyi DB'den yukle (aci baslangicta
        # ag beklemesi arayuzu dondurmasin diye after ile).
        self.after(100, self._initial_load)

    def _initial_load(self):
        """Ilk acilista veriyi yukler. Bugunun ayi DB'de varsa onu, yoksa en
        son dolu ayi acar - boylece kullanicinin saati farkli olsa bile bos
        aya dusup 'veri yok' sanmaz."""
        self.status.set("Veriler yükleniyor…")
        today = datetime.date.today()
        hedef = (today.year, today.month)
        try:
            rows = self.db.select("aylar", params={"select": "yil,ay_no"})
            aylar = {(int(r["yil"]), int(r["ay_no"])) for r in rows}
            if hedef not in aylar and aylar:
                hedef = max(aylar)
        except Exception:
            pass
        self._load_month_from_db(hedef[0], hedef[1], ilk=True)
        # Canli senkronu baslat: bu andan sonraki degisiklikleri izle.
        try:
            self._last_sync = nobet_sync.latest_change_ts(self.db)
        except Exception:
            self._last_sync = None
        self._start_sync()

    # ------------------------------------------------------------- canli senkron
    def _start_sync(self):
        """Arka planda YOKLAMA_SANIYE'de bir DB'yi yoklayip baska PC'lerdeki
        degisiklikleri ceker. Ag cagrisi ayri is parcaciginda yapilir ki arayuz
        (ozellikle yavas internette) donmasin; sonuc ana is parcaciginda islenir."""
        if self.db is None:
            return
        self._sync_stop = threading.Event()
        self._sync_result_q = queue.Queue()
        self._sync_thread = threading.Thread(target=self._sync_loop, name="nobet-sync", daemon=True)
        self._sync_thread.start()
        self.after(1000, self._drain_sync_results)

    def _sync_loop(self):
        while not self._sync_stop.wait(YOKLAMA_SANIYE):
            since = getattr(self, "_last_sync", None)
            if not since:
                continue
            try:
                changes = nobet_sync.poll_changes(self.db, since)
            except Exception:
                continue  # gecici ag hatasi - sonraki turda yeniden dener
            if changes:
                # Isareti ilerlet (bu satirlari aldik). Thread'de yazilir,
                # thread'de okunur; ana is parcacigi sadece uygular.
                self._last_sync = max(c["updated_at"] for c in changes)
                self._sync_result_q.put(changes)

    def _drain_sync_results(self):
        try:
            while True:
                changes = self._sync_result_q.get_nowait()
                try:
                    self.cizelge_tab.apply_remote_changes(changes)
                except Exception:
                    pass
        except queue.Empty:
            pass
        finally:
            if not getattr(self, "_sync_stop", None) or not self._sync_stop.is_set():
                try:
                    self.after(1000, self._drain_sync_results)
                except Exception:
                    pass

    def _stop_sync(self):
        st = getattr(self, "_sync_stop", None)
        if st is not None:
            st.set()

    def _load_month_from_db(self, yil: int, ay: int, ilk: bool = False):
        """Verilen ayin cizelgesini DB'den yukleyip tum sekmeleri tazeler.

        Optimizasyon: kod tanimlari, personel ve bolumler oturum boyunca cok
        nadir degistigi icin bir kez cekilip onbellekte tutulur; ay gecerken
        tekrar sorgulanmaz. Sadece aya ozel veri (ay_personel + nobet_gunleri)
        ve devir hesabi her ay icin cekilir. Personel eklenince/degisince
        onbellek `refresh_caches()` ile tazelenir."""
        try:
            if ilk or not hasattr(self, "_bolum_map"):
                self._bolum_map = nobet_sync.load_bolumler(self.db)
                self._codes_cache = nobet_sync.load_codes(self.db)
                self._personnel_cache = nobet_sync.load_personnel(self.db, self._bolum_map)
                self._lg_cache = {}
            self.data.codes = self._codes_cache
            self.data.personnel = self._personnel_cache
            settings, schedule = nobet_sync.load_month(self.db, yil, ay, self._bolum_map)
            # L.G. devri: her ayin L.G.'si onceki ayin L.C.'sinden gelir.
            # Devir sonucu (yil,ay) bazinda onbelleklenir; bir hucre yazilinca
            # sonraki aylarin devri degisebilecegi icin onbellek temizlenir.
            key = (yil, ay)
            if key not in self._lg_cache:
                code_hours = {}
                for c in self.data.codes:
                    try:
                        code_hours[c["kod"]] = float(c.get("saat") or 0)
                    except (TypeError, ValueError):
                        code_hours[c["kod"]] = 0.0
                self._lg_cache[key] = nobet_sync.carryover_lg_map(self.db, yil, ay, code_hours)
            lg_map = self._lg_cache[key]
            for row in schedule:
                row["lg"] = lg_map.get(row.get("_personel_id"), 0.0)
            self.data.settings.update(settings)
            self.data.schedule = schedule
        except DbError as exc:
            messagebox.showerror(APP_TITLE, f"Veriler yüklenemedi:\n{exc}")
            self.status.set("Yükleme hatası.")
            return
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Beklenmeyen hata:\n{exc}")
            self.status.set("Yükleme hatası.")
            return
        self._refresh_all_tabs()
        dolu = sum(len(r.get("days", {})) for r in schedule)
        self.status.set(
            f"{settings.get('ay_adi','')} {settings.get('yil','')} yüklendi — "
            f"{len(schedule)} personel, {dolu} dolu hücre."
        )

    def _invalidate_lg_cache(self):
        """Bir hucre yazilinca cagrilir. Duzenlenen ayin Toplam'i degistigi icin
        SONRAKI aylarin L.G. devri degisebilir; devir onbellegini temizleyip
        bir sonraki ay geciste yeniden hesaplanmasini sagliyoruz. (Duzenlenen
        ayin kendi L.G.'si onceki aydan geldigi icin degismez, o yuzden mevcut
        goruntuyu yeniden yuklemeye gerek yok.)"""
        if hasattr(self, "_lg_cache"):
            self._lg_cache.clear()

    def refresh_caches(self):
        """Kod/personel/bolum onbelleklerini bir sonraki yuklemede tazeler.
        Personel eklenip cikarilinca cagrilacak (ileride)."""
        for attr in ("_bolum_map", "_codes_cache", "_personnel_cache", "_lg_cache", "_bolum_cache"):
            if hasattr(self, attr):
                delattr(self, attr)

    def _update_login_status(self):
        """Sag alt kosede kimin, hangi rolle giris yaptigini gosterir."""
        p = self.db.profile or {}
        kim = p.get("ad_soyad") or p.get("kullanici_adi") or "?"
        rol = ROL_ETIKET.get(self.db.rol, self.db.rol)
        metin = f"👤 {kim} — {rol}"
        if self.db.is_sorumlu:
            metin += f"   |   Bölüm: {self._bolum_adi(self.db.bolum_id) or '-'}"
        self.login_status.set(metin)

    def _bolum_adi(self, bolum_id):
        """bolum_id -> bölüm adı (tek seferlik sorgu, sonuç önbelleğe alınır)."""
        if not bolum_id:
            return None
        if not hasattr(self, "_bolum_cache"):
            self._bolum_cache = {}
        if bolum_id not in self._bolum_cache:
            try:
                rows = self.db.select("bolumler", params={"select": "ad", "id": f"eq.{bolum_id}"})
                self._bolum_cache[bolum_id] = rows[0]["ad"] if rows else None
            except Exception:
                return None
        return self._bolum_cache[bolum_id]

    # ------------------------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)
        dosya = tk.Menu(menubar, tearoff=0)
        dosya.add_command(label="Yeni", command=self.new_file, accelerator="Ctrl+N")
        dosya.add_separator()
        dosya.add_command(label="Excel'den Aç...", command=self.open_excel, accelerator="Ctrl+O")
        dosya.add_command(label="Excel'e Aktar...", command=self.save_excel, accelerator="Ctrl+S")
        dosya.add_separator()
        dosya.add_command(label="Çıkış", command=self.destroy)
        menubar.add_cascade(label="Dosya", menu=dosya)

        oturum = tk.Menu(menubar, tearoff=0)
        oturum.add_command(label="Çıkış Yap (oturumu kapat)", command=self._logout)
        menubar.add_cascade(label="Oturum", menu=oturum)
        self.config(menu=menubar)

        self.bind("<Control-n>", lambda e: self.new_file())
        self.bind("<Control-o>", lambda e: self.open_excel())
        self.bind("<Control-s>", lambda e: self.save_excel())

    # ------------------------------------------------------------------
    def _current_codes_map(self):
        m = {}
        for c in self.kod_tab.get_codes():
            try:
                m[c["kod"]] = float(c.get("saat") or 0)
            except (TypeError, ValueError):
                m[c["kod"]] = 0.0
        return m

    def _current_hedef_saat(self):
        try:
            return float(self.data.settings.get("hedef_saat") or 0)
        except (TypeError, ValueError):
            return 0.0

    def _on_codes_saved(self):
        """Kod tanimlari (saatler) degisince: onbellegi tazele, ayi yeniden yukle
        ki yeni saatlerle Toplam/Fark/L.C. bastan hesaplansin."""
        self.refresh_caches()
        self._reload_current_month()

    def _on_personnel_changed(self):
        """Personel eklenip/silinince: onbellegi tazele ve ayi yeniden yukle;
        boylece yeni personel Nobet Cizelgesi'ne otomatik duser."""
        self.refresh_caches()
        self._reload_current_month()

    def _reload_current_month(self):
        yil = int(self.data.settings.get("yil"))
        ay = int(self.data.settings.get("ay_no"))
        self._load_month_from_db(yil, ay, ilk=True)

    def _edit_month_settings(self):
        """Bu ay icin hedef saat + resmi tatilleri duzenler (sadece basemsire).
        Eski 'Ayarlar' sekmesinin islevini ay bazinda ustlenir; DB'deki aylar
        tablosuna yazar."""
        if not (self.db is None or self.db.is_bashemsire):
            messagebox.showinfo(APP_TITLE, "Ay ayarlarını yalnızca başhemşire değiştirebilir.")
            return
        yil = int(self.data.settings.get("yil"))
        ay = int(self.data.settings.get("ay_no"))
        gun_sayisi = self.data.days_in_month()
        dlg = AyAyarlariDialog(
            self, yil, ay,
            hedef=self.data.settings.get("hedef_saat", 210),
            tatiller=self.data.settings.get("resmi_tatiller") or [],
            gun_sayisi=gun_sayisi,
        )
        self.wait_window(dlg)
        if dlg.sonuc is None:
            return
        hedef, tatiller = dlg.sonuc
        if self.db is not None:
            try:
                ay_id = self.data.settings.get("_ay_id")
                if not ay_id:
                    ay_id = nobet_sync.ensure_ay(self.db, yil, ay)
                self.db.update("aylar", {"id": f"eq.{ay_id}"},
                               {"hedef_saat": hedef, "resmi_tatiller": tatiller})
            except DbError as exc:
                messagebox.showerror(APP_TITLE, f"Kaydedilemedi:\n{exc}")
                return
        self.data.settings["hedef_saat"] = hedef
        self.data.settings["resmi_tatiller"] = tatiller
        self._invalidate_lg_cache()
        self._reload_current_month()

    def _export_personnel_to_schedule(self, personnel_rows):
        added = self.cizelge_tab.add_personnel_rows(personnel_rows)
        messagebox.showinfo(APP_TITLE, f"{added} personel Nöbet Çizelgesine eklendi.")
        self.notebook.select(self.cizelge_tab)

    def _open_personel_report(self, rows):
        """Cizelgeden gelen bir ya da daha fazla personeli rapor sekmesinde acar."""
        if isinstance(rows, dict):  # eski tek-kisi cagrisiyla uyum
            rows = [rows]
        self.personel_rapor_tab.show_people(rows)
        self.notebook.select(self.personel_rapor_tab)

    def _save_personel_report(self, ad_soyad, days, notes, lg):
        """Rapor sekmesindeki degisiklikleri cizelgeye ve DB'ye yazar.

        Sadece GERCEKTEN degisen hucreler DB'ye gonderilir (31 gunun tamamini
        gondermek gereksiz yazma ve log kalabaligi olurdu). set_cell_data
        programatik oldugu icin _on_modified'a guvenmiyoruz; degisenleri
        acikca kaydediyoruz."""
        tab = self.cizelge_tab
        sheet = tab.sheet
        target_r = None
        for idx, r in enumerate(sheet.get_sheet_data()):
            if r and str(r[0]).strip() == ad_soyad:
                target_r = idx
                break
        if target_r is None:
            # Filtre yuzunden gorunmuyor olabilir - sessizce atlamak yerine bildir.
            messagebox.showerror(
                APP_TITLE,
                f"'{ad_soyad}' şu an çizelgede görünmüyor (filtre?); değişiklikleri kaydedilemedi.",
            )
            return
        degisen_kolonlar = []
        tab._suppress_events = True
        try:
            for d in range(1, self.data.days_in_month() + 1):
                col = tab.DAY_START_IDX + (d - 1)
                eski = str(sheet.get_cell_data(target_r, col) or "").strip()
                yeni = str(days.get(d, "") or "").strip()
                opt = sheet.get_cell_options(key="note").get((target_r, col))
                eski_not = (opt.get("note") if opt else None) or None
                yeni_not = notes.get(d) or None
                if eski != yeni or eski_not != yeni_not:
                    sheet.set_cell_data(target_r, col, yeni, redraw=False)
                    sheet.note(target_r, col, note=yeni_not)
                    degisen_kolonlar.append(col)
        finally:
            tab._suppress_events = False
        tab._recalc_row(target_r)
        if self.db is not None:
            for col in degisen_kolonlar:
                tab._save_cell_to_db(target_r, col)

    def _on_tab_changed(self, event):
        current = self.notebook.select()
        if current == str(self.rapor_tab):
            self.rapor_tab.refresh()
        elif current == str(self.cizelge_tab):
            self.cizelge_tab.recalc_all()

    def _sync_schedule_from_cizelge(self):
        self.data.schedule = self.cizelge_tab.get_schedule()

    def _change_month(self, delta: int):
        # Ay verisi artik DB'de; bellekteki go_to_month yerine hedef ayi
        # dogrudan DB'den yukluyoruz.
        yil = int(self.data.settings.get("yil"))
        ay = int(self.data.settings.get("ay_no")) + delta
        while ay < 1:
            ay += 12
            yil -= 1
        while ay > 12:
            ay -= 12
            yil += 1
        self._load_month_from_db(yil, ay)

    # ------------------------------------------------------------------
    def _sync_data_from_tabs(self):
        self.data.codes = self.kod_tab.get_codes()
        self.data.personnel = self.personel_tab.get_personnel()
        self.data.schedule = self.cizelge_tab.get_schedule()

    def _refresh_all_tabs(self):
        self.kod_tab.refresh()
        self.personel_tab.refresh()
        self.cizelge_tab.refresh()
        self.rapor_tab.refresh()

    def _logout(self):
        if not messagebox.askyesno(APP_TITLE, "Oturumu kapatıp giriş ekranına dönülsün mü?"):
            return
        self._stop_sync()
        self.db.sign_out()
        self.relogin_requested = True
        self.destroy()

    def new_file(self):
        if not messagebox.askyesno(APP_TITLE, "Tüm veriler silinip yeni bir çalışma başlatılsın mı?"):
            return
        self.data = NobetData()
        self._refresh_all_tabs()
        self.personel_rapor_tab.clear()
        self.status.set("Yeni çalışma başlatıldı.")

    def open_excel(self):
        path = filedialog.askopenfilename(
            title="Excel dosyası aç", filetypes=[("Excel dosyaları", "*.xlsx")]
        )
        if not path:
            return
        try:
            self.data.import_from_excel(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Dosya okunamadı:\n{exc}")
            return
        self._refresh_all_tabs()
        self.personel_rapor_tab.clear()
        self.status.set(f"İçe aktarıldı: {os.path.basename(path)}")

    def save_excel(self):
        self._sync_data_from_tabs()
        path = filedialog.asksaveasfilename(
            title="Excel'e Aktar",
            defaultextension=".xlsx",
            filetypes=[("Excel dosyaları", "*.xlsx")],
            initialfile=f"Nobet_Listesi_{dosya_zaman_damgasi()}.xlsx",
        )
        if not path:
            return
        try:
            self.data.export_to_excel(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Kaydedilemedi:\n{exc}")
            return
        self.status.set(f"Excel'e aktarıldı: {os.path.basename(path)}")
        messagebox.showinfo(APP_TITLE, "Excel dosyası oluşturuldu.")


def _enable_dpi_awareness():
    """Windows, DPI-aware olmayan uygulamalari otomatik olarak bulaniklastirip
    buyutur (bitmap stretching); bu da tum piksel olculerini (sutun genislikleri
    dahil) olmasi gerekenden ~1.5-2 kat buyuk gosterir. Bunu onlemek icin
    pencere olusturulmadan once Windows'a bu surecin DPI-aware oldugunu
    bildiriyoruz."""
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _close_splash():
    """PyInstaller onefile acilis ekranini kapatir. Sadece splash'li exe'de
    pyi_splash modulu bulunur; dev ortaminda (python nobet_app.py) yoktur."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    # Giris -> uygulama dongusu. Kullanici "Cikis Yap" derse tekrar giris
    # ekranina doner; pencereyi kaparsa program biter.
    ilk = True
    while True:
        db = SupabaseClient()
        login_root = tk.Tk()
        try:
            login_root.tk.call("tk", "scaling", 1.0)
        except Exception:
            pass
        login_root.withdraw()
        dlg = LoginDialog(login_root, db)
        if ilk:
            # Giris penceresi hazir; artik acilis splash'ini kapat (ilk actan
            # sonra tekrar giris turlarinda splash zaten yok).
            _close_splash()
            ilk = False
        login_root.wait_window(dlg)
        girildi = db.logged_in
        login_root.destroy()
        if not girildi:
            db.close()
            return

        app = NobetApp(db)
        app.mainloop()
        yeniden = app.relogin_requested
        db.close()
        if not yeniden:
            return


if __name__ == "__main__":
    main()
