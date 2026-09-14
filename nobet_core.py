"""Nöbet Planlama Sistemi - veri modeli, hesaplama ve Excel aktarım mantığı.

Bu modül Nobet_Sistemi_Taslak.xlsx dosyasındaki mantığı (Kod_Tanımları
tablosundan saat okuyup COUNTIF/SUMPRODUCT ile toplam saat hesaplama)
Python tarafında yeniden üretir; Excel'e aktarırken de aynı formülleri yazar.
"""
from __future__ import annotations

import calendar
import copy
import re
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DAY_COUNT = 31
FIRST_DAY_COL = 4  # "D" sutunu = 1. gun
LAST_DAY_COL = FIRST_DAY_COL + DAY_COUNT - 1  # "AH" = 31. gun

AY_ADLARI = [
    "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
    "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK",
]

DEFAULT_CODES = [
    ("D", "Gündüz", 10),
    ("N", "Gece", 14),
    ("24", "24 Saat Nöbet", 24),
    ("Yİ", "Yıllık İzin", 10),
    ("R", "Rapor", 10),
    ("RHS", "Resmi Hafta Tatili", 0),
    ("GÇ", "Gebe Çalışan", 7.5),
    ("Sİ", "SüT İzni Kullanan", 10),
    ("Eİ", "Evlilik İzni", 10),
    ("RA", "Radyasyonlu Alan Çalışanı", 7),
    ("DSİ", "Doğum Sonrası İzin", 10),
    ("DÖİ", "Doğum Öncesi İzin", 10),
    ("E", "Eğitim İzni", 10),
    ("İİ", "İdari İzin", 10),
    ("N+1", "Gece+1 Saat", 15),
    ("N+2", "Gece+2 Saat", 16),
    ("İST", "İstifa", 0),
    ("N+1,5", "Gece+1,5", 15.5),
    ("N+3,5", "Gece+3,5", 17.5),
    ("N+4", "Gece+4 Saat", 18),
    ("D+1", "Gündüz+1 Saat", 11),
    ("D+2", "Gündüz+2 Saat", 12),
    ("D+3", "Gündüz+3 Saat", 13),
    ("D+4", "Gündüz+4 Saat", 14),
    ("24+1", "24+1 Saat", 25),
    ("Şİ", "Şua İzni", 10),
    ("RAÖ", "Radyasyonlu Alan 14:30-21:30", 7),
    ("RAG", "Radyasyonlu Alan Gece 21:30-07:30", 10),
    ("İN", "İcap Nöbeti (hücrede İN=SAAT ile özel saat girilir)", 4),
]

# Kod -> hücre rengi. Her kod (ve haftasonu/resmi tatil/farklı bölüm/fark
# renkleri dahil) kasıtlı olarak birbirinden farklı bir renge sahiptir.
CODE_COLORS = {
    "D": "C4DAEE",
    "D+1": "A2C9EB",
    "D+2": "79AFEC",
    "D+3": "ADC6EB",
    "D+4": "B8C6EA",
    "N": "F0C2C6",
    "N+1": "EEA0A5",
    "N+1,5": "EDABB6",
    "N+2": "EDB6C3",
    "N+3,5": "EA9AB5",
    "N+4": "EDB6CD",
    "24": "D1BCE6",
    "24+1": "BA94E0",
    "Yİ": "F1E4B1",
    "R": "EFEAB3",
    "Sİ": "EBEDB6",
    "Eİ": "EFDBB3",
    "DSİ": "EDD6B6",
    "DÖİ": "EDD1B6",
    "E": "E3E8B0",
    "İİ": "F0E8D1",
    "GÇ": "EED5C4",
    "RA": "B3E1E6",
    "RAÖ": "98D3E1",
    "RAG": "BCE6E4",
    "Şİ": "B5E3CC",
    "RHS": "D1D1D1",
    "İST": "E8E8E8",
    "İN": "B6E0AE",
}
UNKNOWN_CODE_COLOR = "FFFFFF"
FARKLI_BOLUM_COLOR = "E8BAD9"
HOLIDAY_HEADER_COLOR = "EB8570"
WEEKEND_HEADER_COLOR = "DCDFE5"
# Saat araligi girisleri: "Ç 10.00-13.30" (saatlik calisma, toplama EKLENIR)
# ve "İ 10.00-13.30" (saatlik izin, toplamdan DUSULUR). Kullanici hucrenin
# saat araligi olup olmadigini renginden anlasin diye ikisine de kendine ozgu
# birer renk ayrilmistir.
CALISMA_SAAT_COLOR = "6FCFC3"  # turkuaz - saatlik calisma
IZIN_SAAT_COLOR = "F5B971"     # kehribar - saatlik izin


# İzin/devamsızlık kodları. Bunlar "çalışma" sayılmaz; resmi tatilde çalışma
# raporunda (kişi tatilde fiilen çalıştı mı?) bu kodlar hariç tutulur.
# Buradaki kodlar dışındaki her kod çalışma kabul edilir (D, N, 24, RA, İN...).
IZIN_KODLARI = {"Yİ", "R", "RHS", "Sİ", "Eİ", "DSİ", "DÖİ", "E", "İİ", "Şİ", "İST"}


def is_calisma_kodu(kod: str) -> bool:
    """Verilen kod fiili çalışmayı mı gösteriyor? (izin/rapor değilse evet)

    Saat aralığı girişlerinde 'Ç ...' çalışma, 'İ ...' izin sayılır."""
    kod = (kod or "").strip()
    if not kod:
        return False
    sa = parse_saat_araligi(kod)
    if sa is not None:
        return sa[0] == "Ç"
    return kod not in IZIN_KODLARI


def tr_upper(s) -> str:
    """Türkçe'ye uygun büyük harf. Python'un .upper()'i 'i'->'I' yapar; biz
    'i'->'İ', 'ı'->'I' istiyoruz (Yİ, İN, Sİ gibi kodlar doğru eşleşsin)."""
    return str(s or "").replace("i", "İ").replace("ı", "I").upper()


# "Ç 10.00-13.30" / "İ 08.00-12.00" biçimi. Ayraç olarak nokta da iki nokta
# da kabul edilir (10:00), çıktı her zaman noktaya çevrilir. İngilizce klavye
# için C -> Ç, I -> İ eşlenir.
_SAAT_ARALIGI_RE = re.compile(
    r"^([ÇCİI])\s*(\d{1,2})[.:](\d{2})\s*-\s*(\d{1,2})[.:](\d{2})$"
)


def parse_saat_araligi(raw) -> Optional[Tuple[str, float, str]]:
    """Saat aralığı hücresini ayrıştırır.

      'Ç 10.00-13.30' -> ('Ç', 3.5,  'Ç 10.00-13.30')   çalışma: EKLENİR
      'İ 10.00-13.30' -> ('İ', 3.5,  'İ 10.00-13.30')   izin:    DÜŞÜLÜR
      'ç10:00-13:30'  -> ('Ç', 3.5,  'Ç 10.00-13.30')   (küçük harf/iki nokta da olur)
      'Ç 22.00-06.00' -> ('Ç', 8.0,  ...)               gece yarısını aşan aralık

    Dönüş: (tip, süre_saat, standart_metin) ya da tanımıyorsa None.
    """
    s = tr_upper(str(raw or "").strip())
    m = _SAAT_ARALIGI_RE.match(s)
    if not m:
        return None
    tip = {"C": "Ç", "I": "İ"}.get(m.group(1), m.group(1))
    sh, sm, eh, em = (int(m.group(i)) for i in range(2, 6))
    if not (0 <= sh <= 23 and 0 <= eh <= 23 and 0 <= sm <= 59 and 0 <= em <= 59):
        return None
    fark = (eh * 60 + em) - (sh * 60 + sm)
    if fark == 0:
        return None  # ayni saat araligi anlamsiz
    if fark < 0:
        fark += 24 * 60  # gece yarisini asan vardiya
    normalized = f"{tip} {sh:02d}.{sm:02d}-{eh:02d}.{em:02d}"
    return tip, fark / 60.0, normalized


def cell_hours(raw, codes_map: Dict[str, float]) -> Optional[float]:
    """Bir gün hücresinin toplam saate NET katkısı.

    Öncelik sırası: saat aralığı (Ç ekler / İ düşer) > KOD=SAAT > kod tablosu.
    Hücre tanınmıyorsa None (katkı yok)."""
    sa = parse_saat_araligi(raw)
    if sa is not None:
        return sa[1] if sa[0] == "Ç" else -sa[1]
    kod, override, _ = parse_cell_code(raw)
    if override is not None:
        return override
    if kod in codes_map:
        return codes_map[kod]
    return None


def normalize_code(raw) -> str:
    """Bir gün hücresi metnini standart biçime getirir: kod kısmı Türkçe büyük
    harfe çevrilir, '@' ön eki ve '=SAAT' son eki korunur.
      'd'      -> 'D'
      '@d'     -> '@D'
      'in=6'   -> 'İN=6'
      'yi'     -> 'Yİ'
    Böylece kullanıcı küçük harf yazsa da kod tanımlarıyla eşleşir."""
    raw = str(raw or "").strip()
    if not raw:
        return ""
    # Saat araligi ise standart bicime cevir: 'ç10:00-13:30' -> 'Ç 10.00-13.30'
    sa = parse_saat_araligi(raw)
    if sa is not None:
        return sa[2]
    prefix = ""
    if raw.startswith("@"):
        prefix = "@"
        raw = raw[1:].strip()
    suffix = ""
    if "=" in raw:
        kod_part, val_part = raw.split("=", 1)
        suffix = "=" + val_part.strip()
        raw = kod_part.strip()
    return prefix + tr_upper(raw) + suffix


def parse_cell_code(raw) -> Tuple[str, Optional[float], bool]:
    """Bir gün hücresindeki metni ayrıştırır.

    Desteklenen biçimler:
      "D"        -> kod=D
      "@D"       -> kod=D, farklı bölümde çalışma
      "İN=6"     -> kod=İN, bu hücreye özel 6 saat (Kod_Tanımları'ndaki
                    sabit saati göz ardı eder)
      "@İN=6"    -> hem farklı bölüm hem özel saat
    Dönüş: (kod, özel_saat_ya_da_None, farklı_bölüm_mü)
    """
    raw = str(raw or "").strip()
    if not raw:
        return "", None, False
    farkli_bolum = False
    if raw.startswith("@"):
        farkli_bolum = True
        raw = raw[1:].strip()
    override = None
    if "=" in raw:
        kod_part, val_part = raw.split("=", 1)
        val_part = val_part.strip().replace(",", ".")
        try:
            override = float(val_part)
            raw = kod_part.strip()
        except ValueError:
            pass
    return raw, override, farkli_bolum


def new_schedule_row(ad_soyad="", gorev="", bolum="", lg: float = 0.0) -> Dict:
    return {
        "ad_soyad": ad_soyad,
        "gorev": gorev,
        "bolum": bolum,
        "days": {},
        "notes": {},
        "lg": lg,
    }


class NobetData:
    def __init__(self) -> None:
        self.settings: Dict = {
            "yil": 2026,
            "ay_no": 1,
            "ay_adi": "OCAK",
            "hedef_saat": 210,
            "resmi_tatiller": [],  # bu aya ait resmi tatil gün numaraları, örn. [1, 23]
        }
        self.codes: List[Dict] = [
            {"kod": k, "aciklama": a, "saat": s} for k, a, s in DEFAULT_CODES
        ]
        self.personnel: List[Dict] = []
        self.schedule: List[Dict] = []  # new_schedule_row(...) listesi
        self.months: Dict[Tuple[int, int], Dict] = {}  # (yıl,ay) -> aya ait anlık görüntü
        self.source_path: Optional[str] = None

    # ------------------------------------------------------------------
    # Hesaplama
    # ------------------------------------------------------------------
    def code_hour_map(self) -> Dict[str, float]:
        m: Dict[str, float] = {}
        for c in self.codes:
            kod = str(c.get("kod") or "").strip()
            if not kod:
                continue
            try:
                m[kod] = float(c.get("saat") or 0)
            except (TypeError, ValueError):
                m[kod] = 0.0
        return m

    def row_total_hours(self, row: Dict) -> float:
        m = self.code_hour_map()
        total = 0.0
        # Tablo her zaman 31 sutunlu; ancak ayin gercekte cekmedigi gunler
        # (orn. Subat 30-31) hesaba katilmaz.
        for d in range(1, self.days_in_month() + 1):
            saat = cell_hours(row.get("days", {}).get(d), m)
            if saat is not None:
                total += saat
        return total

    def row_fark(self, row: Dict) -> float:
        hedef = float(self.settings.get("hedef_saat") or 0)
        return self.row_total_hours(row) - hedef

    def row_lc(self, row: Dict) -> float:
        lg = float(row.get("lg") or 0)
        return lg + self.row_fark(row)

    def code_usage_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        gun_sayisi = self.days_in_month()
        for row in self.schedule:
            for d in range(1, gun_sayisi + 1):
                kod, _, _ = parse_cell_code(row.get("days", {}).get(d))
                if kod:
                    counts[kod] = counts.get(kod, 0) + 1
        return counts

    def toplam_saat_genel(self) -> float:
        return sum(self.row_total_hours(r) for r in self.schedule)

    # ------------------------------------------------------------------
    # Takvim yardımcıları (hafta sonu / resmi tatil renklendirmesi için)
    # ------------------------------------------------------------------
    def day_weekday(self, day: int) -> Optional[int]:
        try:
            return calendar.weekday(int(self.settings["yil"]), int(self.settings["ay_no"]), day)
        except Exception:
            return None

    def is_weekend_day(self, day: int) -> bool:
        wd = self.day_weekday(day)
        return wd is not None and wd >= 5

    def is_holiday_day(self, day: int) -> bool:
        return day in (self.settings.get("resmi_tatiller") or [])

    def days_in_month(self) -> int:
        try:
            return calendar.monthrange(int(self.settings["yil"]), int(self.settings["ay_no"]))[1]
        except Exception:
            return DAY_COUNT

    # ------------------------------------------------------------------
    # Ay geçişleri - önceki/sonraki ay, veriler bellekte tutulur
    # ------------------------------------------------------------------
    def _month_key(self) -> Tuple[int, int]:
        return (int(self.settings.get("yil")), int(self.settings.get("ay_no")))

    def _snapshot_month(self) -> Dict:
        return {
            "schedule": copy.deepcopy(self.schedule),
            "hedef_saat": self.settings.get("hedef_saat"),
            "ay_adi": self.settings.get("ay_adi"),
            "resmi_tatiller": list(self.settings.get("resmi_tatiller") or []),
        }

    def save_current_month(self) -> None:
        self.months[self._month_key()] = self._snapshot_month()

    def _shift_month(self, delta: int) -> Tuple[int, int]:
        yil = int(self.settings.get("yil"))
        ay = int(self.settings.get("ay_no")) + delta
        while ay < 1:
            ay += 12
            yil -= 1
        while ay > 12:
            ay -= 12
            yil += 1
        return yil, ay

    def go_to_month(self, yil: int, ay: int) -> None:
        self.save_current_month()
        key = (yil, ay)
        if key in self.months:
            snap = self.months[key]
            self.schedule = copy.deepcopy(snap["schedule"])
            self.settings["hedef_saat"] = snap["hedef_saat"]
            self.settings["ay_adi"] = snap["ay_adi"]
            self.settings["resmi_tatiller"] = list(snap["resmi_tatiller"])
        else:
            # Yeni ay: aynı personel listesi boş günlerle başlar, geçen ayın
            # L.Ç. değeri yeni ayın L.G. değeri olarak devreder.
            new_schedule = []
            for row in self.schedule:
                new_schedule.append(
                    new_schedule_row(
                        ad_soyad=row.get("ad_soyad", ""),
                        gorev=row.get("gorev", ""),
                        bolum=row.get("bolum", ""),
                        lg=self.row_lc(row),
                    )
                )
            self.schedule = new_schedule
            self.settings["ay_adi"] = AY_ADLARI[ay - 1]
            self.settings["resmi_tatiller"] = []
        self.settings["yil"] = yil
        self.settings["ay_no"] = ay

    def go_to_next_month(self) -> None:
        y, a = self._shift_month(1)
        self.go_to_month(y, a)

    def go_to_prev_month(self) -> None:
        y, a = self._shift_month(-1)
        self.go_to_month(y, a)

    # ------------------------------------------------------------------
    # Excel'den içe aktarma (Nobet_Sistemi_Taslak.xlsx formatı)
    # ------------------------------------------------------------------
    def import_from_excel(self, path: str) -> None:
        wb = openpyxl.load_workbook(path, data_only=False)

        # Onceki oturumda gezinilen ay/yil gibi kalintilar yeni dosyaya
        # tasinmasin diye ayarlari once varsayilana sifirla.
        self.settings = {
            "yil": 2026,
            "ay_no": 1,
            "ay_adi": "OCAK",
            "hedef_saat": 210,
            "resmi_tatiller": [],
        }

        if "Ayarlar" in wb.sheetnames:
            ws = wb["Ayarlar"]
            self.settings["yil"] = ws["B3"].value or self.settings["yil"]
            self.settings["ay_no"] = ws["B4"].value or self.settings["ay_no"]
            self.settings["hedef_saat"] = ws["B6"].value or self.settings["hedef_saat"]

        # Ay adi daima ay_no'dan turetilir. Excel dosyasindaki "Ay Adi"
        # hucresi elle girildigi icin tutarsiz olabiliyor (orn. Ay No=7
        # iken Ay Adi='OCAK' yazan dosyalar gorduk); bu sekilde ay adi/ay
        # no her zaman birbiriyle uyumlu kalir.
        try:
            ay_no_int = int(self.settings["ay_no"])
            if 1 <= ay_no_int <= 12:
                self.settings["ay_adi"] = AY_ADLARI[ay_no_int - 1]
        except (TypeError, ValueError):
            pass

        if "Kod_Tanimlari" in wb.sheetnames:
            ws = wb["Kod_Tanimlari"]
            codes = []
            seen_kodlar = set()
            r = 2
            while ws.cell(r, 1).value not in (None, ""):
                kod = str(ws.cell(r, 1).value).strip()
                aciklama = ws.cell(r, 2).value or ""
                saat = ws.cell(r, 3).value or 0
                if kod not in seen_kodlar:
                    codes.append({"kod": kod, "aciklama": aciklama, "saat": saat})
                    seen_kodlar.add(kod)
                r += 1
            if codes:
                self.codes = codes

        # İN kodu (İcap Nöbeti, hücrede İN=SAAT ile özel saat girilir) bu
        # uygulamaya ozgudur; disaridan gelen bir dosyada tanimli degilse ekle.
        if not any(c["kod"] == "İN" for c in self.codes):
            self.codes.append(
                {"kod": "İN", "aciklama": "İcap Nöbeti (hücrede İN=SAAT ile özel saat girilir)", "saat": 4}
            )

        if "Personeller" in wb.sheetnames:
            ws = wb["Personeller"]
            personnel = []
            r = 2
            while r <= ws.max_row:
                ad = ws.cell(r, 2).value
                if ad not in (None, ""):
                    personnel.append(
                        {
                            "sicil_no": ws.cell(r, 1).value or "",
                            "ad_soyad": ad,
                            "birim": ws.cell(r, 3).value or "",
                            "unvan": ws.cell(r, 4).value or "",
                            "aktif": ws.cell(r, 5).value or "Aktif",
                        }
                    )
                r += 1
            self.personnel = personnel

        if "Nobet_Cizelgesi" in wb.sheetnames:
            ws = wb["Nobet_Cizelgesi"]
            # L.G./L.Ç. sütunları uygulamanın kendi ihracatında bulunur;
            # başlık satırında ararsak eski/yabancı dosyalarla da uyumlu kalırız.
            # Gun sutunlari basliktaki gun numarasindan bulunur. Ay 31 gun
            # cekmiyorsa dosyada daha az gun sutunu olur; sabit bir sutun
            # araligi varsaymak Toplam/Fark/L.C. sutunlarini gun sanmaya yol
            # acardi.
            lg_col = None
            day_col_by_num: Dict[int, int] = {}
            header_row = 2
            for c in range(1, ws.max_column + 1):
                v = ws.cell(header_row, c).value
                if v is None:
                    continue
                text = str(v).strip()
                if text.upper() in ("L.G.", "LG"):
                    lg_col = c
                elif text.isdigit() and 1 <= int(text) <= DAY_COUNT:
                    day_col_by_num.setdefault(int(text), c)
            if not day_col_by_num:
                # Basliksiz/yabanci dosyalar icin eski davranisa geri don.
                base = lg_col + 1 if lg_col is not None else FIRST_DAY_COL
                day_col_by_num = {d + 1: base + d for d in range(DAY_COUNT)}

            schedule = []
            r = 3
            while r <= ws.max_row:
                ad = ws.cell(r, 1).value
                if isinstance(ad, str) and not ad.startswith("="):
                    days = {}
                    notes = {}
                    for gun, col in day_col_by_num.items():
                        cell = ws.cell(r, col)
                        v = cell.value
                        if v not in (None, "") and not (isinstance(v, str) and v.startswith("=")):
                            days[gun] = str(v).strip()
                        if cell.comment is not None and cell.comment.text:
                            notes[gun] = cell.comment.text
                    lg_val = 0.0
                    if lg_col is not None:
                        try:
                            lg_val = float(ws.cell(r, lg_col).value or 0)
                        except (TypeError, ValueError):
                            lg_val = 0.0
                    row = new_schedule_row(
                        ad_soyad=ad,
                        gorev=ws.cell(r, 2).value or "",
                        bolum=ws.cell(r, 3).value or "",
                        lg=lg_val,
                    )
                    row["days"] = days
                    row["notes"] = notes
                    schedule.append(row)
                r += 1
            self.schedule = schedule

        # Bazı dosyalarda Personeller sayfası eksik/kısmi doldurulmuş olabilir
        # (gerçek isim listesi sadece Nobet_Cizelgesi'nde bulunur). Bu durumda
        # çizelgede olup personel listesinde olmayan herkesi de listeye ekle.
        existing_names = {str(p.get("ad_soyad", "")).strip() for p in self.personnel}
        for row in self.schedule:
            ad = str(row.get("ad_soyad", "")).strip()
            if ad and ad not in existing_names:
                self.personnel.append(
                    {
                        "sicil_no": "",
                        "ad_soyad": ad,
                        "birim": row.get("bolum", ""),
                        "unvan": row.get("gorev", ""),
                        "aktif": "Aktif",
                    }
                )
                existing_names.add(ad)

        self.months.clear()
        self.source_path = path

    # ------------------------------------------------------------------
    # Excel'e dışa aktarma - aynı sayfa yapısı ve aynı formül mantığı
    # ------------------------------------------------------------------
    def export_to_excel(self, path: str) -> None:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        # Cikti, ayin gercekte cektigi gun kadar sutun icerir (Subat'ta 28/29).
        gun_sayisi = self.days_in_month()

        header_fill = PatternFill("solid", fgColor="305496")
        header_font = Font(bold=True, color="FFFFFF")
        weekend_fill = PatternFill("solid", fgColor=WEEKEND_HEADER_COLOR)
        holiday_fill = PatternFill("solid", fgColor=HOLIDAY_HEADER_COLOR)

        # --- Ayarlar ---
        ws = wb.create_sheet("Ayarlar")
        ws["A1"] = "NÖBET PLANLAMA SİSTEMİ - AYARLAR"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A3"] = "Yıl"; ws["B3"] = self.settings["yil"]
        ws["A4"] = "Ay No"; ws["B4"] = self.settings["ay_no"]
        ws["A5"] = "Ay Adı"; ws["B5"] = self.settings["ay_adi"]
        ws["A6"] = "Aylık Hedef Saat"; ws["B6"] = self.settings["hedef_saat"]
        ws["A7"] = "Resmi Tatil Günleri"
        ws["B7"] = ", ".join(str(d) for d in (self.settings.get("resmi_tatiller") or []))
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 14

        # --- Kod_Tanimlari ---
        ws = wb.create_sheet("Kod_Tanimlari")
        for col, title in enumerate(["Kod", "Açıklama", "Saat"], start=1):
            c = ws.cell(1, col, title)
            c.font = header_font
            c.fill = header_fill
        for i, code in enumerate(self.codes, start=2):
            ws.cell(i, 1, code["kod"])
            ws.cell(i, 2, code["aciklama"])
            ws.cell(i, 3, code["saat"])
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 10

        # --- Personeller ---
        ws = wb.create_sheet("Personeller")
        headers = ["Sicil No", "Ad Soyad", "Birim", "Unvan", "Aktif/Pasif"]
        for col, title in enumerate(headers, start=1):
            c = ws.cell(1, col, title)
            c.font = header_font
            c.fill = header_fill
        for i, p in enumerate(self.personnel, start=2):
            ws.cell(i, 1, p.get("sicil_no", ""))
            ws.cell(i, 2, p.get("ad_soyad", ""))
            ws.cell(i, 3, p.get("birim", ""))
            ws.cell(i, 4, p.get("unvan", ""))
            ws.cell(i, 5, p.get("aktif", "Aktif"))
        for col_letter, width in zip("ABCDE", (10, 26, 18, 18, 12)):
            ws.column_dimensions[col_letter].width = width

        # --- Nobet_Cizelgesi ---
        ws = wb.create_sheet("Nobet_Cizelgesi")
        ws["A1"] = "=Ayarlar!B3"
        ws["C1"] = "Ay"
        ws["D1"] = "=Ayarlar!B5"
        ws["F1"] = "Hedef Saat"
        ws["H1"] = "=Ayarlar!B6"

        header_row = 2
        first_data_row = 3
        static_headers = ["Personel", "Görev", "BÖLÜM"]
        for col, title in enumerate(static_headers, start=1):
            c = ws.cell(header_row, col, title)
            c.font = header_font
            c.fill = header_fill
        # Sütun sırası: Ad Soyad, Görev, Bölüm, L.G., 1..31, Toplam Saat, Fark, L.Ç.
        lg_col = FIRST_DAY_COL
        c = ws.cell(header_row, lg_col, "L.G.")
        c.font = header_font
        c.fill = header_fill
        day_first_col = lg_col + 1
        for d in range(gun_sayisi):
            col = day_first_col + d
            c = ws.cell(header_row, col, str(d + 1))
            c.font = header_font
            c.alignment = Alignment(horizontal="center")
            if self.is_holiday_day(d + 1):
                c.fill = holiday_fill
            elif self.is_weekend_day(d + 1):
                c.fill = weekend_fill
            else:
                c.fill = header_fill
        toplam_col = day_first_col + gun_sayisi
        fark_col = toplam_col + 1
        lc_col = fark_col + 1
        for col, title in ((toplam_col, "Toplam Saat"), (fark_col, "Fark"), (lc_col, "L.Ç.")):
            c = ws.cell(header_row, col, title)
            c.font = header_font
            c.fill = header_fill

        code_range = "Kod_Tanimlari!$A$2:$A$%d" % (len(self.codes) + 1)
        hour_range = "Kod_Tanimlari!$C$2:$C$%d" % (len(self.codes) + 1)

        for i, row in enumerate(self.schedule):
            r = first_data_row + i
            ws.cell(r, 1, row.get("ad_soyad", ""))
            ws.cell(r, 2, row.get("gorev", ""))
            ws.cell(r, 3, row.get("bolum", ""))
            lg_cell = ws.cell(r, lg_col, float(row.get("lg") or 0))
            lg_cell.alignment = Alignment(horizontal="center")
            for d in range(gun_sayisi):
                raw = row.get("days", {}).get(d + 1)
                cell = ws.cell(r, day_first_col + d)
                if raw:
                    cell.value = raw
                    kod, _, farkli_bolum = parse_cell_code(raw)
                    if farkli_bolum:
                        color = FARKLI_BOLUM_COLOR
                    else:
                        color = CODE_COLORS.get(kod, UNKNOWN_CODE_COLOR)
                    if color != UNKNOWN_CODE_COLOR:
                        cell.fill = PatternFill("solid", fgColor=color)
                note_text = row.get("notes", {}).get(d + 1)
                if note_text:
                    cell.comment = Comment(note_text, "Nobet Uygulamasi")
                cell.alignment = Alignment(horizontal="center")
            first_col_letter = get_column_letter(day_first_col)
            last_col_letter = get_column_letter(day_first_col + gun_sayisi - 1)
            toplam_formula = (
                f"=SUMPRODUCT(COUNTIF({first_col_letter}{r}:{last_col_letter}{r},{code_range}),{hour_range})"
            )
            toplam_cell = ws.cell(r, toplam_col, toplam_formula)
            toplam_col_letter = get_column_letter(toplam_col)
            fark_col_letter = get_column_letter(fark_col)
            lg_col_letter = get_column_letter(lg_col)
            fark_cell = ws.cell(r, fark_col, f"={toplam_col_letter}{r}-Ayarlar!$B$6")
            lc_cell = ws.cell(r, lc_col, f"={lg_col_letter}{r}+{fark_col_letter}{r}")
            for cc in (toplam_cell, fark_cell, lc_cell):
                cc.alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 24
        ws.column_dimensions[get_column_letter(lg_col)].width = 8
        for d in range(gun_sayisi):
            ws.column_dimensions[get_column_letter(day_first_col + d)].width = 5
        for col in (toplam_col, fark_col, lc_col):
            ws.column_dimensions[get_column_letter(col)].width = 10
        ws.freeze_panes = ws.cell(first_data_row, day_first_col)

        # --- Raporlar ---
        ws = wb.create_sheet("Raporlar")
        ws["A1"] = "NÖBET RAPOR ÖZETİ"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A3"] = "Yıl"; ws["B3"] = "=Ayarlar!B3"
        ws["A4"] = "Ay"; ws["B4"] = "=Ayarlar!B5"
        ws["A5"] = "Toplam Personel"; ws["B5"] = "=COUNTA(Personeller!B2:B10000)"
        last_row = first_data_row + max(len(self.schedule) - 1, 0)
        ws["A6"] = "Toplam Saat"
        ws["B6"] = f"=SUM(Nobet_Cizelgesi!{get_column_letter(toplam_col)}{first_data_row}:{get_column_letter(toplam_col)}{last_row})"

        ws["D3"] = "Kod"; ws["E3"] = "Açıklama"; ws["F3"] = "Saat"; ws["G3"] = "Toplam Kullanım"
        for i in range(len(self.codes)):
            rr = 4 + i
            kt_row = 2 + i
            ws.cell(rr, 4, f"=Kod_Tanimlari!A{kt_row}")
            ws.cell(rr, 5, f"=Kod_Tanimlari!B{kt_row}")
            ws.cell(rr, 6, f"=Kod_Tanimlari!C{kt_row}")
            ws.cell(rr, 7, (
                f"=COUNTIF(Nobet_Cizelgesi!${get_column_letter(day_first_col)}${first_data_row}:"
                f"${get_column_letter(day_first_col + gun_sayisi - 1)}${last_row},D{rr})"
            ))
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["D"].width = 10
        ws.column_dimensions["E"].width = 26
        ws.column_dimensions["F"].width = 10
        ws.column_dimensions["G"].width = 16

        wb.save(path)

    # ------------------------------------------------------------------
    # Yazdırma için basit, biçimlendirilmiş HTML üretimi (tarayıcıda önizleme)
    # ------------------------------------------------------------------
    def render_print_html(self) -> str:
        hedef = float(self.settings.get("hedef_saat") or 0)
        gun_sayisi = self.days_in_month()
        rows_html = []
        for row in self.schedule:
            cells = [f"<td class='name'>{_esc(row.get('ad_soyad',''))}</td>",
                     f"<td>{_esc(row.get('gorev',''))}</td>",
                     f"<td>{_esc(row.get('bolum',''))}</td>",
                     f"<td class='num'>{_fmt(row.get('lg', 0))}</td>"]
            for d in range(1, gun_sayisi + 1):
                raw = row.get("days", {}).get(d, "")
                kod, _, farkli_bolum = parse_cell_code(raw)
                sa = parse_saat_araligi(raw)
                if sa is not None:
                    bg = CALISMA_SAAT_COLOR if sa[0] == "Ç" else IZIN_SAAT_COLOR
                elif farkli_bolum and raw:
                    bg = FARKLI_BOLUM_COLOR
                else:
                    bg = CODE_COLORS.get(kod, "")
                style = f" style='background:#{bg}'" if bg else ""
                title = row.get("notes", {}).get(d, "")
                title_attr = f" title='{_esc(title)}'" if title else ""
                cells.append(f"<td{style}{title_attr}>{_esc(raw)}</td>")
            total = self.row_total_hours(row)
            fark = self.row_fark(row)
            lc = self.row_lc(row)
            fark_color = "#c00" if fark < 0 else ("#06c" if fark > 0 else "#333")
            cells.append(f"<td class='num'>{_fmt(total)}</td>")
            cells.append(f"<td class='num' style='color:{fark_color}'>{_fmt(fark)}</td>")
            cells.append(f"<td class='num'>{_fmt(lc)}</td>")
            rows_html.append("<tr>" + "".join(cells) + "</tr>")

        day_headers = []
        for d in range(1, gun_sayisi + 1):
            cls = "holiday" if self.is_holiday_day(d) else ("weekend" if self.is_weekend_day(d) else "")
            cls_attr = f" class='{cls}'" if cls else ""
            day_headers.append(f"<th{cls_attr}>{d}</th>")

        return f"""<!DOCTYPE html>
<html lang="tr"><head><meta charset="utf-8">
<title>Nöbet Çizelgesi - {_esc(self.settings.get('ay_adi',''))} {self.settings.get('yil','')}</title>
<style>
  body {{ font-family: Arial, sans-serif; font-size: 11px; }}
  h1 {{ font-size: 16px; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th, td {{ border: 1px solid #999; padding: 2px 4px; text-align: center; white-space: nowrap; }}
  th {{ background: #305496; color: #fff; }}
  th.weekend {{ background: {WEEKEND_HEADER_COLOR and '#'+WEEKEND_HEADER_COLOR}; color: #333; }}
  th.holiday {{ background: #{HOLIDAY_HEADER_COLOR}; color: #333; }}
  td.name {{ text-align: left; font-weight: bold; white-space: normal; }}
  td.num {{ font-weight: bold; }}
  @media print {{ body {{ margin: 0; }} }}
</style></head>
<body>
<h1>Nöbet Çizelgesi - {_esc(self.settings.get('ay_adi',''))} {self.settings.get('yil','')} (Hedef: {_fmt(hedef)} saat)</h1>
<table>
<thead><tr>
<th>Ad Soyad</th><th>Görev</th><th>Bölüm</th><th>L.G.</th>
{''.join(day_headers)}
<th>Toplam</th><th>Fark</th><th>L.Ç.</th>
</tr></thead>
<tbody>
{''.join(rows_html)}
</tbody>
</table>
</body></html>"""


def export_personel_raporu_excel(
    path: str,
    people: List[Dict],
    codes: List[Dict],
    hedef: float,
    ay_adi: str,
    yil,
    gun_sayisi: int,
    tatil_gunleri,
    hafta_sonu_gunleri,
    usage_rows: List[List],
    tatil_rows: List[List],
) -> None:
    """Personel Raporu ekranini 3 sayfalik bir Excel dosyasina aktarir:
    Rapor (gun izelgesi), Vardiya Dagilimi, Resmi Tatilde Calisma.

    Renkler ekrandakiyle ayni (kod renkleri, hafta sonu/tatil basliklari) ki
    ciktiya bakan kisi ayni gorseli gorsun.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(bold=True, color="FFFFFF")
    weekend_fill = PatternFill("solid", fgColor=WEEKEND_HEADER_COLOR)
    holiday_fill = PatternFill("solid", fgColor=HOLIDAY_HEADER_COLOR)
    center = Alignment(horizontal="center")

    tatil_set = set(tatil_gunleri or [])
    hs_set = set(hafta_sonu_gunleri or [])

    # ---------------- 1) Rapor ----------------
    ws = wb.create_sheet("Rapor")
    ws["A1"] = f"PERSONEL RAPORU — {ay_adi} {yil}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Hedef Saat: {_fmt(hedef)}"
    ws["A2"].font = Font(italic=True, color="555555")

    header_row = 4
    basliklar = ["Ad Soyad", "Görev", "Bölüm", "L.G."] + [str(d) for d in range(1, gun_sayisi + 1)] + [
        "Toplam Saat", "Fark", "L.Ç."
    ]
    for col, title in enumerate(basliklar, start=1):
        c = ws.cell(header_row, col, title)
        c.font = header_font
        c.alignment = center
        gun_no = col - 4  # 4 sabit sutundan sonrasi gun
        if 1 <= gun_no <= gun_sayisi:
            if gun_no in tatil_set:
                c.fill = holiday_fill
            elif gun_no in hs_set:
                c.fill = weekend_fill
            else:
                c.fill = header_fill
        else:
            c.fill = header_fill

    kod_saat = {}
    for cd in codes:
        try:
            kod_saat[cd["kod"]] = float(cd.get("saat") or 0)
        except (TypeError, ValueError):
            kod_saat[cd["kod"]] = 0.0

    r = header_row + 1
    for p in people:
        ws.cell(r, 1, p.get("ad_soyad", ""))
        ws.cell(r, 2, p.get("gorev", ""))
        ws.cell(r, 3, p.get("bolum", ""))
        lg = float(p.get("lg") or 0)
        ws.cell(r, 4, lg).alignment = center
        toplam = 0.0
        for gun in range(1, gun_sayisi + 1):
            raw = p.get("days", {}).get(gun, "")
            cell = ws.cell(r, 4 + gun)
            cell.alignment = center
            if raw:
                cell.value = raw
                kod, override, farkli = parse_cell_code(raw)
                sa = parse_saat_araligi(raw)
                if sa is not None:
                    renk = CALISMA_SAAT_COLOR if sa[0] == "Ç" else IZIN_SAAT_COLOR
                elif farkli:
                    renk = FARKLI_BOLUM_COLOR
                else:
                    renk = CODE_COLORS.get(kod, UNKNOWN_CODE_COLOR)
                if renk != UNKNOWN_CODE_COLOR:
                    cell.fill = PatternFill("solid", fgColor=renk)
                katki = cell_hours(raw, kod_saat)
                if katki is not None:
                    toplam += katki
            else:
                # Bos hucre: hafta sonu/tatil sutunu renklensin (ekranla ayni)
                if gun in tatil_set:
                    cell.fill = holiday_fill
                elif gun in hs_set:
                    cell.fill = weekend_fill
            note = p.get("notes", {}).get(gun)
            if note:
                cell.comment = Comment(note, "Nobet Uygulamasi")
        fark = toplam - float(hedef or 0)
        for col, deger in ((4 + gun_sayisi + 1, toplam), (4 + gun_sayisi + 2, fark), (4 + gun_sayisi + 3, lg + fark)):
            cc = ws.cell(r, col, deger)
            cc.alignment = center
            cc.font = Font(bold=True)
        r += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 8
    for gun in range(1, gun_sayisi + 1):
        ws.column_dimensions[get_column_letter(4 + gun)].width = 5
    for i in range(1, 4):
        ws.column_dimensions[get_column_letter(4 + gun_sayisi + i)].width = 12
    ws.freeze_panes = ws.cell(header_row + 1, 5)

    # ---------------- 2) Vardiya Dagilimi ----------------
    ws2 = wb.create_sheet("Vardiya Dagilimi")
    for col, title in enumerate(["Kod", "Açıklama", "Saat", "Gün Sayısı", "Toplam Saat Katkısı"], start=1):
        c = ws2.cell(1, col, title)
        c.font = header_font
        c.fill = header_fill
    for i, satir in enumerate(usage_rows, start=2):
        for col, deger in enumerate(satir, start=1):
            ws2.cell(i, col, deger)
        kod = satir[0] if satir else ""
        if kod == "Ç (saatlik)":
            renk = CALISMA_SAAT_COLOR
        elif kod == "İ (saatlik)":
            renk = IZIN_SAAT_COLOR
        else:
            renk = CODE_COLORS.get(kod)
        if renk:
            ws2.cell(i, 1).fill = PatternFill("solid", fgColor=renk)
    for col_letter, width in zip("ABCDE", (10, 30, 10, 12, 20)):
        ws2.column_dimensions[col_letter].width = width

    # ---------------- 3) Resmi Tatilde Calisma ----------------
    ws3 = wb.create_sheet("Resmi Tatilde Calisma")
    for col, title in enumerate(["Ad Soyad", "Gün", "Kod", "Saat"], start=1):
        c = ws3.cell(1, col, title)
        c.font = header_font
        c.fill = holiday_fill
    for i, satir in enumerate(tatil_rows, start=2):
        for col, deger in enumerate(satir, start=1):
            ws3.cell(i, col, deger)
    for col_letter, width in zip("ABCD", (26, 8, 12, 10)):
        ws3.column_dimensions[col_letter].width = width
    if not tatil_rows:
        ws3.cell(2, 1, "Bu ayda resmi tatilde çalışma yok.")

    wb.save(path)


def _esc(v) -> str:
    return str(v if v is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _fmt(x) -> str:
    try:
        x = float(x)
    except (TypeError, ValueError):
        return str(x)
    if x == int(x):
        return str(int(x))
    return str(round(x, 2))
