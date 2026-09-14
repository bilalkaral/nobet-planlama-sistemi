"""Veritabani <-> uygulama veri donusumu (okuma yolu).

nobet_db.SupabaseClient ile DB'den kod tanimlari, personel ve bir ayin
cizelgesini cekip, nobet_core/arayuzun bekledigi sozluk yapilarina cevirir.

DB modeli (personeller + ay_personel + nobet_gunleri) uygulamanin tek-satir
"schedule row" modeline burada birlestirilir. Yazma sirasinda hangi DB
kaydina dokunulacagi bilinsin diye her satira gizli anahtarlar eklenir:
    _personel_id : personeller.id
    _ap_id       : ay_personel.id (o ay icin; yoksa None)
    _bolum_id    : personelin bolum_id'si (yetki kontrolu icin)
"""
from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from nobet_core import AY_ADLARI, DAY_COUNT, cell_hours, parse_cell_code
from nobet_config import SADECE_EKSIK_DEVRET
from nobet_db import SupabaseClient

# nobet_gunleri sorgusunda tek istekte kac ay_personel_id sorulacak (URL uzunlugu
# sinirina takilmamak icin parcalanir).
_IN_CHUNK = 80


def load_codes(db: SupabaseClient) -> List[Dict]:
    """kod_tanimlari -> [{kod, aciklama, saat}] (sira sutununa gore)."""
    rows = db.select(
        "kod_tanimlari",
        params={"select": "kod,aciklama,saat,sira", "order": "sira.asc"},
    )
    codes = []
    for r in rows:
        codes.append({
            "kod": r.get("kod") or "",
            "aciklama": r.get("aciklama") or "",
            "saat": r.get("saat") or 0,
        })
    return codes


def load_bolumler(db: SupabaseClient) -> Dict[str, str]:
    """bolumler -> {id: ad}."""
    rows = db.select("bolumler", params={"select": "id,ad"})
    return {r["id"]: r["ad"] for r in rows}


def load_personnel(db: SupabaseClient, bolum_map: Optional[Dict[str, str]] = None) -> List[Dict]:
    """personeller -> Personeller sekmesinin bekledigi sozluk listesi."""
    if bolum_map is None:
        bolum_map = load_bolumler(db)
    rows = db.select(
        "personeller",
        params={
            "select": "id,sicil_no,ad_soyad,unvan,bolum_id,aktif",
            "order": "ad_soyad.asc",
        },
    )
    people = []
    for r in rows:
        people.append({
            "sicil_no": r.get("sicil_no") or "",
            "ad_soyad": r.get("ad_soyad") or "",
            "birim": bolum_map.get(r.get("bolum_id"), "") if r.get("bolum_id") else "",
            "unvan": r.get("unvan") or "",
            "aktif": "Aktif" if r.get("aktif", True) else "Pasif",
            "_id": r["id"],
            "_bolum_id": r.get("bolum_id"),
        })
    return people


def _chunked(seq: List, n: int):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def load_month(
    db: SupabaseClient,
    yil: int,
    ay: int,
    bolum_map: Optional[Dict[str, str]] = None,
) -> Tuple[Dict, List[Dict]]:
    """Bir ayin ayarlarini ve cizelge satirlarini DB'den okur.

    Doner: (settings_partial, schedule_rows)
      settings_partial: {yil, ay_no, ay_adi, hedef_saat, resmi_tatiller, _ay_id}
      schedule_rows   : her aktif personel icin bir satir (gun verisi o aya ait)
    """
    if bolum_map is None:
        bolum_map = load_bolumler(db)

    # 1) Ay satiri (yoksa varsayilan; _ay_id None kalir)
    ay_rows = db.select(
        "aylar",
        params={
            "select": "id,yil,ay_no,hedef_saat,resmi_tatiller",
            "yil": f"eq.{yil}",
            "ay_no": f"eq.{ay}",
        },
    )
    if ay_rows:
        ay_row = ay_rows[0]
        ay_id = ay_row["id"]
        hedef = ay_row.get("hedef_saat") or 210
        tatiller = list(ay_row.get("resmi_tatiller") or [])
    else:
        ay_id = None
        hedef = 210
        tatiller = []

    settings = {
        "yil": yil,
        "ay_no": ay,
        "ay_adi": AY_ADLARI[ay - 1] if 1 <= ay <= 12 else "",
        "hedef_saat": hedef,
        "resmi_tatiller": tatiller,
        "_ay_id": ay_id,
    }

    # 2) Personel listesi (aktif olanlar cizelgede gorunur)
    personeller = db.select(
        "personeller",
        params={
            "select": "id,ad_soyad,unvan,bolum_id,aktif",
            "aktif": "eq.true",
            "order": "bolum_id.asc.nullslast,ad_soyad.asc",
        },
    )

    # 3) Bu aya ait ay_personel (lg + ap_id) ve nobet_gunleri
    ap_by_personel: Dict[str, Dict] = {}
    gun_by_ap: Dict[str, Dict[int, Tuple[str, Optional[str]]]] = {}
    if ay_id:
        ap_rows = db.select(
            "ay_personel",
            params={"select": "id,personel_id,lg", "ay_id": f"eq.{ay_id}"},
        )
        ap_by_personel = {r["personel_id"]: r for r in ap_rows}
        ap_ids = [r["id"] for r in ap_rows]

        for parca in _chunked(ap_ids, _IN_CHUNK):
            liste = ",".join(parca)
            gun_rows = db.select(
                "nobet_gunleri",
                params={
                    "select": "ay_personel_id,gun,kod,aciklama",
                    "ay_personel_id": f"in.({liste})",
                },
            )
            for g in gun_rows:
                apid = g["ay_personel_id"]
                gun = g["gun"]
                kod = (g.get("kod") or "").strip()
                aciklama = g.get("aciklama")
                if kod == "" and not aciklama:
                    continue  # temizlenmis hucre - bos gecilir
                gun_by_ap.setdefault(apid, {})[gun] = (kod, aciklama)

    # 4) Cizelge satirlarini olustur
    schedule: List[Dict] = []
    for p in personeller:
        pid = p["id"]
        ap = ap_by_personel.get(pid)
        ap_id = ap["id"] if ap else None
        lg = float(ap.get("lg") or 0) if ap else 0.0

        days: Dict[int, str] = {}
        notes: Dict[int, str] = {}
        if ap_id and ap_id in gun_by_ap:
            for gun, (kod, aciklama) in gun_by_ap[ap_id].items():
                if 1 <= gun <= DAY_COUNT:
                    if kod:
                        days[gun] = kod
                    if aciklama:
                        notes[gun] = aciklama

        schedule.append({
            "ad_soyad": p.get("ad_soyad") or "",
            "gorev": p.get("unvan") or "",
            "bolum": bolum_map.get(p.get("bolum_id"), "") if p.get("bolum_id") else "",
            "lg": lg,
            "days": days,
            "notes": notes,
            "_personel_id": pid,
            "_ap_id": ap_id,
            "_bolum_id": p.get("bolum_id"),
        })

    return settings, schedule


# ======================================================================
#  YAZMA yolu - arayuzdeki degisiklikleri DB'ye aktarir
# ======================================================================

def ensure_ay(db: SupabaseClient, yil: int, ay: int) -> str:
    """Ay satirini bulur; yoksa olusturur ve id'sini doner.

    Yeni ay olusturmak RLS geregi basemsire yetkisi ister; sorumlu var olan
    aya (import'la olusmus Tem/Agu/Eyl gibi) yazabilir ama yeni ay acamaz.
    """
    rows = db.select(
        "aylar", params={"select": "id", "yil": f"eq.{yil}", "ay_no": f"eq.{ay}"}
    )
    if rows:
        return rows[0]["id"]
    created = db.insert(
        "aylar",
        [{"yil": yil, "ay_no": ay, "hedef_saat": 210, "resmi_tatiller": []}],
        upsert=True, on_conflict="yil,ay_no",
    )
    return created[0]["id"]


def ensure_ay_personel(db: SupabaseClient, ay_id: str, personel_id: str) -> str:
    """(ay, personel) satirini bulur; yoksa olusturur. id'yi doner.

    ONEMLI: upsert ile lg gondermiyoruz; var olan satiri upsert etsek lg=0
    yazip devreden L.G.'yi silerdik. O yuzden once select, yoksa insert.
    """
    rows = db.select(
        "ay_personel",
        params={"select": "id", "ay_id": f"eq.{ay_id}", "personel_id": f"eq.{personel_id}"},
    )
    if rows:
        return rows[0]["id"]
    created = db.insert("ay_personel", [{"ay_id": ay_id, "personel_id": personel_id, "lg": 0}])
    return created[0]["id"]


def save_cell(db: SupabaseClient, ap_id: str, gun: int, kod: str, aciklama: Optional[str] = None) -> None:
    """Bir gun hucresini yazar (upsert). Hucre temizlenirken satir silinmez,
    kod='' yazilir (delta yoklamasi silmeyi goremezdi - bkz. schema.sql)."""
    db.insert(
        "nobet_gunleri",
        [{"ay_personel_id": ap_id, "gun": gun, "kod": kod or "", "aciklama": aciklama or None}],
        upsert=True, on_conflict="ay_personel_id,gun",
    )


def save_lg(db: SupabaseClient, ap_id: str, lg: float) -> None:
    """ay_personel.lg (devreden L.G.) gunceller."""
    db.update("ay_personel", {"id": f"eq.{ap_id}"}, {"lg": lg})


def add_personel(
    db: SupabaseClient, ad_soyad: str, unvan: str = "", bolum_id: Optional[str] = None,
    sicil_no: str = "",
) -> Dict:
    """Yeni personel ekler, olusan kaydi doner."""
    payload = {
        "ad_soyad": ad_soyad, "unvan": unvan or None,
        "bolum_id": bolum_id, "sicil_no": sicil_no or None, "aktif": True,
    }
    rows = db.insert("personeller", [payload])
    return rows[0] if rows else {}


def delete_personel(db: SupabaseClient, personel_id: str) -> None:
    """Personeli siler. ON DELETE CASCADE ile ay_personel + nobet_gunleri de
    silinir."""
    db.delete("personeller", {"id": f"eq.{personel_id}"})


# ======================================================================
#  L.G. DEVRI - bir ayin L.G.'si onceki ayin L.C.'sinden gelir
# ======================================================================
#  L.C. = L.G. + Fark,  Fark = Toplam - Hedef.
#  Bir sonraki ayin L.G.'si = bu ayin L.C.'si (devreden bakiye).
#  Onceki ayda kisinin verisi yoksa L.G. = 0.

def _person_totals(db: SupabaseClient, ay_id: str, code_hours: Dict[str, float]):
    """Bir aydaki her personelin toplam saatini ve veri girilip girilmedigini
    hesaplar. Doner: ({personel_id: toplam}, {personel_id: veri_var_mi})."""
    ap = db.select(
        "ay_personel", params={"select": "id,personel_id", "ay_id": f"eq.{ay_id}"}
    )
    apid_to_pid = {r["id"]: r["personel_id"] for r in ap}
    totals: Dict[str, float] = {}
    hasdata: Dict[str, bool] = {}
    for parca in _chunked(list(apid_to_pid), _IN_CHUNK):
        liste = ",".join(parca)
        rows = db.select(
            "nobet_gunleri",
            params={"select": "ay_personel_id,gun,kod", "ay_personel_id": f"in.({liste})"},
        )
        for g in rows:
            pid = apid_to_pid.get(g["ay_personel_id"])
            if pid is None:
                continue
            raw = g.get("kod") or ""
            kod, override, _ = parse_cell_code(raw)
            if not kod and override is None:
                continue
            hasdata[pid] = True
            # cell_hours saat araligini (Ç ekler / İ duser) da hesaba katar.
            hrs = cell_hours(raw, code_hours)
            totals[pid] = totals.get(pid, 0.0) + (hrs or 0.0)
    return totals, hasdata


def latest_change_ts(db: SupabaseClient) -> Optional[str]:
    """nobet_gunleri'ndeki en son degisiklik zamani (ISO). Canli senkronun
    baslangic isaretidir; bundan SONRA olan degisiklikler cekilir."""
    rows = db.select(
        "nobet_gunleri",
        params={"select": "updated_at", "order": "updated_at.desc", "limit": "1"},
    )
    return rows[0]["updated_at"] if rows else None


def poll_changes(db: SupabaseClient, since_iso: str, limit: int = 1000) -> List[Dict]:
    """since_iso'dan (ISO zaman damgasi) SONRA degisen gun hucrelerini doner.
    Tum aylardaki degisiklikleri getirir; cagiran taraf yalniz acik aya ait
    olanlari uygular. `nobet_gunleri_updated_idx` sayesinde hizli."""
    return db.select(
        "nobet_gunleri",
        params={
            "select": "ay_personel_id,gun,kod,aciklama,updated_at",
            "updated_at": f"gt.{since_iso}",
            "order": "updated_at.asc",
            "limit": str(limit),
        },
    )


def carryover_lg_map(
    db: SupabaseClient, yil: int, ay: int, code_hours: Dict[str, float], _depth: int = 0
) -> Dict[str, float]:
    """(yil, ay) ayi icin {personel_id: L.G.} doner.

    L.G. = onceki ayin L.C.'si. Onceki ay yoksa ya da kisinin orada verisi
    yoksa 0. Zincir: onceki ayin L.C.'si de kendi L.G.'si (bir onceki aydan)
    + Fark'tan olusur; bu yuzden ozyinelemeli. Var olan aylar boyunca geriye
    gider, ilk ayda durur.

    SADECE_EKSIK_DEVRET acikken yalnizca EKSIK mesai (negatif L.C.) devreder;
    ayi fazla mesai ile kapatan kisi sonraki aya 0 ile baslar. Kural zincirin
    her adiminda uygulanir, yani fazla mesai birikip sonraki aylarin eksigini
    kapatmaz."""
    if _depth > 24:
        return {}
    pyil, pay = (yil, ay - 1) if ay > 1 else (yil - 1, 12)
    prev = db.select(
        "aylar",
        params={"select": "id,hedef_saat", "yil": f"eq.{pyil}", "ay_no": f"eq.{pay}"},
    )
    if not prev:
        return {}  # onceki ay yok -> devir yok, herkes 0
    prev_id = prev[0]["id"]
    prev_hedef = float(prev[0].get("hedef_saat") or 0)

    prev_lg = carryover_lg_map(db, pyil, pay, code_hours, _depth + 1)
    totals, hasdata = _person_totals(db, prev_id, code_hours)

    result: Dict[str, float] = {}
    for pid in hasdata:  # sadece onceki ayda verisi olanlar devreder
        toplam = totals.get(pid, 0.0)
        lc = prev_lg.get(pid, 0.0) + (toplam - prev_hedef)
        if SADECE_EKSIK_DEVRET and lc >= 0:
            continue  # ayi fazla/tam kapatmis - devir yok, L.G. 0 baslar
        result[pid] = lc
    return result
