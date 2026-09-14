"""Supabase baglanti katmani - httpx ile dogrudan REST/GoTrue.

supabase-py PAKETI KULLANILMIYOR. O paket pydantic/gotrue/realtime/storage3
gibi kalabalik bir bagimlilik agaci getirir; PyInstaller'da sorun cikarir ve
exe'yi sisirir. Ihtiyacimiz olan tek sey giris yapmak ve tablo okuyup yazmak;
bunu zaten kurulu olan httpx ile yapiyoruz.

Anlik senkron icin websocket YERINE periyodik yoklama (delta cekme) kullaniyoruz;
hastane aglari websocket'i sik sik kapatir, yoklama her yerde calisir. Bu modul
sadece veri erisimini saglar; yoklama dongusunu arayuz (nobet_app) kurar.
"""
from __future__ import annotations

import threading
import time
from typing import Any, Dict, List, Optional

import httpx

from nobet_config import EMAIL_DOMAIN, SUPABASE_ANON_KEY, SUPABASE_URL


class DbError(Exception):
    """Veritabani/istek hatasi - kullaniciya gosterilebilir Turkce mesaj tasir."""


class AuthError(DbError):
    """Giris/kimlik dogrulama hatasi."""


# Yetki reddi (RLS) hatalarini ayirt etmek icin PostgREST kodlari.
_YETKI_KODLARI = {"42501", "PGRST301", "PGRST116"}


def _readable_error(resp: httpx.Response) -> str:
    """Supabase hata govdesinden okunabilir bir mesaj cikarir."""
    try:
        data = resp.json()
    except Exception:
        return f"Sunucu hatasi ({resp.status_code})."
    if isinstance(data, dict):
        # GoTrue: {"msg": ...} / {"error_description": ...}
        # PostgREST: {"message": ..., "code": ...}
        msg = (
            data.get("msg")
            or data.get("message")
            or data.get("error_description")
            or data.get("error")
        )
        code = str(data.get("code") or "")
        if code in _YETKI_KODLARI or resp.status_code in (401, 403):
            return "Bu islem icin yetkiniz yok."
        if msg:
            return str(msg)
    return f"Sunucu hatasi ({resp.status_code})."


class SupabaseClient:
    """Tek kullanicinin oturumunu tutan istemci.

    Kullanim:
        db = SupabaseClient()
        db.sign_in("bashemsire", "sifre")
        db.profile        -> {"rol": ..., "bolum_id": ..., ...}
        db.select("personeller", params={"select": "*"})
    """

    def __init__(self) -> None:
        self._http = httpx.Client(base_url=SUPABASE_URL, timeout=20.0)
        self._access_token: Optional[str] = None
        self._refresh_token: Optional[str] = None
        self._expires_at: float = 0.0
        self.user_id: Optional[str] = None
        self.profile: Optional[Dict[str, Any]] = None
        # Yazma islemleri arka plan is parcaciginda calisiyor (arayuz donmasin);
        # httpx.Client es zamanli kullanim icin guvenli degil, o yuzden tum
        # istekleri bu kilitle siraya sokuyoruz.
        self._lock = threading.RLock()

    # ------------------------------------------------------------------ oturum
    @property
    def logged_in(self) -> bool:
        return self._access_token is not None

    def _auth_headers(self, use_user_token: bool = True) -> Dict[str, str]:
        token = self._access_token if (use_user_token and self._access_token) else SUPABASE_ANON_KEY
        return {
            "apikey": SUPABASE_ANON_KEY,
            "Authorization": f"Bearer {token}",
        }

    def sign_in(self, username: str, password: str) -> Dict[str, Any]:
        """Kullanici adi + sifre ile giris. Basarida profili yukler ve dondurur."""
        username = (username or "").strip()
        if not username or not password:
            raise AuthError("Kullanici adi ve sifre bos birakilamaz.")
        email = username if "@" in username else f"{username}@{EMAIL_DOMAIN}"
        try:
            resp = self._http.post(
                "/auth/v1/token",
                params={"grant_type": "password"},
                headers={"apikey": SUPABASE_ANON_KEY, "Authorization": f"Bearer {SUPABASE_ANON_KEY}"},
                json={"email": email, "password": password},
            )
        except httpx.RequestError as exc:
            raise AuthError(f"Sunucuya baglanilamadi:\n{exc}") from exc
        if resp.status_code != 200:
            body = resp.json() if resp.headers.get("content-type", "").startswith("application/json") else {}
            if body.get("error_code") == "invalid_credentials" or resp.status_code == 400:
                raise AuthError("Kullanici adi veya sifre hatali.")
            raise AuthError(_readable_error(resp))

        data = resp.json()
        self._store_session(data)
        self._load_profile()
        return self.profile

    def _store_session(self, data: Dict[str, Any]) -> None:
        self._access_token = data.get("access_token")
        self._refresh_token = data.get("refresh_token")
        expires_in = data.get("expires_in") or 3600
        self._expires_at = time.time() + float(expires_in)
        user = data.get("user") or {}
        self.user_id = user.get("id")

    def _refresh_session(self) -> bool:
        """Suresi dolan access token'i yeniler. Basarisizsa False."""
        if not self._refresh_token:
            return False
        try:
            resp = self._http.post(
                "/auth/v1/token",
                params={"grant_type": "refresh_token"},
                headers={"apikey": SUPABASE_ANON_KEY, "Authorization": f"Bearer {SUPABASE_ANON_KEY}"},
                json={"refresh_token": self._refresh_token},
            )
        except httpx.RequestError:
            return False
        if resp.status_code != 200:
            return False
        self._store_session(resp.json())
        return True

    def _ensure_token(self) -> None:
        """Token bitmek uzereyse (60 sn kala) proaktif yeniler."""
        if self._access_token and time.time() > self._expires_at - 60:
            self._refresh_session()

    def sign_out(self) -> None:
        self._access_token = None
        self._refresh_token = None
        self._expires_at = 0.0
        self.user_id = None
        self.profile = None

    def _load_profile(self) -> None:
        rows = self.select(
            "profiller",
            params={"select": "id,kullanici_adi,ad_soyad,rol,bolum_id", "id": f"eq.{self.user_id}"},
        )
        if not rows:
            raise AuthError(
                "Giris yapildi ama profil bulunamadi. Yoneticinizle iletisime gecin "
                "(profiller tablosunda kaydiniz olmayabilir)."
            )
        self.profile = rows[0]

    # ------------------------------------------------------------------ roller
    @property
    def rol(self) -> str:
        return (self.profile or {}).get("rol", "personel")

    @property
    def bolum_id(self) -> Optional[str]:
        return (self.profile or {}).get("bolum_id")

    @property
    def is_bashemsire(self) -> bool:
        return self.rol == "bashemsire"

    @property
    def is_sorumlu(self) -> bool:
        return self.rol == "sorumlu"

    def can_edit_bolum(self, bolum_id: Optional[str]) -> bool:
        """Bu istemci verilen bolum uzerinde yazma yetkisine sahip mi?

        NOT: Asil zorlama veritabaninda (RLS). Bu sadece arayuzu kilitlemek,
        yani kullaniciya yetkisi olmayan hucreyi hic acmamak icin.
        """
        if self.is_bashemsire:
            return True
        if self.is_sorumlu:
            return bolum_id is not None and bolum_id == self.bolum_id
        return False

    # ------------------------------------------------------------- REST erisim
    def _request(self, method: str, path: str, **kwargs) -> httpx.Response:
        # Kilit: ana is parcacigi (okuma) ile arka plan yazici ayni anda
        # httpx.Client'a dokunmasin.
        with self._lock:
            self._ensure_token()
            headers = self._auth_headers()
            headers.update(kwargs.pop("headers", {}))
            try:
                resp = self._http.request(method, path, headers=headers, **kwargs)
            except httpx.RequestError as exc:
                raise DbError(f"Sunucuya baglanilamadi:\n{exc}") from exc
            # Token yeni suresi dolduysa bir kez yenileyip tekrar dene.
            if resp.status_code == 401 and self._refresh_session():
                headers = self._auth_headers()
                headers.update(kwargs.get("headers", {}))
                try:
                    resp = self._http.request(method, path, headers=headers, **kwargs)
                except httpx.RequestError as exc:
                    raise DbError(f"Sunucuya baglanilamadi:\n{exc}") from exc
            return resp

    def select(self, table: str, params: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
        resp = self._request("GET", f"/rest/v1/{table}", params=params or {})
        if resp.status_code != 200:
            raise DbError(_readable_error(resp))
        return resp.json()

    def insert(
        self, table: str, rows: Any, *, upsert: bool = False, on_conflict: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        prefer = ["return=representation"]
        if upsert:
            prefer.append("resolution=merge-duplicates")
        params: Dict[str, Any] = {}
        if on_conflict:
            params["on_conflict"] = on_conflict
        resp = self._request(
            "POST", f"/rest/v1/{table}", params=params, json=rows,
            headers={"Prefer": ",".join(prefer), "Content-Type": "application/json"},
        )
        if resp.status_code not in (200, 201):
            raise DbError(_readable_error(resp))
        return resp.json() if resp.content else []

    def update(self, table: str, params: Dict[str, Any], values: Dict[str, Any]) -> List[Dict[str, Any]]:
        resp = self._request(
            "PATCH", f"/rest/v1/{table}", params=params, json=values,
            headers={"Prefer": "return=representation", "Content-Type": "application/json"},
        )
        if resp.status_code not in (200, 204):
            raise DbError(_readable_error(resp))
        return resp.json() if resp.content else []

    def delete(self, table: str, params: Dict[str, Any]) -> None:
        resp = self._request("DELETE", f"/rest/v1/{table}", params=params)
        if resp.status_code not in (200, 204):
            raise DbError(_readable_error(resp))

    def close(self) -> None:
        try:
            self._http.close()
        except Exception:
            pass
