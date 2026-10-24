# -*- coding: utf-8 -*-
"""Sauvegarde durable sur Supabase Storage (gratuit) : sur Streamlit Community Cloud le disque est effacé à chaque redémarrage
(Reboot, mise à jour GitHub, mise en veille) ; ce module garde la base SQLite (comptes, paramètres, attributions) et les fichiers
attribués dans un bucket privé, et les restaure au démarrage.

Configuration (Secrets Streamlit / variables d'environnement) :
  SUPABASE_URL     https://xxxx.supabase.co
  SUPABASE_KEY     clé « service_role » (Project Settings → API) — secrète, jamais dans GitHub
  SUPABASE_BUCKET  nom du bucket privé (défaut : timar-data ; créé automatiquement s'il n'existe pas)
Sans ces variables : rien ne change (disque local seulement)."""

from __future__ import annotations

import os
import sqlite3
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

DB_KEY = "timar.db"
_lock = threading.Lock()                 # une sauvegarde à la fois
_sched = threading.Lock()                # planification (très court : ne bloque jamais l'application)
_timer: threading.Timer | None = None
_dirty = False
_state = {"last_ok": "", "last_error": ""}


def _cfg() -> tuple[str, str, str]:
    url = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
    key = os.environ.get("SUPABASE_KEY", "").strip()
    return url, key, os.environ.get("SUPABASE_BUCKET", "").strip() or "timar-data"


def enabled() -> bool:
    url, key, _ = _cfg()
    return bool(url and key)


def status() -> dict:
    return {"enabled": enabled(), **_state}


def _request(method: str, path: str, data: bytes | None = None, headers: dict | None = None, timeout: int = 25) -> tuple[int, bytes]:
    url, key, _ = _cfg()
    req = urllib.request.Request(f"{url}/storage/v1/{path}", data=data, method=method)
    if not key.startswith("sb_"):                       # clé legacy (JWT « eyJ… ») : aussi en Bearer ; clés « sb_secret_… » : en-tête apikey seul
        req.add_header("Authorization", f"Bearer {key}")
    req.add_header("apikey", key)
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read()
    except Exception as exc:  # noqa: BLE001
        _state["last_error"] = f"{type(exc).__name__}: {exc}"
        return 0, b""


def _obj(key: str) -> str:
    return f"object/{urllib.parse.quote(_cfg()[2])}/{urllib.parse.quote(key.replace(os.sep, '/'), safe='/')}"


def ensure_bucket() -> bool:
    if not enabled():
        return False
    bucket = _cfg()[2]
    code, _ = _request("GET", f"bucket/{urllib.parse.quote(bucket)}")
    if code == 200:
        return True
    code, body = _request("POST", "bucket", f'{{"id":"{bucket}","name":"{bucket}","public":false}}'.encode(),
                          {"Content-Type": "application/json"})
    return code in (200, 201) or b"already exists" in body.lower()


def put_file(key: str, data: bytes) -> bool:
    """Envoie (ou remplace) un fichier."""
    if not enabled():
        return False
    ensure_bucket()
    code, body = _request("POST", _obj(key), data, {"x-upsert": "true", "Content-Type": "application/octet-stream"})
    if code in (200, 201):
        _state["last_ok"] = key
        _state["last_error"] = ""
        return True
    _state["last_error"] = f"envoi de {key} : HTTP {code} {body[:120]!r}"
    return False


def get_file(key: str) -> bytes | None:
    if not enabled():
        return None
    code, body = _request("GET", _obj(key))
    _state["last_get_code"] = code
    return body if code == 200 else None


def delete_file(key: str) -> bool:
    if not enabled():
        return False
    code, _ = _request("DELETE", _obj(key))
    return code in (200, 204)


# ----------------------------------------------------------------------------- base SQLite
def backup_db(db_path: Path) -> bool:
    """Copie cohérente de la base (API de sauvegarde SQLite) puis envoi."""
    if not enabled() or not Path(db_path).is_file() or blocked():
        return False
    with _lock:
        tmp = Path(tempfile.mkdtemp()) / "snapshot.db"
        try:
            src, dst = sqlite3.connect(db_path, timeout=15), sqlite3.connect(tmp)
            try:
                src.backup(dst)
            finally:
                dst.close()
                src.close()
            return put_file(DB_KEY, tmp.read_bytes())
        except Exception as exc:  # noqa: BLE001
            _state["last_error"] = f"sauvegarde : {type(exc).__name__}: {exc}"
            return False
        finally:
            try:
                tmp.unlink(missing_ok=True)
                tmp.parent.rmdir()
            except OSError:
                pass


def restore_db(db_path: Path, force: bool = False) -> bool:
    """Au démarrage : si la base locale n'existe pas (disque effacé), la récupère depuis Supabase.
    Supabase injoignable (réseau, panne) : les envois sont bloqués pour ne JAMAIS écraser la sauvegarde par une base vide ;
    une nouvelle tentative est faite au prochain démarrage de session."""
    db_path = Path(db_path)
    if not enabled() or (db_path.exists() and not force):
        return False
    data = get_file(DB_KEY)
    code = _state.get("last_get_code")
    if not data:
        # 400/404 = aucune sauvegarde encore (1re utilisation) : on peut démarrer à vide ; sinon : blocage
        _state["blocked"] = code not in (400, 404)
        if _state["blocked"]:
            _state["last_error"] = f"restauration impossible (HTTP {code}) : sauvegarde bloquée pour protéger les données"
        return False
    db_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = db_path.with_suffix(".restore")
    tmp.write_bytes(data)
    tmp.replace(db_path)
    _state["blocked"] = False
    return True


def blocked() -> bool:
    return bool(_state.get("blocked"))


def _run(db_path: Path, delay: float):
    global _dirty, _timer
    with _sched:
        _dirty = False
    backup_db(db_path)
    with _sched:                                   # écriture survenue pendant l'envoi : on recommence
        if _dirty:
            _timer = threading.Timer(delay, _run, args=(db_path, delay))
            _timer.daemon = True
            _timer.start()


def mark_dirty(db_path: Path, delay: float = 2.0):
    """Planifie une sauvegarde de la base ~2 s après une écriture (regroupe les écritures rapprochées, ne bloque jamais l'application)."""
    global _timer, _dirty
    if not enabled():
        return
    with _sched:
        _dirty = True
        if _timer is None or not _timer.is_alive():
            _timer = threading.Timer(delay, _run, args=(db_path, delay))
            _timer.daemon = True
            _timer.start()
