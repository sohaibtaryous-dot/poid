# -*- coding: utf-8 -*-
"""IMPORT — référentiel NGP enregistré dans le site.

- Chaque tableau de référence déposé remplace le référentiel enregistré (dernière version = celle qui fait foi).
- Sans dépôt, le dernier référentiel enregistré est utilisé automatiquement.
- « Ajouts du site » : NGP saisis dans le site pour des références absentes du tableau ; ils complètent le tableau
  (le tableau déposé reste prioritaire) et sont gardés pour les dossiers suivants.
Tout est copié sur Supabase (si configuré) : rien n'est perdu au redémarrage."""

from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path

import database
import persist

DIR = Path(database.DATA_DIR) / "referentiel"
DATA_F, META_F, ADD_F = DIR / "current.bin", DIR / "meta.json", DIR / "ajouts.json"
SHEET_ADD = "Ajouts du site"


def _key(p: Path) -> str:
    return f"referentiel/{p.name}"


def _read(p: Path) -> bytes | None:
    if p.is_file():
        return p.read_bytes()
    data = persist.get_file(_key(p))                     # disque effacé (redémarrage) : copie Supabase
    if data:
        DIR.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)
    return data


def _write(p: Path, data: bytes):
    DIR.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    persist.put_file(_key(p), data)


def meta() -> dict:
    raw = _read(META_F)
    try:
        return json.loads(raw.decode("utf-8")) if raw else {}
    except Exception:  # noqa: BLE001
        return {}


def current() -> tuple[bytes, str] | None:
    """(contenu, nom de fichier) du référentiel enregistré, ou None."""
    m = meta()
    data = _read(DATA_F) if m else None
    return (data, m.get("filename", "referentiel.xlsx")) if data else None


def save(data: bytes, filename: str, count: int, user: str = ""):
    """Enregistre le tableau déposé (s'il a changé) comme référentiel courant."""
    sha = hashlib.sha256(data).hexdigest()
    m = meta()
    if m.get("sha") == sha:
        return False
    _write(DATA_F, data)
    _write(META_F, json.dumps({"filename": filename, "sha": sha, "count": count, "user": user,
                               "date": datetime.now().strftime("%d/%m/%Y %H:%M")}, ensure_ascii=False).encode("utf-8"))
    return True


def additions() -> dict:
    raw = _read(ADD_F)
    try:
        return json.loads(raw.decode("utf-8")) if raw else {}
    except Exception:  # noqa: BLE001
        return {}


def add(rows: list[dict], user: str = "") -> int:
    """rows : [{"ref": …, "ngp": …, "desc": …}] — NGP saisis dans le site (8 ou 10 chiffres)."""
    import app as core
    cur = additions()
    n = 0
    for r in rows:
        ref, ngp = core.norm_ref(r.get("ref")), core.normalize_ngp(r.get("ngp"))
        if not ref or len(ngp) not in (8, 10) or not ngp.isdigit():
            continue
        cur[ref] = {"ref": str(r.get("ref")).strip(), "ngp": ngp, "desc": str(r.get("desc") or "")[:80], "user": user,
                    "date": datetime.now().strftime("%d/%m/%Y %H:%M")}
        n += 1
    if n:
        _write(ADD_F, json.dumps(cur, ensure_ascii=False, indent=0).encode("utf-8"))
    return n


def remove(refs: list[str]) -> int:
    cur = additions()
    n = sum(1 for r in refs if cur.pop(r, None) is not None)
    if n:
        _write(ADD_F, json.dumps(cur, ensure_ascii=False, indent=0).encode("utf-8"))
    return n


def merge(refdb: dict, adds: dict | None = None) -> dict:
    """Ajoute les NGP saisis dans le site aux références absentes du tableau (le tableau reste prioritaire)."""
    import app as core
    adds = additions() if adds is None else adds
    for key, a in adds.items():
        if key in refdb["exact"]:
            continue
        refdb["exact"][key] = a["ngp"]
        refdb.setdefault("sheet_names", {})[key] = SHEET_ADD
        ck = core.compact_ref(key)
        if ck:
            refdb["compact"].setdefault(ck, a["ngp"])
            refdb["ocr"].setdefault(core.ocr_ref(key), a["ngp"])
    refdb["count"] = len(refdb["exact"])
    refdb["added"] = len(adds)
    return refdb


def refdb_for(ref_file, log: list, user: str = "") -> dict:
    """Référentiel d'un dossier : tableau déposé (enregistré au passage) sinon dernier enregistré, + ajouts du site."""
    import app as core
    empty = {"exact": {}, "compact": {}, "ocr": {}, "sheet_names": {}, "count": 0}
    src = ""
    if ref_file is not None:
        data, name = ref_file.getvalue(), ref_file.name
        refdb = core.load_reference(data, name, log)
        if refdb.get("count"):
            try:
                if save(data, name, refdb["count"], user):
                    log.append(f"💾 Référentiel « {name} » enregistré : il sera utilisé automatiquement pour les prochains dossiers.")
            except Exception as exc:  # noqa: BLE001
                log.append(f"⚠️ Référentiel non enregistré ({type(exc).__name__}: {exc}).")
        src = name
    else:
        try:
            cur = current()
        except Exception:  # noqa: BLE001
            cur = None
        if cur:
            refdb = core.load_reference(cur[0], cur[1], log)
            src = cur[1]
            log.append(f"📗 Référentiel enregistré utilisé : « {cur[1]} » ({meta().get('date', '')}).")
        else:
            refdb = empty
    try:
        refdb = merge(refdb)
    except Exception:  # noqa: BLE001
        pass
    refdb["source"] = src
    return refdb
