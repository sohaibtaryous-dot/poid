# -*- coding: utf-8 -*-
"""EXPORT — poids net lus dans la packing list (unités de manutention).

Une packing list Alstom liste des unités de manutention (carton, palette…) avec « Poids total » (brut) et « Poids de chargement » (net),
puis les articles qu'elles contiennent. Quand une unité ne contient qu'UNE référence, le poids net de cette référence est connu avec
certitude (pas d'estimation) ; quand elle en contient plusieurs, le poids de l'unité n'est connu que pour le groupe.

  hu_map(docs)   -> {n° BL: [{"hu", "net", "gross", "items": [(référence, quantité), …]}, …]}
  attach(df, m)  -> copie du tableau BL avec les colonnes `_hu` (unité de manutention de la ligne) et `_hu_net` (poids net de l'unité)
                    que ventilation_export.build_ventilation utilise à la place de l'estimation « au prorata de la valeur ».
"""

from __future__ import annotations

import re

import pandas as pd

BL_NO_RE = re.compile(r"(?:bon\s+de\s+livraison|delivery\s+note)\s*(?:n[°o]\.?)?\s*[:#]?\s*(\d{6,12})\b", re.I)
TABLE_RE = re.compile(r"(?:unit[ée]\s+de\s+(?:manutention|article)|handling\s+unit)", re.I)
HU_RE = re.compile(r"^\s*(\d{6,8})\s+[A-Za-zÀ-ÿ]")
HU_INLINE = re.compile(r"(\d[\d.,]*)\s*KG\s+(\d[\d.,]*)\s*KG\s*$", re.I)
W_LINE = re.compile(r"^\s*(\d[\d.,]*)\s+(\d[\d.,]*)\s*$")
MAT_RE = re.compile(r"^\s*([A-Z]{1,4}\d{6,})\s+.*?\s(\d+(?:[.,]\d+)?)\s*/\s*(PC|EA|ST|PCE|PCS|SET|U|M|KG)\b", re.I)


def _num(s: str, dec: str = ",") -> float | None:
    s = str(s).strip().replace(" ", "")
    try:
        return float(s.replace(".", "").replace(",", ".")) if dec == "," else float(s.replace(",", ""))
    except ValueError:
        return None


def parse_units(text: str, dec: str = ",") -> list[dict]:
    """Unités de manutention d'une packing list : [{"hu", "net", "gross", "items": [(ref, qty)]}] (poids net = poids de chargement)."""
    m = TABLE_RE.search(text or "")
    if not m:
        return []
    lines = (text or "")[m.start():].split("\n")
    units: list[dict] = []
    cur, pending = None, None
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        h = HU_RE.match(s)
        if h:
            w = HU_INLINE.search(s)
            g = n = None
            if w:
                g, n = _num(w.group(1), dec), _num(w.group(2), dec)
            elif re.search(r"KG\s+KG\s*$", s, re.I) and pending:
                g, n = pending
            cur = {"hu": h.group(1), "net": n, "gross": g, "items": []}
            units.append(cur)
            pending = None
            continue
        w2 = W_LINE.match(s)
        if w2 and not MAT_RE.match(s):
            pending = (_num(w2.group(1), dec), _num(w2.group(2), dec))
            continue
        mm = MAT_RE.match(s)
        if mm and cur is not None:                                  # une ligne d'article avant tout en-tête d'unité n'est pas rattachée
            cur["items"].append((mm.group(1).upper(), float(mm.group(2).replace(",", "."))))
    return [u for u in units if u["items"] or u["net"]]


def hu_map(docs) -> dict:
    out: dict = {}
    for d in docs:
        if getattr(d, "kind", "") != "PACKING":
            continue
        t = d.text or ""
        bl = BL_NO_RE.search(t)
        if not bl:
            continue
        units = parse_units(t, getattr(d, "decimal", ",") or ",")
        if units:
            out.setdefault(bl.group(1), []).extend(units)
    return out


def attach(df: pd.DataFrame, m: dict) -> pd.DataFrame:
    """Ajoute `_hu` / `_hu_net` : pour chaque BL, chaque ligne du tableau est rattachée (dans l'ordre) à l'unité de manutention qui contient
    sa référence ; les références répétées restent des lignes distinctes, jamais fusionnées."""
    import ventilation_export as vx
    out = df.copy()
    out["_hu"] = None
    out["_hu_net"] = float("nan")
    if not m:
        return out
    bl_col = "BL Number (Bon de Livraison)"
    for bl, idx in out.groupby(out[bl_col].astype(str), sort=False).groups.items():
        units = m.get(bl)
        if not units:
            continue
        free = {i: vx.norm_ref(out.at[i, "Reference"]) for i in idx}
        for u in units:
            if u["net"] is None:
                continue
            for ref, _q in u["items"]:
                key = vx.norm_ref(ref)
                hit = next((i for i, r in free.items() if r == key), None)
                if hit is not None:
                    out.at[hit, "_hu"] = u["hu"]
                    out.at[hit, "_hu_net"] = float(u["net"])
                    free.pop(hit)
    return out
