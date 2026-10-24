# -*- coding: utf-8 -*-
"""Contrôles du dossier IMPORT : complétude par client, manifeste déposé seul, tableau d'engagement (valeur par client)."""

from __future__ import annotations

import difflib
import io
import re

import pandas as pd

LEGAL = re.compile(r"\b(sas|sarl|s\.a\.s|s\.a\.r\.l|s\.a|sa|gmbh|ag|ltd|limited|llc|inc|bv|b\.v|nv|spa|s\.p\.a|srl|s\.r\.l|co|kg|oy|ab|as|plc|se)\b\.?", re.I)


def norm_name(s) -> str:
    t = re.sub(r"[^\w\s]", " ", str(s or "").casefold())
    t = LEGAL.sub(" ", t)
    return re.sub(r"\s+", " ", t).strip()


def match_name(name: str, candidates: list[str], cutoff: float = 0.62) -> str | None:
    """Nom le plus proche (insensible à la casse, aux formes juridiques et à l'ordre des mots)."""
    n = norm_name(name)
    if not n:
        return None
    best, score = None, 0.0
    for c in candidates:
        m = norm_name(c)
        if not m:
            continue
        if n == m:
            return c
        s = difflib.SequenceMatcher(None, n, m).ratio()
        if n in m or m in n:
            s = max(s, 0.85)
        ts, tm = set(n.split()), set(m.split())
        if (ts & tm) and min(len(ts), len(tm)) <= 2:
            s = max(s, 0.9 * len(ts & tm) / min(len(ts), len(tm)))
        if s > score:
            best, score = c, s
    return best if score >= cutoff else None


def _num(v):
    import app as core
    if isinstance(v, str):
        return core.parse_number(v, ".")
    return None if v is None or v != v else float(v)


def client_totals(final: pd.DataFrame) -> pd.DataFrame:
    d = final.copy()
    d["Valeur HT"] = pd.to_numeric(d["Valeur HT"], errors="coerce").fillna(0.0)
    d["Client"] = d["Client"].map(lambda v: str(v or "").strip() or "(sans nom)")
    g = d.groupby("Client", sort=False).agg(Lignes=("Valeur HT", "size"), Valeur=("Valeur HT", "sum"))
    if "_devise" in d:
        g["Devise"] = d.groupby("Client", sort=False)["_devise"].agg(lambda s: ", ".join(sorted({str(x) for x in s if str(x) not in ("", "nan", "None")})))
    return g.reset_index()


def dossier_totals(final: pd.DataFrame) -> dict:
    """Totaux du dossier lus dans les factures : valeur, poids brut et colis (comptés une fois par facture)."""
    d = final.copy()
    key = d["Client"].map(lambda v: str(v or "")) + "§" + d["N° facture"].map(lambda v: str(v or ""))
    first = lambda c: (pd.to_numeric(d[c], errors="coerce").groupby(key).agg(lambda s: s.dropna().iloc[0] if s.notna().any() else 0.0).sum()  # noqa: E731
                       if c in d else 0.0)
    return {"valeur": float(pd.to_numeric(d["Valeur HT"], errors="coerce").fillna(0).sum()),
            "brut": float(first("Poids brut facture (kg)")), "colis": float(first("Nb colis / palettes (facture)"))}


def completeness(final: pd.DataFrame) -> pd.DataFrame:
    """Par client : lignes auxquelles il manque NGP, pays d'origine, poids net, quantité ou valeur."""
    import ventilation_import as vi
    d = final.copy()
    n = len(d)
    empty = lambda v: pd.Series([v] * n, index=d.index)  # noqa: E731
    d["Client"] = d["Client"].map(lambda v: str(v or "").strip() or "(sans nom)")
    ngp = d["NGP (REF_PAR_FRS)"].map(vi.ngp_digits) if "NGP (REF_PAR_FRS)" in d else empty("")
    ngp_f = d["NGP/code douanier (facture)"].map(vi.ngp_digits) if "NGP/code douanier (facture)" in d else empty("")
    org = d["Origine"].map(vi.origin_code) if "Origine" in d else empty("")
    num = lambda c: pd.to_numeric(d[c], errors="coerce") if c in d else empty(float("nan"))  # noqa: E731
    tab = pd.DataFrame({"Client": d["Client"], "ngp": ((ngp == "") & (ngp_f == "")), "org": org == "",
                        "net": ~(num("Poids net ligne (kg)") > 0), "brut": ~(num("Poids brut facture (kg)") > 0),
                        "qty": ~(num("Qté") > 0), "val": ~(num("Valeur HT") > 0)})
    g = tab.groupby("Client", sort=False).agg(Lignes=("ngp", "size"), a=("ngp", "sum"), b=("org", "sum"), c=("net", "sum"),
                                              e=("qty", "sum"), f=("val", "sum"), h=("brut", "sum")).reset_index()
    g.columns = ["Client", "Lignes", "Sans NGP", "Sans origine", "Sans poids net", "Sans quantité", "Sans valeur", "Sans poids brut facture"]
    miss = g[["Sans NGP", "Sans origine", "Sans quantité", "Sans valeur"]].sum(axis=1)
    g["Statut"] = ["Complet" if (m == 0 and k == 0) else ("Poids net à estimer" if m == 0 else "À compléter")
                   for m, k in zip(miss, g["Sans poids net"])]
    return g


def read_table(data: bytes, name: str) -> pd.DataFrame:
    """Tableau Excel / CSV -> DataFrame ; entête = 1re ligne contenant au moins 2 textes."""
    if name.lower().endswith((".csv", ".txt")):
        raw = pd.read_csv(io.BytesIO(data), header=None, sep=None, engine="python", dtype=object)
    else:
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)
        raw = max(sheets.values(), key=lambda s: int(s.notna().sum().sum()))
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    hdr = 0
    for i in range(min(len(raw), 15)):
        if sum(isinstance(v, str) and bool(v.strip()) for v in raw.iloc[i]) >= 2:
            hdr = i
            break
    cols = [str(c).strip() if (c == c and str(c).strip()) else f"Colonne {j + 1}" for j, c in enumerate(raw.iloc[hdr])]
    out = raw.iloc[hdr + 1:].reset_index(drop=True)
    out.columns = cols
    return out


def guess_columns(df: pd.DataFrame) -> tuple[str, str]:
    cols = list(df.columns)
    val = next((c for c in cols if re.search(r"valeur|montant|engag|total|amount|value", str(c), re.I)), None)
    cli = next((c for c in cols if re.search(r"client|fournisseur|exp[ée]diteur|supplier|shipper|nom|vendor", str(c), re.I)), None)
    if val is None:
        val = max(cols, key=lambda c: sum(_num(v) is not None for v in df[c]))
    if cli is None:
        cli = next((c for c in cols if c != val), cols[0])
    return cli, val


def compare_engagement(final: pd.DataFrame, eng: pd.DataFrame, client_col: str, value_col: str, tol: float = 0.005) -> pd.DataFrame:
    """Σ Valeur HT lue dans les factures de chaque client  vs  valeur du tableau d'engagement."""
    docs = client_totals(final)
    eng = eng[[client_col, value_col]].copy()
    eng["_v"] = eng[value_col].map(_num)
    eng = eng[eng[client_col].map(lambda v: str(v or "").strip() != "") & eng["_v"].notna()]
    eng_names = eng[client_col].astype(str).tolist()
    rows, used = [], set()
    for _, r in docs.iterrows():
        m = match_name(r["Client"], eng_names)
        used.add(m)
        ev = float(eng.loc[eng[client_col].astype(str) == m, "_v"].sum()) if m else None
        diff = (r["Valeur"] - ev) if ev is not None else None
        pct = (diff / ev * 100) if ev else None
        status = "Absent de l'engagement" if m is None else ("OK" if abs(diff) <= max(1.0, tol * abs(ev)) else "ÉCART")
        rows.append({"Client (factures)": r["Client"], "Client (engagement)": m or "—", "Devise": r.get("Devise", ""),
                     "Valeur factures": round(r["Valeur"], 2), "Valeur engagement": None if ev is None else round(ev, 2),
                     "Écart": None if diff is None else round(diff, 2), "Écart %": None if pct is None else round(pct, 2), "Statut": status})
    for _, e in eng.iterrows():
        if str(e[client_col]) not in used:
            rows.append({"Client (factures)": "—", "Client (engagement)": str(e[client_col]), "Devise": "", "Valeur factures": 0.0,
                         "Valeur engagement": round(float(e["_v"]), 2), "Écart": round(-float(e["_v"]), 2), "Écart %": -100.0,
                         "Statut": "Pas de facture lue"})
    return pd.DataFrame(rows)


def compare_manifest(totals: dict, man: pd.DataFrame, tol: float = 0.005) -> pd.DataFrame:
    """Totaux du dossier (factures) vs totaux du manifeste (valeur, poids brut, colis)."""
    def total(col):
        if col not in man.columns:
            return None
        s = pd.to_numeric(man[col], errors="coerce")
        return float(s.sum()) if s.notna().any() else None
    rows = []
    for label, mine, theirs, unit in (("Valeur (Goods Value)", totals["valeur"], total("Goods Value"), ""),
                                      ("Poids brut", totals["brut"], total("Poids (kg)"), "kg"),
                                      ("Colis", totals["colis"], total("Colis"), "")):
        lab = label + (f" ({unit})" if unit else "")
        if theirs is None:
            rows.append({"Contrôle": lab, "Dossier (factures)": round(mine, 3), "Manifeste": None, "Écart": None, "Statut": "Non trouvé dans le manifeste"})
            continue
        diff = mine - theirs
        ok = abs(diff) <= max(0.05 if unit else 1.0, tol * abs(theirs))
        rows.append({"Contrôle": lab, "Dossier (factures)": round(mine, 3), "Manifeste": round(theirs, 3),
                     "Écart": round(diff, 3), "Statut": "OK" if ok else "ÉCART"})
    return pd.DataFrame(rows)


def factor_guess(ratio: float, tol: float = 0.01) -> float:
    """Facteur 10/100/1000 (ou son inverse) qui explique l'écart de valeur ; 1.0 si aucun."""
    if not ratio or ratio <= 0:
        return 1.0
    for k in (1000.0, 100.0, 10.0):
        if abs(ratio / k - 1) <= tol:
            return k
        if abs(ratio * k - 1) <= tol:
            return 1.0 / k
    return 1.0


def compare_manifest_clients(final: pd.DataFrame, man: pd.DataFrame, tol: float = 0.005) -> pd.DataFrame:
    docs = client_totals(final)
    names = [str(x) for x in man["Expéditeur"]] if "Expéditeur" in man.columns else []
    rows = []
    for _, r in docs.iterrows():
        m = match_name(r["Client"], names)
        sub = man[man["Expéditeur"].astype(str) == m] if m else man.iloc[0:0]
        mv = float(pd.to_numeric(sub["Goods Value"], errors="coerce").sum()) if (len(sub) and "Goods Value" in sub) else None
        diff = None if mv is None else r["Valeur"] - mv
        ratio = (r["Valeur"] / mv) if mv else None
        f = factor_guess(ratio) if ratio else 1.0
        ok = m is not None and mv is not None and abs(diff) <= max(1.0, tol * abs(mv or 0))
        rows.append({"Client (factures)": r["Client"], "Expéditeur (manifeste)": m or "—", "Valeur factures": round(r["Valeur"], 2),
                     "Goods Value manifeste": None if mv is None else round(mv, 2), "Écart": None if diff is None else round(diff, 2),
                     "Rapport": None if ratio is None else round(ratio, 3),
                     "Erreur probable": ("" if f == 1.0 else (f"valeur lue ×{f:g} trop grande" if f > 1 else f"valeur lue ÷{1 / f:g} trop petite")),
                     "Statut": "Absent du manifeste" if m is None else ("OK" if ok else "ÉCART")})
    return pd.DataFrame(rows)


def value_fixes(final: pd.DataFrame, man: pd.DataFrame) -> dict[str, float]:
    """{client: facteur de division} pour les clients dont l'écart avec le manifeste est exactement un facteur 10/100/1000."""
    out = {}
    for _, r in compare_manifest_clients(final, man).iterrows():
        rt = r["Rapport"]
        if rt and r["Erreur probable"]:
            out[r["Client (factures)"]] = factor_guess(rt)
    return out


def build_checks(final: pd.DataFrame, man: pd.DataFrame | None, eng: pd.DataFrame | None = None,
                 eng_cols: tuple[str, str] | None = None) -> dict[str, pd.DataFrame]:
    """Tous les contrôles du dossier IMPORT (feuilles prêtes à afficher / à ajouter au classeur Excel)."""
    sheets: dict[str, pd.DataFrame] = {"Complétude clients": completeness(final)}
    if man is not None and len(man) and pd.to_numeric(man.get("Goods Value"), errors="coerce").notna().any():
        sheets["Manifeste vs dossier"] = compare_manifest(dossier_totals(final), man)
        if not man["Expéditeur"].astype(str).str.upper().str.startswith("TOTAL").all():       # manifeste sans détail par expéditeur : totaux seulement
            sheets["Manifeste par client"] = compare_manifest_clients(final, man)
    if eng is not None and eng_cols:
        sheets["Engagement vs factures"] = compare_engagement(final, eng, *eng_cols)
    return sheets


def _fill_sheets(wb, sheets: dict[str, pd.DataFrame], prefix: str = ""):
    from openpyxl.styles import Font, PatternFill
    ok, bad, warn = (PatternFill("solid", start_color=c, end_color=c) for c in ("C6EFCE", "FFC7CE", "FFEB9C"))
    for name, df in sheets.items():
        if df is None or df.empty:
            continue
        title = (prefix + name)[:31]
        if title in wb.sheetnames:
            del wb[title]
        ws = wb.create_sheet(title)
        ws.append(list(df.columns))
        for c in ws[1]:
            c.font = Font(bold=True)
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and v != v) else v for v in row])
        if "Statut" in df.columns:
            j = list(df.columns).index("Statut") + 1
            for r in range(2, ws.max_row + 1):
                v = str(ws.cell(r, j).value)
                ws.cell(r, j).fill = ok if v in ("OK", "Complet") else (bad if ("CART" in v or "Absent" in v or "Pas de" in v or "compléter" in v) else warn)
        for j, col in enumerate(df.columns, start=1):
            ws.column_dimensions[ws.cell(1, j).column_letter].width = max(14, min(42, len(str(col)) + 6))
        ws.freeze_panes = "A2"


def checks_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    _fill_sheets(wb, sheets)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def add_sheets_to_xlsx(xlsx: bytes, sheets: dict[str, pd.DataFrame], prefix: str = "VERIF ") -> bytes:
    """Ajoute les feuilles de contrôle (complétude, manifeste, engagement) au classeur TOUS_CLIENTS."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx))
    _fill_sheets(wb, sheets, prefix)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
