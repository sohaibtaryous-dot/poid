# -*- coding: utf-8 -*-
"""EXPORT — calcul multi-méthodes (« vote ») des chiffres clés.

Chaque chiffre est recalculé par plusieurs méthodes indépendantes, directement sur le texte des documents :

  Colis d'un BL       : « No. of packages / Nombre de colis » du bon de livraison · unités de manutention de la packing list
                        · valeur retenue par la lecture principale
  Poids brut / net    : « Total Weight / Poids brut total » · somme des poids des unités de manutention · lecture principale
  Quantité d'un BL    : lignes de la packing list · lignes du bon de livraison · lignes Excel
  Montant d'une facture : total imprimé · somme des lignes · somme des quantité × prix

Règle : la valeur retenue est celle sur laquelle au moins 2 méthodes sont d'accord. Si la valeur du fichier Excel n'est pas
celle de la majorité, elle est corrigée (et signalée). Si aucune majorité ne se dégage, ⚠ : contrôle manuel.
"""

from __future__ import annotations

import re
from collections import defaultdict

import pandas as pd

import core_export as core

HEAD_RE = re.compile(r"^[ \t]*(delivery\s+note|bon\s+de\s+livraison|packing\s+list)[ \t]*(?:n[°o]\.?)?[ \t]*[:#]?\s*(\d{6,12})\b", re.I | re.M)   # en début de ligne : « N° bon de livraison : … » d'une packing list n'est pas un en-tête
PKG_RE = re.compile(r"(?:no\.?\s*of\s*packages|nombre\s+de\s+colis)\s*:?\s*([^\n]*)", re.I)
GROSS_RE = re.compile(r"(?:total\s+weight|poids\s+brut\s+total)\s*:?\s*(\d[\d\s.,]*\d|\d)\s*kg", re.I)
NET_RE = re.compile(r"(?:net\s+weight|poids\s+net\s+total)\s*:?\s*(\d[\d\s.,]*\d|\d)\s*kg", re.I)
HU_LINE_RE = re.compile(r"^(\d{6,8})\s+[A-Za-zÀ-ÿ].*?$", re.M)
HU_W_INLINE = re.compile(r"(\d[\d.,]*)\s*KG\s+(\d[\d.,]*)\s*KG\s*$", re.I)
HU_W_SPLIT = re.compile(r"^\s*(\d[\d.,]*)\s+(\d[\d.,]*)\s*$")
MAT_QTY_RE = re.compile(r"^\s*([A-Z]{1,3}\d{6,})\s+.*?\s(\d+(?:[.,]\d+)?)\s*/\s*(PC|EA|ST|PCE|PCS|M|KG|SET|U)\b", re.I | re.M)
DN_QTY_RE = re.compile(r"^\s*\d{1,5}\s+\d{1,5}\s+.*?\s(\d+(?:[.,]\d+)?)\s*/?\s*(PC|EA|ST|PCE|PCS|SET|U)\s*$", re.I | re.M)


def _num(v, dec=","):
    return core.parse_number(str(v).replace(" ", ""), dec) if v not in (None, "") else None


def _same(a, b, tol_rel=0.005, tol_abs=0.01):
    return a is not None and b is not None and abs(a - b) <= max(tol_abs, tol_rel * max(abs(a), abs(b)))


def vote(cands: dict):
    """{méthode: valeur} -> (valeur majoritaire | None, nb de voix, nb de méthodes)."""
    vals = [(k, v) for k, v in cands.items() if v is not None]
    best, best_n = None, 0
    for _, v in vals:
        n = sum(1 for _, w in vals if _same(v, w))
        if n > best_n:
            best, best_n = v, n
    return (best if best_n >= 2 else None), best_n, len(vals)


def segments(docs) -> dict:
    """Texte des documents découpé par BL : {n° BL: {"dn": texte bon de livraison, "pl": texte packing list}}."""
    out: dict = defaultdict(lambda: {"dn": "", "pl": ""})
    for d in docs:
        if getattr(d, "kind", "") != "PACKING":
            continue
        t = d.text or ""
        heads = list(HEAD_RE.finditer(t))
        for i, m in enumerate(heads):
            seg = t[m.end(): heads[i + 1].start() if i + 1 < len(heads) else len(t)]
            kind = "pl" if m.group(1).lower().startswith("packing") else "dn"
            out[m.group(2)][kind] += "\n" + seg
    return out


def hu_weights(pl_text: str, dec: str = ","):
    """Unités de manutention d'une packing list : {id: (poids total, poids de chargement)} (« 1528939 … 171,000 KG 154,000 KG »
    ou poids sur la ligne précédente « 38,000 28,000 » puis « 1528940 … KG KG »)."""
    lines = (pl_text or "").split("\n")
    hus: dict = {}
    for i, ln in enumerate(lines):
        m = HU_LINE_RE.match(ln.strip())
        if not m:
            continue
        hu = m.group(1)
        w = HU_W_INLINE.search(ln)
        g = n = None
        if w:
            g, n = _num(w.group(1), dec), _num(w.group(2), dec)
        elif re.search(r"KG\s+KG\s*$", ln, re.I) and i > 0 and (s := HU_W_SPLIT.match(lines[i - 1])):
            g, n = _num(s.group(1), dec), _num(s.group(2), dec)
        if hu not in hus or (hus[hu][0] is None and g is not None):
            hus[hu] = (g, n)
    return hus


def analyse(invoices, bls, bl_df: pd.DataFrame, fact_df: pd.DataFrame, docs) -> tuple[pd.DataFrame, pd.DataFrame, list[str], pd.DataFrame]:
    """Renvoie (bl_df corrigé, fact_df corrigé, notes, tableau de contrôle multi-méthodes)."""
    bl_df, fact_df = bl_df.copy(), fact_df.copy()
    segs = segments(docs)
    dec = next((d.decimal for d in docs if getattr(d, "kind", "") == "PACKING"), ",") or ","
    rows, notes = [], []
    col_bl = "BL Number (Bon de Livraison)"
    n = lambda s: pd.to_numeric(s, errors="coerce")  # noqa: E731

    def record(objet, chiffre, cands, current, apply=None, fmt="{:g}"):
        best, votes, nmeth = vote(cands)
        verdict = "—"
        if nmeth < 2:
            verdict = "1 seule méthode" if nmeth else "aucune donnée"
        elif best is None:
            verdict = "⚠ désaccord"
            notes.append(f"⚠ {objet} — {chiffre} : les méthodes ne s'accordent pas ("
                         + ", ".join(f"{k} {fmt.format(v)}" for k, v in cands.items() if v is not None) + ") — contrôle manuel.")
        elif current is None or not _same(current, best):
            if apply is not None:
                apply(best)
                verdict = f"✔ corrigé ({votes}/{nmeth})"
                notes.append(f"✔ {objet} — {chiffre} corrigé par vote : {fmt.format(best)} (au lieu de "
                             f"{'vide' if current is None else fmt.format(current)}), {votes} méthodes sur {nmeth} d'accord.")
            else:
                verdict = f"⚠ Excel ≠ majorité ({votes}/{nmeth})"
                notes.append(f"⚠ {objet} — {chiffre} : Excel {'vide' if current is None else fmt.format(current)} ≠ "
                             f"{fmt.format(best)} trouvé par {votes} méthodes sur {nmeth} — contrôle manuel.")
        else:
            verdict = f"✔ OK ({votes}/{nmeth})"
        rows.append({"Objet": objet, "Chiffre": chiffre, **{k: v for k, v in cands.items()},
                     "Excel": current, "Retenu": best if best is not None else current, "Verdict": verdict})

    # ------------------------------------------------------------- BL : colis, poids, quantité
    first_idx = {}
    for i, b in bl_df[col_bl].items():
        first_idx.setdefault(str(b), i)
    for bl in bls:
        if not bl.number or bl.number not in first_idx:
            continue
        s = segs.get(bl.number, {"dn": "", "pl": ""})
        i0 = first_idx[bl.number]
        # facture rattachée à ce seul BL (et ce BL à cette seule facture) : ses totaux sont une méthode indépendante
        solo = bl.invoices[0] if len(bl.invoices) == 1 and sum(1 for b in bls if bl.invoices[0] in b.invoices) == 1 else None
        sel = bl_df[col_bl].astype(str) == bl.number

        def setter(col, i=i0):
            return lambda v: bl_df.__setitem__(col, bl_df[col].where(bl_df.index != i, v))
        # colis
        pk_txt = (PKG_RE.search(s["dn"]) or PKG_RE.search(s["pl"]))
        part_of = bool(pk_txt and re.match(r"(?i)\s*partie\s+du\s+colis", pk_txt.group(1)))
        hus = hu_weights(s["pl"], dec)
        cur_pk = n(pd.Series([bl_df.at[i0, "Number of Packages (Nombre de Colis)"]])).iloc[0]
        cur_pk = None if pd.isna(cur_pk) else float(cur_pk)
        if not part_of:
            dn_pk = None
            if pk_txt and (m := re.match(r"\s*(\d{1,6})\b", pk_txt.group(1))):
                dn_pk = float(int(m.group(1)))
            record(f"BL {bl.number}", "Colis", {"Bon de livraison": dn_pk, "Unités de manutention": float(len(hus)) if hus else None,
                                               "Facture": None if not (solo and solo.packages) else float(solo.packages)},
                   cur_pk, setter("Number of Packages (Nombre de Colis)"))
        # poids
        for chiffre, rx, k, col, attr in (("Poids brut (kg)", GROSS_RE, 0, "Total Gross Weight of the BL", "gross"),
                                          ("Poids net (kg)", NET_RE, 1, "Total Net Weight of the BL", "net")):
            m = rx.search(s["pl"]) or rx.search(s["dn"])
            hw = [w[k] for w in hus.values()]
            cur = n(pd.Series([bl_df.at[i0, col]])).iloc[0]
            record(f"BL {bl.number}", chiffre, {"Total imprimé": _num(m.group(1), dec) if m else None,
                                               "Σ unités de manutention": round(sum(hw), 3) if hw and all(x is not None for x in hw) else None,
                                               "Facture": getattr(solo, attr, None) if solo else None},
                   None if pd.isna(cur) else float(cur), setter(col), "{:,.3f}")
        # quantité (pas de correction automatique : ce sont des lignes)
        pl_q = [float(x[1].replace(",", ".")) for x in MAT_QTY_RE.findall(s["pl"])]
        dn_q = [float(x[0].replace(",", ".")) for x in DN_QTY_RE.findall(s["dn"])]
        if bl.orphan:
            continue                                   # BL sans facture : la quantité vient des factures (non comptée)
        inv_q = sum(float(l["qty"] or 0) for l in solo.lines) if solo and solo.lines else None
        record(f"BL {bl.number}", "Quantité", {"Lignes packing list": sum(pl_q) if pl_q else None,
                                              "Lignes bon de livraison": sum(dn_q) if dn_q else None,
                                              "Facture": inv_q},
               float(n(bl_df.loc[sel, "Quantity"]).sum()))

    # ------------------------------------------------------------- factures : montant
    for inv in invoices:
        if not inv.number:
            continue
        idx = fact_df.index[fact_df["Numéro de facture"].astype(str).str.strip() == str(inv.number).strip()]
        if not len(idx):
            continue
        lines = inv.lines + inv.fees
        s_lines = round(sum(l["amount"] for l in lines if l.get("amount") is not None), 2) if any(l.get("amount") is not None for l in lines) else None
        qp = [float(l["qty"]) * float(l["price"]) / float(l.get("price_per") or 1) for l in lines
              if l.get("qty") is not None and l.get("price") is not None]
        s_qp = round(sum(qp), 2) if qp and len(qp) == len(lines) else None
        printed = inv.total_ht if inv.total_src not in ("", "Somme des lignes") else None
        cur = n(pd.Series([fact_df.at[idx[0], "Montant en devise"]])).iloc[0]

        def set_amt(v, i=idx[0]):
            fact_df.at[i, "Montant en devise"] = round(v, 2)
        record(f"Facture {inv.number}", "Montant", {"Total imprimé": printed, "Σ lignes": s_lines, "Σ qté × prix": s_qp},
               None if pd.isna(cur) else float(cur), set_amt, "{:,.2f}")

    table = pd.DataFrame(rows)
    return bl_df, fact_df, notes, table
