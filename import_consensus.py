# -*- coding: utf-8 -*-
"""IMPORT — calcul multi-méthodes (« vote ») des chiffres clés de chaque facture et de chaque client.

  Montant d'une facture  : somme des lignes · total imprimé sur la facture · somme des quantité × prix unitaire
  Valeur d'un client     : somme des lignes · somme des totaux imprimés · GOODS VALUE de l'avis d'arrivée / manifeste
  Poids brut d'un client : poids brut des factures / packing lists · poids de l'avis d'arrivée / manifeste
  Colis d'un client      : colis des factures / packing lists · colis de l'avis d'arrivée / manifeste

La valeur retenue est celle sur laquelle au moins 2 méthodes sont d'accord. Montant d'une facture : si « quantité × prix » et le
total imprimé sont d'accord contre la somme des lignes, les montants de ligne faux sont corrigés (quantité × prix). Sinon ⚠.
"""

from __future__ import annotations

import pandas as pd

import app as core

C_GROSS, C_COLIS = "Poids brut facture (kg)", "Nb colis / palettes (facture)"


def _same(a, b, rel=0.005, abs_=0.05):
    return a is not None and b is not None and abs(a - b) <= max(abs_, rel * max(abs(a), abs(b)))


TOL = {"money": (0.0001, 0.05), "weight": (0.005, 0.05), "count": (0.0, 0.0)}


def vote(cands: dict, tol=TOL["money"]):
    vals = [(k, v) for k, v in cands.items() if v is not None]
    best, best_n = None, 0
    for _, v in vals:
        n = sum(1 for _, w in vals if _same(v, w, *tol))
        if n > best_n:
            best, best_n = v, n
    return (best if best_n >= 2 else None), best_n, len(vals)


def _num(s):
    return pd.to_numeric(s, errors="coerce")


def _first_per_invoice(df: pd.DataFrame, col: str) -> float | None:
    if col not in df:
        return None
    v = df.assign(_v=_num(df[col])).groupby("N° facture")["_v"].first()
    v = v.dropna()
    return float(v.sum()) if len(v) else None


def analyse(lines: pd.DataFrame, manifest: pd.DataFrame | None, inv_totals: dict | None):
    """Renvoie (lignes corrigées, notes, tableau de contrôle)."""
    import import_checks as ic
    df = lines.copy()
    inv_totals = inv_totals or {}
    rows, notes = [], []

    def record(objet, chiffre, cands, current, verdict_fix=None, kind="money"):
        tol = TOL[kind]
        best, votes, nmeth = vote(cands, tol)
        if nmeth < 2:
            verdict = "1 seule méthode" if nmeth else "aucune donnée"
        elif best is None:
            verdict = "⚠ désaccord"
            notes.append(f"⚠ {objet} — {chiffre} : les méthodes ne s'accordent pas ("
                         + ", ".join(f"{k} {v:,.2f}" for k, v in cands.items() if v is not None) + ") — contrôle manuel.")
        elif current is None or not _same(current, best, *tol):
            if verdict_fix:
                verdict = f"✔ corrigé ({votes}/{nmeth})"
                notes.append(f"✔ {objet} — {chiffre} corrigé par vote : {best:,.2f} (au lieu de "
                             f"{'vide' if current is None else f'{current:,.2f}'}) — {verdict_fix}")
            else:
                verdict = f"⚠ Excel ≠ majorité ({votes}/{nmeth})"
                notes.append(f"⚠ {objet} — {chiffre} : tableau {'vide' if current is None else f'{current:,.2f}'} ≠ {best:,.2f} "
                             f"trouvé par {votes} méthodes sur {nmeth} — contrôle manuel.")
        else:
            verdict = f"✔ OK ({votes}/{nmeth})"
        rows.append({"Objet": objet, "Chiffre": chiffre, **cands, "Tableau": current,
                     "Retenu": best if best is not None else current, "Verdict": verdict})
        return best

    if df.empty or "N° facture" not in df:
        return df, notes, pd.DataFrame()
    qty, price, val = _num(df["Qté"]), _num(df["Prix unitaire"]), _num(df["Valeur HT"])
    src = df["_source"].astype(str) if "_source" in df else pd.Series([""] * len(df), index=df.index)
    key = df["N° facture"].map(lambda v: core.compact_ref(v)) + "§" + src
    # ------------------------------------------------------------ factures
    printed_by_client: dict[str, list] = {}
    for k, idx in key.groupby(key).groups.items():
        idx = list(idx)
        num = str(df.at[idx[0], "N° facture"])
        client = str(df.at[idx[0], "Client"])
        if not num.strip():
            continue
        s_lines = float(val.loc[idx].fillna(0).sum())
        has_qp = {i: pd.notna(qty.loc[i]) and pd.notna(price.loc[i]) for i in idx}
        qp = {i: round(float(qty.loc[i] * price.loc[i]), 2) for i in idx if has_qp[i]}
        s_qp = round(sum(qp.values()), 2) if len(qp) == len(idx) else None
        suspects = [i for i in idx if i in qp and (pd.isna(val.loc[i]) or not _same(float(val.loc[i]), qp[i], 0.001, 0.05))]
        sus_txt = [f"{df.at[i, 'Réf article'] or df.at[i, 'Item']} (qté × prix {qp[i]:,.2f} / montant "
                   f"{'vide' if pd.isna(val.loc[i]) else f'{float(val.loc[i]):,.2f}'})" for i in suspects]
        printed = inv_totals.get(k)
        printed_by_client.setdefault(client, []).append(printed)
        objet = f"Facture {num} ({client})"
        cands = {"Σ lignes": s_lines, "Total imprimé": printed, "Σ qté × prix": s_qp}
        # correction sûre : remplacer par quantité × prix la (les) ligne(s) dont la correction redonne exactement le total imprimé
        fixed = None
        if printed is not None and not _same(s_lines, printed, *TOL["money"]) and suspects:
            amt = lambda i: 0.0 if pd.isna(val.loc[i]) else float(val.loc[i])  # noqa: E731
            pow10 = lambda i: amt(i) and any(_same(qp[i] / amt(i), f, 0.01, 0) for f in (10, 100, 1000, 0.1, 0.01, 0.001))  # noqa: E731
            tries = [[i] for i in suspects] + [[i for i in suspects if not pow10(i)], suspects]
            for grp in tries:
                if grp and _same(s_lines - sum(amt(i) for i in grp) + sum(qp[i] for i in grp), printed, *TOL["money"]):
                    fixed = grp
                    break
        if fixed:
            for i in fixed:
                df.at[i, "Valeur HT"] = qp[i]
            refs = ", ".join(str(df.at[i, "Réf article"] or df.at[i, "Item"]) for i in fixed)
            notes.append(f"✔ {objet} — Montant corrigé : {len(fixed)} ligne(s) recalculée(s) en quantité × prix ({refs}) ; "
                         f"la somme des lignes redonne le total imprimé {printed:,.2f} (au lieu de {s_lines:,.2f}).")
            rows.append({"Objet": objet, "Chiffre": "Montant", **cands, "Tableau": s_lines, "Retenu": printed,
                         "Verdict": "✔ corrigé (ligne × prix = total imprimé)"})
            continue
        n0 = len(notes)
        record(objet, "Montant", cands, s_lines)
        if len(notes) > n0 and notes[-1].startswith("⚠") and sus_txt:
            notes[-1] += " Lignes suspectes : " + " ; ".join(sus_txt[:5])
    # ------------------------------------------------------------ clients (avis d'arrivée / manifeste)
    val2 = _num(df["Valeur HT"])
    man = manifest if manifest is not None and len(manifest) else None
    names = [str(x) for x in man["Expéditeur"]] if man is not None and "Expéditeur" in man else []
    for client, sub in df.groupby("Client"):
        mrow = None
        if names:
            m = ic.match_name(str(client), names)
            if m and not m.upper().startswith("TOTAL"):
                mrow = man[man["Expéditeur"].astype(str) == m]
        if mrow is None and man is not None and df["Client"].nunique() == 1 and len(man) == 1:
            mrow = man                                            # un seul client, un seul avis : c'est le sien
        g = lambda c: float(_num(mrow[c]).sum()) if mrow is not None and c in mrow and _num(mrow[c]).notna().any() else None  # noqa: E731
        pr = printed_by_client.get(str(client), [])
        s_pr = round(sum(pr), 2) if pr and all(p is not None for p in pr) else None
        s_val = float(val2.loc[sub.index].fillna(0).sum())
        record(f"Client {client}", "Valeur", {"Σ lignes": s_val, "Σ totaux imprimés": s_pr, "Avis / manifeste": g("Goods Value")}, s_val)
        gross = _first_per_invoice(sub, C_GROSS)
        record(f"Client {client}", "Poids brut (kg)", {"Factures / PL": gross, "Avis / manifeste": g("Poids (kg)")}, gross, kind="weight")
        col = _first_per_invoice(sub, C_COLIS)
        record(f"Client {client}", "Colis", {"Factures / PL": col, "Avis / manifeste": g("Colis")}, col, kind="count")
    return df, notes, pd.DataFrame(rows)
