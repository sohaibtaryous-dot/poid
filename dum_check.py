# -*- coding: utf-8 -*-
"""EXPORT — vérification d'une déclaration (DUM de BADR, modèle « D.U.M 2014 ») avec les données du dossier.

  parse_dum(texte)  -> {numéro, date, destinataire, devise, montant, taux, fret, valeur_declaree, brut, net, articles[…], pied{…}}
  self_checks(dum)  -> contrôles internes de la déclaration (sommes des articles, montant × taux, pied de page…)
  compare(dum, vent, rate, fret, fact_df, bl_df) -> lignes de contrôle (déclaration ↔ ventilation / factures / packing lists)
La lecture se fait sur le texte du PDF (pas d'OCR) : une DUM scannée n'est pas lue."""

from __future__ import annotations

import re

NUM = r"\d[\d ]*(?:\.\d+)?"


def _f(s) -> float | None:
    try:
        return float(str(s).replace(" ", "").replace(" ", ""))
    except ValueError:
        return None


def pdf_text(data: bytes) -> str:
    import io
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def _split_declarations(text: str) -> list[str]:
    """Un PDF peut contenir plusieurs DUM (une par page) : on découpe sur « Mod. D.U.M »."""
    parts = re.split(r"(?=MINISTERE DE L'ECONOMIE ET DES FINANCES)", text)
    return [p for p in parts if "Mod. D.U.M" in p or "Code marchandises" in p]


def parse_dum(text: str) -> dict:
    t = text or ""
    d: dict = {"articles": [], "pied": {}}
    m = re.search(r"(\d{7})\s+([A-Z])\s+(\d{2}/\d{2}/\d{4})", t)
    if m:
        d["numero"], d["date"] = f"{m.group(1)} {m.group(2)}", m.group(3)
    # destinataire : lignes entre « 8 Importateur / Destinataire » et « 10 Déclarant », hors cases d'identification
    m = re.search(r"Importateur\s*/\s*Destinataire(.*?)(?:10\s+D[ée]clarant)", t, re.S)
    if m:
        lines = [ln.strip() for ln in m.group(1).split("\n") if ln.strip()]
        keep = [ln for ln in lines if not re.match(r"^(N°|Centre|ICE|,|9\s+Autre)", ln) and "Autre personne" not in ln]
        d["destinataire"] = keep[0] if keep else ""
        d["pays_dest"] = keep[1] if len(keep) > 1 else ""
    m = re.search(r"\b([A-Z]{3})\s+(" + NUM + r"\.\d{3})\s+(\d+\.\d{6})\s+(" + NUM + r"\.\d{3})", t)
    if m:
        d["devise"], d["montant"], d["taux"], d["fret"] = m.group(1), _f(m.group(2)), _f(m.group(3)), _f(m.group(4))
    m = re.search(r"Valeur totale d[ée]clar[ée]e(.*?)(?:\n\s*24\s+Date)", t, re.S)
    if m:
        v = re.findall(r"(\d[\d ]*\.\d{3})", m.group(1))
        if v:
            d["valeur_declaree"] = _f(v[-1])
    m = re.search(r"Poids net total \(kg\)(.*?)(?:8\s+Importateur)", t, re.S)
    if m:
        blk = re.sub(r"\b\d+\s*/\s*\d+\b", " ", m.group(1))                # « formules 1 / 1 » ne fait pas partie des poids
        w = [_f(x) for x in re.findall(r"(?<!\d)\d{1,3}(?: \d{3})*\.\d{2}(?!\d)", blk)]
        w = [x for x in w if x is not None]
        if w:
            d["brut"], d["net"] = max(w), min(w)                      # les deux valeurs sont imprimées l'une sous l'autre : brut ≥ net
    m = re.search(r"Incoterm|Conditions de livraison\s*\n?\s*([A-Z]{3})\b", t)
    if m and m.groups():
        d["incoterm"] = m.group(1)
    # articles
    pos = [x for x in re.finditer(r"^\s*(\d{1,2})\s+(\d{8,10})\s+(" + NUM + r"\.\d{3})\s*$", t, re.M)]
    for k, mm in enumerate(pos):
        blk = t[mm.start(): pos[k + 1].start() if k + 1 < len(pos) else len(t)]
        before = t[max(0, mm.start() - 260): mm.start()]
        des = re.search(r"de l'art\.\s+(.+?)\s+(\d+(?:\.\d+)?)(?:\s+(NB|U|KG|M|PC))?\s*$", before.split("\n")[-2] if "\n" in before else before, re.M) \
            or re.search(r"de l'art\.\s+(.+?)\s+(\d+(?:\.\d+)?)(?:\s+(NB|U|KG|M|PC))?\s*$", before, re.M)
        art = {"n": int(mm.group(1)), "ngp": mm.group(2), "valeur": _f(mm.group(3)), "designation": "", "qte": None, "unite": ""}
        if des:
            art["designation"], art["qte"], art["unite"] = des.group(1).strip(), _f(des.group(2)), des.group(3) or ""
        if not art["unite"]:
            u = re.match(r"\s*(NB|U|KG|M|PC)\b", blk[mm.end() - mm.start():])
            if u:
                art["unite"] = u.group(1)
        w = re.search(r"(" + NUM + r"\.\d)\s+(KG|U|NB|M|PC)\s+(" + NUM + r"\.\d{2})\s+(AP|SP)", blk)
        if w:
            art["unites_compl"], art["unite_compl"], art["net"] = _f(w.group(1)), w.group(2), _f(w.group(3))
        o = re.search(r"(?:^|\s)([A-ZÉ' ]{3,})\s+([A-Z]{2})\s*$", blk, re.M)
        po = re.search(r"Pays d'origine[^\n]*\n(?:[^\n]*?)\b([A-ZÉ]{3,})\s+([A-Z]{2})\b", blk)
        if po:
            art["origine"] = po.group(2)
        a = re.search(r"Accord\s*:\s*([A-Z]+)", blk)
        art["accord"] = a.group(1) if a else ""
        d["articles"].append(art)
    # pied de page : QTE / VAL / factures
    m = re.search(r"QTE\s*\.\s*(" + NUM + r"\.\d{3})\s*KG", t)
    if m:
        d["pied"]["qte_net"] = _f(m.group(1))
    m = re.search(r"VAL\s*\.\s*(" + NUM + r"\.\d{3})\s*([A-Z]{3})", t)
    if m:
        d["pied"]["val"], d["pied"]["devise"] = _f(m.group(1)), m.group(2)
    m = re.search(r"Nombre de factures\s*\((\d+)\)\s*N°\s*:\s*([^\n]*)", t)
    if m:
        d["pied"]["n_factures"] = int(m.group(1))
        d["pied"]["truncated"] = "..." in m.group(2) or "…" in m.group(2)
        d["pied"]["factures"] = re.findall(r"\d{6,}", m.group(2))
    return d


def parse_all(data: bytes) -> list[dict]:
    txt = pdf_text(data)
    return [parse_dum(p) for p in (_split_declarations(txt) or [txt])]


def _row(rows, niveau, controle, decl, site, tol=0.0, fmt="{:,.2f}", ok_txt="✔", note=""):
    if decl is None and site is None:
        return
    if decl is None or site is None:
        rows.append({"Niveau": niveau, "Contrôle": controle, "Déclaration": "—" if decl is None else fmt.format(decl).replace(",", " "),
                     "Dossier (site)": "—" if site is None else fmt.format(site).replace(",", " "), "Écart": "", "Résultat": "ℹ non comparable"})
        return
    ec = decl - site
    ok = abs(ec) <= tol
    rows.append({"Niveau": niveau, "Contrôle": controle, "Déclaration": fmt.format(decl).replace(",", " "),
                 "Dossier (site)": fmt.format(site).replace(",", " "), "Écart": "" if ok else fmt.format(ec).replace(",", " "),
                 "Résultat": ok_txt if ok else ("⚠ écart" + (f" — {note}" if note else ""))})


def self_checks(d: dict) -> list[dict]:
    rows: list[dict] = []
    arts = d.get("articles", [])
    if arts and d.get("net") is not None:
        _row(rows, "Déclaration", "Σ poids net des articles = poids net total", sum(a.get("net") or 0 for a in arts), d["net"], 0.05)
    if arts and d.get("valeur_declaree") is not None:
        _row(rows, "Déclaration", "Σ valeur déclarée des articles = valeur totale déclarée", sum(a.get("valeur") or 0 for a in arts), d["valeur_declaree"], 0.5, "{:,.3f}")
    if d.get("montant") and d.get("taux") and d.get("valeur_declaree") is not None:
        _row(rows, "Déclaration", "montant facturé × taux ≈ valeur totale déclarée", d["montant"] * d["taux"], d["valeur_declaree"], max(2.0, len(arts) * 1.0), "{:,.2f}")
    p = d.get("pied", {})
    if p.get("qte_net") is not None and d.get("net") is not None:
        _row(rows, "Déclaration", "pied de page : QTE net = poids net total", p["qte_net"], d["net"], 0.05)
    if p.get("val") is not None and d.get("montant") is not None:
        _row(rows, "Déclaration", "pied de page : VAL = montant facturé", p["val"], d["montant"], 0.01)
    return rows


def compare(d: dict, vent, rate: float, fret: float, fact_df=None, bl_df=None) -> list[dict]:
    """Déclaration ↔ données du site (ventilation en cours, factures, BL / packing lists)."""
    rows = self_checks(d)
    t = vent.totals
    _row(rows, "Dossier", "Poids net total (kg)", d.get("net"), float(t.get("net", 0.0)), 0.5)
    _row(rows, "Dossier", "Poids brut total (kg)", d.get("brut"), float(t.get("brut", 0.0)), 0.5)
    _row(rows, "Dossier", f"Montant total facturé ({d.get('devise', '')})", d.get("montant"), float(t.get("valeur", 0.0)), 0.05, "{:,.3f}")
    _row(rows, "Dossier", "Taux de change", d.get("taux"), float(rate), 0.00005, "{:.6f}")
    _row(rows, "Dossier", "Fret (devise)", d.get("fret"), float(fret or 0.0), 0.005, "{:,.3f}")
    n_art = len([a for a in vent.articles if a.lines])
    _row(rows, "Dossier", "Nombre d'articles", float(len(d.get("articles", []))), float(n_art), 0.0, "{:g}")
    if bl_df is not None and "_currency" in bl_df:
        cur = sorted({str(c) for c in bl_df["_currency"] if str(c) not in ("", "nan", "None")})
        if cur and d.get("devise"):
            rows.append({"Niveau": "Dossier", "Contrôle": "Devise", "Déclaration": d["devise"], "Dossier (site)": ", ".join(cur), "Écart": "",
                         "Résultat": "✔" if cur == [d["devise"]] else "⚠ écart"})
    if fact_df is not None and "Numéro de facture" in fact_df:
        site_inv = [str(x).strip() for x in fact_df["Numéro de facture"].dropna()]
        pied = d.get("pied", {})
        if pied.get("n_factures") is not None:
            _row(rows, "Dossier", "Nombre de factures", float(pied["n_factures"]), float(len(set(site_inv))), 0.0, "{:g}")
            missing = [x for x in pied.get("factures", []) if x not in site_inv]
            if pied.get("factures"):
                rows.append({"Niveau": "Dossier", "Contrôle": "Factures citées dans la déclaration", "Déclaration": f"{len(pied['factures'])} citée(s)"
                             + (" (liste tronquée)" if pied.get("truncated") else ""), "Dossier (site)": f"{len(site_inv) - len(missing)} retrouvée(s)",
                             "Écart": ", ".join(missing), "Résultat": "✔" if not missing else "⚠ absente(s) du dossier"})
    # articles : regroupés par code marchandises (NGP)
    site: dict = {}
    for a in vent.articles:
        g = site.setdefault(str(a.ngp or "?"), {"qte": 0.0, "net": 0.0, "val": 0.0, "name": a.template_name})
        g["qte"] += a.qty
        g["net"] += a.net
        g["val"] += a.value
    seen = set()
    for a in d.get("articles", []):
        niv = f"Article {a['n']} · {a['ngp']}"
        g = site.get(a["ngp"])
        seen.add(a["ngp"])
        if g is None:
            rows.append({"Niveau": niv, "Contrôle": "Code marchandises (NGP)", "Déclaration": a["ngp"] + " — " + a.get("designation", ""),
                         "Dossier (site)": "absent de la ventilation", "Écart": "", "Résultat": "⚠ NGP absent"})
            continue
        rows.append({"Niveau": niv, "Contrôle": "Code marchandises (NGP)", "Déclaration": a["ngp"], "Dossier (site)": f"{a['ngp']} — {g['name']}", "Écart": "", "Résultat": "✔"})
        _row(rows, niv, "Quantité", a.get("qte"), g["qte"], 0.0, "{:g}")
        _row(rows, niv, "Poids net (kg)", a.get("net"), g["net"], 0.5)
        _row(rows, niv, "Valeur déclarée (DH)", a.get("valeur"), g["val"] * float(d.get("taux") or rate), 1.5, "{:,.3f}")
    for ngp, g in site.items():
        if ngp not in seen:
            rows.append({"Niveau": f"Article · {ngp}", "Contrôle": "Présent dans la ventilation", "Déclaration": "absent de la déclaration",
                         "Dossier (site)": f"{g['name']} · qté {g['qte']:g} · {g['net']:.2f} kg", "Écart": "", "Résultat": "⚠ article non déclaré"})
    return rows


def render(vent, rate: float, fret: float, bl_df=None, fact_df=None, key: str = "vx_dum"):
    """Case « Vérification de la déclaration » (EXPORT) : dépôt de la DUM téléchargée de BADR, contrôle automatique avec la ventilation,
    les factures et les packing lists du dossier."""
    import io

    import pandas as pd
    import streamlit as st
    st.markdown("#### ✅ Vérification de la déclaration (DUM de BADR)")
    st.caption("Déposez la déclaration téléchargée de BADR (PDF) : le site la compare avec la ventilation ci-dessus, les factures et les packing lists de ce dossier.")
    ups = st.file_uploader("Déclaration DUM (PDF)", type=["pdf"], accept_multiple_files=True, key=key)
    if not ups:
        return
    decls = []
    for f in ups:
        try:
            for d in parse_all(f.getvalue()):
                d["_file"] = f.name
                decls.append(d)
        except Exception as exc:  # noqa: BLE001
            st.error(f"{f.name} : lecture impossible ({type(exc).__name__}: {exc}).")
    decls = [d for d in decls if d.get("articles") or d.get("net") is not None]
    if not decls:
        st.warning("Aucune déclaration lisible : le PDF doit être la DUM imprimée par BADR (texte, pas une image scannée).")
        return
    t = vent.totals
    # plusieurs DUM déposées : on propose d'abord celle qui ressemble le plus au dossier (poids net / montant)
    score = lambda d: abs((d.get("net") or 0) - float(t.get("net", 0))) + abs((d.get("montant") or 0) - float(t.get("valeur", 0))) / 100  # noqa: E731
    order = sorted(range(len(decls)), key=lambda i: score(decls[i]))
    labels = [f"{decls[i].get('numero', '?')} — {decls[i].get('destinataire', '')} ({decls[i]['_file']})" for i in order]
    pick = st.selectbox("Déclaration à vérifier", labels, key=f"{key}_pick") if len(decls) > 1 else labels[0]
    d = decls[order[labels.index(pick)]]
    rows = compare(d, vent, rate, fret, fact_df, bl_df)
    bad = [r for r in rows if r["Résultat"].startswith("⚠")]
    info = [r for r in rows if r["Résultat"].startswith("ℹ")]
    ok = len(rows) - len(bad) - len(info)
    head = f"DUM {d.get('numero', '')} du {d.get('date', '')} · {d.get('destinataire', '')}"
    if bad:
        st.warning(f"{head} — **{len(bad)} point(s) à vérifier**, {ok} contrôle(s) conforme(s).")
    else:
        st.success(f"{head} — **conforme** : {ok} contrôle(s) réussi(s), aucun écart.")
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="Vérification DUM")
    st.download_button("Télécharger la vérification (.xlsx)", data=buf.getvalue(), file_name=f"VERIFICATION_DUM_{str(d.get('numero', '')).replace(' ', '_')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{key}_dl", icon=":material/download:")
