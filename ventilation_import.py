# -*- coding: utf-8 -*-
"""Ventilation IMPORT (DUM) : remplit le modèle VENTI_IMPORT à partir du tableau TOUS_CLIENTS (factures lues).

Différences avec l'export (voir ventilation_export.py, dont on réutilise la logique de classement) :
  * les NGP viennent de la feuille numérotée du destinataire (ATTELIER col. C -> feuille « 6 » pour ALSTOM), NGPN est vide ;
  * un article de DUM = une désignation/NGP **et** un pays d'origine (chaque ligne de facture a son pays) ;
  * VAL DHS d'un article = valeur devise × coefficient, où coefficient = VTD / valeur devise et
    VTD = ACC (acconage, DH) + ASS (assurance = 0,5 % de (valeur × taux + fret DH)) + valeur DH + FRET DH ;
  * pas de code accord (laissé vide) ; les pays de l'UE d'un même NGP sont regroupés (origine « DIVERS »), les autres pays restent séparés ;
  * ANNEXE : paiement « SP », VALEUR = ROUNDUP(VAL DHS).
"""

from __future__ import annotations

import io
import math
import re

import pandas as pd
from openpyxl import load_workbook

import ventilation_export as vx

KIND = "import"
EU = {"AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT",
      "RO", "SK", "SI", "ES", "SE"}
COUNTRIES = {
    "AT": "austria autriche österreich osterreich", "BE": "belgium belgique belgica belgië", "BG": "bulgaria bulgarie",
    "HR": "croatia croatie croacia", "CY": "cyprus chypre", "CZ": "czech tchèque tcheque checa tschechien", "DK": "denmark danemark dinamarca",
    "EE": "estonia estonie", "FI": "finland finlande finlandia", "FR": "france francia frankreich", "DE": "germany allemagne alemania deutschland",
    "GR": "greece grèce grece grecia", "HU": "hungary hongrie hungria ungarn", "IE": "ireland irlande irlanda", "IT": "italy italie italia italien",
    "LV": "latvia lettonie letonia", "LT": "lithuania lituanie lituania", "LU": "luxembourg luxemburgo", "MT": "malta malte",
    "NL": "netherlands pays-bas holland hollande paises niederlande", "PL": "poland pologne polonia polen", "PT": "portugal",
    "RO": "romania roumanie rumania rumänien", "SK": "slovakia slovaquie eslovaquia", "SI": "slovenia slovénie slovenie eslovenia",
    "ES": "spain espagne españa espana spanien", "SE": "sweden suède suede suecia schweden", "GB": "united kingdom royaume-uni great britain england uk gbr reino unido",
    "CH": "switzerland suisse suiza schweiz", "NO": "norway norvège norvege noruega", "TR": "turkey turquie turquia türkiye turkiye",
    "CN": "china chine mainland", "HK": "hong kong", "TW": "taiwan taïwan", "JP": "japan japon", "KR": "korea corée coree south korea",
    "IN": "india inde", "US": "usa united states états-unis etats-unis estados unidos america", "CA": "canada", "MX": "mexico mexique",
    "BR": "brazil brésil bresil brasil", "VN": "vietnam viêt", "TH": "thailand thaïlande thailande", "MY": "malaysia malaisie",
    "ID": "indonesia indonésie indonesie", "SG": "singapore singapour", "EG": "egypt egipto égypte egypte", "TN": "tunisia tunisie",
    "MA": "morocco maroc marruecos", "ZA": "south africa afrique du sud", "UA": "ukraine", "RS": "serbia serbie", "BA": "bosnia bosnie",
    "MK": "macedonia macédoine", "AL": "albania albanie", "RU": "russia russie", "IL": "israel israël", "AE": "emirates émirats",
}
_NAME_TO_CODE = {w: c for c, names in COUNTRIES.items() for w in names.split()}
_NAME_TO_CODE.update({"hong kong": "HK", "united kingdom": "GB", "south korea": "KR", "united states": "US", "south africa": "ZA",
                      "czech republic": "CZ", "great britain": "GB", "pays-bas": "NL", "royaume-uni": "GB", "états-unis": "US"})


def origin_code(value) -> str:
    """Code pays ISO (2 lettres) depuis « DE », « Germany », « MADE IN GERMANY », « Allemagne »… ; chaîne vide si inconnu."""
    s = str(value or "").strip()
    if not s or s.lower() in ("nan", "none"):
        return ""
    if len(s) == 2 and s.isalpha():
        return s.upper()
    low = s.casefold()
    for name in sorted(_NAME_TO_CODE, key=len, reverse=True):          # noms composés d'abord
        if len(name) > 3 and name in low:
            return _NAME_TO_CODE[name]
    for tok in re.findall(r"[A-Za-zÀ-ÿ\-]+", low):
        if tok in _NAME_TO_CODE:
            return _NAME_TO_CODE[tok]
    m = re.search(r"\b([A-Z]{2})\b", s)
    return m.group(1) if m and m.group(1) in COUNTRIES else ""


DIVERS = "DIVERS"


def accord_for(code: str) -> str:
    return "UE" if code in EU else ("UK" if code == "GB" else "")


# ----------------------------------------------------------------------------- données TOUS_CLIENTS -> lignes normalisées
C_NET, C_NETF, C_GROSSF, C_COLIS = "Poids net ligne (kg)", "Poids net facture (kg)", "Poids brut facture (kg)", "Nb colis / palettes (facture)"


def ngp_digits(value) -> str:
    """« 3926.90.92.90 » / « 3926909290 » -> « 3926909290 » ; chaîne vide si ce n'est pas un code NGP (ex. « NON TROUVÉ »)."""
    d = re.sub(r"\D", "", str(value or ""))
    return d if len(d) >= 8 else ""


HEADING_FAMILLE = {"8544": "8544499090", "8546": "8546100000", "8547": "8546100000",
                   "8413": "8413110000", "8414": "8413110000", "8481": "8413110000"}
CHAPTER_FAMILLE = {35: "3506109000", 39: "3926909290", 40: "4016999890", 48: "4821900000", 59: "5911908090",
                   73: "7326909990", 74: "7326909990", 76: "7326909990", 83: "8307100090", 85: "8537100090", 90: "9031800080"}


def regroupe(ngp) -> str:
    """NGP « classement regroupé » du dossier ALSTOM pour n'importe quel NGP (10 chiffres du référentiel ou 8 chiffres d'une facture) :
    exact, sinon par position (4 chiffres), sinon par chapitre. Vide si le code n'est pas rattachable (ex. numéro d'article pris pour un NGP)."""
    d = ngp_digits(ngp)
    if not d:
        return ""
    for k in vx.IMPORT_FAMILLES:
        if k == d or k[:8] == d[:8]:
            return k
    if d[:4] in HEADING_FAMILLE:
        return HEADING_FAMILLE[d[:4]]
    try:
        return CHAPTER_FAMILLE.get(int(d[:2]), "")
    except ValueError:
        return ""


def famille_of(ngp: str, reference: vx.Reference | None = None) -> str:
    """Famille de produits (désignation d'article) d'un NGP : tableau de classement, sinon feuille du destinataire."""
    if not ngp:
        return ""
    if ngp in vx.IMPORT_FAMILLES:
        return vx.IMPORT_FAMILLES[ngp]
    if reference is not None and ngp in reference.ngp_to_name:
        return reference.ngp_to_name[ngp]
    return f"NGP {ngp}"


def _first(series: pd.Series):
    s = series.dropna()
    return s.iloc[0] if len(s) else float("nan")


def default_ratio(final: pd.DataFrame) -> float:
    """Ratio net/brut observé sur les factures qui ont les deux poids (sinon 90 %)."""
    try:
        d = final[["N° facture", C_NET, C_GROSSF]].copy()
        d[C_NET] = pd.to_numeric(d[C_NET], errors="coerce")
        d[C_GROSSF] = pd.to_numeric(d[C_GROSSF], errors="coerce")
        g = d.groupby("N° facture").agg(net=(C_NET, "sum"), brut=(C_GROSSF, _first), n=(C_NET, "count"))
        g = g[(g["n"] > 0) & (g["brut"] > 0) & (g["net"] > 0)]
        if g["brut"].sum() > 0:
            return float(min(max(g["net"].sum() / g["brut"].sum(), 0.5), 1.0))
    except Exception:  # noqa: BLE001
        pass
    return 0.9


def normalize(final: pd.DataFrame, reference: vx.Reference, ratio: float | None = None) -> tuple[pd.DataFrame, dict, list[str]]:
    """Tableau TOUS_CLIENTS (après corrections) -> lignes de marchandises + totaux.
    Lignes de frais exclues ; poids net manquant estimé : (poids net facture, sinon poids brut × ratio) réparti sur les lignes
    sans poids au prorata de la quantité."""
    notes: list[str] = []
    df = final.copy().reset_index(drop=True)

    def is_fee(v) -> bool:                                   # « False » (texte) ne doit pas compter comme vrai
        if isinstance(v, str):
            return v.strip().lower() in ("true", "1", "oui", "yes", "vrai")
        return bool(v) if v is not None and v == v else False
    fee = df["_fee"].map(is_fee) if "_fee" in df else pd.Series([False] * len(df))
    if len(fee) and fee.all():                               # tout marqué « frais » : probablement une erreur, on garde tout
        notes.append("Toutes les lignes étaient marquées « frais » : elles sont conservées comme marchandises.")
        fee = pd.Series([False] * len(df))
    if fee.any():
        notes.append(f"{int(fee.sum())} ligne(s) de frais (transport, emballage…) exclue(s) des articles : saisissez le fret dans « Fret (devise) ».")
    df = df[~fee.values].reset_index(drop=True)
    n = len(df)
    num = lambda c: pd.to_numeric(df[c], errors="coerce") if c in df else pd.Series([float("nan")] * n)  # noqa: E731
    txt = lambda c: df[c].map(lambda v: str(v or "").strip()) if c in df else pd.Series([""] * n)  # noqa: E731
    ref_main, ref_alt = txt("Réf article"), txt("_réf_alstom")
    pick = [m if vx.norm_ref(m) in reference.base or not a or vx.norm_ref(a) not in reference.base else a for m, a in zip(ref_main, ref_alt)]
    inv_raw = df["N° facture"].map(lambda v: str(v or "").strip() or "(sans facture)")
    clients = txt("Client")
    multi = inv_raw.groupby(inv_raw).transform(lambda s: clients[s.index].nunique() > 1)
    inv = [f"{c} · {i}" if d else i for c, i, d in zip(clients, inv_raw, multi)]
    key = pd.Series([f"{c}§{i}" for c, i in zip(clients, inv_raw)])
    ngp_ref = df["NGP (REF_PAR_FRS)"].map(ngp_digits) if "NGP (REF_PAR_FRS)" in df else pd.Series([""] * n)
    ngp_fac = df["NGP/code douanier (facture)"].map(ngp_digits) if "NGP/code douanier (facture)" in df else pd.Series([""] * n)
    out = pd.DataFrame({
        "Reference": pick, "Description": txt("Désignation"), "_client": clients, "_inv": inv, "_bl": inv,
        "_val": num("Valeur HT").fillna(0.0), "_qty": num("Qté").fillna(0.0),
        "_origin": df["Origine"].map(origin_code) if "Origine" in df else [""] * n,
        "_cur": txt("_devise")})
    ngp_final, src = [], []
    for r, f in zip(ngp_ref, ngp_fac):
        g = regroupe(r) or (r if r and len(r) >= 10 else "")            # référentiel : fait foi, même hors tableau de familles
        if g:
            ngp_final.append(g)
            src.append("référentiel")
            continue
        ngp_final.append("")                                              # le NGP vient UNIQUEMENT du tableau de référence
        src.append("")
    out["_ngp"], out["_ngp_src"] = ngp_final, src
    # ---- poids net : ligne imprimée, sinon estimation par facture
    net = num(C_NET)
    gross_inv = num(C_GROSSF).groupby(key).transform(_first)
    netf_inv = num(C_NETF).groupby(key).transform(_first)
    r = default_ratio(final) if ratio is None else ratio
    est = pd.Series([False] * n)
    net_out = net.copy()
    for k, idx in key.groupby(key).groups.items():
        idx = list(idx)
        known = net.loc[idx]
        missing = [i for i in idx if not (net.loc[i] > 0)]
        if not missing:
            continue
        target = netf_inv.loc[idx[0]] if netf_inv.loc[idx[0]] > 0 else (gross_inv.loc[idx[0]] * r if gross_inv.loc[idx[0]] > 0 else float("nan"))
        if not target == target:
            continue
        remaining = max(target - float(known[known > 0].sum()), 0.0)
        w = out.loc[missing, "_qty"].abs().tolist()
        if not sum(w):
            w = out.loc[missing, "_val"].abs().tolist()
        if not sum(w):
            w = [1.0] * len(missing)
        for i, v in zip(missing, cleaning_allocate(remaining, w)):
            net_out.loc[i] = v
            est.loc[i] = True
    # ---- repli : densité du dossier (kg par unité de valeur) pour les lignes encore sans poids
    still = [i for i in range(n) if not (net_out.iloc[i] > 0)]
    if still:
        have = [i for i in range(n) if net_out.iloc[i] > 0 and out["_val"].iloc[i] > 0]
        tv = float(sum(out["_val"].iloc[i] for i in have))
        if tv > 0:
            dens = float(sum(net_out.iloc[i] for i in have)) / tv
            for i in still:
                if out["_val"].iloc[i] > 0:
                    net_out.iloc[i] = round(out["_val"].iloc[i] * dens, 3)
                    est.iloc[i] = True
            notes.append(f"{len(still)} ligne(s) sans aucun poids : estimées par la densité moyenne du dossier ({dens:.4f} kg/€) — à vérifier.")
    out["_net"] = net_out.fillna(0.0)
    out["_net_est"] = est
    # ---- champs manquants : quantité / valeur / origine
    if (out["_qty"] <= 0).any() and (out["_val"] > 0).any():
        out.loc[out["_qty"] <= 0, "_qty"] = 1.0
        notes.append("Quantité absente sur certaines lignes : mise à 1 (à vérifier).")
    main_or = out.loc[out["_origin"] != "", "_origin"].groupby(out["_client"]).agg(lambda s: s.mode().iat[0])
    miss = out["_origin"] == ""
    if miss.any() and len(main_or):
        out.loc[miss, "_origin"] = out.loc[miss, "_client"].map(main_or).fillna("")
    n_est, n_zero = int(est.sum()), int((out["_net"] <= 0).sum())
    if n_est:
        notes.append(f"Poids net estimé pour {n_est} ligne(s) : {'poids net de la facture' if (netf_inv > 0).any() else f'poids brut × {r * 100:.0f} %'} "
                     "réparti sur les lignes sans poids au prorata de la quantité (à vérifier).")
    if n_zero:
        notes.append(f"{n_zero} ligne(s) restent sans poids net (ni net ni brut de facture connu) : saisissez-le dans le tableau TOUS_CLIENTS.")
    if (out["_origin"] == "").any():
        k = int((out["_origin"] == "").sum())
        unk = sorted({str(o).strip() for o, c in zip(df.get("Origine", []), out["_origin"]) if not c and str(o or "").strip()})
        notes.append(f"{k} ligne(s) sans pays d'origine reconnu" + (f" ({', '.join(unk[:6])})" if unk else "")
                     + " : complétez la colonne « Origine » dans le tableau TOUS_CLIENTS (ex. DE, FR, CN).")
    cur = sorted({c for c in out["_cur"] if c})
    if len(cur) > 1:
        notes.append(f"Plusieurs devises dans le dossier ({', '.join(cur)}) : le modèle n'a qu'un taux — vérifiez les montants.")
    pk = pd.DataFrame({"k": key, "colis": num(C_COLIS), "brut": num(C_GROSSF)}).groupby("k").agg(colis=("colis", _first), brut=("brut", _first))
    totals = {"colis": float(pk["colis"].sum()), "brut": float(pk["brut"].sum()), "net": float(out["_net"].sum()),
              "valeur": float(out["_val"].sum()), "quantite": float(out["_qty"].sum()), "lignes": len(out)}
    return out, totals, notes


def cleaning_allocate(total: float, weights: list[float]) -> list[float]:
    """Répartit `total` au prorata de `weights` (somme exacte au gramme)."""
    import cleaning
    return cleaning._allocate(total, weights)


# ----------------------------------------------------------------------------- poids net imposé par article (mode manuel)
def apply_fixed_net(vent: vx.Ventilation, df: pd.DataFrame, manual: dict) -> list[str]:
    """Poids net saisi à la main pour certains articles (cas où la packing list ne donne que le total de plusieurs pièces).

    manual["fixed_net"] = {n° d'article (0 = article 1, …): kg par facture (ou par référence)} ; manual["fixed_unit"] = "inv" | "ref".
    Poids de l'article = kg saisi × nombre de factures (références) de l'article. Le reste du poids net du dossier est réparti
    sur les autres articles. Le total du dossier ne change pas."""
    fixed = {int(k): float(v) for k, v in manual["fixed_net"].items() if v and float(v) > 0}
    if not fixed:
        return []
    by_ref = manual.get("fixed_unit") == "ref"
    n_art = len(manual["articles"])
    assign, assign_ref = manual.get("assign", {}), manual.get("assign_ref")
    used = []                                                       # n° d'article (0-based) de chaque article conservé (non vide)
    seen = set()
    for _, r in df.iterrows():
        k = int(assign_ref.get(vx.norm_ref(r["Reference"]), 1)) - 1 if assign_ref is not None else int(assign.get(r["_inv"], 1)) - 1
        seen.add(min(max(k, 0), n_art - 1))
    used = sorted(seen)
    if len(used) != len(vent.articles):
        return []
    notes: list[str] = []
    total_net = float(vent.totals.get("net", 0.0))
    fixed_total, free = 0.0, []
    for pos, idx in enumerate(used):
        art = vent.articles[pos]
        if idx in fixed:
            units = len({str(l["ref"]).strip() if by_ref else l["inv"] for l in art.lines}) or 1
            target = fixed[idx] * units
            w = [abs(l["value"]) for l in art.lines]
            tot = sum(w) or float(len(art.lines))
            for l, x in zip(art.lines, w if sum(w) else [1.0] * len(art.lines)):
                l["net"] = target * x / tot
            fixed_total += target
            notes.append(f"Article {idx + 1} : poids net imposé = {fixed[idx]:g} kg × {units} {'référence(s)' if by_ref else 'facture(s)'} = {target:g} kg.")
        else:
            free.append(art)
    rest = total_net - fixed_total
    if rest < -0.0005:
        notes.append(f"⚠ Les poids imposés ({fixed_total:g} kg) dépassent le poids net du dossier ({total_net:g} kg) : vérifiez les valeurs saisies.")
        rest = 0.0
    if free:
        cur = [abs(l["net"]) for a in free for l in a.lines]
        base = cur if sum(cur) else [abs(l["value"]) for a in free for l in a.lines]
        tot = sum(base) or float(len(base))
        if not sum(base):
            base = [1.0] * len(base)
        j = 0
        for a in free:
            for l in a.lines:
                l["net"] = rest * base[j] / tot
                j += 1
        notes.append(f"Poids net restant ({rest:g} kg) réparti sur les autres articles.")
    else:
        notes.append(f"⚠ Tous les articles ont un poids imposé : total {fixed_total:g} kg pour un poids net de dossier de {total_net:g} kg.")
    return notes


# ----------------------------------------------------------------------------- ventilation (articles)
def build(final: pd.DataFrame, reference: vx.Reference, guess: bool = True, manual: dict | None = None,
          split_origin: bool = False, ratio: float | None = None) -> vx.Ventilation:
    """Automatique : un article par NGP (les désignations de même NGP sont regroupées), désignation = famille de produits.
    split_origin=True : un article par NGP **et** pays d'origine."""
    df, totals, notes = normalize(final, reference, ratio)
    if manual:
        vent = vx._build_manual(df, reference, manual, totals)
        vent.notes = notes + vent.notes
        if manual.get("fixed_net"):
            vent.notes += apply_fixed_net(vent, df, manual)
        return vent
    articles: dict[tuple, vx.Article] = {}
    missing: dict[str, dict] = {}
    for _, r in df.iterrows():
        ref = str(r["Reference"]).strip()
        ngp, how, name = r["_ngp"], "", ""
        if ngp:
            name = famille_of(ngp, reference)
        org = r["_origin"]
        key = (ngp or "?", org if (split_origin or org not in EU) else "UE")       # UE regroupée ; CH, GB, US, CN… séparés
        if key not in articles:
            articles[key] = vx.Article(ngp, name, name, origin=r["_origin"]) if ngp else vx.Article("", vx.UNCLASSIFIED, vx.UNCLASSIFIED, origin=r["_origin"])
        articles[key].lines.append({"ref": ref, "desc": r["Description"], "qty": float(r["_qty"]), "inv": r["_inv"], "net": float(r["_net"]),
                                    "value": float(r["_val"]), "bl": r["_bl"], "suggested": bool(how), "origin": r["_origin"],
                                    "client": r["_client"], "net_est": bool(r["_net_est"])})
        if not ngp or how:
            m = missing.setdefault(ref, {"Référence": ref, "Client": r["_client"], "Description": r["Description"][:60], "Quantité": 0.0,
                                         "Valeur": 0.0, "Classement": (name + " (suggéré)") if how else "NGP introuvable"})
            m["Quantité"] += float(r["_qty"])
            m["Valeur"] += float(r["_val"])
    ordered = sorted(articles.values(), key=lambda a: (a.ngp == "", a.ngp))
    for a in ordered:                                                     # origine de l'article : le pays, ou « DIVERS » pour plusieurs pays UE
        origs = sorted({l["origin"] for l in a.lines if l.get("origin")})
        a.origin = origs[0] if len(origs) == 1 else (DIVERS if origs else "")
    src = df.groupby("_ngp_src").size().to_dict()
    if missing:
        notes.append(f"{len(missing)} référence(s) sans NGP dans le tableau de référence — "
                     "ajoutez-les au référentiel Excel.")
    if len(ordered) > vx.MAX_ARTICLES:
        notes.append(f"Plus de {vx.MAX_ARTICLES} articles : le modèle n'en accepte que {vx.MAX_ARTICLES}.")
    notes.append("NGP lus : " + ", ".join(f"{k or 'introuvable'} ({v})" for k, v in src.items()) + " ligne(s).")
    return vx.Ventilation(ordered, totals, list(missing.values()), notes)


# ----------------------------------------------------------------------------- calculs du modèle (VENTILATION F2:F5, H3)
def costs(vent: vx.Ventilation, rate: float, fret: float, acc: float) -> dict:
    v = vent.totals["valeur"]
    val_dh = rate * v
    fret_dh = rate * (fret or 0.0)
    ass = (val_dh + fret_dh) * 0.005
    vtd = (acc or 0.0) + ass + val_dh + fret_dh
    return {"valeur_dh": val_dh, "fret_dh": fret_dh, "ass": ass, "acc": acc or 0.0, "vtd": vtd, "coef": (vtd / v) if v else 0.0}


# ----------------------------------------------------------------------------- remplissage du modèle
def fill_template(template_bytes: bytes, vent: vx.Ventilation, destinataire: str, rate: float, fret: float | None,
                  acc: float | None, matricule: str = "") -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), keep_links=False)
    v = wb["VENTILATION"]
    t = vent.totals
    v["C2"], v["C3"], v["C4"] = t["colis"], t["brut"], t["net"]
    v["C5"], v["C6"] = round(t["valeur"], 2), t["quantite"]
    v["I5"] = destinataire
    v["G8"] = float(rate)                                          # remplace =[1]Feuil1!$B$2
    v["L8"] = float(fret or 0.0)                                   # FRET (devise) ; C7 = taux × fret reste en formule
    v["F4"] = float(acc or 0.0)                                    # ACC : acconage en DH
    if matricule:
        v["I7"] = matricule
    # cases de contrôle (total du dossier − somme des articles) : arrondies, vertes quand elles valent 0, rouges sinon
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    v["E11"], v["G11"], v["H11"] = "=ROUND(C6-SUM(E13:E32),2)", "=ROUND(G33-C4,2)", "=ROUND(H33-C5,0)"
    green, red = PatternFill("solid", start_color="92D050", end_color="92D050"), PatternFill("solid", start_color="FF7C80", end_color="FF7C80")
    for c in ("E11", "G11", "H11"):
        v.conditional_formatting.add(c, FormulaRule(formula=[f"ABS({c})<0.005"], fill=green, stopIfTrue=True))
        v.conditional_formatting.add(c, FormulaRule(formula=[f"ABS({c})>=0.005"], fill=red))
    desc = wb["DESCIPTION"]
    row_desc = 3
    for n, art in enumerate(vent.articles[:vx.MAX_ARTICLES], start=1):
        ws = wb[f"ART{n}"]
        ws["A2"] = art.template_name                                # famille de produits
        if art.ngp:
            ws["B2"] = int(art.ngp)                                    # NGP écrit directement (la famille n'est pas dans la feuille du destinataire)
        ws["C2"] = "NB"
        ws["B4"] = None                                              # pas de code accord
        for k, line in enumerate(art.lines[:vx.LAST_LINE_ROW - vx.FIRST_LINE_ROW + 1]):
            r = vx.FIRST_LINE_ROW + k
            org = line.get("origin") or art.origin
            price = round(line["value"] / line["qty"], 4) if line["qty"] else 0
            ws.cell(r, 4, line["ref"])
            ws.cell(r, 6, org)
            ws.cell(r, 7, round(line["net"], 3))
            ws.cell(r, 8, line["qty"])
            ws.cell(r, 9, price)
            ws.cell(r, 10, round(line["value"], 2))
            desc.cell(row_desc, 2, line["ref"])
            desc.cell(row_desc, 4, org)
            desc.cell(row_desc, 5, round(line["net"], 3))
            desc.cell(row_desc, 6, line["qty"])
            desc.cell(row_desc, 7, price)
            desc.cell(row_desc, 8, round(line["value"], 2))
            row_desc += 1
        v.cell(12 + n, 4, art.origin or None)                      # ORGN (le modèle la déduit d'un tableau limité à 22 pays)
        v.cell(12 + n, 11, None)                                    # pas de code accord
    wb.template = False
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def annexe_rows(vent: vx.Ventilation, rate: float, fret: float, acc: float, matricule: str = "") -> list[list]:
    c = costs(vent, rate, fret, acc)
    rows = []
    for i, a in enumerate(vent.articles[:vx.MAX_ARTICLES], start=1):
        value_mad = math.ceil(round(a.value * c["coef"], 6))             # ROUNDUP(VAL DHS, 0)
        ngp = int(a.ngp) if a.ngp else None
        qn = a.qty if (i == 1 and ngp == 8531800000) else round(a.net, 3)
        rows.append([None, i, vent.totals["colis"] if i == 1 else None, 216 if i == 1 else "000", matricule if i == 1 else None,
                     ngp, a.template_name, a.origin or None, "SP", "NON", value_mad, "MAD", a.qty, "002", round(a.net, 3), qn,
                     None, None, None, None])
    return rows


def build_annexe(vent: vx.Ventilation, rate: float, fret: float, acc: float, matricule: str = "",
                 template_bytes: bytes | None = None) -> bytes:
    return vx.build_annexe_xlsx(vent, rate, matricule, template_bytes, rows=annexe_rows(vent, rate, fret, acc, matricule))
