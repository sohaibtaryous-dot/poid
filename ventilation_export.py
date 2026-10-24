# -*- coding: utf-8 -*-
"""Ventilation EXPORT (DUM) : remplit automatiquement le modèle VENTI_EXPORT à partir des BL / factures.

Principe (identique au travail manuel dans le modèle) :
  1. chaque référence est cherchée dans la feuille « BASE DONNER » du modèle -> désignation précise (ex. « Conducteurs électriques »,
     « Parties d'armoires électriques ») ;
  2. la désignation donne le NGP via la feuille « NGPN » ;
  3. les lignes d'un même NGP forment un article de DUM (feuille ART1…ART20) : quantité, poids net, valeur devise ;
  4. les totaux et le destinataire / N° DUM alimentent la feuille « VENTILATION » ; les formules du modèle (VAL DHS, ANNEXE…) restent vivantes.
Les références absentes de BASE DONNER sont regroupées dans un article « À CLASSER » (jamais perdues, jamais devinées).
Le modèle (.xltx) est dans templates/VENTI_EXPORT.xltx ; l'Admin peut le remplacer depuis l'application."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import cleaning

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "templates" / "VENTI_EXPORT.xltx"
TEMPLATE_FILES = {"export": "VENTI_EXPORT.xltx", "import": "VENTI_IMPORT.xltx"}
UNCLASSIFIED = "À CLASSER (référence introuvable dans BASE DONNER)"
MAX_ARTICLES, FIRST_LINE_ROW, LAST_LINE_ROW = 20, 3, 472        # feuilles ART1..ART20 : lignes 3 à 472
ORIGIN_NAME, ORIGIN_CODE = "MAROC", "MA"


def override_path(kind: str = "export") -> Path:
    import database
    return database.DATA_DIR / "templates" / TEMPLATE_FILES[kind]


def get_template_bytes(kind: str = "export") -> bytes | None:
    """Modèle téléversé par l'Admin s'il existe, sinon le modèle livré avec l'application."""
    ov = override_path(kind)
    if not ov.is_file():                                           # modèle de l'admin perdu au redémarrage : récupéré sur Supabase
        try:
            import persist
            data = persist.get_file(f"templates/{TEMPLATE_FILES[kind]}")
            if data:
                ov.parent.mkdir(parents=True, exist_ok=True)
                ov.write_bytes(data)
        except Exception:  # noqa: BLE001
            pass
    for p in (ov, Path(__file__).resolve().parent / "templates" / TEMPLATE_FILES[kind]):
        try:
            if p.is_file():
                return p.read_bytes()
        except OSError:
            continue
    return None


def save_template_override(data: bytes, kind: str = "export"):
    p = override_path(kind)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    try:
        import persist
        persist.put_file(f"templates/{TEMPLATE_FILES[kind]}", data)
    except Exception:  # noqa: BLE001
        pass


@dataclass
class Reference:
    base: dict[str, dict]                 # REF normalisée -> {designation, annexe, ngp}
    ngpn: dict[str, dict]                 # désignation (casefold) -> {designation, ngp, rubrique}
    ngp_to_name: dict[str, str]           # NGP -> désignation du modèle (1re ligne NGPN avec ce NGP)
    destinations: dict[str, str]          # destinataire (nom nettoyé) -> préfixe N° DUM
    dest_raw: dict[str, str] = field(default_factory=dict)   # nom nettoyé -> nom EXACT de la cellule ATTELIER (espaces finaux compris)
    kind: str = "export"                  # « import » : NGP lus dans la feuille numérotée du destinataire (ATTELIER col. C), NGPN vide


def norm_ref(x) -> str:
    return re.sub(r"\s+", "", str(x if x is not None else "")).upper()


def _ngp(v) -> str:
    if v is None or v == "":
        return ""
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return str(v).strip()


# Import : « Famille de produits » et « Classement regroupé » (NGP) — tableau du dossier de classement ALSTOM
IMPORT_FAMILLES = {
    "3506109000": "Colles", "3926909290": "Ouvrages en matières plastiques", "4016999890": "Ouvrages en caoutchouc",
    "4821900000": "Etiquettes et carton", "5911908090": "Produits textiles", "7326909990": "Ouvrages en métaux",
    "8307100090": "Flexibles et serrures", "8413110000": "Pompes et robinetteries", "8537100090": "Appareils électriques",
    "8544499090": "Câblages", "8546100000": "Isolateurs", "9031800080": "Instruments de mesure",
}

EXTRA_ARTICLES = [("Armoires électriques et leurs parties", "Armoires électriques")]


def load_reference(template_bytes: bytes) -> Reference:
    """Lit BASE DONNER, NGPN et ATTELIER dans le modèle (lecture seule, valeurs)."""
    wb = load_workbook(io.BytesIO(template_bytes), read_only=True, data_only=True)
    base: dict[str, dict] = {}
    for row in wb["BASE DONNER"].iter_rows(min_row=2, values_only=True):
        ref = norm_ref(row[2]) if len(row) > 2 else ""
        if ref and ref not in base:
            base[ref] = {"designation": str(row[9] or "").strip(), "annexe": str(row[10] or "").strip(), "ngp": _ngp(row[4])}
    ngpn: dict[str, dict] = {}
    ngp_to_name: dict[str, str] = {}
    destinations, dest_raw, codes = {}, {}, []
    for r in wb["ATTELIER"].iter_rows(min_row=1, values_only=True):
        if r[0]:
            clean = str(r[0]).strip()
            destinations[clean] = str(r[1] or "").strip()
            dest_raw[clean] = str(r[0])          # « Allemagne » est écrit « Allemagne␠ » dans le modèle : Excel exige le nom exact
            if len(r) > 2 and r[2] not in (None, ""):
                codes.append(_ngp(r[2]))
    def read_ngp_sheet(name):
        for row in wb[name].iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                n, ngp = str(row[0]).strip(), _ngp(row[1])
                ngpn.setdefault(n.casefold(), {"designation": n, "ngp": ngp, "rubrique": str(row[2] or n).strip()})
                ngp_to_name.setdefault(ngp, n)
    read_ngp_sheet("NGPN")
    kind = "export"
    if not ngpn:                                              # modèle IMPORT : NGPN est vide, chaque destinataire a sa feuille « 6 », « 7 »…
        kind = "import"
        for code in dict.fromkeys(codes):
            if code in wb.sheetnames:
                read_ngp_sheet(code)
    if kind == "import":                                      # les familles de produits sont proposées comme désignations d'article
        for ngp, fam in IMPORT_FAMILLES.items():
            ngpn.setdefault(fam.casefold(), {"designation": fam, "ngp": ngp, "rubrique": fam, "famille": True})
    for extra, like in EXTRA_ARTICLES:                       # désignations ajoutées au site (même NGP qu'une ligne du modèle)
        src = ngpn.get(like.casefold())
        if src and extra.casefold() not in ngpn:
            ngpn[extra.casefold()] = {"designation": extra, "ngp": src["ngp"], "rubrique": extra, "extra": True}
    wb.close()
    return Reference(base, ngpn, ngp_to_name, destinations, dest_raw, kind)


def classify(ref: str, reference: Reference) -> dict | None:
    """Article de DUM d'une référence : {'ngp', 'name', 'rubrique'} ou None si introuvable."""
    info = reference.base.get(norm_ref(ref))
    if not info:
        return None
    j = info["designation"]
    if "conducteur" in j.casefold():                       # « CONDUCTEUR ( NB ) vérifie » -> « Conducteurs électriques »
        j = "Conducteurs électriques"
    row = reference.ngpn.get(j.casefold())
    if row:
        return {"ngp": row["ngp"], "name": row["designation"], "rubrique": row["rubrique"]}
    if info["ngp"]:
        return {"ngp": info["ngp"], "name": j or info["annexe"], "rubrique": info["annexe"] or j}
    return None


CONDUCTOR_RE = re.compile(r"(?:^|[\s(])H_[A-Z]|HARNAIS|FAISCEAU|C[ÂA]BL(?:AGE|E)\b|CONDUCTEUR", re.I)
CABINET_RE = re.compile(r"ARMOIRE|CABINET|COFFRET|PUPITRE|\bBAIE\b", re.I)


def guess_from_description(desc: str, reference: Reference) -> dict | None:
    """Suggestion (à vérifier) pour une référence absente de BASE DONNER, d'après sa description."""
    text = str(desc or "")
    for rx, name in ((CABINET_RE, "armoires électriques"), (CONDUCTOR_RE, "conducteurs électriques")):
        if rx.search(text) and name in reference.ngpn:
            row = reference.ngpn[name]
            return {"ngp": row["ngp"], "name": row["designation"], "rubrique": row["rubrique"]}
    return None


@dataclass
class Article:
    ngp: str
    designation: str                  # libellé DUM (rubrique annexe)
    template_name: str                # désignation à poser en A2 de la feuille ART (pour le VLOOKUP du NGP)
    lines: list[dict] = field(default_factory=list)
    origin: str = ""                  # code pays d'origine (import) ; vide = MA (export)

    @property
    def suggested(self) -> int: return sum(1 for l in self.lines if l.get("suggested"))
    @property
    def qty(self) -> float: return float(sum(l["qty"] for l in self.lines))
    @property
    def net(self) -> float: return float(sum(l["net"] for l in self.lines))
    @property
    def value(self) -> float: return float(sum(l["value"] for l in self.lines))


@dataclass
class Ventilation:
    articles: list[Article]
    totals: dict
    missing: list[dict]
    notes: list[str]

    def summary_df(self, rate: float = 0.0, coef: float = 1.0) -> pd.DataFrame:
        return pd.DataFrame([{
            "N°": i, "Désignation": a.designation, "NGP": a.ngp or "—", "Lignes": len(a.lines),
            "Quantité": round(a.qty, 3), "Poids net (kg)": round(a.net, 3), "Valeur devise": round(a.value, 2),
            "Valeur DH": int(-(-a.value * rate * coef // 1)) if rate else None, "Origine": a.origin or ORIGIN_CODE,
            "Source": ("BASE DONNER" if not a.suggested else f"BASE DONNER + {a.suggested} ligne(s) suggérée(s)") if a.ngp else "À classer",
        } for i, a in enumerate(self.articles, start=1)])


def _apply_hu_weights(df: pd.DataFrame) -> list[str]:
    """Poids net lus dans la packing list (colonnes `_hu` / `_hu_net`, voir pl_weights) : une référence seule dans son unité de manutention
    garde son poids exact ; les références d'une même unité se partagent le poids de l'unité (au prorata de la valeur) ; seules les lignes
    sans unité connue restent estimées sur le reste du poids net du BL. Modifie df["_net"] ; renvoie les notes."""
    if "_hu" not in df.columns or "_hu_net" not in df.columns:
        return []
    known = df["_hu"].notna() & df["_hu_net"].notna()
    if not known.any():
        return []
    n_single = n_group = n_est = 0
    warn: list[str] = []
    for bl, idx in df.groupby("_bl", sort=False).groups.items():
        idx = list(idx)
        total = float(df.loc[idx, "_netbl"].sum(skipna=True))
        kidx = [i for i in idx if known[i]]
        if not kidx:
            continue
        groups: dict = {}
        for i in kidx:
            groups.setdefault(str(df.at[i, "_hu"]), []).append(i)
        hu_total = sum(float(df.at[g[0], "_hu_net"]) for g in groups.values())
        unknown = [i for i in idx if not known[i]]
        scale = 1.0
        if not unknown and total > 0 and abs(hu_total - total) > 0.005 * total:
            scale = total / hu_total if hu_total > 0 else 1.0           # le total du BL fait foi
        for g in groups.values():
            gnet = float(df.at[g[0], "_hu_net"]) * scale
            for i, v in zip(g, cleaning._allocate(gnet, df.loc[g, "_val"].abs().tolist() or [1.0])):
                df.at[i, "_net"] = v
            n_single, n_group = (n_single + 1, n_group) if len(g) == 1 else (n_single, n_group + len(g))
        if unknown:
            rest = total - hu_total
            if rest < 0:
                warn.append(f"⚠ BL {bl} : les poids des unités de manutention ({hu_total:g} kg) dépassent le poids net du BL ({total:g} kg).")
                rest = 0.0
            for i, v in zip(unknown, cleaning._allocate(rest, df.loc[unknown, "_val"].abs().tolist() or [1.0])):
                df.at[i, "_net"] = v
            n_est += len(unknown)
    notes = [f"Poids net lus dans la packing list (unités de manutention) : {n_single} référence(s) seule(s) dans son unité (poids exact) "
             f"+ {n_group} référence(s) groupées (poids de l'unité partagé au prorata de la valeur)"
             + (f" ; {n_est} ligne(s) estimée(s) sur le reste du BL." if n_est else ".")]
    return notes + warn


def build_ventilation(bl_df: pd.DataFrame, reference: Reference, guess: bool = True, manual: dict | None = None) -> Ventilation:
    """Regroupe les lignes d'export en articles de DUM. `bl_df` : tableau BL References (colonnes standard de l'export)."""
    num = lambda s: pd.to_numeric(s, errors="coerce")  # noqa: E731
    df = bl_df.copy().reset_index(drop=True)
    df["_bl"] = df["BL Number (Bon de Livraison)"].astype(str)
    inv_col = df["Invoice Number (Facture)"] if "Invoice Number (Facture)" in df else pd.Series([""] * len(df))
    df["_inv"] = inv_col.map(lambda v: str(v).strip() if str(v).strip() not in ("", "nan", "None") else "(sans facture)")
    df["_val"] = num(df["Prix unitaire total (Hors-taxe)"]).fillna(0.0)
    df["_qty"] = num(df["Quantity"]).fillna(0.0)
    df["_netbl"] = num(df["Total Net Weight of the BL"])
    # poids net du BL réparti sur ses lignes au prorata de la valeur (somme exacte)
    df["_net"] = 0.0
    for _bl, idx in df.groupby("_bl", sort=False).groups.items():
        idx = list(idx)
        total = float(df.loc[idx, "_netbl"].sum(skipna=True))
        if total > 0:
            for i, v in zip(idx, cleaning._allocate(total, df.loc[idx, "_val"].abs().tolist() or [1.0])):
                df.at[i, "_net"] = v
    hu_note = _apply_hu_weights(df)
    if manual:                                                    # classement choisi par l'utilisateur : article -> désignation, facture -> article
        vent = _build_manual(df, reference, manual)
        vent.notes = hu_note + vent.notes
        return vent
    articles: dict[str, Article] = {}
    missing: dict[str, dict] = {}
    for _, r in df.iterrows():
        ref = str(r["Reference"]).strip()
        art = classify(ref, reference)
        suggested = False
        if not art and guess:
            art = guess_from_description(r.get("Description", ""), reference)
            suggested = art is not None
        key = art["ngp"] if art else "?"
        if key not in articles:
            articles[key] = Article(art["ngp"], art["rubrique"], art["name"]) if art else Article("", UNCLASSIFIED, UNCLASSIFIED)
        articles[key].lines.append({"ref": ref, "desc": str(r.get("Description", "") or ""), "qty": float(r["_qty"]), "inv": r["_inv"],
                                    "net": float(r["_net"]), "value": float(r["_val"]), "bl": r["_bl"],
                                    "suggested": suggested})
        if not art or suggested:
            m = missing.setdefault(ref, {"Référence": ref, "Description": str(r.get("Description", "") or "")[:60], "Quantité": 0.0,
                                         "Valeur": 0.0, "Classement": (art["name"] + " (suggéré)") if suggested else "À classer"})
            m["Quantité"] += float(r["_qty"])
            m["Valeur"] += float(r["_val"])
    ordered = sorted(articles.values(), key=lambda a: (a.ngp == "", -a.value))
    totals = {"colis": float(num(df["Number of Packages (Nombre de Colis)"]).sum()),
              "brut": float(num(df["Total Gross Weight of the BL"]).sum()), "net": float(df["_netbl"].sum(skipna=True)),
              "valeur": float(df["_val"].sum()), "quantite": float(df["_qty"].sum()), "lignes": len(df)}
    notes = []
    if missing:
        n_sug = sum(1 for m in missing.values() if m["Classement"] != "À classer")
        notes.append(f"{len(missing)} référence(s) absente(s) de BASE DONNER"
                     + (f" ({n_sug} classée(s) par suggestion d'après la description : à vérifier)" if n_sug else "")
                     + " — ajoutez-les à la base du modèle pour un classement sûr.")
    if len(ordered) > MAX_ARTICLES:
        notes.append(f"Plus de {MAX_ARTICLES} articles : le modèle n'en accepte que {MAX_ARTICLES}.")
    cur = df["_currency"].dropna().astype(str).unique().tolist() if "_currency" in df else []
    if len([c for c in cur if c]) > 1:
        notes.append(f"Plusieurs devises dans le dossier ({', '.join(c for c in cur if c)}) : la ventilation additionne les montants tels quels.")
    return Ventilation(ordered, totals, list(missing.values()), hu_note + notes)


def _totals(df: pd.DataFrame) -> dict:
    num = lambda s: pd.to_numeric(s, errors="coerce")  # noqa: E731
    return {"colis": float(num(df["Number of Packages (Nombre de Colis)"]).sum()),
            "brut": float(num(df["Total Gross Weight of the BL"]).sum()), "net": float(df["_netbl"].sum(skipna=True)),
            "valeur": float(df["_val"].sum()), "quantite": float(df["_qty"].sum()), "lignes": len(df)}


def _build_manual(df: pd.DataFrame, reference: Reference, manual: dict, totals: dict | None = None) -> Ventilation:
    """manual = {'articles': [désignation NGPN article 1, 2, …], 'assign': {n° facture: n° d'article (1..n)}}"""
    names = manual["articles"]
    arts: list[Article] = []
    for n in names:
        row = reference.ngpn.get(str(n).casefold())
        arts.append(Article(row["ngp"], row["rubrique"], row["designation"]) if row else Article("", str(n), str(n)))
    assign = manual.get("assign", {})                              # n° facture -> article
    assign_ref = manual.get("assign_ref")                          # référence -> article (prioritaire : choix par référence)
    notes: list[str] = []
    for _, r in df.iterrows():
        if assign_ref is not None:
            k = int(assign_ref.get(norm_ref(r["Reference"]), 1)) - 1
        else:
            k = int(assign.get(r["_inv"], 1)) - 1
        k = min(max(k, 0), len(arts) - 1)
        arts[k].lines.append({"ref": str(r["Reference"]).strip(), "desc": str(r.get("Description", "") or ""), "qty": float(r["_qty"]),
                              "inv": r["_inv"], "net": float(r["_net"]), "value": float(r["_val"]), "bl": r["_bl"], "suggested": False,
                              "origin": str(r.get("_origin", "") or "")})
    for a in arts:
        origs = sorted({l["origin"] for l in a.lines if l.get("origin")})
        if not origs:
            continue
        if len(origs) == 1:
            a.origin = origs[0]
        else:
            from ventilation_import import DIVERS, EU                  # import module : plusieurs pays UE = « DIVERS »
            if all(o in EU for o in origs):
                a.origin = DIVERS
            else:
                top: dict[str, float] = {}
                for l in a.lines:
                    top[l["origin"]] = top.get(l["origin"], 0.0) + abs(l["value"])
                a.origin = max((o for o in origs), key=lambda o: top.get(o, 0.0))
                notes.append(f"L'article « {a.template_name} » mélange des pays hors UE ({', '.join(origs)}) : le modèle n'en retient qu'un "
                             f"({a.origin}) — séparez-les en articles distincts.")
    empty = [i + 1 for i, a in enumerate(arts) if not a.lines]
    if empty:
        notes.append(f"Article(s) {', '.join(map(str, empty))} sans aucune facture : ils seront laissés vides.")
    kept = [a for a in arts if a.lines]
    return Ventilation(kept, totals or _totals(df), [], notes)


def apply_ref_net(vent: Ventilation, picks: set, kg_per_ref: float) -> list[str]:
    """Poids connu de certaines lignes de référence (la packing list ne donne que le total de plusieurs pièces) : chaque ligne choisie
    (identifiée par (n° d'article, n° de ligne), 0-based) reçoit `kg_per_ref` kg ; le reste du poids net du dossier est réparti
    sur toutes les autres lignes (tous articles). Les références répétées ne sont jamais fusionnées."""
    if kg_per_ref <= 0 or not picks:
        return []
    chosen = [l for ai, a in enumerate(vent.articles) for li, l in enumerate(a.lines) if (ai, li) in picks]
    others = [l for ai, a in enumerate(vent.articles) for li, l in enumerate(a.lines) if (ai, li) not in picks]
    if not chosen:
        return []
    total = float(vent.totals.get("net", 0.0))
    fixed = kg_per_ref * len(chosen)
    for l in chosen:
        l["net"] = kg_per_ref
    notes = [f"Poids net imposé : {kg_per_ref:g} kg × {len(chosen)} ligne(s) de référence choisie(s) = {fixed:g} kg."]
    rest = total - fixed
    if rest < -0.0005:
        notes.append(f"⚠ Le poids imposé ({fixed:g} kg) dépasse le poids net du dossier ({total:g} kg) : vérifiez le poids saisi.")
        rest = 0.0
    if others:
        base = [abs(l["net"]) for l in others]
        if not sum(base):
            base = [abs(l["value"]) for l in others]
        if not sum(base):
            base = [1.0] * len(others)
        for l, x in zip(others, base):
            l["net"] = rest * x / sum(base)
        notes.append(f"Poids net restant ({rest:g} kg) réparti sur les {len(others)} autre(s) ligne(s) (au prorata de leur poids calculé).")
    else:
        notes.append(f"⚠ Toutes les lignes ont un poids imposé : {fixed:g} kg pour un poids net de dossier de {total:g} kg.")
    return notes


def invoice_table(vent: Ventilation) -> pd.DataFrame:
    """Une ligne par facture avec l'article proposé (celui qui porte la plus grande valeur de cette facture)."""
    rows: dict[str, dict] = {}
    for idx, art in enumerate(vent.articles, start=1):
        for l in art.lines:
            r = rows.setdefault(l["inv"], {"Facture": l["inv"], "BL": set(), "Lignes": 0, "Quantité": 0.0, "Valeur": 0.0, "_best": {}})
            r["BL"].add(l["bl"])
            r["Lignes"] += 1
            r["Quantité"] += l["qty"]
            r["Valeur"] += l["value"]
            r["_best"][idx] = r["_best"].get(idx, 0.0) + l["value"]
    out = []
    for r in rows.values():
        best = max(r["_best"], key=r["_best"].get) if r["_best"] else 1
        out.append({"Facture": r["Facture"], "BL": ", ".join(sorted(r["BL"])), "Lignes": r["Lignes"],
                    "Quantité": round(r["Quantité"], 3), "Valeur": round(r["Valeur"], 2), "Article": best})
    return pd.DataFrame(out)


def reference_table(vent: Ventilation) -> pd.DataFrame:
    """Une ligne par référence distincte avec l'article proposé (celui qui porte la plus grande valeur de cette référence)."""
    rows: dict[str, dict] = {}
    for idx, art in enumerate(vent.articles, start=1):
        for l in art.lines:
            key = norm_ref(l["ref"])
            r = rows.setdefault(key, {"Référence": l["ref"], "Description": "", "Factures": set(), "Lignes": 0,
                                      "Quantité": 0.0, "Valeur": 0.0, "_best": {}})
            if not r["Description"] and l["desc"] and not l["desc"].startswith("Société"):
                r["Description"] = l["desc"][:60]
            r["Factures"].add(l["inv"])
            r["Lignes"] += 1
            r["Quantité"] += l["qty"]
            r["Valeur"] += l["value"]
            r["_best"][idx] = r["_best"].get(idx, 0.0) + l["value"]
    out = []
    for r in rows.values():
        best = max(r["_best"], key=r["_best"].get) if r["_best"] else 1
        out.append({"Référence": r["Référence"], "Description": r["Description"], "Factures": len(r["Factures"]),
                    "Lignes": r["Lignes"], "Quantité": round(r["Quantité"], 3), "Valeur": round(r["Valeur"], 2), "Article": best})
    return pd.DataFrame(out)


def fill_template(template_bytes: bytes, vent: Ventilation, destinataire: str, rate: float, fret: float | None = None,
                  matricule: str = "") -> bytes:
    """Remplit une copie du modèle (formules conservées : VAL DHS, ANNEXE, N° DUM… se calculent à l'ouverture dans Excel)."""
    wb = load_workbook(io.BytesIO(template_bytes), keep_links=False)       # keep_links=False : supprime le lien externe du taux
    v = wb["VENTILATION"]
    t = vent.totals
    v["C2"], v["C3"], v["C4"] = t["colis"], t["brut"], t["net"]
    v["C5"], v["C6"] = round(t["valeur"], 2), t["quantite"]
    v["I5"] = destinataire
    v["G8"] = float(rate)                                                    # remplace =[1]Feuil1!$B$2
    if matricule:
        v["I7"] = matricule                                                  # REMORQUE (matricule) : repris par la colonne MARQUE de l'ANNEXE
    if fret is not None:
        v["L8"] = float(fret)
    desc = wb["DESCIPTION"]
    row_desc = 3
    for n, art in enumerate(vent.articles[:MAX_ARTICLES], start=1):
        ws = wb[f"ART{n}"]
        ws["A2"] = art.template_name
        if art.template_name in {e for e, _ in EXTRA_ARTICLES} and art.ngp:
            ws["B2"] = int(art.ngp)                                          # désignation absente de NGPN : le NGP est écrit directement
        ws["C2"] = "NB"
        ws["B4"] = "UE"
        for k, line in enumerate(art.lines[:LAST_LINE_ROW - FIRST_LINE_ROW + 1]):
            r = FIRST_LINE_ROW + k
            ws.cell(r, 4, line["ref"])
            ws.cell(r, 6, ORIGIN_NAME)
            ws.cell(r, 7, round(line["net"], 3))
            ws.cell(r, 8, line["qty"])
            ws.cell(r, 9, round(line["value"] / line["qty"], 4) if line["qty"] else 0)
            ws.cell(r, 10, round(line["value"], 2))
            desc.cell(row_desc, 2, line["ref"])
            desc.cell(row_desc, 4, ORIGIN_NAME)
            desc.cell(row_desc, 5, round(line["net"], 3))
            desc.cell(row_desc, 6, line["qty"])
            desc.cell(row_desc, 7, round(line["value"] / line["qty"], 4) if line["qty"] else 0)
            desc.cell(row_desc, 8, round(line["value"], 2))
            row_desc += 1
        v.cell(12 + n, 4, ORIGIN_CODE)                                       # ORGN : les exportations Alstom sont d'origine MAROC
        v.cell(12 + n, 11, "UE")                                             # Code accord (les formules du modèle ne le reprennent que pour l'article 1)
    wb.template = False
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


ANNEXE_COLUMNS = ["ID ART", "ORDR ART", "NBR COLIS", "TYPE", "MARQUE", "NGP", "DESIGNATION", "PAYS ORIGINE", "PAIEMENT", "OCCASION",
                  "VALEUR", "DEVISE", "QTE ART", "UN MESURE", "POIDS NET", "QTE NORMALISE", "ACCORD", "FRANCHISE", "NOM ET PRENOM", "CIN"]


def annexe_rows(vent: Ventilation, rate: float, matricule: str = "") -> list[list]:
    """Lignes de la feuille ANNEXE (mêmes règles que les formules du modèle), uniquement les articles utilisés."""
    import math
    rows = []
    for i, a in enumerate(vent.articles[:MAX_ARTICLES], start=1):
        value_mad = math.ceil(round(a.value * rate, 6))                      # ROUNDUP(valeur devise × taux, 0)
        rows.append([None, i, vent.totals["colis"] if i == 1 else None, 216 if i == 1 else "000", matricule if i == 1 else None,
                     int(a.ngp) if a.ngp else None, a.template_name, ORIGIN_CODE, "AP", "NON", value_mad, "MAD",
                     a.qty, "002", round(a.net, 3), a.qty, "UE", None, None, None])
    return rows


def build_annexe_xlsx(vent: Ventilation, rate: float, matricule: str = "", template_bytes: bytes | None = None,
                      rows: list[list] | None = None) -> bytes:
    """Fichier Excel ne contenant que la feuille ANNEXE, strictement identique à celle du modèle (mêmes polices, couleurs, largeurs),
    avec les valeurs à la place des formules et sans les lignes vides."""
    if template_bytes is None:
        template_bytes = get_template_bytes()
    if template_bytes is None:
        return _plain_annexe_xlsx(vent, rate, matricule, rows)
    wb = load_workbook(io.BytesIO(template_bytes), keep_links=False)
    ws = wb["ANNEXE "] if "ANNEXE " in wb.sheetnames else wb[[n for n in wb.sheetnames if n.strip().upper() == "ANNEXE"][0]]
    for name in list(wb.sheetnames):
        if wb[name] is not ws:
            wb.remove(wb[name])
    for dn in list(wb.defined_names.keys()):                       # les noms définis pointaient vers les feuilles supprimées
        del wb.defined_names[dn]
    ws.title = "ANNEXE"
    rows = rows if rows is not None else annexe_rows(vent, rate, matricule)
    last_col = len(ANNEXE_COLUMNS)
    for r in range(2, max(ws.max_row, len(rows) + 1) + 1):         # efface les formules du modèle
        for c in range(1, last_col + 1):
            ws.cell(r, c).value = None
    for i, row in enumerate(rows, start=2):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j).value = v
    if ws.max_row > len(rows) + 1:                                 # supprime les lignes vides (« 0 ») du modèle
        ws.delete_rows(len(rows) + 2, ws.max_row - len(rows) - 1)
    wb.template = False
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _plain_annexe_xlsx(vent: Ventilation, rate: float, matricule: str = "", rows: list[list] | None = None) -> bytes:
    """Fichier Excel ne contenant que la feuille ANNEXE (valeurs, sans lignes vides)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "ANNEXE"
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    tan, green = PatternFill("solid", start_color="D9C8A0", end_color="D9C8A0"), PatternFill("solid", start_color="92D050", end_color="92D050")
    for j, name in enumerate(ANNEXE_COLUMNS, start=1):
        c = ws.cell(1, j, name)
        c.font, c.border = Font(name="Times New Roman", size=14, bold=True), border
        c.fill = green if name in ("QTE ART", "QTE NORMALISE") else tan
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 38
    for r, row in enumerate(rows if rows is not None else annexe_rows(vent, rate, matricule), start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(r, j, v)
            c.font, c.border = Font(name="Times New Roman", size=16), border
            c.alignment = Alignment(horizontal="center", vertical="center")
            if ANNEXE_COLUMNS[j - 1] in ("VALEUR", "QTE ART", "QTE NORMALISE"):
                c.number_format = "#,##0.###"
            elif ANNEXE_COLUMNS[j - 1] == "POIDS NET":
                c.number_format = "#,##0.000"
        ws.row_dimensions[r].height = 26
    for j, w in enumerate([12, 16, 16, 10, 24, 20, 40, 16, 14, 14, 18, 12, 14, 16, 16, 18, 12, 16, 24, 10], start=1):
        ws.column_dimensions[ws.cell(1, j).column_letter].width = w
    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
