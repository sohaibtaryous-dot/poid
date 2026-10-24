# -*- coding: utf-8 -*-
"""Section « Ventilation DUM » de l'IMPORT (copie indépendante de celle de l'EXPORT : modifier l'un ne touche pas l'autre)."""
from __future__ import annotations

import io
import re

import pandas as pd
import streamlit as st

import app as core
import ventilation_export as vx
import ventilation_import as vi

def _vx_reference(template_bytes: bytes):
    import streamlit as st
    import ventilation_export as vx

    @st.cache_resource(show_spinner="Lecture du modèle de ventilation…")
    def _load(data: bytes):
        return vx.load_reference(data)
    return _load(template_bytes)


def _render_ventilation_section(bl_final: pd.DataFrame, kind: str = "export"):
    """Ventilation DUM : remplit automatiquement le modèle VENTI_EXPORT / VENTI_IMPORT (articles, NGP, poids, valeurs).
    kind='export' : depuis les BL / factures de l'export ; kind='import' : depuis le tableau TOUS_CLIENTS (avec ACC, fret, taux)."""
    import streamlit as st
    import database
    import ventilation_export as vx
    import ventilation_import as vi
    import ui
    imp = kind == "import"
    P = "vi" if imp else "vx"                                       # préfixe des clés de session : les deux modules ne se mélangent pas
    NAME = "VENTI_IMPORT" if imp else "VENTI_EXPORT"
    build_vent = vi.build if imp else vx.build_ventilation
    ui.section(f"Ventilation DUM — modèle {NAME}", "🧮")
    stored = vx.get_template_bytes(kind)
    with st.expander(f"Modèle Excel de ventilation ({NAME})", expanded=stored is None):
        up = st.file_uploader(f"Déposez ici votre fichier {NAME} (.xltx ou .xlsx)", type=["xltx", "xlsx"], key=f"{P}_tpl_session")
        if up is not None:
            st.session_state[f"{P}_tpl_bytes"] = up.getvalue()
            st.session_state[f"{P}_tpl_name"] = up.name
        if st.session_state.get(f"{P}_tpl_bytes"):
            st.success(f"Modèle utilisé : {st.session_state.get(f'{P}_tpl_name', 'fichier déposé')} (valable pour cette session).")
        elif stored:
            st.caption("Modèle utilisé : celui enregistré sur le site. Déposez un fichier ici pour utiliser le vôtre à la place.")
    tpl = st.session_state.get(f"{P}_tpl_bytes") or stored
    if not tpl:
        st.info(f"Déposez votre fichier {NAME} ci-dessus pour lancer la ventilation.")
        return
    try:
        ref = _vx_reference(tpl)
        if imp and ref.kind != "import":
            raise ValueError("c'est le modèle EXPORT (NGPN rempli) ; déposez VENTI_IMPORT")
    except Exception as exc:  # noqa: BLE001
        st.error(f"Ce fichier n'est pas un modèle {NAME} valide (feuilles BASE DONNER / ATTELIER / feuille du destinataire attendues) : {exc}")
        return
    try:
        default_rate = float(database.get_setting("EUR_MAD_RATE") or 10.9272)
    except ValueError:
        default_rate = 10.9272
    c1, c2, c3, *c4 = st.columns([2, 1, 1, 1] if imp else [2, 1, 1])
    matricule = st.text_input("Matricule de la remorque / du camion", key=f"{P}_mat", placeholder="Ex. 12345-A-6",
                              help="Écrit dans VENTILATION (REMORQUE) et dans la colonne MARQUE de l'ANNEXE.").strip()
    names = list(ref.destinations)
    dest = c1.selectbox("Destinataire (donne le N° de DUM)", names, index=0 if (imp and len(names) == 1) else None,
                        placeholder="Choisir le destinataire…", key=f"{P}_dest",
                        format_func=lambda n: f"{n} — {ref.destinations[n]}")
    rate = c2.number_input("Taux de change (DH pour 1 devise)", min_value=0.0, value=default_rate, step=0.0001, format="%.4f", key=f"{P}_rate")
    fret = c3.number_input("Fret (devise)", min_value=0.0, value=0.0, step=1.0, key=f"{P}_fret")
    fret_note = ""
    if imp:                                                          # le fret peut être saisi en dirhams ou dans la devise du dossier
        _cur = sorted({str(c) for c in (bl_final["_devise"] if "_devise" in bl_final else []) if str(c) not in ("", "nan", "None")})
        dossier_cur = _cur[0] if _cur else "EUR"
        fret_cur = c3.selectbox("Devise du fret", ["MAD (dirhams)", f"{dossier_cur} (devise du dossier)"], index=1, key=f"{P}_fretcur")
        fret_in = fret
        if fret_cur.startswith("MAD") and rate:
            fret = fret_in / rate                                    # le modèle attend le fret en devise (L8) : C7 = taux × L8
        fret_note = f"Fret : {fret_in:,.2f} {'MAD' if fret_cur.startswith('MAD') else dossier_cur} = {fret * rate:,.2f} DH (≈ {fret:,.2f} {dossier_cur})".replace(",", " ")
    acc = c4[0].number_input("ACC — acconage (DH)", min_value=0.0, value=0.0, step=10.0, key=f"{P}_acc",
                             help="Frais d'acconage en dirhams (case ACC du modèle). Entre dans la valeur taxable VTD.") if imp else 0.0
    if imp:
        mode = st.radio("Classement des articles", ["Automatique (un article par NGP, désignation = famille de produits)",
                                                    "Manuel (je choisis les articles et leurs factures)"], horizontal=True, key=f"{P}_mode", index=0)
        o1, o2 = st.columns(2)
        split = o1.checkbox("Séparer aussi les pays de l'UE entre eux", value=False, key=f"{P}_split",
                            help="Décoché (conseillé) : les pays de l'UE d'un même NGP sont regroupés dans un article d'origine « DIVERS » ; "
                                 "les pays hors UE (Suisse, Royaume-Uni, USA, Chine, Japon…) restent toujours des articles séparés.")
        import ventilation_import as _vi
        ratio = o2.number_input("Ratio net / brut estimé (%)", min_value=50.0, max_value=100.0, step=1.0, key=f"{P}_ratio",
                                value=round(_vi.default_ratio(bl_final) * 100, 1),
                                help="Utilisé seulement quand une facture n'a pas de poids net : net = brut × ratio, réparti sur ses lignes au prorata de la quantité.") / 100.0
        _b = vi.build
        build_vent = lambda df, rf, **kw: _b(df, rf, split_origin=split, ratio=ratio, **kw)  # noqa: E731
    else:
        mode = st.radio("Classement des articles", ["Automatique (par référence, via BASE DONNER)", "Manuel (je choisis les articles et leurs factures)"],
                        horizontal=True, key=f"{P}_mode", index=1)
    manual = None
    guess = True
    if mode.startswith("Automatique"):
        guess = st.checkbox("Classer les références inconnues d'après leur description (suggestion : à vérifier)", value=True, key=f"{P}_guess")
        vent = build_vent(bl_final, ref, guess=guess)
    else:
        auto = build_vent(bl_final, ref, guess=True)
        if not any(a.lines for a in auto.articles):
            st.warning("Aucune ligne de marchandise à ventiler dans ce dossier (tableau vide ou lignes sans montant).")
            for n in auto.notes:
                st.caption(n)
            return
        inv_df = vx.invoice_table(auto)
        ref_df = vx.reference_table(auto)
        names = sorted({row["designation"] for row in ref.ngpn.values()}, key=str.casefold)
        defaults = [a.template_name for a in auto.articles if a.ngp]
        n_art = int(st.number_input("Nombre d'articles de la DUM", min_value=1, max_value=vx.MAX_ARTICLES,
                                    value=1, step=1, key=f"{P}_nart"))
        by = st.radio("Pour séparer les articles, je choisis…", ["les factures (numéros de facture)", "les références (détectées dans les factures)"],
                      horizontal=True, key=f"{P}_by2")
        by_inv = by.startswith("les factures")
        if by_inv:
            keys = inv_df["Facture"].tolist()
            labels = {r["Facture"]: f"{r['Facture']} — {int(r['Lignes'])} ligne(s) · qté {r['Quantité']:g} · {r['Valeur']:,.2f}".replace(",", " ")
                      for _, r in inv_df.iterrows()}
            noun, tag = "factures", "inv"
        else:
            keys = ref_df["Référence"].tolist()
            labels = {r["Référence"]: f"{r['Référence']} — {(r['Description'] or '—')[:40]} · qté {r['Quantité']:g} · {r['Valeur']:,.2f}".replace(",", " ")
                      for _, r in ref_df.iterrows()}
            noun, tag = "références", "ref"
        st.caption(f"{len(keys)} {noun} détectée(s). Tout ce que vous ne cochez pas dans les articles 2, 3… reste dans l'article 1.")
        chosen, assigned = [], {}
        for i in range(n_art):
            pool = defaults + ["Armoires électriques", "Conducteurs électriques"] + names
            dflt = next((p for p in pool if p in names and p not in chosen), names[0])       # chaque article démarre avec une désignation différente
            with st.container(border=True):
                pick = st.selectbox(f"Article {i + 1} — désignation", names, index=names.index(dflt) if dflt in names else 0,
                                    format_func=lambda n: f"{n} — NGP {ref.ngpn[n.casefold()]['ngp']}", key=f"{P}_art_{i}")
                chosen.append(pick)
                taken = set(assigned)
                options = [k for k in keys if k not in taken]
                skey = f"{P}_sel_{tag}_{i}"
                st.session_state[skey] = [v for v in st.session_state.get(skey, []) if v in options]   # retire ce qu'un article précédent a pris
                label = f"{noun.capitalize()} de l'article {i + 1}" + (" (facultatif : le reste y va)" if i == 0 else "")
                sel = st.multiselect(label, options, format_func=lambda k: labels[k], key=skey, placeholder="Choisir…")
                for k in sel:
                    assigned[k] = i + 1
        rest = len(keys) - len(assigned)
        st.caption(f"Répartition : {', '.join(f'article {i + 1} = ' + str(sum(1 for v in assigned.values() if v == i + 1) + (rest if i == 0 else 0)) for i in range(n_art))} {noun}.")
        if by_inv:
            manual = {"articles": chosen, "assign": {k: assigned.get(k, 1) for k in keys}}
        else:
            manual = {"articles": chosen, "assign_ref": {vx.norm_ref(k): assigned.get(k, 1) for k in keys}}
        vent = build_vent(bl_final, ref, manual=manual)
    cost = vi.costs(vent, rate, fret, acc) if imp else None
    st.dataframe(vent.summary_df(rate, cost["coef"] if cost else 1.0), use_container_width=True, hide_index=True,
                 column_config={"Poids net (kg)": st.column_config.NumberColumn(format="%.3f"),
                                "Valeur devise": st.column_config.NumberColumn(format="%.2f")})
    t = vent.totals
    st.caption(f"Totaux du dossier : {t['colis']:g} colis · {t['brut']:.3f} kg brut · {t['net']:.3f} kg net · "
               f"{t['quantite']:g} pièces · valeur {t['valeur']:,.2f}".replace(",", " "))
    if fret_note:
        st.caption(fret_note)
    if cost:
        st.caption(f"Valeur DH {cost['valeur_dh']:,.2f} + fret {cost['fret_dh']:,.2f} + assurance (0,5 %) {cost['ass']:,.2f} + ACC {cost['acc']:,.2f} "
                   f"= VTD {cost['vtd']:,.2f} DH · coefficient {cost['coef']:.4f}".replace(",", " "))
    for n in vent.notes:
        st.warning(n)
    if vent.missing:
        with st.expander(f"Références à vérifier ({len(vent.missing)})"):
            st.dataframe(pd.DataFrame(vent.missing), use_container_width=True, hide_index=True)
    sig = (dest, round(rate, 4), round(fret, 2), round(acc, 2), guess, (split, round(ratio, 4)) if imp else None, len(bl_final), round(t["valeur"], 2), round(t["net"], 3),
           repr(manual) if manual else "auto", matricule)
    suffix = "IMPORT" if imp else "DUM"
    if st.button(f"Générer VENTILATION_{suffix}.xlsx + ANNEXE_{suffix}.xlsx", type="primary", icon=":material/calculate:", key=f"{P}_go", disabled=not dest):
        with st.spinner("Remplissage du modèle et de l'annexe (≈ 1 min)…"):
            if imp:
                data = vi.fill_template(tpl, vent, ref.dest_raw.get(dest, dest), rate, fret, acc, matricule)
                annexe = vi.build_annexe(vent, rate, fret, acc, matricule, tpl)
            else:
                data = vx.fill_template(tpl, vent, ref.dest_raw.get(dest, dest), rate, fret or None, matricule)
                annexe = vx.build_annexe_xlsx(vent, rate, matricule, tpl)
        st.session_state[f"{kind[:3]}_vent"] = {"sig": sig, "bytes": data, "annexe": annexe, "dest": dest}
        try:
            database.set_setting("EUR_MAD_RATE", f"{rate:.4f}")
        except Exception:  # noqa: BLE001
            pass
    cur = st.session_state.get(f"{kind[:3]}_vent")
    if cur and cur["sig"] == sig:
        if not imp and "exp_files" in st.session_state:
            st.session_state["exp_files"][f"VENTILATION_{suffix}.xlsx"] = cur["bytes"]      # inclus dans l'attribution aux clients
            st.session_state["exp_files"][f"ANNEXE_{suffix}.xlsx"] = cur["annexe"]
        d1, d2 = st.columns(2)
        d1.download_button(f"VENTILATION_{suffix}.xlsx — {cur['dest']}", data=cur["bytes"], file_name=f"VENTILATION_{suffix}.xlsx",
                           icon=":material/download:", type="primary", key=f"dl_{P}", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button(f"ANNEXE_{suffix}.xlsx — feuille ANNEXE seule", data=cur["annexe"], file_name=f"ANNEXE_{suffix}.xlsx",
                           icon=":material/download:", type="primary", key=f"dl_annexe_{P}", use_container_width=True,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           help="Même présentation que la feuille ANNEXE du modèle, valeurs remplies, sans lignes vides")
        st.caption("Ouvrez-le dans Excel : les formules du modèle (valeur DH, ANNEXE, N° DUM) se calculent à l'ouverture.")
    elif cur:
        st.info("Les paramètres ont changé : régénérez le fichier.")


