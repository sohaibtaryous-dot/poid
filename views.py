# -*- coding: utf-8 -*-
"""Page IMPORT : téléversement multi-formats, contrôle des fichiers, tableau TOUS_CLIENTS (25 colonnes), alertes,
groupes de documents, synthèse par fournisseur et export Excel / CSV."""

from __future__ import annotations

import hashlib
import traceback
from datetime import datetime

import pandas as pd
import streamlit as st

import app as core
import fileio
import shipment
import ui

NGP_COLS = ("NGP (REF_PAR_FRS)", "Onglet REF_PAR_FRS", "Contrôle NGP")
SEV_CSS = {0: "background-color: rgba(220,38,38,.22)", 1: "background-color: rgba(245,158,11,.16)"}


def _size(n: int) -> str:
    return f"{n / 1e6:.1f} Mo" if n >= 1e6 else f"{max(n / 1e3, 0.1):.0f} Ko"


def files_overview(files) -> list[str]:
    """Affiche le contrôle de chaque fichier avant traitement ; renvoie les noms des fichiers inutilisables."""
    rows, bad = [], []
    for f in files:
        try:
            data = f.getvalue()
            kind = fileio.validate_upload(f.name, data)
            rows.append({"name": f.name, "type": fileio.KIND_LABEL.get(kind, kind), "size": _size(len(data)),
                         "ok": True, "message": ""})
        except fileio.FileReadError as exc:
            bad.append(f.name)
            rows.append({"name": f.name, "type": "?", "size": "", "ok": False, "message": str(exc)})
        except Exception as exc:  # fichier impossible à lire côté navigateur
            bad.append(f.name)
            rows.append({"name": f.name, "type": "?", "size": "", "ok": False, "message": f"lecture impossible ({type(exc).__name__})"})
    ui.file_list(rows)
    return bad


def clean_edited(edited: pd.DataFrame) -> pd.DataFrame:
    """Tableau modifié par l'utilisateur -> types propres (nombres, entiers, textes)."""
    edited = edited.copy()
    for c in core.EDITABLE_COLUMNS:
        if c not in edited.columns:
            edited[c] = None
    for c in core.NUMERIC_COLUMNS:
        if c in edited.columns:
            edited[c] = pd.to_numeric(edited[c], errors="coerce")
    for c in core.INT_COLUMNS:
        if c in edited.columns:
            edited[c] = pd.to_numeric(edited[c], errors="coerce").fillna(0).astype(int)
    for c in core.EDITABLE_COLUMNS:
        if c not in core.NUMERIC_COLUMNS and c not in core.INT_COLUMNS and c in edited.columns:
            edited[c] = edited[c].map(core.clean_cell)
    cols = [c for c in core.EDITABLE_COLUMNS if c in edited.columns] + [c for c in core.INTERNAL_COLUMNS if c in edited.columns]
    return edited[cols]


def _column_config(n_pages: int) -> dict:
    cfg = {"Qté": st.column_config.NumberColumn("Qté", format="%.3f"),
           "Prix unitaire": st.column_config.NumberColumn("Prix unitaire", format="%.4f"),
           "Valeur HT": st.column_config.NumberColumn("Valeur HT", format="%.2f"),
           core.C_NETU: st.column_config.NumberColumn(core.C_NETU, format="%.4f"),
           core.C_COLIS: st.column_config.NumberColumn(core.C_COLIS, format="%d"),
           "Page PDF": st.column_config.NumberColumn("Page PDF", format="%d"),
           core.PAGE_IN_PDF: st.column_config.NumberColumn(
               f"{core.PAGE_IN_PDF} ({n_pages} p.)" if n_pages else core.PAGE_IN_PDF, format="%d"),
           core.C_BASE: st.column_config.TextColumn(core.C_BASE, width="large")}
    for c in (core.C_NET, core.C_NETF, core.C_GROSSF, core.C_MANIF):
        cfg[c] = st.column_config.NumberColumn(c, format="%.3f")
    return cfg


def _style_view(view: pd.DataFrame, sev: pd.Series):
    """Surlignage ligne à ligne : rouge = critique, orange = à vérifier ; NGP comme avant."""
    fmt2 = lambda v: "" if pd.isna(v) else f"{v:,.2f}".replace(",", " ")  # noqa: E731
    fmt3 = lambda v: "" if pd.isna(v) else f"{v:,.3f}".replace(",", " ")  # noqa: E731
    fmt = {c: fmt2 for c in ("Prix unitaire", "Valeur HT")}
    fmt.update({c: fmt3 for c in (core.C_NET, core.C_NETU, core.C_NETF, core.C_GROSSF, core.C_MANIF, "Qté")})

    def row_style(row):
        styles = core.highlight_rows(row)
        css = SEV_CSS.get(int(sev.get(row.name, 9)))
        if css:
            cols = list(row.index)
            for k, c in enumerate(cols):
                if c in ("Item", "Réf article", "Remarque", core.C_BASE, core.C_NET, core.C_NETF, core.C_GROSSF):
                    styles[k] = f"{styles[k]};{css};font-weight:600" if styles[k] else f"{css};font-weight:600"
        return styles
    return view.style.apply(row_style, axis=1).format(fmt, na_rep="")


def render_import_checks(checks: dict, eng_loaded: bool, final: pd.DataFrame | None = None, man_final: pd.DataFrame | None = None):
    """Affiche les contrôles faits avant la ventilation : complétude par client, manifeste, engagement (déposés en haut de page)."""
    import import_checks as ic
    ui.section("Contrôles du dossier : clients, manifeste, engagement", "✅")
    comp = checks.get("Complétude clients")
    if comp is not None:
        n_bad = int((comp["Statut"] == "À compléter").sum())
        with st.expander(f"Données par client — {len(comp) - n_bad}/{len(comp)} complet(s)", expanded=bool(n_bad)):
            st.caption("Par client : nombre de lignes auxquelles il manque le NGP, le pays d'origine, le poids net, la quantité ou la valeur. "
                       "Complétez-les dans « Tableau TOUS_CLIENTS » (le poids net manquant est estimé à partir du brut).")
            st.dataframe(comp, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2, gap="medium")
    with c1:
        st.markdown("**Manifeste ↔ factures**")
        if "Manifeste vs dossier" in checks:
            t1 = checks["Manifeste vs dossier"]
            n_ko = int((t1["Statut"] != "OK").sum())
            (st.success if not n_ko else st.warning)("Totaux conformes au manifeste." if not n_ko else f"{n_ko} écart(s) avec le manifeste.")
            st.dataframe(t1, use_container_width=True, hide_index=True)
            with st.expander("Par client / expéditeur", expanded=bool(n_ko)) if "Manifeste par client" in checks else st.container():
                if "Manifeste par client" in checks:
                    st.dataframe(checks["Manifeste par client"], use_container_width=True, hide_index=True)
                else:
                    st.caption("Ce manifeste ne donne que les totaux (pas de détail par expéditeur).")
            fixes = ic.value_fixes(final, man_final) if (final is not None and man_final is not None) else {}
            if fixes:
                st.warning("L'écart de valeur de ces clients est exactement un facteur 10 / 100 / 1000 (virgule ou séparateur de milliers mal lu) : "
                           + " · ".join(f"{c} (÷{f:g})" if f > 1 else f"{c} (×{1 / f:g})" for c, f in fixes.items()))
                if st.button("Corriger automatiquement ces valeurs", key="imp_fix_vals", icon=":material/build:",
                             help="Divise « Valeur HT » et « Prix unitaire » des lignes de ces clients ; vos modifications non validées du tableau sont rechargées."):
                    base = st.session_state["lines_edit"].copy()
                    cl = base["Client"].map(core.clean_cell)
                    for c, f in fixes.items():
                        m = cl == c
                        for col in ("Valeur HT", "Prix unitaire"):
                            base.loc[m, col] = pd.to_numeric(base.loc[m, col], errors="coerce") / f
                    st.session_state["lines_edit"] = base
                    st.session_state["editor_version"] = st.session_state.get("editor_version", 0) + 1
                    st.session_state.pop("xlsx_sig", None)
                    st.rerun()
        else:
            st.info(st.session_state.get("imp_man_msg") or "Aucun manifeste chargé : déposez-le dans la carte « Manifeste (seul) » en haut de page, puis relancez.")
    with c2:
        st.markdown("**Engagement ↔ factures (par client)**")
        if "Engagement vs factures" in checks:
            t3 = checks["Engagement vs factures"]
            n_ko = int((~t3["Statut"].eq("OK")).sum())
            (st.success if not n_ko else st.warning)(f"{len(t3) - n_ko}/{len(t3)} client(s) conforme(s) à l'engagement.")
            st.dataframe(t3, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun tableau d'engagement chargé : déposez-le dans la carte « Tableau d'engagement » en haut de page.")
    if len(checks) > 1:
        st.download_button("Télécharger les contrôles (Excel)", data=ic.checks_xlsx(checks), file_name="CONTROLES_DOSSIER.xlsx",
                           icon=":material/download:", key="dl_checks", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Ces contrôles sont aussi ajoutés (feuilles « VERIF … ») au classeur TOUS_CLIENTS téléchargé.")


def render_import_douane(final: pd.DataFrame, man_final: pd.DataFrame):
    """Ventilation douane Alstom Railways Maroc : un classeur Excel par fournisseur (feuilles « Ventilation » + « Contrôle BL »)
    + récapitulatif consolidé (Détail / Total par Fournisseur / Total Global)."""
    import ventilation_douane as vd
    import ventilation_import as vi
    ui.section("Tableaux Excel par fournisseur — Ventilation + Contrôle BL", "📑")
    st.caption("Un classeur par fournisseur : feuille « Ventilation » (une ligne par ligne de facture, références jamais fusionnées, formules vivantes, "
               "NGP du référentiel) et feuille « Contrôle BL » (écarts poids / valeur, rapprochement avec le manifeste, anomalies). "
               "Plus un récapitulatif consolidé de tout le dossier.")
    names = list(dict.fromkeys(str(v or "").strip() or "(sans nom)" for v in final["Client"]))
    c1, c2 = st.columns([1, 2], gap="medium")
    ratio = c1.number_input("Ratio net / brut estimé (%)", min_value=50.0, max_value=100.0, step=1.0, key="dou_ratio",
                            value=round(vi.default_ratio(final) * 100, 1),
                            help="Seulement pour les factures sans poids net : net = brut × ratio, réparti au prorata de la quantité.") / 100.0
    grp = c2.multiselect("Mettre ces fournisseurs dans le même tableau (sinon : un fichier par fournisseur)", names, key="dou_group",
                         placeholder="Aucun — un fichier par fournisseur",
                         help="À utiliser seulement si vous voulez regrouper volontairement plusieurs fournisseurs dans un seul classeur.")
    man = man_final if (man_final is not None and len(man_final)) else None
    sig = (round(ratio, 4), tuple(grp), len(final), round(float(pd.to_numeric(final["Valeur HT"], errors="coerce").fillna(0).sum()), 2),
           0 if man is None else len(man))
    if st.button("Générer les tableaux Excel", type="primary", icon=":material/table_view:", key="dou_go"):
        with st.spinner("Construction des classeurs…"):
            res = vd.build_all(final, man, ratio, [grp] if len(grp) > 1 else None)
        st.session_state["imp_douane"] = {"sig": sig, "res": res}
    cur = st.session_state.get("imp_douane")
    if not cur:
        return
    if cur["sig"] != sig:
        st.info("Les données ou les options ont changé : régénérez les tableaux.")
    res = cur["res"]
    sm = res["summary"].rename(columns={"supplier": "Fournisseur", "lignes": "Lignes", "factures": "Factures", "qté": "Qté", "net": "Poids net (kg)",
                                        "brut": "Poids brut (kg)", "ht": "Total HT", "frais": "Frais", "devise": "Devise",
                                        "avis_valeur": "Valeur manifeste", "avis_brut": "Brut manifeste", "file": "Fichier"})
    st.dataframe(sm, use_container_width=True, hide_index=True)
    flags = [(s_["supplier"], f) for s_ in res["stats"] for f in s_["flags"]]
    if flags:
        with st.expander(f"Points à vérifier ({len(flags)})", expanded=True):
            for sup, f in flags[:40]:
                st.markdown(f"- **{sup}** : {f}")
    d1, d2 = st.columns(2)
    d1.download_button("Tout télécharger (.zip : un classeur par fournisseur + récapitulatif)", data=res["zip"], file_name=f"VENTILATIONS_{res['stamp']}.zip",
                       icon=":material/download:", type="primary", use_container_width=True, key="dl_dou_zip", mime="application/zip")
    d2.download_button("RECAP_CONSOLIDE.xlsx (Détail · Total par fournisseur · Total global)", data=res["recap"], file_name="RECAP_CONSOLIDE.xlsx",
                       icon=":material/download:", use_container_width=True, key="dl_dou_recap",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    pick = st.selectbox("Télécharger un seul fournisseur", list(res["files"]), key="dou_pick")
    st.download_button(pick, data=res["files"][pick], file_name=pick, icon=":material/download:", key="dl_dou_one",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_import(ocr, ai_cfg, ai_label: str, basis: str):
    result0 = st.session_state.get("result")

    # ---------------- En-tête
    if result0:
        s0 = core.compute_summary(core.apply_matching(result0["lines"], result0["refdb"]))
        stats = {"lignes ventilées": s0["lignes"], "factures": s0["factures"], "fournisseurs": s0["fournisseurs"],
                 "alertes": sum(1 for a in result0.get("alerts", []) if a["Sévérité"] != "Info")}
    else:
        stats = {"formats": "PDF · Excel · Word", "langues": "toutes", "lecture": "IA + contrôles" if ai_cfg else "règles",
                 "feuille Excel": "25 colonnes"}
    ui.hero_3d(stats, ai_label=ai_label)

    # ---------------- Imports
    c1, c2 = st.columns(2, gap="medium")
    groups_raw: list[tuple[str, list]] = []
    files = []
    with c1:
        with st.container(key="card_upl_pdf"):
            ui.card_title("📂", "Documents de la remorque",
                          "Un dépôt par client (conseillé) ou un dossier global — factures, packing lists, BL, manifeste, tous formats",
                          ["PDF", "PDF scanné", "Excel", "Word", "CSV", "Images"])
            mode = st.radio("Dépôt des documents", ["Par client (un dépôt par client)", "Dossier global (un seul dépôt)"],
                            horizontal=True, key="imp_mode", label_visibility="collapsed")
            by_client = mode.startswith("Par client")
            if by_client:
                n_cl = int(st.number_input("Nombre de clients", min_value=1, max_value=40, value=2, step=1, key="imp_ncl",
                                           help="Une case de dépôt par client : le tableau Excel garde chaque client séparé, avec sa couleur."))
            else:
                files = st.file_uploader("Dossier complet (PDF, Excel, Word, CSV, images)", type=fileio.UPLOAD_TYPES,
                                         accept_multiple_files=True, key="imp_files")
    with c2:
        with st.container(key="card_upl_ref"):
            ui.card_title("📗", "Référentiel articles & NGP",
                          "Importer le fichier Excel de référence (Références & NGP)",
                          ["Réf Article", "Réf Alstom (PN)", "Code NGP", ".xlsx / .xls / .csv"])
            ref_file = st.file_uploader("Référentiel Excel (Références & NGP)", type=["xlsx", "xlsm", "xls", "csv"],
                                        key="imp_ref")
            try:
                import referentiel_store as rs
                r_meta, r_add = rs.meta(), len(rs.additions())
            except Exception:  # noqa: BLE001
                r_meta, r_add = {}, 0
            if ref_file is not None:
                st.caption(":material/save: Ce tableau sera enregistré et remplacera le référentiel actuel.")
            elif r_meta:
                st.caption(f":material/check_circle: Référentiel enregistré utilisé automatiquement : **{r_meta.get('filename')}** — "
                           f"{r_meta.get('count', 0)} réf. · {r_meta.get('date', '')}"
                           + (f" · + {r_add} NGP ajouté(s) dans le site" if r_add else ""))
            else:
                st.caption("Aucun référentiel enregistré : déposez votre tableau, il sera gardé pour les dossiers suivants.")
            if r_add:
                with st.expander(f"NGP ajoutés dans le site ({r_add})"):
                    adds = rs.additions()
                    st.dataframe(pd.DataFrame([{"Référence": a.get("ref"), "NGP": a.get("ngp"), "Désignation": a.get("desc"),
                                                "Date": a.get("date")} for a in adds.values()]), hide_index=True, use_container_width=True)
                    rm = st.multiselect("Supprimer", list(adds), format_func=lambda k: f"{adds[k].get('ref')} → {adds[k].get('ngp')}",
                                        key="ngp_add_rm")
                    if rm and st.button("Supprimer la sélection", key="ngp_add_rm_btn"):
                        rs.remove(rm)
                        st.rerun()
    if by_client:
        cols_cl = st.columns(2, gap="medium")
        for i in range(n_cl):
            with cols_cl[i % 2]:
                with st.container(key=f"card_upl_cl_{i}", border=True):
                    st.markdown(f"**Client {i + 1}**")
                    nm = st.text_input("Nom du client (facultatif — sinon détecté dans la facture)", key=f"imp_cl_name_{i}",
                                       placeholder="Ex. JUMO, EAO, Phoenix Contact…")
                    fl = st.file_uploader(f"Documents du client {i + 1}", type=fileio.UPLOAD_TYPES, accept_multiple_files=True,
                                          key=f"imp_cl_files_{i}", label_visibility="collapsed")
                    groups_raw.append((nm.strip(), list(fl or [])))
        files = [f for _, fl in groups_raw for f in fl]
    import import_checks as ic
    c3, c4 = st.columns(2, gap="medium")
    with c3:
        with st.container(key="card_upl_man"):
            ui.card_title("🧾", "Manifeste (seul)", "Facultatif : sert à comparer valeur, poids brut et colis avec les factures",
                          ["PDF", "Scan", "Excel"])
            man_files = st.file_uploader("Manifeste (PDF, scan, Excel…)", type=fileio.UPLOAD_TYPES, accept_multiple_files=True,
                                         key="imp_man_only")
    eng_df, eng_cols = None, None
    with c4:
        with st.container(key="card_upl_eng"):
            ui.card_title("📊", "Tableau d'engagement", "Facultatif : valeur attendue par client, comparée aux factures",
                          [".xlsx / .xls", ".csv"])
            ef = st.file_uploader("Engagement (Excel ou CSV)", type=["xlsx", "xls", "xlsm", "csv"], key="imp_eng")
            if ef is not None:
                try:
                    eng_df = ic.read_table(ef.getvalue(), ef.name)
                    g_cli, g_val = ic.guess_columns(eng_df)
                    cols = list(eng_df.columns)
                    a, b = st.columns(2)
                    cli = a.selectbox("Colonne client", cols, index=cols.index(g_cli), key="imp_eng_cli")
                    val = b.selectbox("Colonne valeur", cols, index=cols.index(g_val), key="imp_eng_val")
                    eng_cols = (cli, val)
                except Exception as exc:  # noqa: BLE001
                    eng_df = None
                    st.error(f"Tableau d'engagement illisible ({type(exc).__name__}: {exc}).")
    bad = files_overview(files) if files else []
    usable = [f for f in (files or []) if f.name not in bad]
    groups = ([(nm, [f for f in fl if f.name not in bad]) for nm, fl in groups_raw] if by_client else [("", usable)])
    groups = [(nm, fl) for nm, fl in groups if fl]
    if files and bad:
        st.warning(f"{len(bad)} fichier(s) inutilisable(s) seront ignorés : corrigez-les ou retirez-les de la liste.")

    done = "result" in st.session_state
    ui.steps_strip(2 if done else (1 if usable else 0))
    launch = st.button("Lancer la ventilation du dossier", type="primary", use_container_width=True,
                       disabled=not usable, icon=":material/play_arrow:")
    if usable and not ref_file and not r_meta and not r_add:
        st.caption("Sans fichier de référence, toutes les lignes seront « NON TROUVÉ ».")

    if launch:
        progress = st.progress(0.0, text="Initialisation...")
        try:
            if by_client:
                result = core.run_pipeline_clients(groups, ref_file, ocr, basis, progress, ai_cfg)
            else:
                result = core.run_pipeline(usable, ref_file, ocr, basis, progress, ai_cfg)
        except Exception as exc:  # garde-fou : l'application ne doit jamais « planter »
            progress.empty()
            st.error(f"Le traitement n'a pas pu aboutir ({type(exc).__name__}: {exc}). Vos fichiers n'ont pas été "
                     "modifiés. Retirez le fichier en cause (voir la liste ci-dessus) ou convertissez-le en PDF / Excel, "
                     "puis relancez.")
            with st.expander("Détail technique (à transmettre au support)"):
                st.code(traceback.format_exc())
            return
        st.session_state.pop("imp_man_msg", None)
        if man_files:                                      # manifeste déposé seul : remplace celui lu dans le dossier
            try:
                man_extra, man_log = core.manifest_from_files(man_files, ocr, ai_cfg)
            except Exception as exc:  # noqa: BLE001
                man_extra, man_log = pd.DataFrame(), [f"{type(exc).__name__}: {exc}"]
            if man_extra.empty:
                st.session_state["imp_man_msg"] = ("Le manifeste déposé n'a pas pu être lu automatiquement (aucun tableau exploitable) : "
                                                   "saisissez ses lignes dans l'onglet « Manifeste ». " + " ".join(man_log[:2]))
            else:
                result["manifest"] = man_extra
        try:                                               # calcul multi-méthodes (vote) : corrections sûres + ⚠
            import import_consensus
            result["lines"], c_notes, c_tab = import_consensus.analyse(result["lines"], result["manifest"], result.get("inv_totals"))
        except Exception as exc:  # noqa: BLE001
            c_notes, c_tab = [f"⚠ Contrôle multi-méthodes indisponible ({type(exc).__name__}: {exc})."], pd.DataFrame()
        result["consensus"] = c_tab
        result.setdefault("alerts", []).extend(
            shipment.make_alert("Avertissement" if n.startswith("⚠") else "Info", "Vote multi-méthodes", n) for n in c_notes)
        final0 = core.apply_matching(result["lines"], result["refdb"])
        st.session_state["result"] = result
        st.session_state["lines_edit"] = final0
        st.session_state["manifest_edit"] = result["manifest"]
        st.session_state["editor_version"] = st.session_state.get("editor_version", 0) + 1
        st.session_state.pop("xlsx_sig", None)
        progress.progress(1.0, text="✅ Traitement terminé.")
        st.rerun()

    if "result" not in st.session_state:
        return
    result = st.session_state["result"]
    ver = st.session_state.get("editor_version", 0)
    n_pages = result.get("n_pages", 0)
    issues = result.get("issues") or []

    top = st.container()
    tabs = st.tabs([":material/dashboard: Résumé", ":material/table_rows: Tableau TOUS_CLIENTS",
                    ":material/calculate: Ventilation DUM", ":material/table_view: Tableaux par fournisseur",
                    ":material/visibility: Contrôle visuel", ":material/notification_important: Alertes",
                    ":material/hub: Groupes de documents", ":material/receipt_long: Manifeste",
                    ":material/fact_check: Analyse IA", ":material/folder_open: Fichiers & journal"])
    t_sum, t_edit, t_vent, t_dou, t_view, t_alert, t_grp, t_man, t_ai, t_files = tabs

    # ---------------- Tableau modifiable + manifeste
    with t_edit:
        st.caption("Corrigez ou complétez : NGP, contrôles, alertes et classeur Excel sont recalculés automatiquement. "
                   "Colonnes jaunes (NGP référentiel) calculées. Lignes ajoutables / supprimables.")
        edited = st.data_editor(
            st.session_state["lines_edit"], num_rows="dynamic", use_container_width=True, height=540,
            key=f"lines_editor_{ver}", column_order=core.COLUMNS, disabled=list(NGP_COLS),
            column_config=_column_config(n_pages))
    with t_man:
        st.caption("Poids brut et Goods Value du manifeste par fournisseur (extraits si le manifeste contient un "
                   "tableau, sinon à saisir). Rattachement automatique par nom proche.")
        man_edit = st.data_editor(
            st.session_state["manifest_edit"], num_rows="dynamic", use_container_width=True, key=f"manifest_editor_{ver}",
            column_config={"Goods Value": st.column_config.NumberColumn(format="%.2f"),
                           "Poids (kg)": st.column_config.NumberColumn(format="%.3f")})

    final = core.apply_matching(clean_edited(edited), result["refdb"])
    man_final = man_edit.copy()
    for c in ("Goods Value", "Poids (kg)", "Colis"):
        if c in man_final.columns:
            man_final[c] = pd.to_numeric(man_final[c], errors="coerce")
    for c in ("Expéditeur", "Shipment", "Devise"):
        if c in man_final.columns:
            man_final[c] = man_final[c].map(core.clean_cell)
    alerts_df = shipment.collect_alerts(result.get("alerts", []), shipment.live_alerts(final), issues)
    sev_map = shipment.severity_map(alerts_df)
    sev = final.apply(lambda r: shipment.row_severity(r, sev_map), axis=1) if len(final) else pd.Series(dtype=int)

    # ---------------- Classeur Excel (reconstruit seulement si le contenu change)
    try:
        checks = ic.build_checks(final, man_final, eng_df, eng_cols)
    except Exception as exc:  # noqa: BLE001
        checks = {}
        st.warning(f"Contrôles manifeste / engagement indisponibles ({type(exc).__name__}: {exc}).")
    sig = hashlib.md5((final.astype(str).to_csv() + man_final.astype(str).to_csv() + alerts_df.astype(str).to_csv()
                       + "".join(df.astype(str).to_csv() for df in checks.values())).encode()).hexdigest()
    if st.session_state.get("xlsx_sig") != sig:
        try:
            wb_bytes = core.build_workbook(final, man_final, result.get("docs"), n_pages, alerts_df, result.get("groups"))
            if checks:
                wb_bytes = ic.add_sheets_to_xlsx(wb_bytes, checks)        # feuilles « VERIF … » : complétude, manifeste, engagement
            st.session_state["xlsx"] = wb_bytes
            st.session_state["xlsx_sig"] = sig
        except Exception as exc:
            st.session_state["xlsx"] = None
            st.session_state["xlsx_sig"] = None
            st.error(f"Génération du classeur Excel impossible ({type(exc).__name__}: {exc}).")
    xlsx = st.session_state.get("xlsx")
    s = core.compute_summary(final)
    n_crit = int((alerts_df["Sévérité"] == "Critique").sum()) if not alerts_df.empty else 0
    n_warn = int((alerts_df["Sévérité"] == "Avertissement").sum()) if not alerts_df.empty else 0
    n_info = int((alerts_df["Sévérité"] == "Info").sum()) if not alerts_df.empty else 0

    # ---------------- Bandeau de statut + téléchargements (1 clic)
    with top:
        n_err = sum(1 for i in issues if i.get("Gravité") == "Erreur")
        if n_err:
            ui.status_banner("ko", "", f"{n_err} fichier(s) ignoré(s) ou illisible(s)",
                             "Voir l'onglet « Fichiers & journal » pour la cause et la marche à suivre.")
        if not s["lignes"]:
            ui.status_banner("ko", "", "Aucune ligne de facture extraite",
                             "Consultez « Fichiers & journal » ou ajoutez les lignes à la main dans le tableau.")
        elif n_crit:
            ui.status_banner("ko", "", f"{n_crit} alerte(s) critique(s) à traiter avant déclaration",
                             f"{n_warn} avertissement(s) · voir l'onglet « Alertes ». Les lignes concernées sont surlignées.")
        elif n_warn or s["non_trouves"] or s["ecarts"]:
            ui.status_banner("warn", "", "Dossier à vérifier",
                             f"{n_warn} avertissement(s) · {s['non_trouves']} NON TROUVÉ · {s['ecarts']} écart(s) NGP.")
        else:
            ui.status_banner("ok", "", "Dossier conforme", "Aucune alerte : prêt pour la déclaration.")
        if s["lignes"] and s["non_trouves"] >= 0.9 * s["lignes"]:
            rd = result.get("refdb") or {}
            ex_fact = ", ".join(dict.fromkeys(str(v) for v in final["Réf article"].head(40) if str(v).strip() and str(v) != core.NON_TROUVE))
            ex_fact = ", ".join(ex_fact.split(", ")[:4])
            ex_ref = ", ".join(list(rd.get("exact", {}))[:4])
            st.warning(
                f"**Presque aucune référence trouvée dans le référentiel** ({s['non_trouves']}/{s['lignes']} NON TROUVÉ). "
                f"Référentiel lu : **{rd.get('count', 0)}** référence(s)" + (f" — exemples : `{ex_ref}`" if ex_ref else " — **le fichier n'a pas pu être lu** (voir « Fichiers & journal »)")
                + f". Références lues dans les factures — exemples : `{ex_fact}`. "
                "Comparez les deux : si elles ne se ressemblent pas, ce n'est pas le bon fichier / la bonne colonne ; "
                "si elles se ressemblent (préfixe, tiret, zéros), dites-le pour adapter la correspondance.")
        if s["non_trouves"]:
            render_ngp_additions(final, result)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        import ventilation_douane as vd
        dsig = (sig, tuple(sorted(st.session_state.get("dou_group", []) or [])))
        cache = st.session_state.get("dou_auto")
        if not cache or cache["sig"] != dsig:
            try:
                man_for = man_final if (man_final is not None and len(man_final)) else None
                cache = {"sig": dsig, "res": vd.build_all(final, man_for, None, None), "err": ""}
            except Exception as exc:  # noqa: BLE001
                cache = {"sig": dsig, "res": None, "err": f"{type(exc).__name__}: {exc}"}
            st.session_state["dou_auto"] = cache
        dres = cache["res"]
        d1, d2, d3 = st.columns([2.2, 1.3, 0.8])
        with d1:
            if dres and len(dres["files"]) == 1:
                fname, fdata = next(iter(dres["files"].items()))
                st.download_button("Télécharger la ventilation (Excel : Ventilation + Contrôle BL)", data=fdata, file_name=fname,
                                   icon=":material/download:", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   type="primary", use_container_width=True, key="dl_dou_main")
            elif dres:
                st.download_button(f"Télécharger les ventilations ({len(dres['files'])} fournisseurs + récapitulatif) — .zip", data=dres["zip"],
                                   file_name=f"VENTILATIONS_{dres['stamp']}.zip", icon=":material/download:", mime="application/zip",
                                   type="primary", use_container_width=True, key="dl_dou_main")
            else:
                st.error(f"Ventilation Excel indisponible ({cache['err']}).")
        with d2:
            st.download_button("TOUS_CLIENTS (.xlsx) — tableau d'analyse", data=xlsx or b"", disabled=not xlsx,
                               file_name=f"TOUS_CLIENTS_{stamp}.xlsx", icon=":material/download:",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, key="dl_xlsx")
        with d3:
            csv_bytes = final[core.COLUMNS].to_csv(sep=";", index=False, decimal=",").encode("utf-8-sig")
            st.download_button("CSV", data=csv_bytes, file_name=f"TOUS_CLIENTS_{stamp}.csv",
                               mime="text/csv", use_container_width=True, icon=":material/csv:", key="dl_csv")

    # ---------------- Résumé
    with t_sum:
        pct_ok = round(100 * s["ok"] / s["lignes"]) if s["lignes"] else 0
        ui.kpi_cards([
            {"icon": "list", "label": "Lignes extraites", "value": s["lignes"], "sub": "individualisées", "tone": "blue"},
            {"icon": "file", "label": "Factures", "value": s["factures"], "sub": f"{result.get('n_files', 0)} fichier(s) · {n_pages} p.", "tone": "cyan"},
            {"icon": "building", "label": "Fournisseurs", "value": s["fournisseurs"], "sub": "expéditeurs", "tone": "gold"},
            {"icon": "alert", "label": "Alertes critiques", "value": n_crit, "sub": f"{n_warn} avertissement(s)",
             "tone": "bad" if n_crit else ("warn" if n_warn else "good")},
            {"icon": "help", "label": "NON TROUVÉ", "value": s["non_trouves"],
             "sub": "à créer dans le référentiel" if s["non_trouves"] else "aucun", "tone": "warn" if s["non_trouves"] else "good"},
            {"icon": "check", "label": "Contrôle NGP OK", "value": f"{pct_ok} %", "sub": f"{s['ok']} / {s['lignes']} lignes",
             "tone": "good" if pct_ok == 100 else "blue"},
        ])
        cards2 = [{"icon": "banknote", "label": f"Valeur HT {dev or '(devise ?)'}", "value": round(float(tot), 2), "decimals": 2,
                   "sub": "total des lignes", "tone": "blue"} for dev, tot in s["par_devise"].items()][:3]
        cards2 += [{"icon": "scale", "label": "Poids net total", "value": round(s["poids_net"], 3), "decimals": 3,
                    "sub": "kg · Σ lignes", "tone": "cyan"},
                   {"icon": "package", "label": "Poids brut factures", "value": round(s["poids_brut"], 3), "decimals": 3,
                    "sub": "kg · Σ factures", "tone": "gold"}]
        if cards2:
            ui.kpi_cards(cards2[:6])
        render_consensus(result.get("consensus"))
        ui.section("Synthèse par fournisseur / client", "building")
        ui.supplier_cards(shipment.supplier_summary(final, alerts_df))
        if len(final):
            final_for_ui = final.rename(columns={"Client": "Fournisseur / Expéditeur", "N° facture": "N° Facture",
                                                 "NGP (REF_PAR_FRS)": "NGP (Excel Réf)"}).assign(
                Devise=final["_devise"] if "_devise" in final.columns else "")
            g1, g2 = st.columns([1.5, 1], gap="medium")
            with g1:
                with st.container(key="card_chart_val"):
                    ui.card_title("💶", "Valeur HT par fournisseur", "Survolez une barre pour le détail")
                    try:
                        ui.chart_value_by_supplier(final_for_ui)
                    except Exception:
                        st.write(final_for_ui.groupby("Fournisseur / Expéditeur")["Valeur HT"].sum().to_frame())
            with g2:
                with st.container(key="card_chart_ngp"):
                    ui.card_title("🎯", "Contrôle NGP", "Répartition des lignes par statut")
                    try:
                        ui.chart_ngp_status(final_for_ui)
                    except Exception:
                        st.write(final_for_ui["Contrôle NGP"].value_counts().to_frame())

    # ---------------- Contrôle visuel (lecture seule, filtres, surlignage)
    with t_vent:
        import import_ventilation_ui
        if not len(final):
            st.info("Aucune ligne de facture : lancez d'abord le traitement du dossier.")
        else:
            try:
                render_import_checks(checks, eng_df is not None, final, man_final)
                import_ventilation_ui._render_ventilation_section(final, kind="import")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Ventilation DUM impossible ({type(exc).__name__}: {exc}).")
                st.code(traceback.format_exc())

    with t_dou:
        if not len(final):
            st.info("Aucune ligne de facture : lancez d'abord le traitement du dossier.")
        else:
            try:
                render_import_douane(final, man_final)
            except Exception as exc:  # noqa: BLE001
                st.error(f"Génération des tableaux impossible ({type(exc).__name__}: {exc}).")
                st.code(traceback.format_exc())

    with t_view:
        f1, f2, f3 = st.columns([2, 1.2, 1.4])
        sups = sorted({core.clean_cell(c) or "À compléter" for c in final["Client"]}) if len(final) else []
        pick = f1.multiselect("Fournisseur / client", sups, default=[], placeholder="Tous", key=f"f_sup_{ver}")
        only = f2.toggle("Lignes en alerte seulement", value=False, key=f"f_alert_{ver}")
        query = f3.text_input("Recherche (réf, désignation, n° facture)", key=f"f_q_{ver}")
        view = final[core.COLUMNS].copy()
        mask = pd.Series(True, index=view.index)
        if pick:
            mask &= final["Client"].map(lambda c: core.clean_cell(c) or "À compléter").isin(pick)
        if only and len(sev):
            mask &= sev.lt(9)
        if query.strip():
            q = query.strip().lower()
            mask &= view[["Réf article", "Désignation", "N° facture"]].astype(str).apply(
                lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)
        view = view[mask]
        st.caption(f"{len(view)} ligne(s) affichée(s) sur {len(final)} — rouge : alerte critique · orange : à vérifier. "
                   "Survolez le N° de facture / la colonne « Base / remarque poids » pour l'origine de chaque poids.")
        if len(view):
            st.dataframe(_style_view(view, sev), use_container_width=True, height=560, hide_index=True)
        else:
            st.info("Aucune ligne ne correspond aux filtres.")

    # ---------------- Alertes
    with t_alert:
        ui.kpi_cards([
            {"icon": "alert", "label": "Critiques", "value": n_crit, "sub": "à traiter avant déclaration", "tone": "bad" if n_crit else "good"},
            {"icon": "alert", "label": "Avertissements", "value": n_warn, "sub": "à vérifier", "tone": "warn" if n_warn else "good"},
            {"icon": "help", "label": "Infos", "value": n_info, "sub": "pour information", "tone": "blue"},
        ])
        if alerts_df.empty:
            ui.status_banner("ok", "", "Aucune alerte", "Poids, calculs, liaison des documents et fichiers : tout est cohérent.")
        else:
            sev_pick = st.multiselect("Sévérité", list(shipment.SEV_ORDER), default=["Critique", "Avertissement"],
                                      key=f"a_sev_{ver}")
            types = sorted(alerts_df["Type"].unique())
            typ_pick = st.multiselect("Type", types, default=[], placeholder="Tous", key=f"a_typ_{ver}")
            av = alerts_df[alerts_df["Sévérité"].isin(sev_pick)]
            if typ_pick:
                av = av[av["Type"].isin(typ_pick)]
            av = av.copy()
            av.insert(0, " ", av["Sévérité"].map(shipment.SEV_ICON))
            st.dataframe(av, use_container_width=True, hide_index=True, height=480,
                         column_config={" ": st.column_config.TextColumn(" ", width="small"),
                                        "Message": st.column_config.TextColumn("Message", width="large")})
            st.caption("Écarts de poids : facture ↔ packing list / BL ↔ manifeste (tolérance 50 g ou 0,5 %). "
                       "Corrigez le tableau : les alertes se recalculent.")

    # ---------------- Groupes de documents
    with t_grp:
        groups = result.get("groups")
        if groups is None or groups.empty:
            st.info("Aucun groupe de documents (aucune facture détectée).")
        else:
            n_ok = int(groups["Statut"].astype(str).str.startswith("✅").sum())
            ui.kpi_cards([
                {"icon": "layers", "label": "Groupes (factures)", "value": len(groups), "sub": "facture + documents liés", "tone": "blue"},
                {"icon": "check", "label": "Complets", "value": n_ok, "sub": "facture + packing list/BL cohérents", "tone": "good"},
                {"icon": "alert", "label": "À vérifier", "value": len(groups) - n_ok, "sub": "documents manquants ou écart", "tone": "warn" if len(groups) - n_ok else "good"},
            ])
            st.caption("Chaque facture est reliée automatiquement à son BL / packing list (n° cité, fournisseur, proximité) "
                       "et au manifeste (nom du fournisseur). Le poids manquant sur une ligne est repris de la packing list / du BL.")
            gcfg = {c: st.column_config.NumberColumn(c, format="%.3f") for c in groups.columns if c.startswith("Poids")}
            gcfg["Packing list / BL liés"] = st.column_config.TextColumn("Packing list / BL liés", width="large")
            st.dataframe(groups, use_container_width=True, hide_index=True, height=360, column_config=gcfg)
            ui.section("Par fournisseur", "building")
            for sup, g in groups.groupby("Fournisseur", sort=False):
                ok = g["Statut"].astype(str).str.startswith("✅").all()
                with st.expander(f"{'✅' if ok else '⚠️'}  {core.clean_cell(sup) or 'À compléter'} — {len(g)} facture(s)"):
                    for _, r in g.iterrows():
                        st.markdown(f"**Facture {r['N° facture'] or '—'}** · {r['Facture']}  \n"
                                    f"Packing list / BL : {r['Packing list / BL liés']}  \n"
                                    f"Manifeste : {r['Manifeste']}  \n{r['Statut']}")

    with t_ai:
        core.render_audit(result.get("audit") or [])

    with t_files:
        if issues:
            ui.section("Problèmes de fichiers", "alert")
            st.dataframe(pd.DataFrame(issues), use_container_width=True, hide_index=True,
                         column_config={"Problème": st.column_config.TextColumn(width="large"),
                                        "Que faire": st.column_config.TextColumn(width="large")})
        else:
            st.success("Tous les fichiers ont été lus sans problème.")
        ui.section("Documents détectés dans le dossier", "layers")
        st.dataframe(result["docs"], use_container_width=True, hide_index=True)
        st.caption(f"Référentiel chargé : {result['refdb']['count']} références.")
        ui.section("Journal de traitement", "list")
        for msg in result["log"]:
            st.write(msg)


# =============================================================================
# COMPTES UTILISATEURS : connexion, inscription, espace utilisateur, espace Admin, dispatch
# =============================================================================

import database  # noqa: E402
import mailer  # noqa: E402

PLAN_CHOICES = {"Gratuit (consultation)": ("free", None), "Pro — 1 mois": ("pro", 1), "Pro — 3 mois": ("pro", 3),
                "Pro — 6 mois": ("pro", 6), "Pro — 12 mois": ("pro", 12), "Pro — sans limite": ("pro", None),
                "Collaborateur (outils Pro offerts, sans accès Admin)": ("collab", None)}
PRO_MONTHS = {k: v[1] for k, v in PLAN_CHOICES.items() if v[0] == "pro" and v[1]}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DISPATCH_FILES = ("FACTURES.xlsx", "BL_REFERENCES_SUMMARY.xlsx")


def current_user() -> dict | None:
    return st.session_state.get("auth_user")


def logout():
    st.session_state.clear()
    st.rerun()


def sidebar_account(user: dict):
    ui.sidebar_cap("Compte", "shield")
    ui.sidebar_pill(ui.GOOD, f"<b>{user['email']}</b>")
    st.caption("Administrateur" if user["role"] == "admin" else "Utilisateur")
    if st.button("Se déconnecter", use_container_width=True, icon=":material/logout:", key="btn_logout"):
        logout()
    account_password_box(user)
    if user["role"] == "admin":
        admin_pending_watcher()


def _is_local() -> bool:
    """Le formulaire « Créer l'Admin » n'est proposé que sur la machine locale (jamais sur un site publié)."""
    try:
        host = (st.context.headers.get("Host") or "").split(":")[0].lower()
    except Exception:  # noqa: BLE001
        return False
    return host in ("localhost", "127.0.0.1", "::1", "[::1]")


def _send_code(email: str) -> tuple[bool, str]:
    """Génère et envoie un code de vérification. Sans SMTP : affiché à l'écran uniquement sur le poste local (mode test)."""
    code, msg = database.issue_verification_code(email)
    if code is None:
        return False, msg
    if mailer.is_configured():
        ok, info = mailer.send_verification_code(email, code)
        return ok, ("Un code de vérification a été envoyé à " + email + "." if ok else info)
    if _is_local():
        st.session_state["dev_code"] = code
        return True, "Mode test (SMTP non configuré) : le code est affiché ci-dessous."
    return False, "L'envoi d'e-mails n'est pas configuré sur ce serveur : contactez l'administrateur."


def _verify_step(email: str):
    """Saisie du code reçu par e-mail ; une fois validé, le compte attend l'approbation de l'Admin."""
    ui.card_title("✉️", "Vérifiez votre e-mail", f"Un code à 6 chiffres a été envoyé à {email}")
    if st.session_state.get("dev_code"):
        st.warning(f"Mode test — code : **{st.session_state['dev_code']}**")
    with st.form("form_verify"):
        code = st.text_input("Code de vérification (6 chiffres)", max_chars=6)
        go = st.form_submit_button("Valider le code", type="primary", use_container_width=True)
    if go:
        ok, msg = database.check_verification_code(email, code)
        if ok:
            st.session_state.pop("verify_email", None)
            st.session_state.pop("dev_code", None)
            st.session_state["verified_ok"] = True
            st.rerun()
        else:
            st.error(msg)
    c1, c2 = st.columns(2)
    if c1.button("Renvoyer le code", use_container_width=True, key="btn_resend"):
        ok, msg = _send_code(email)
        (st.success if ok else st.error)(msg)
    if c2.button("Retour", use_container_width=True, key="btn_verify_back"):
        st.session_state.pop("verify_email", None)
        st.session_state.pop("dev_code", None)
        st.rerun()


SITE_SETTINGS = [  # (clé, libellé, valeur par défaut, aide)
    ("PRO_PRICE_TEXT", "Prix affiché de la formule Pro", "50 $ / mois", "Ex. : 99 $ / mois"),
    ("CONTACT_WHATSAPP", "Numéro WhatsApp (avec indicatif, sans +)", "", "Ex. : 212612345678 — le bouton de contact s'affiche s'il est rempli"),
    ("CONTACT_LABEL", "Texte du bouton de contact", "Pour info, contacter TIMAR", ""),
    ("PAYMENT_INSTRUCTIONS", "Comment payer (RIB, nom, banque…)", "", "Affiché aux clients sous « Passer à Pro »"),
    ("HERO_VIDEO_URL", "Vidéo de l'en-tête (lien direct .mp4)", "", "Optionnel. Vide = vidéo du dossier static/hero.mp4 si elle existe, sinon camion animé"),
    ("MAINTENANCE_MESSAGE", "Message affiché pendant la maintenance", "Le site est temporairement indisponible. Merci de votre patience.", ""),
]


def maintenance_on() -> bool:
    """Mode maintenance : activé par l'Admin dans l'application (ou MAINTENANCE_MODE=1 dans les Secrets)."""
    return str(_site("MAINTENANCE_MODE", "")).strip().lower() in ("1", "true", "oui", "on")


def maintenance_message() -> str:
    return _site("MAINTENANCE_MESSAGE", "Le site est temporairement indisponible. Merci de votre patience.")


def _site(key: str, default: str = "") -> str:
    """Réglage du site : modifié par l'Admin dans l'application, sinon Secrets / variables d'environnement, sinon défaut."""
    try:
        val = database.get_setting(key)
    except Exception:  # noqa: BLE001
        val = ""
    return val or mailer._setting(key, default)


def _price_text() -> str:
    return _site("PRO_PRICE_TEXT", "50 $ / mois")


def _contact_link() -> str:
    num = "".join(ch for ch in _site("CONTACT_WHATSAPP", "") if ch.isdigit())
    return f"https://wa.me/{num}" if num else ""


def plans_card():
    """Présentation des deux formules (page de connexion et espace gratuit)."""
    c1, c2 = st.columns(2, gap="small")
    with c1:
        with st.container(key="card_plan_free"):
            ui.card_title("📂", "Gratuit", "Consultation et téléchargement",
                          ["Dossiers attribués", "Téléchargement Excel"])
            st.caption("Recevez et téléchargez les fichiers préparés pour vous par TIMAR.")
    with c2:
        with st.container(key="card_plan_pro"):
            ui.card_title("⚡", f"Pro — {_price_text()}", "Les mêmes outils que l'administrateur",
                          ["Import / ventilation", "Export BL & Factures", "Lecture IA", "Vos propres fichiers Excel"])
            st.caption("Préparez vous-même vos fichiers FACTURES.xlsx, BL_REFERENCES_SUMMARY.xlsx et TOUS_CLIENTS.xlsx.")


def upgrade_card():
    link = _contact_link()
    with st.container(key="card_upgrade"):
        ui.card_title("⚡", "Passer à Pro", f"{_price_text()} — créez vos propres fichiers (ventilation, export, lecture IA)")
        st.caption("Contactez l'administrateur pour activer la formule Pro sur votre compte, puis réglez l'abonnement avec lui.")
        if link:
            st.link_button(_site("CONTACT_LABEL", "Pour info, contacter TIMAR"), link, icon=":material/chat:")
        pay = _site("PAYMENT_INSTRUCTIONS", "")
        if pay:
            st.markdown("**Comment payer**")
            st.text(pay)


def render_auth():
    """Page d'accueil : connexion / inscription (ou création du premier compte Admin)."""
    ui.hero_3d({"accès": "sécurisé", "validation": "par l'Admin", "espace": "personnel"},
               badge="TIMAR TANGER · ESPACE SÉCURISÉ", title="Connexion", highlight="Dossiers dédouanés",
               desc="Connectez-vous pour consulter les fichiers qui vous ont été attribués. Les nouveaux comptes sont "
                    "activés par l'administrateur.")
    if maintenance_on():
        ui.status_banner("warn", "🟠", "Site en maintenance", maintenance_message() + " (seul l'administrateur peut se connecter.)")
    ui.section("Nos formules", "⚡")
    plans_card()
    import ui_login_lamp
    ui_login_lamp.inject()
    _, mid, _ = st.columns([0.6, 2, 0.6])
    with mid:
        with st.container(key="card_auth_blue"):
            if st.session_state.get("verify_email"):
                _verify_step(st.session_state["verify_email"])
                return
            lamp_on = st.session_state.get("lamp_on", False)
            left, right = st.columns([0.8, 1.2], gap="large")
            with left:
                st.markdown(ui_login_lamp.lamp_scene(lamp_on), unsafe_allow_html=True)
                if st.button("Éteindre la lampe" if lamp_on else "Tirer le cordon  💡", key="lamp_pull", use_container_width=True):
                    st.session_state["lamp_on"] = not lamp_on
                    st.rerun()
                if not lamp_on:
                    st.markdown('<div class="lamp-hint">Tirez le cordon pour allumer et vous connecter</div>', unsafe_allow_html=True)
            with right:
                if not lamp_on:
                    st.markdown('<div class="lamp-ghost"><h3>Connexion</h3><p>E-mail</p><p>Mot de passe</p></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="lamp-form"></div>', unsafe_allow_html=True)
                    _auth_forms()


def _auth_forms():
    """Onglets Connexion / S'inscrire (design d'origine, inchangé)."""
    if st.session_state.pop("verified_ok", False):
        st.success("E-mail vérifié. Votre compte est en attente de validation de l'Admin.")
    no_admin = not database.admin_exists()
    setup_ok = no_admin and _is_local()
    tab_in, tab_up, *tab_setup = st.tabs(["Connexion", "S'inscrire"] + (["Créer l'Admin"] if setup_ok else []))
    if no_admin and not setup_ok:
        st.warning("Aucun administrateur n'est encore configuré : définissez TIMAR_ADMIN_EMAIL et TIMAR_ADMIN_PASSWORD "
                   "(Secrets) puis redémarrez l'application.")
    if setup_ok:
        with tab_setup[0]:
            st.caption("Première utilisation : créez le compte Administrateur (accès complet).")
            with st.form("form_setup_admin"):
                email = st.text_input("E-mail de l'administrateur")
                pw = st.text_input("Mot de passe (8 caractères minimum)", type="password")
                pw2 = st.text_input("Confirmer le mot de passe", type="password")
                ok = st.form_submit_button("Créer le compte Admin", type="primary", use_container_width=True)
            if ok:
                if pw != pw2:
                    st.error("Les deux mots de passe sont différents.")
                else:
                    done, msg = database.create_user(email, pw, role="admin", status="active")
                    if done:
                        st.success("Compte Admin créé. Vous pouvez vous connecter (onglet Connexion).")
                    else:
                        st.error(msg)
    with tab_in:
        with st.form("form_login"):
            email = st.text_input("E-mail", key="login_email")
            pw = st.text_input("Mot de passe", type="password", key="login_pw")
            go = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
        if go:
            user, code = database.authenticate(email, pw)
            if code == "ok" and maintenance_on() and user["role"] != "admin":
                st.warning(maintenance_message())
            elif code == "ok":
                st.session_state["auth_user"] = user
                st.rerun()
            elif code == "unverified":
                st.session_state["verify_email"] = user["email"]
                ok, msg = _send_code(user["email"])
                if not ok:
                    st.session_state.pop("verify_email", None)
                    st.error(msg)
                else:
                    st.rerun()
            elif code == "pending":
                st.warning("Votre compte n'a pas encore été validé par l'administrateur.")
            elif code == "locked":
                st.error("Trop de tentatives échouées. Réessayez dans 15 minutes.")
            else:
                st.error("E-mail ou mot de passe incorrect.")
    with tab_up:
        with st.form("form_signup"):
            email = st.text_input("E-mail (Gmail)", key="su_email")
            pw = st.text_input("Mot de passe (8 caractères minimum)", type="password", key="su_pw")
            pw2 = st.text_input("Confirmer le mot de passe", type="password", key="su_pw2")
            go = st.form_submit_button("S'inscrire", type="primary", use_container_width=True)
        if go:
            if pw != pw2:
                st.error("Les deux mots de passe sont différents.")
            else:
                if not mailer.is_configured():
                    # Pas d'e-mail configuré : inscription directe, validée ensuite par l'Admin
                    done, msg = database.create_user(email, pw, role="user", status="pending", verified=True)
                    if done:
                        st.success("Compte créé avec succès. En attente de validation de l'Admin.")
                    else:
                        st.error(msg)
                else:
                    done, msg = database.create_user(email, pw, role="user", status="pending", verified=False)
                    if not done:
                        st.error(msg)
                    else:
                        ok, info = _send_code(email)
                        if ok:
                            st.session_state["verify_email"] = database.normalize_email(email)
                            st.rerun()
                        else:
                            database.delete_user(email)          # pas de compte orphelin si l'e-mail n'a pas pu partir
                            st.error(info)


# ----------------------------------------------------------------------------- espace utilisateur
def _zip_dossier(row: dict, email: str) -> bytes:
    import io
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname in row["file_paths"]:
            data = database.read_dispatch_file(row["id"], fname, email)
            if data is not None:
                z.writestr(fname, data)
    return buf.getvalue()


def render_user_space(user: dict):
    """Consultation et téléchargement uniquement. La liste se met à jour toute seule (toutes les 5 s) :
    un dossier attribué par l'Admin apparaît sans rafraîchir la page, avec une notification."""
    ui.hero_3d({"compte": "actif", "accès": "personnel"}, badge="TIMAR TANGER · ESPACE PERSONNEL",
               title="Mes Fichiers", highlight="Dédouanés",
               desc="Retrouvez ici les dossiers attribués par l'administrateur et téléchargez vos fichiers Excel.")

    @st.fragment(run_every=5)
    def live():
        sub = database.subscription_info(database.get_user(user["email"]) or user)
        if sub["state"] == "expire":
            st.warning(f"Votre abonnement Pro a pris fin le {sub['end']} : vous gardez l'accès à vos dossiers, mais plus aux outils.")
        elif sub["state"] in ("actif", "bientot"):
            (st.warning if sub["state"] == "bientot" else st.caption)(
                f"Abonnement Pro actif jusqu'au {sub['end']} ({sub['days_left']} jour(s) restant(s))."
                + (" Pensez à le renouveler." if sub["state"] == "bientot" else ""))
        rows = database.list_dispatch(user["email"])
        ids = {r["id"] for r in rows}
        seen = st.session_state.get("seen_dispatch")
        fresh = st.session_state.setdefault("new_dispatch", set())
        if seen is not None:
            for r in rows:
                if r["id"] not in seen:
                    fresh.add(r["id"])
                    st.toast(f"Nouveau dossier : {r['dossier_name']}", icon="📁")
        st.session_state["seen_dispatch"] = ids
        fresh &= ids
        ui.section(f"Mes Fichiers Dédouanés ({len(rows)})", "📂")
        if fresh:
            ui.status_banner("ok", "🟢", f"{len(fresh)} nouveau(x) dossier(s) reçu(s)", "Ils sont signalés « Nouveau » ci-dessous.")
        if not rows:
            ui.status_banner("warn", "🟠", "Aucun dossier pour le moment",
                             "Cette page se met à jour automatiquement dès que l'administrateur vous attribue un dossier.")
            return
        for r in rows:
            with st.container(key=f"card_dos_{r['id']}"):
                tags = (["🆕 Nouveau"] if r["id"] in fresh else []) + list(r["file_paths"])
                ui.card_title("📁", r["dossier_name"], f"Attribué le {r['created_at'].replace('T', ' à ')}", tags)
                names = list(r["file_paths"])
                cols = st.columns(len(names) + 1 if len(names) > 1 else 1)
                for col, fname in zip(cols, names):
                    data = database.read_dispatch_file(r["id"], fname, user["email"])
                    with col:
                        if data is None:
                            st.caption(f"{fname} : fichier introuvable")
                            continue
                        st.download_button(f"📥 Télécharger {fname}", data=data, file_name=fname, mime=XLSX_MIME,
                                           type="primary", use_container_width=True, key=f"dl_{r['id']}_{fname}")
                if len(names) > 1:
                    with cols[-1]:
                        st.download_button("📦 Tout (.zip)", data=_zip_dossier(r, user["email"]),
                                           file_name=f"{r['dossier_name']}.zip", mime="application/zip",
                                           use_container_width=True, key=f"zip_{r['id']}")
    live()
    if database.effective_plan(database.get_user(user["email"]) or user) == "free":
        upgrade_card()


def account_password_box(user: dict):
    """Changement de mot de passe (barre latérale)."""
    with st.expander("Changer mon mot de passe"):
        with st.form("form_chg_pw", clear_on_submit=True):
            old = st.text_input("Mot de passe actuel", type="password")
            new = st.text_input("Nouveau mot de passe (8 caractères min.)", type="password")
            go = st.form_submit_button("Enregistrer", use_container_width=True)
        if go:
            ok, msg = database.change_password(user["email"], old, new)
            (st.success if ok else st.error)(msg)


def admin_pending_watcher():
    """Barre latérale (Admin) : nombre de demandes d'accès en attente, mis à jour toutes les 8 s, avec notification."""
    @st.fragment(run_every=8)
    def _w():
        n = len(database.list_users(status="pending", role="user"))
        prev = st.session_state.get("pending_prev")
        if prev is not None and n > prev:
            st.toast(f"Nouvelle demande d'accès ({n} en attente)", icon="🔔")
        st.session_state["pending_prev"] = n
        ui.sidebar_pill(ui.WARNING if n else ui.GOOD,
                        f"<b>{n}</b>&nbsp;demande(s) d'accès en attente" if n else "Aucune demande en attente")
    _w()


# ----------------------------------------------------------------------------- dispatch (Admin)
def dispatch_box(files: dict[str, bytes], default_name: str = "", key: str = "dsp"):
    """Choisit les utilisateurs actifs qui recevront les fichiers générés, et enregistre l'attribution."""
    ui.section("Attribuer ce dossier à des utilisateurs", "📤")
    active = database.list_users(status="active", role="user")
    if not active:
        st.info("Aucun utilisateur actif. Approuvez d'abord une demande dans « Gestion des Accès & Dispatch ».")
        return
    c1, c2 = st.columns([1, 1.3], gap="medium")
    name = c1.text_input("Nom du dossier", value=default_name or datetime.now().strftime("Dossier_%Y%m%d_%H%M"), key=f"{key}_name")
    dests = c2.multiselect("Utilisateur(s) destinataire(s)", [u["email"] for u in active], key=f"{key}_dest")
    if st.button("Envoyer le dossier aux utilisateurs choisis", type="primary", icon=":material/send:", key=f"{key}_go",
                 disabled=not dests):
        ok, msg = database.dispatch_dossier(name, dests, files)
        (st.success if ok else st.error)(msg)


def render_admin_panel():
    ui.hero_3d({"demandes": len(database.list_users(status="pending", role="user")),
                "utilisateurs actifs": len(database.list_users(status="active", role="user"))},
               badge="TIMAR TANGER · ADMINISTRATION", title="Gestion des Accès", highlight="& Dispatch",
               desc="Validez les demandes d'inscription et attribuez les fichiers FACTURES.xlsx et "
                    "BL_REFERENCES_SUMMARY.xlsx aux utilisateurs.")
    import persist
    if not persist.enabled():
        st.error("**Sauvegarde non configurée** : les comptes acceptés, paramètres et fichiers attribués sont **effacés à chaque "
                 "Reboot / mise à jour GitHub**. Ajoutez SUPABASE_URL et SUPABASE_KEY dans les Secrets (onglet « Paramètres du site »).",
                 icon=":material/cloud_off:")
    elif persist.blocked():
        st.error("**Sauvegarde Supabase injoignable au démarrage** : pour protéger vos données, rien n'est envoyé tant que la "
                 "restauration n'a pas réussi. Évitez d'accepter des comptes maintenant ; faites « Reboot app » dans quelques minutes.",
                 icon=":material/cloud_off:")
    pending = database.list_users(status="pending", role="user")
    t1, t2, t3, t4, t5 = st.tabs([f"Demandes d'Accès & Attribution ({len(pending)})", "Utilisateurs actifs",
                                  "Attribuer des fichiers", "Historique des dispatchs", "Paramètres du site"])
    with t1:
        ui.section("Demandes d'accès en attente", "🕓")
        if not pending:
            st.info("Aucune demande en attente.")
        for u in pending:
            c1, cd, c2, c3 = st.columns([2.2, 1.8, 1, 1.3], vertical_alignment="center")
            c1.markdown(f"<b>{u['email']}</b><br><span style='opacity:.6;font-size:.82rem'>inscrit le {u['created_at'].replace('T', ' ')}</span>",
                        unsafe_allow_html=True)
            dur = cd.selectbox("Formule", list(PLAN_CHOICES), index=0, key=f"dur_{u['id']}", label_visibility="collapsed")
            if c2.button("Approuver", type="primary", key=f"ap_{u['id']}", icon=":material/check:", use_container_width=True):
                plan, months = PLAN_CHOICES[dur]
                database.approve_user(u["email"], months, plan)
                st.toast(f"{u['email']} approuvé.")
                st.rerun()
            if c3.button("Refuser / Supprimer", key=f"rf_{u['id']}", icon=":material/delete:", use_container_width=True):
                database.delete_user(u["email"])
                st.toast(f"{u['email']} supprimé.")
                st.rerun()
    with t2:
        ui.section("Créer un utilisateur", "➕")
        with st.form("form_admin_create_user", clear_on_submit=True):
            c1, c2, c3 = st.columns([2, 2, 1])
            new_email = c1.text_input("E-mail de l'utilisateur")
            new_pw = c2.text_input("Mot de passe (8 caractères minimum)", type="password")
            new_dur = c3.selectbox("Formule", list(PLAN_CHOICES), index=0)
            go = st.form_submit_button("Créer le compte (actif immédiatement)", type="primary")
        if go:
            ok, msg = database.create_user(new_email, new_pw, role="user", status="active", verified=True)
            if ok:
                database.set_plan(new_email, PLAN_CHOICES[new_dur][0], PLAN_CHOICES[new_dur][1])
            if ok:
                st.success(f"Compte créé et actif : {database.normalize_email(new_email)}. Transmettez-lui son e-mail et son mot de passe.")
            else:
                st.error(msg)
        ui.section("Utilisateurs actifs", "👥")
        actives = database.list_users(status="active", role="user")
        if not actives:
            st.info("Aucun utilisateur actif.")
        for u in actives:
            c1, c2, c3 = st.columns([3, 1, 1.3], vertical_alignment="center")
            sub = database.subscription_info(u)
            sub_txt = {"gratuit": "Gratuit", "pro_illimite": "⚡ Pro sans limite", "collab": "🤝 Collaborateur (outils Pro offerts)",
                       "actif": f"⚡ Pro jusqu'au {sub['end']} ({sub['days_left']} j)",
                       "bientot": f"⚠ Pro jusqu'au {sub['end']} ({sub['days_left']} j)",
                       "expire": f"⛔ Pro expiré le {sub['end']} (repassé en Gratuit)"}[sub["state"]]
            c1.markdown(f"<b>{u['email']}</b><br><span style='opacity:.6;font-size:.82rem'>{len(u['assigned_files'])} dossier(s) attribué(s) · {sub_txt}</span>",
                        unsafe_allow_html=True)
            with c1.popover("Formule / abonnement", icon=":material/event_repeat:"):
                dd = st.selectbox("Activer / prolonger Pro de", list(PRO_MONTHS), key=f"rn_{u['id']}")
                if st.button("Prolonger", key=f"rn_go_{u['id']}"):
                    new_end = database.set_plan(u["email"], "pro", PRO_MONTHS[dd])
                    st.success(f"Pro jusqu'au {new_end}.")
                if st.button("Pro sans limite", key=f"rn_inf_{u['id']}"):
                    database.set_plan(u["email"], "pro", None)
                    st.success("Pro sans limite.")
                if st.button("Collaborateur (offert)", key=f"rn_col_{u['id']}"):
                    database.set_plan(u["email"], "collab")
                    st.success("Collaborateur : outils Pro offerts, sans accès à la gestion des accès.")
                if st.button("Repasser en Gratuit", key=f"rn_free_{u['id']}"):
                    database.set_plan(u["email"], "free")
                    st.success("Formule Gratuite.")
            with c1.popover("Nouveau mot de passe", icon=":material/key:"):
                npw = st.text_input("Nouveau mot de passe", type="password", key=f"npw_{u['id']}")
                if st.button("Enregistrer", key=f"npw_go_{u['id']}"):
                    ok, msg = database.set_password(u["email"], npw)
                    (st.success if ok else st.error)(msg)
            if c2.button("Suspendre", key=f"sp_{u['id']}", icon=":material/pause:", use_container_width=True):
                database.suspend_user(u["email"])
                st.rerun()
            if c3.button("Supprimer", key=f"del_{u['id']}", icon=":material/delete:", use_container_width=True):
                database.delete_user(u["email"])
                st.rerun()
    with t3:
        ui.section("Fichiers générés dans cette session", "📦")
        files = dict(st.session_state.get("exp_files") or {})
        if files:
            st.success("Les fichiers du dernier export (module EXPORT) sont prêts à être attribués.")
        else:
            st.caption("Aucun export dans cette session : générez-les dans le module EXPORT (le bloc d'attribution y apparaît "
                       "sous les téléchargements), ou déposez-les ici.")
        up = st.file_uploader("Ou déposez FACTURES.xlsx et BL_REFERENCES_SUMMARY.xlsx", type=["xlsx"],
                              accept_multiple_files=True, key="dsp_upload")
        if up:
            files = {f.name: f.getvalue() for f in up}
        if files:
            dispatch_box(files, key="dsp_panel")
    with t4:
        ui.section("Historique des attributions", "🗂️")
        hist = database.list_dispatch()
        if not hist:
            st.info("Aucun dossier attribué pour le moment.")
        else:
            st.dataframe(pd.DataFrame([{"Dossier": h["dossier_name"], "Utilisateur": h["assigned_to_email"],
                                        "Fichiers": ", ".join(h["file_paths"]), "Date": h["created_at"].replace("T", " ")}
                                       for h in hist]), use_container_width=True, hide_index=True)
            pick = st.selectbox("Retirer une attribution", [""] + [f"{h['id']} — {h['dossier_name']} → {h['assigned_to_email']}" for h in hist],
                                key="dsp_rm")
            if pick and st.button("Retirer l'attribution", icon=":material/delete:", key="dsp_rm_go"):
                database.delete_dispatch(int(pick.split(" — ")[0]))
                st.rerun()
    with t5:
        ui.section("Paramètres du site", "⚙️")
        st.caption("Ces réglages s'affichent aux clients (carte « Passer à Pro »). Laisser vide = valeur par défaut / Secrets.")
        maint = st.toggle("Mode maintenance : bloquer l'accès aux clients (l'administrateur garde l'accès)",
                          value=database.get_setting("MAINTENANCE_MODE") == "1", key="tg_maint")
        if maint != (database.get_setting("MAINTENANCE_MODE") == "1"):
            database.set_setting("MAINTENANCE_MODE", "1" if maint else "")
            st.rerun()
        if maint:
            st.warning("Le site est EN MAINTENANCE : les clients ne peuvent plus se connecter et sont déconnectés.")
        with st.form("form_site_settings"):
            vals = {}
            for key, label, default, hint in SITE_SETTINGS:
                cur = database.get_setting(key)
                vals[key] = (st.text_area if key == "PAYMENT_INSTRUCTIONS" else st.text_input)(
                    label, value=cur, placeholder=default, help=hint or None)
            saved = st.form_submit_button("Enregistrer les paramètres", type="primary")
        if saved:
            for key, val in vals.items():
                database.set_setting(key, val)
            st.success("Paramètres enregistrés : visibles tout de suite par les clients.")
        import persist
        ui.section("Sauvegarde durable (Supabase)", "☁️")
        ps = persist.status()
        if ps["enabled"]:
            st.success("Sauvegarde active : comptes, paramètres, fichiers attribués et modèles Excel sont conservés même après un redémarrage du site.")
            if ps["last_error"]:
                st.error(f"Dernière erreur : {ps['last_error']}")
            if st.button("Sauvegarder maintenant", key="persist_now", icon=":material/cloud_upload:"):
                ok = persist.backup_db(database.DB_PATH)
                (st.success if ok else st.error)("Base sauvegardée." if ok else f"Échec : {persist.status()['last_error'] or 'voir les Secrets SUPABASE_URL / SUPABASE_KEY'}")
        else:
            st.warning("Sauvegarde non configurée : les comptes et fichiers sont perdus à chaque redémarrage (Reboot, mise à jour GitHub). "
                       "Ajoutez dans Secrets : SUPABASE_URL et SUPABASE_KEY (clé service_role). Voir le README (« Sauvegarde Supabase »).")
        ui.section("Modèle de ventilation EXPORT (VENTI_EXPORT)", "🧮")
        import ventilation_export as vx
        st.caption("Modèle Excel rempli automatiquement dans le module EXPORT (BASE DONNER, NGPN, ATTELIER…). "
                   "Téléversez une nouvelle version après l'avoir mise à jour (nouvelles références).")
        up = st.file_uploader("Remplacer le modèle (.xltx ou .xlsx)", type=["xltx", "xlsx"], key="vx_tpl_up")
        if up is not None and st.button("Enregistrer ce modèle", key="vx_tpl_save"):
            vx.save_template_override(up.getvalue())
            st.cache_resource.clear()
            st.success("Modèle remplacé.")
        st.caption("Modèle actuel : " + ("téléversé par l'administrateur" if vx.override_path().is_file() else "modèle livré avec l'application"))
        ui.section("Modèle de ventilation IMPORT (VENTI_IMPORT)", "🧮")
        st.caption("Modèle rempli dans le module IMPORT (acconage, fret, taux, NGP du destinataire…).")
        upi = st.file_uploader("Remplacer le modèle IMPORT (.xltx ou .xlsx)", type=["xltx", "xlsx"], key="vi_tpl_up")
        if upi is not None and st.button("Enregistrer ce modèle IMPORT", key="vi_tpl_save"):
            vx.save_template_override(upi.getvalue(), "import")
            st.cache_resource.clear()
            st.success("Modèle IMPORT remplacé.")
        st.caption("Modèle actuel : " + ("téléversé par l'administrateur" if vx.override_path("import").is_file() else "modèle livré avec l'application"))


def render_ngp_additions(final: pd.DataFrame, result: dict):
    """Références NON TROUVÉ : saisie du NGP dans le site, gardé pour les dossiers suivants (complète le référentiel)."""
    import referentiel_store as rs
    nf = final[final["NGP (REF_PAR_FRS)"].astype(str) == core.NON_TROUVE] if "NGP (REF_PAR_FRS)" in final else final.iloc[0:0]
    if nf.empty:
        return
    tab = (nf.assign(_r=nf["Réf article"].astype(str).str.strip()).drop_duplicates("_r")
           [["Réf article", "Désignation", "Client", "NGP/code douanier (facture)"]]
           .rename(columns={"NGP/code douanier (facture)": "HS facture (indication)"}))
    tab.insert(1, "NGP à enregistrer", "")
    with st.expander(f":material/edit_note: Compléter les NGP manquants ({len(tab)} référence(s) NON TROUVÉ) — "
                     "enregistrés pour les prochains dossiers", expanded=False):
        st.caption("Saisissez le NGP (8 ou 10 chiffres) des références absentes de votre tableau. Ils complètent le référentiel "
                   "(votre tableau reste prioritaire) et seront retrouvés automatiquement dans les dossiers suivants.")
        ed = st.data_editor(tab, hide_index=True, use_container_width=True, key=f"ngp_add_{st.session_state.get('editor_version', 0)}",
                            disabled=["Réf article", "Désignation", "Client", "HS facture (indication)"])
        if st.button("Enregistrer ces NGP et recalculer", icon=":material/save:", key="ngp_add_btn"):
            rows = [{"ref": r["Réf article"], "ngp": r["NGP à enregistrer"], "desc": r["Désignation"]}
                    for _, r in ed.iterrows() if str(r["NGP à enregistrer"] or "").strip()]
            n = rs.add(rows, (st.session_state.get("user") or {}).get("email", "") if isinstance(st.session_state.get("user"), dict) else "")
            if not n:
                st.warning("Aucun NGP valide (8 ou 10 chiffres) saisi.")
                return
            refdb = rs.merge(result.get("refdb") or {"exact": {}, "compact": {}, "ocr": {}, "sheet_names": {}, "count": 0})
            result["refdb"] = refdb
            st.session_state["lines_edit"] = core.apply_matching(st.session_state["lines_edit"], refdb)
            st.session_state["editor_version"] = st.session_state.get("editor_version", 0) + 1
            st.session_state.pop("xlsx_sig", None)
            st.session_state.pop("dou_auto", None)
            st.success(f"{n} NGP enregistré(s).")
            st.rerun()


def render_consensus(cons):
    """Tableau du calcul multi-méthodes (vote) : chaque chiffre clé recalculé de 2 à 3 façons indépendantes."""
    if cons is None or not len(cons):
        return
    v = cons["Verdict"].astype(str)
    n_ok, n_fix, n_bad = int(v.str.startswith("✔ OK").sum()), int(v.str.startswith("✔ corrigé").sum()), int(v.str.startswith("⚠").sum())
    ui.section("Calcul multi-méthodes (vote)", "check")
    st.caption("Montant de chaque facture, valeur / poids brut / colis de chaque client : recalculés par 2 à 3 méthodes "
               "indépendantes (lignes, total imprimé, quantité × prix, avis d'arrivée / manifeste). "
               "La valeur retenue est celle sur laquelle au moins 2 méthodes sont d'accord.")
    c1, c2, c3 = st.columns(3)
    c1.metric("Confirmés", n_ok)
    c2.metric("Corrigés par vote", n_fix)
    c3.metric("Désaccords (à contrôler)", n_bad)
    only = st.toggle("Afficher seulement les corrections et désaccords", value=bool(n_fix or n_bad), key="imp_cons_only")
    show = cons[~v.str.startswith("✔ OK") & ~v.isin(["1 seule méthode", "aucune donnée"])] if only else cons
    st.dataframe(show.dropna(axis=1, how="all").astype(object).where(lambda x: x.notna(), ""), use_container_width=True, hide_index=True)
