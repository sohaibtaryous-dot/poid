# -*- coding: utf-8 -*-
"""Décorations animées (CSS pur, aucune dépendance) inspirées de vos vidéos :
  1. Dépôt de fichiers  : un document tombe dans une caisse (comme le ballon dans le panier) ; une fois les fichiers reçus : caisse verte + coche.
  2. Connexion « Most Secure Login » : cercle cadenas + étapes Identifiants -> Contrôle de sécurité -> Authentification -> Bienvenue.
  3. Connexion / inscription en panneau coulissant (« Hello Friend ») : voir views.render_auth + CSS ci-dessous."""

from __future__ import annotations

import html
import time

import streamlit as st

CRATE = ("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='78' height='64' viewBox='0 0 78 64'>"
         "<path d='M6 26h66l-5 32a4 4 0 0 1-4 3H15a4 4 0 0 1-4-3z' fill='%23182338' stroke='%233B82F6' stroke-width='2.4' stroke-linejoin='round'/>"
         "<path d='M4 26h70' stroke='%233B82F6' stroke-width='3.2' stroke-linecap='round'/>"
         "<path d='M24 32v22M39 32v22M54 32v22' stroke='%233B82F6' stroke-opacity='.45' stroke-width='2'/></svg>")
CRATE_OK = CRATE.replace("%233B82F6", "%2322C55E")
DOC = ("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='30' height='38' viewBox='0 0 30 38'>"
       "<path d='M3 2h16l8 8v24a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z' fill='%23F3F6FD' stroke='%2393A8D6' stroke-width='1.6'/>"
       "<path d='M19 2v8h8' fill='none' stroke='%2393A8D6' stroke-width='1.6'/><path d='M7 17h16M7 22h16M7 27h10' stroke='%233B82F6' stroke-width='2' stroke-linecap='round'/></svg>")

_UPLOAD_KEYS = ("card_upl_pdf", "card_upl_exp", "card_upl_ref")
_SEL = ", ".join(f"div.st-key-{k}" for k in _UPLOAD_KEYS) + ", div[class*='st-key-card_upl_cl_']"
_SEL_BEFORE = ", ".join(f"{s}::before" for s in _SEL.split(", "))
_SEL_AFTER = ", ".join(f"{s}::after" for s in _SEL.split(", "))
_SEL_OK_B = ", ".join(f"{s}:has([data-testid='stFileChip'])::before" for s in _SEL.split(", "))
_SEL_OK_A = ", ".join(f"{s}:has([data-testid='stFileChip'])::after" for s in _SEL.split(", "))

_PAD = ", ".join(f"div.st-key-{k} .card-title" for k in _UPLOAD_KEYS)

CSS = f"""
<style>
{_SEL} {{ position:relative; }}
{_PAD} {{ padding-right:88px; }}
div[class*='st-key-card_upl_cl_'] [data-testid='stMarkdownContainer'] p:first-child {{ padding-right:80px; }}
{_SEL_BEFORE} {{ content:""; position:absolute; right:16px; top:12px; width:62px; height:52px; pointer-events:none; opacity:.85;
  background:url("{CRATE}") center bottom/contain no-repeat; transition:filter .3s; z-index:2; }}
{_SEL_AFTER} {{ content:""; position:absolute; right:30px; top:-26px; width:30px; height:38px; pointer-events:none; z-index:3;
  background:url("{DOC}") center/contain no-repeat; animation: drop-doc 2.8s cubic-bezier(.5,0,.8,.4) infinite; }}
@keyframes drop-doc {{
  0% {{ transform:translate(14px,-6px) rotate(-18deg); opacity:0 }} 10% {{ opacity:1 }}
  45% {{ transform:translate(6px,34px) rotate(10deg); opacity:1 }} 58% {{ transform:translate(0,44px) rotate(0) scale(.7); opacity:0 }}
  100% {{ transform:translate(0,44px) scale(.7); opacity:0 }} }}
/* fichiers reçus : caisse verte, plus de chute, coche */
{_SEL_OK_B} {{ background-image:url("{CRATE_OK}"); filter:drop-shadow(0 0 10px rgba(34,197,94,.55)); opacity:1; }}
{_SEL_OK_A} {{ content:"✓"; animation:pop-ok .5s ease-out both; color:#22C55E; font-weight:800; font-size:1.25rem; width:auto; height:auto;
  background:none; right:44px; top:22px; text-shadow:0 0 10px rgba(34,197,94,.8); }}
@keyframes pop-ok {{ 0% {{ transform:scale(0); opacity:0 }} 70% {{ transform:scale(1.35); opacity:1 }} 100% {{ transform:scale(1); opacity:1 }} }}

/* ---------- Connexion / inscription : panneau coulissant ---------- */
@keyframes auth-in-left  {{ from {{ transform:translateX(-26px); opacity:0 }} to {{ transform:none; opacity:1 }} }}
@keyframes auth-in-right {{ from {{ transform:translateX(34px); opacity:0 }}  to {{ transform:none; opacity:1 }} }}
@keyframes auth-float {{ 0%,100% {{ transform:translateY(0) rotate(0) }} 50% {{ transform:translateY(-14px) rotate(8deg) }} }}
div[class*="st-key-auth_form_"] {{ animation: auth-in-left .55s cubic-bezier(.2,.7,.2,1); }}
div[class*="st-key-auth_side_"] {{ animation: auth-in-right .6s cubic-bezier(.2,.7,.2,1); position:relative; overflow:hidden; border-radius:20px; min-height:360px;
  display:flex; flex-direction:column; justify-content:center; padding:1.6rem 1.4rem; text-align:center;
  background: radial-gradient(90% 70% at 20% 10%, rgba(255,214,122,.35), transparent 60%),
              linear-gradient(150deg, rgba(224,177,58,.30), rgba(120,70,10,.55) 55%, rgba(18,13,6,.9)); border:1px solid rgba(245,210,122,.35);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.18); }}
div[class*="st-key-auth_side_"]::before, div[class*="st-key-auth_side_"]::after {{ content:""; position:absolute; border-radius:22px; border:1px solid rgba(255,255,255,.22);
  background:linear-gradient(135deg, rgba(255,255,255,.13), rgba(255,255,255,.02)); animation:auth-float 6s ease-in-out infinite; pointer-events:none; }}
div[class*="st-key-auth_side_"]::before {{ width:90px; height:90px; right:-22px; top:-18px; transform:rotate(18deg); }}
div[class*="st-key-auth_side_"]::after {{ width:70px; height:70px; left:-18px; bottom:-14px; animation-delay:-3s; }}
div[class*="st-key-auth_side_"] h3 {{ margin:.2rem 0 .5rem; font-size:1.45rem; color:#fff; letter-spacing:.02em; }}
div[class*="st-key-auth_side_"] p {{ font-size:.85rem; color:rgba(255,245,225,.8) !important; margin:0 0 1rem; }}

/* ---------- Connexion : étapes « Most Secure Login » ---------- */
.sec-wrap {{ display:grid; grid-template-columns:1fr 1fr; gap:1rem; align-items:center; min-height:190px; margin:.3rem 0 .6rem; }}
.sec-ring {{ position:relative; width:104px; height:104px; margin:0 auto; border-radius:50%; display:grid; place-items:center; color:#F5D27A;
  border:2px solid rgba(224,177,58,.55); box-shadow:0 0 30px rgba(224,177,58,.25), inset 0 0 20px rgba(224,177,58,.12); }}
.sec-ring::after {{ content:""; position:absolute; inset:-9px; border-radius:50%; border:2px solid transparent; border-top-color:#F5D27A; border-right-color:rgba(245,210,122,.35);
  animation: otp-spin .8s linear 2 forwards; }}
.sec-ring.ok {{ color:#4ADE80; border-color:rgba(34,197,94,.8); box-shadow:0 0 34px rgba(34,197,94,.4); }}
.sec-ring.ok::after {{ display:none; }}
.sec-steps {{ display:flex; flex-direction:column; gap:.75rem; }}
.sec-step {{ display:flex; align-items:center; gap:.6rem; font-size:.86rem; color:rgba(243,235,221,.55); animation: sec-lit .35s var(--d) forwards; }}
.sec-step i {{ width:20px; height:20px; border-radius:50%; border:1.5px solid rgba(255,255,255,.3); display:grid; place-items:center; font-style:normal; font-size:.7rem; font-weight:800;
  animation: sec-tick .35s var(--d) forwards; }}
@keyframes sec-lit {{ to {{ color:#F3EBDD }} }}
@keyframes sec-tick {{ to {{ background:#22C55E; border-color:#22C55E; color:#04210f; box-shadow:0 0 12px rgba(34,197,94,.6) }} }}
/* Allègement : plus d'animation permanente sur la page entière (fluidité, surtout sur PC modestes / connexion lente) */
.stApp, .stApp::before, .stApp::after {{ animation:none !important; }}
.block-container [data-testid="stVerticalBlock"] > div {{ animation:none !important; }}
.stButton > button[kind="primary"]:not(:disabled), .stButton > button[kind="primary"]:not(:disabled)::after {{ animation:none !important; }}
[data-testid="stFileUploaderDropzone"], [data-testid="stProgress"] div, [data-testid="stProgress"] > div > div {{ animation:none !important; }}
@media (prefers-reduced-motion: reduce) {{ *, *::before, *::after {{ animation-duration:.01ms !important; animation-iteration-count:1 !important; transition-duration:.01ms !important; }} }}
</style>
"""

LOCK = ('<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
        '<rect x="5" y="11" width="14" height="9" rx="2.5"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>')
CHECK = ('<svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">'
         '<path d="M5 12.5 10 17.5 19 7.5"/></svg>')


def inject():
    st.markdown(CSS, unsafe_allow_html=True)


def secure_login(name: str = "", steps=("Identifiants", "Contrôle de sécurité", "Authentification"), hold: float = 1.5, after: float = 0.6):
    """Cercle cadenas + 3 étapes qui se cochent une à une, puis « Bienvenue » (≈3,7 s)."""
    li = "".join(f'<div class="sec-step" style="--d:{0.2 + k * 0.4:.1f}s"><i>✓</i>{html.escape(s)}</div>' for k, s in enumerate(steps))
    box = st.empty()
    box.markdown('<div class="otp-head"><h3>Connexion <b>sécurisée</b></h3><p>Vérification de votre accès…</p></div>'
                 f'<div class="sec-wrap"><div class="sec-ring">{LOCK}</div><div class="sec-steps">{li}</div></div>'
                 '<div class="otp-btn busy">Authentification…</div>', unsafe_allow_html=True)
    time.sleep(hold)
    who = f", {html.escape(name)}" if name else ""
    box.markdown(f'<div class="otp-head otp-ok"><h3>Bienvenue<b>{who} !</b></h3><p>Vous êtes connecté en toute sécurité.</p></div>'
                 f'<div class="sec-wrap" style="grid-template-columns:1fr"><div class="sec-ring ok">{CHECK}</div></div>'
                 '<div class="otp-btn" style="background:#22C55E;color:#04210f">Connecté et sécurisé</div>', unsafe_allow_html=True)
    time.sleep(after)
