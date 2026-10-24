# -*- coding: utf-8 -*-
"""Connexion « lampe » aux couleurs du site : une lampe éteinte ; « tirer le cordon » l'allume et fait apparaître le formulaire
(onglets Connexion / S'inscrire inchangés)."""

from __future__ import annotations

import streamlit as st

BLUE, BLUE_2, SKY, NAVY = "#3B82F6", "#2563EB", "#BFDBFE", "#0A1020"

CSS = f"""
<style>
div.st-key-card_auth_blue {{
  position:relative; overflow:hidden; border-radius:26px; padding:1.4rem 1.6rem 1.2rem;
  background:
    radial-gradient(90% 70% at 15% 0%, rgba(59,130,246,.28), transparent 60%),
    radial-gradient(80% 70% at 100% 100%, rgba(37,99,235,.22), transparent 60%),
    linear-gradient(160deg, #0e1730 0%, #0a1020 65%, #0c1428 100%) !important;
  border:1px solid rgba(96,165,250,.30) !important; box-shadow:0 30px 70px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.08) !important; }}
div.st-key-card_auth_blue p, div.st-key-card_auth_blue label, div.st-key-card_auth_blue span {{ color:#E6EBF5; }}
div.st-key-card_auth_blue h3 {{ color:#fff; }}
div.st-key-card_auth_blue [data-baseweb="input"], div.st-key-card_auth_blue [data-baseweb="base-input"] {{
  background:rgba(255,255,255,.06) !important; border:1px solid rgba(147,197,253,.25) !important; border-radius:12px !important; }}
div.st-key-card_auth_blue [data-baseweb="input"]:focus-within {{ border-color:rgba(96,165,250,.9) !important; box-shadow:0 0 0 2px rgba(59,130,246,.35), 0 0 22px rgba(59,130,246,.25) !important; }}
div.st-key-card_auth_blue input {{ color:#fff !important; background:transparent !important; caret-color:{SKY}; }}
div.st-key-card_auth_blue [data-testid^="stBaseButton-primary"] {{ background:linear-gradient(180deg,#4F95F8,{BLUE_2}) !important; border:none !important; border-radius:12px !important;
  box-shadow:0 8px 22px rgba(37,99,235,.35); }}
div.st-key-card_auth_blue [data-testid^="stBaseButton-primary"] * {{ color:#fff !important; font-weight:700 !important; }}
div.st-key-card_auth_blue [data-testid^="stBaseButton-secondary"] {{ background:rgba(255,255,255,.06) !important; border:1px solid rgba(147,197,253,.28) !important; border-radius:12px !important; }}
/* ------------------------------ LAMPE ------------------------------ */
@keyframes lamp-on {{ 0% {{ opacity:0 }} 12% {{ opacity:.9 }} 20% {{ opacity:.25 }} 34% {{ opacity:1 }} 100% {{ opacity:1 }} }}
@keyframes lamp-form {{ from {{ opacity:0; transform:translateY(14px) scale(.98) }} to {{ opacity:1; transform:none }} }}
@keyframes lamp-swing {{ 0%,100% {{ transform:rotate(0) }} 25% {{ transform:rotate(1.4deg) }} 75% {{ transform:rotate(-1.4deg) }} }}
.lamp-scene {{ position:relative; height:330px; display:flex; justify-content:center; }}
.lamp-scene svg {{ position:relative; z-index:2; transform-origin:50% 0; animation: lamp-swing 5s ease-in-out infinite; }}
.lamp-cone {{ position:absolute; top:112px; left:50%; width:300px; height:230px; margin-left:-150px; z-index:1; pointer-events:none;
  clip-path:polygon(34% 0, 66% 0, 100% 100%, 0 100%);
  background:linear-gradient(180deg, rgba(191,219,254,.75), rgba(96,165,250,.28) 55%, rgba(59,130,246,0)); animation: lamp-on 1.1s ease-out both; filter:blur(.4px); }}
.lamp-glow {{ position:absolute; top:70px; left:50%; width:190px; height:130px; margin-left:-95px; z-index:1; pointer-events:none; border-radius:50%;
  background:radial-gradient(closest-side, rgba(191,219,254,.9), rgba(96,165,250,.0)); animation: lamp-on 1.1s ease-out both; }}
.lamp-hint {{ text-align:center; font-size:.82rem; color:rgba(230,235,245,.55) !important; margin-top:-.3rem; }}
.lamp-form {{ animation: lamp-form .7s .25s ease-out both; }}
.lamp-ghost {{ opacity:.18; filter:blur(3px); pointer-events:none; }}
div.st-key-lamp_pull button {{ border-radius:999px !important; background:rgba(255,255,255,.07) !important; border:1px solid rgba(147,197,253,.45) !important; }}
div.st-key-lamp_pull button:hover {{ border-color:#93C5FD !important; box-shadow:0 0 18px rgba(96,165,250,.35); }}
</style>
"""

LAMP_SVG = """
<svg width="200" height="330" viewBox="0 0 200 330" aria-hidden="true">
  <line x1="100" y1="0" x2="100" y2="56" stroke="#7aa2e0" stroke-width="3"/>
  <path d="M52 112 Q52 66 100 62 Q148 66 148 112 Z" fill="%(shade)s" stroke="#93C5FD" stroke-width="2.4"/>
  <rect x="44" y="110" width="112" height="9" rx="4.5" fill="#1d2f57" stroke="#60A5FA" stroke-width="1.6"/>
  <ellipse cx="100" cy="124" rx="17" ry="9" fill="%(bulb)s" %(bulbfx)s/>
  <line x1="138" y1="118" x2="138" y2="210" stroke="#93C5FD" stroke-width="2.2"/>
  <circle cx="138" cy="216" r="7.5" fill="#BFDBFE" stroke="#3B82F6" stroke-width="2"/>
</svg>"""


def inject():
    st.markdown(CSS, unsafe_allow_html=True)


def lamp_scene(on: bool) -> str:
    svg = LAMP_SVG % {"shade": "#2563EB" if on else "#14234a", "bulb": "#EFF6FF" if on else "#3a4a6e",
                      "bulbfx": 'style="filter:drop-shadow(0 0 14px #BFDBFE)"' if on else ""}
    return f'<div class="lamp-scene">{svg}' + ('<div class="lamp-glow"></div><div class="lamp-cone"></div>' if on else "") + "</div>"
