from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Perth Azzurri Dashboard", layout="wide")

st.markdown("""
<style>
  .block-container { padding: 0 !important; margin: 0 !important; max-width: 100% !important; }
  .stApp { overflow-x: hidden !important; }
  section.main > div { padding: 0 !important; max-width: 100% !important; }
  header { display: none !important; }
  footer { display: none !important; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).parent
HTML_FILE = BASE_DIR / "perth_azzurri_painel.html"

# perth_azzurri_painel.html already has all data (team stats, player stats,
# NPL rankings, shot images) baked in by atualizar_painel.py — this app just
# serves it as-is, so there's a single source of truth to keep updated.
html = HTML_FILE.read_text(encoding="utf-8")
components.html(html, height=15000, scrolling=False)
