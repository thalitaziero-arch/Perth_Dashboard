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

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

components.html(html, height=6000, scrolling=False)
