from pathlib import Path
import json
import re

import pandas as pd
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
EXCEL_FILE = BASE_DIR / "team_stats_perth.xlsx"

team_df = pd.read_excel(EXCEL_FILE, sheet_name="TeamStats")
team_data = team_df.to_dict(orient="records")

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

html = re.sub(
    r"const team = \[.*?\];",
    "const team = " + json.dumps(team_data) + ";",
    html,
    flags=re.S
)

components.html(html, height=6000, scrolling=False)
