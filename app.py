from pathlib import Path
import base64
import json
import re

import fitz  # pymupdf
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

BASE_DIR    = Path(__file__).parent
HTML_FILE   = BASE_DIR / "perth_azzurri_painel.html"
EXCEL_FILE  = BASE_DIR / "team_stats_perth.xlsx"
PERTH_PDF   = BASE_DIR / "Perth_SC.pdf"
NPL_PDF     = BASE_DIR / "NPL_Comparison.pdf"


# ── 1. Gerar DashboardData automaticamente do TeamStats ──────────────────────
def build_dashboard_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name="TeamStats", header=None)
    C = {
        'Date':65,'Match':66,'Competition':67,'Duration':68,'Team':69,
        'Scheme':70,'Goals':71,'xG':72,'Shots':73,'SOT':74,
        'Shots_acc_pct':75,'Possession_pct':76,
        'Losses_total':77,'Losses_low':78,'Losses_medium':79,'Losses_high':80,
        'Rec_total':81,'Rec_low':82,'Rec_mid':83,'Rec_high':84,
        'Duels_total':85,'Duels_won':86,'Duels_won_pct':87,
        'Fwd_passes_pct':128,
    }
    rows = df.iloc[3:].copy()
    # forward-fill Date and Match (new Wyscout format only fills Perth's row)
    rows.iloc[:, 65] = rows.iloc[:, 65].ffill()
    rows.iloc[:, 66] = rows.iloc[:, 66].ffill()
    matches = {}
    for i, row in rows.iterrows():
        k = (str(row.iloc[65])[:10], str(row.iloc[66]))
        matches.setdefault(k, []).append(i)
    round_map = {k: r+1 for r, k in enumerate(sorted(matches))}
    records = []
    for i, row in rows.iterrows():
        k = (str(row.iloc[65])[:10], str(row.iloc[66]))
        def g(c, _row=row):
            v = _row.iloc[c]
            return None if str(v) in ('nan','NaN','None') else v
        records.append({
            'Round': round_map[k],
            'Date': k[0], 'Match': k[1],
            'Competition': g(C['Competition']),
            'Duration': int(g(C['Duration']) or 90),
            'Team': g(C['Team']), 'Scheme': g(C['Scheme']),
            'Goals': int(g(C['Goals']) or 0),
            'xG': round(float(g(C['xG']) or 0), 2),
            'Shots': int(g(C['Shots']) or 0),
            'Shots_on_target': int(g(C['SOT']) or 0),
            'Shots_acc_pct': round(float(g(C['Shots_acc_pct']) or 0), 2),
            'Passes': 0, 'Passes_accurate': 0,
            'Passes_acc_pct': round(float(g(C['Fwd_passes_pct']) or 0), 2),
            'Possession_pct': round(float(g(C['Possession_pct']) or 0), 2),
            'Losses_total': int(g(C['Losses_total']) or 0),
            'Losses_low': int(g(C['Losses_low']) or 0),
            'Losses_medium': int(g(C['Losses_medium']) or 0),
            'Losses_high': int(g(C['Losses_high']) or 0),
            'Recoveries_total': int(g(C['Rec_total']) or 0),
            'Recoveries_low': int(g(C['Rec_low']) or 0),
            'Recoveries_medium': int(g(C['Rec_mid']) or 0),
            'Recoveries_high': int(g(C['Rec_high']) or 0),
            'Duels_total': int(g(C['Duels_total']) or 0),
            'Duels_won': int(g(C['Duels_won']) or 0),
            'Duels_won_pct': round(float(g(C['Duels_won_pct']) or 0), 2),
        })
    return sorted(records, key=lambda r: (r['Round'], r['Team']))


# ── 2. Extrair figura recortada de um PDF como base64 ────────────────────────
def pdf_crop_to_base64(pdf_path, page_num, x0, y0, x1, y1, scale=2.0):
    doc = fitz.open(str(pdf_path))
    page = doc[page_num - 1]
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, clip=fitz.Rect(x0, y0, x1, y1))
    return base64.b64encode(pix.tobytes("png")).decode()


# ── 3. Carregar dados ─────────────────────────────────────────────────────────
team_data = build_dashboard_data(EXCEL_FILE)

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

html = re.sub(
    r"const DATA = \[.*?\];",
    "const DATA = " + json.dumps(team_data) + ";",
    html,
    flags=re.S
)

if PERTH_PDF.exists():
    doc = fitz.open(str(PERTH_PDF))
    pg = 14 if len(doc) <= 18 else 19
    doc.close()
    # Fig 1: shots on goalkeeper (goal face) — top-left of the finishing page
    img1 = pdf_crop_to_base64(PERTH_PDF, pg, x0=0,   y0=55,  x1=298, y1=245)
    # Fig 2: field shot map — middle-left of the finishing page
    img2 = pdf_crop_to_base64(PERTH_PDF, pg, x0=0,   y0=245, x1=298, y1=475)
    html = html.replace('const FINISHING_IMG1 = "";', f'const FINISHING_IMG1 = "{img1}";')
    html = html.replace('const FINISHING_IMG2 = "";', f'const FINISHING_IMG2 = "{img2}";')

components.html(html, height=15000, scrolling=False)
