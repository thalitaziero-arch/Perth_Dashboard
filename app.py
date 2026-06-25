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

# Column layout of the "TeamStats" sheet (current Wyscout export format).
# Keep this in sync with TS_COLS in atualizar_painel.py — if Wyscout changes
# their export columns again, update both.
C = {
    'Date': 0, 'Match': 1, 'Competition': 2, 'Duration': 3, 'Team': 4, 'Scheme': 5,
    'Goals': 6, 'xG': 7, 'Shots': 8, 'SOT': 9, 'Shots_acc_pct': 10,
    'Possession_pct': 11,
    'Losses_total': 12, 'Losses_low': 13, 'Losses_medium': 14, 'Losses_high': 15,
    'Rec_total': 16, 'Rec_low': 17, 'Rec_mid': 18, 'Rec_high': 19,
    'Duels_total': 20, 'Duels_won': 21, 'Duels_won_pct': 22,
    'Fwd_passes_pct': 63,
}


# ── 1. Gerar DashboardData automaticamente do TeamStats ──────────────────────
def build_dashboard_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name="TeamStats", header=None)
    rows = df.iloc[3:].copy()  # pula as 3 linhas de cabeçalho/médias
    # o Wyscout só preenche Date/Match na linha do Perth; a linha do
    # adversário do mesmo jogo vem em branco nessas colunas
    rows[C['Date']] = rows[C['Date']].ffill()
    rows[C['Match']] = rows[C['Match']].ffill()
    rows = rows[rows[C['Team']].notna()]  # ignora linhas totalmente vazias no fim da planilha

    matches = {}
    for i, row in rows.iterrows():
        k = (str(row.iloc[C['Date']])[:10], str(row.iloc[C['Match']]))
        matches.setdefault(k, []).append(i)
    round_map = {k: r + 1 for r, k in enumerate(sorted(matches))}

    records = []
    for i, row in rows.iterrows():
        k = (str(row.iloc[C['Date']])[:10], str(row.iloc[C['Match']]))

        def g(c, _row=row):
            v = _row.iloc[c]
            return None if str(v) in ('nan', 'NaN', 'None') else v

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
def pdf_crop_to_base64(pdf_path, page_num, x0, y0, x1, y1, scale=4.5):
    doc = fitz.open(str(pdf_path))
    page = doc[page_num - 1]
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, clip=fitz.Rect(x0, y0, x1, y1))
    doc.close()
    return base64.b64encode(pix.tobytes("png")).decode()


def find_shots_page(pdf_path):
    """Locate the 'Shots' page regardless of report length/layout."""
    doc = fitz.open(str(pdf_path))
    pg = None
    for i in range(len(doc)):
        if doc[i].get_text().startswith("Shots\nGoal\nOn target\nMiss"):
            pg = i + 1
            break
    doc.close()
    return pg


# ── 3. Carregar dados ─────────────────────────────────────────────────────────
team_data = build_dashboard_data(EXCEL_FILE)

html = HTML_FILE.read_text(encoding="utf-8")

html = re.sub(
    r"const DATA = \[.*?\];",
    "const DATA = " + json.dumps(team_data) + ";",
    html,
    flags=re.S,
)

# Inject correct shooting totals from Excel
perth_rows = [r for r in team_data if r["Team"] == "Perth"]
sd_shots = sum(r["Shots"] for r in perth_rows)
sd_sot = sum(r["Shots_on_target"] for r in perth_rows)
sd_pct = round(sd_sot / sd_shots * 100, 1) if sd_shots else 0
sd_xg = round(sum(r["xG"] for r in perth_rows), 2)
sd_goals = sum(r["Goals"] for r in perth_rows)
sd_total = f'{{"shots":{sd_shots},"on_target":{sd_sot},"pct":{sd_pct},"xg":{sd_xg},"goals":{sd_goals}}}'
html = re.sub(r"const SD_TOTAL = \{.*?\};", f"const SD_TOTAL = {sd_total};", html, flags=re.S)

if PERTH_PDF.exists():
    pg = find_shots_page(PERTH_PDF)
    if pg:
        img1 = pdf_crop_to_base64(PERTH_PDF, pg, x0=0, y0=55, x1=298, y1=245)
        img2 = pdf_crop_to_base64(PERTH_PDF, pg, x0=0, y0=245, x1=298, y1=475)
        html = re.sub(r'const FINISHING_IMG1 = "[^"]*";', f'const FINISHING_IMG1 = "{img1}";', html)
        html = re.sub(r'const FINISHING_IMG2 = "[^"]*";', f'const FINISHING_IMG2 = "{img2}";', html)

components.html(html, height=15000, scrolling=False)
