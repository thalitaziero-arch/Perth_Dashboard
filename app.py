from pathlib import Path
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

st.set_page_config(page_title="Perth Dashboard", layout="wide")

BASE_DIR = Path(__file__).parent

EXCEL_FILE = BASE_DIR / "team_stats_perth.xlsx"
HTML_FILE = BASE_DIR / "perth_azzurri_painel.html"



def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data
def load_stats():
    df = pd.read_excel(EXCEL_FILE)
    df = clean_columns(df)

    # keep only Perth rows if Team column exists
    if "Team" in df.columns:
        df = df[df["Team"].astype(str).str.strip().str.lower() == "perth"]

    return df


def safe_sum(df, col):
    if col not in df.columns:
        return 0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def safe_mean(df, col):
    if col not in df.columns:
        return 0
    return pd.to_numeric(df[col], errors="coerce").dropna().mean()


df = load_stats()

matches = len(df)
goals = int(safe_sum(df, "Goals"))
xg = round(safe_sum(df, "xG"), 2)

possession = round(safe_mean(df, "Possession, %"), 1)

passes = safe_sum(df, "Passes")
accurate = safe_sum(df, "Accurate")
pass_accuracy = round((accurate / passes) * 100, 1) if passes else 0

duels = safe_sum(df, "Duels")
duels_won = safe_sum(df, "Won")
duel_win = round((duels_won / duels) * 100, 1) if duels else 0

# basic W-D-L from Match text
wins = draws = losses = 0
for match in df.get("Match", []):
    text = str(match)
    score_part = text.split()[-1] if text else ""
    if ":" in score_part:
        try:
            a, b = score_part.split(":")
            a, b = int(a), int(b)
            if a > b:
                wins += 1
            elif a == b:
                draws += 1
            else:
                losses += 1
        except:
            pass

record = f"{wins}-{draws}-{losses}"

with open(HTML_FILE, "r", encoding="utf-8") as f:
    html = f.read()

# replace existing dashboard numbers
html = html.replace("10", str(matches), 1)
html = html.replace("7-3-0", record)
html = html.replace("17.80", str(xg))
html = html.replace("54.6%", f"{possession}%")
html = html.replace("75.4%", f"{pass_accuracy}%")
html = html.replace("52.0%", f"{duel_win}%")

components.html(html, height=5000, scrolling=False)