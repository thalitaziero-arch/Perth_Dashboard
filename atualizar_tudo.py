#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
atualizar_tudo.py  —  Atualização semanal AUTOMÁTICA do painel Perth Azzurri
============================================================================

TODA SEMANA, só faça isto:
  1. Baixe do Wyscout e SUBSTITUA na pasta (mesmos nomes):
       • team_stats_perth.xlsx   (export "Team Stats" — pode ter 2025 junto, o
                                   script filtra sozinho pra temporada atual)
       • Perth_SC.pdf            (Team Report do Perth — últimos jogos)
       • NPL_Comparison.pdf      (Season Report / NPL Comparison)
  2. (Opcional) Se quiser atualizar as FIGURAS DA GOLEIRA, salve os prints na
       pasta como:
         • sch_shots_against.png       (Shots against / xG2 / penalties)
         • sch_crosses_setpieces.png   (Crosses / set pieces / distribution)
  3. Rode:
         python3 atualizar_tudo.py
     Isso atualiza TUDO e publica no GitHub (o site online atualiza em 1-2 min).

O QUE ELE FAZ SOZINHO:
  • Limpa temporadas antigas (2025 etc.) e remonta a classificação/rodadas
  • Standings, pitch (recuperações/perdas), lista de jogos, badge da rodada
  • Estatísticas das jogadoras (últimos 10 jogos do relatório) + goleiras
    (temporada completa, contando os jogos antigos que já estão no sistema)
  • GOLEADORAS da temporada — lê quem fez cada gol direto do PDF (ícone de bola)
    e guarda num "livro-caixa" (goalscorers_season.json) pra nunca perder jogo
    e nunca duplicar
  • Figura + tabela de finalização (Shooting) do PDF novo
  • Classificação e rankings da NPLW
  • Figuras da goleira (dos prints que você salvar)
  • git add/commit/push

NÃO É PRECISO EDITAR NADA À MÃO. Se algum PDF estiver faltando, ele avisa e pula
aquela parte, sem quebrar o resto.

Requisitos:  pip install pandas openpyxl pymupdf pdfplumber --break-system-packages
"""

import base64
import datetime
import json
import math
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import fitz          # pymupdf
import pandas as pd

import atualizar_painel as ap      # reaproveita NPL / shooting-crop / GPS já testados
import extract_wyscout as ew

BASE = Path(__file__).parent
HTML = BASE / "perth_azzurri_painel.html"
EXCEL = BASE / "team_stats_perth.xlsx"
PDF_NEW = BASE / "Perth_SC.pdf"          # relatório rolante (últimos ~10 jogos)
PDF_OLD = BASE / "perth_sc_old.pdf"      # relatório antigo fixo (jogos iniciais)
NPL_PDF = BASE / "NPL_Comparison.pdf"
NPL_JSON = BASE / "npl_comparison_data.json"
LEDGER = BASE / "goalscorers_season.json"   # livro-caixa de gols por jogo (não apagar!)
GK_IMG_SHOTS = BASE / "sch_shots_against.png"
GK_IMG_CROSSES = BASE / "sch_crosses_setpieces.png"

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Colunas do export TeamStats (0-index)
TS = ap.TS_COLS


# ─────────────────────────────────────────────────────────────────────────────
# 1) DASHBOARDDATA — filtra a temporada atual e remonta rodadas (sem 2025, sem dup)
# ─────────────────────────────────────────────────────────────────────────────
def rebuild_dashboard_data():
    print("1) Remontando DashboardData (temporada atual, sem duplicar)...")
    ts = pd.read_excel(EXCEL, sheet_name="TeamStats", header=None)
    rows = ts.iloc[3:].copy()

    # descobre a temporada atual = ano do jogo mais recente
    perth = rows[rows.iloc[:, TS["Team"]] == "Perth"]
    dates = pd.to_datetime(perth.iloc[:, TS["Date"]], errors="coerce").dropna()
    season_year = int(dates.max().year)
    print(f"   Temporada detectada: {season_year}")

    # coleta jogos do Perth dessa temporada, em ordem cronológica
    info = []
    for idx in rows.index:
        team = rows.at[idx, TS["Team"]]
        match = rows.at[idx, TS["Match"]]
        date = rows.at[idx, TS["Date"]]
        if team != "Perth" or pd.isna(match) or pd.isna(date):
            continue
        d = str(date)[:10]
        if not d.startswith(str(season_year)):
            continue
        info.append((d, match, idx))
    info.sort(key=lambda x: x[0])

    def g(row, c):
        return row.iloc[c]

    def build(row, rnd, match=None, date=None, comp=None, dur=None):
        return {
            "Round": rnd,
            "Date": date if date else str(g(row, TS["Date"]))[:10],
            "Match": match if match else g(row, TS["Match"]),
            "Competition": comp if comp else g(row, TS["Competition"]),
            "Duration": dur if dur else int(g(row, TS["Duration"])),
            "Team": g(row, TS["Team"]), "Scheme": g(row, TS["Scheme"]),
            "Goals": int(g(row, TS["Goals"])), "xG": round(float(g(row, TS["xG"])), 2),
            "Shots": int(g(row, TS["Shots"])), "Shots_on_target": int(g(row, TS["SOT"])),
            "Shots_acc_pct": round(float(g(row, TS["Shots_acc_pct"])), 2),
            "Passes": int(g(row, TS["Passes"])), "Passes_accurate": int(g(row, TS["Passes_accurate"])),
            "Passes_acc_pct": round(float(g(row, TS["Passes_acc_pct"])), 2),
            "Possession_pct": round(float(g(row, TS["Possession_pct"])), 2),
            "Losses_total": int(g(row, TS["Losses_total"])), "Losses_low": int(g(row, TS["Losses_low"])),
            "Losses_medium": int(g(row, TS["Losses_medium"])), "Losses_high": int(g(row, TS["Losses_high"])),
            "Recoveries_total": int(g(row, TS["Rec_total"])), "Recoveries_low": int(g(row, TS["Rec_low"])),
            "Recoveries_medium": int(g(row, TS["Rec_mid"])), "Recoveries_high": int(g(row, TS["Rec_high"])),
            "Duels_total": int(g(row, TS["Duels_total"])), "Duels_won": int(g(row, TS["Duels_won"])),
            "Duels_won_pct": round(float(g(row, TS["Duels_won_pct"])), 2),
        }

    records = []
    for rnd, (d, match, pidx) in enumerate(info, start=1):
        prow = rows.loc[pidx]
        records.append(build(prow, rnd))
        # linha do adversário = próxima linha (Match em branco no export)
        oidx = pidx + 1
        if oidx in rows.index and pd.isna(rows.at[oidx, TS["Match"]]):
            orow = rows.loc[oidx]
            records.append(build(orow, rnd, match=match, date=d,
                                 comp=str(prow.iloc[TS["Competition"]]),
                                 dur=int(prow.iloc[TS["Duration"]])))

    dd = pd.DataFrame(records)
    shutil.copy(EXCEL, EXCEL.with_suffix(".bak.xlsx"))
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL)
    if "DashboardData" in wb.sheetnames:
        del wb["DashboardData"]
    wb.save(EXCEL)
    with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a") as w:
        dd.to_excel(w, sheet_name="DashboardData", index=False)
    n_games = dd[dd["Team"] == "Perth"].shape[0]
    print(f"   {n_games} jogos (Rounds 1-{n_games}), só {season_year}. Backup: {EXCEL.with_suffix('.bak.xlsx').name}")
    return dd


# ─────────────────────────────────────────────────────────────────────────────
# 2) DATA / SD_TOTAL / badge no HTML
# ─────────────────────────────────────────────────────────────────────────────
def inject_team_data(html, dd):
    records = dd.to_dict("records")
    records = [{k: (None if isinstance(v, float) and math.isnan(v) else v) for k, v in r.items()}
               for r in records]
    for r in records:
        r["Date"] = str(r["Date"])[:10]
    html = re.sub(r"const DATA = \[.*?\];", "const DATA = " + json.dumps(records) + ";", html, flags=re.S)

    perth = [r for r in records if r["Team"] == "Perth"]
    last = max(perth, key=lambda r: r["Round"])
    ld = datetime.date.fromisoformat(last["Date"])
    html = re.sub(r'(id="homeRoundBadge"[^>]*>)Round \d+ · WA NPLW \d+(<)',
                  rf'\g<1>Round {last["Round"]} · WA NPLW {ld.year}\g<2>', html)
    html = re.sub(r'(id="homeUpdatedBadge"[^>]*>)Last updated: [^<]+(<)',
                  rf'\g<1>Last updated: {ld.day} {MESES[ld.month - 1]} {ld.year}\g<2>', html)

    shots = sum(r["Shots"] for r in perth); sot = sum(r["Shots_on_target"] for r in perth)
    pct = round(sot / shots * 100, 1) if shots else 0
    xg = round(sum(r["xG"] for r in perth), 2); goals = sum(r["Goals"] for r in perth)
    html = re.sub(r"const SD_TOTAL = \{.*?\};",
                  f'const SD_TOTAL = {{"shots":{shots},"on_target":{sot},"pct":{pct},"xg":{xg},"goals":{goals}}};',
                  html, flags=re.S)
    print(f"2) Time: Round {last['Round']} · {ld.year} · {goals} gols / {shots} chutes")
    return html, last["Round"], ld, goals


# ─────────────────────────────────────────────────────────────────────────────
# 3) GOLEADORAS — lê quem fez cada gol nos PDFs (ícone de bola) + livro-caixa
# ─────────────────────────────────────────────────────────────────────────────
def _extract_goals_from_page(page, roster, surnames):
    words = page.get_text("words")
    name_rows = [(surnames[w[4]], w[0], (w[1] + w[3]) / 2)
                 for w in words if w[4] in surnames and (w[1] + w[3]) / 2 < 420]
    left = [r for r in name_rows if r[1] < 290]
    right = [r for r in name_rows if r[1] >= 290]
    side = left if len(left) >= len(right) else right
    xmin, xmax = (0, 290) if side is left else (297, 595)
    times = [w for w in words if w[4].endswith("'") and xmin <= w[0] < xmax
             and re.match(r'^\d', w[4]) and (w[1] + w[3]) / 2 < 420]

    def is_goal(w):
        clip = fitz.Rect(w[0] - 14, (w[1] + w[3]) / 2 - 5, w[0] - 4, (w[1] + w[3]) / 2 + 5)
        pix = page.get_pixmap(matrix=fitz.Matrix(10, 10), clip=clip)
        n = pix.width * pix.height
        blk = red = yel = grn = 0
        for py in range(pix.height):
            for px in range(pix.width):
                r, gg, b = pix.pixel(px, py)
                if r < 70 and gg < 70 and b < 70: blk += 1
                elif r > 150 and gg < 110 and b < 110: red += 1
                elif r > 170 and gg > 140 and b < 120: yel += 1
                elif gg > 120 and r < 130 and b < 130: grn += 1
        return blk > n * 0.015 and red < 60 and yel < 100 and grn < 60

    goals = {}
    for w in times:
        ty = (w[1] + w[3]) / 2
        cand = [nr for nr in side if abs(nr[2] - ty) < 7 and nr[1] < w[0]]
        if cand and is_goal(w):
            pl = min(cand, key=lambda nr: abs(nr[2] - ty))[0]
            goals[pl] = goals.get(pl, 0) + 1
    return goals


def _match_pages(path, roster, surnames):
    """{date: goals_dict} lido das páginas 'Matches' de um relatório."""
    out = {}
    if not path.exists():
        return out
    doc = fitz.open(str(path))
    for i in range(len(doc)):
        t = doc[i].get_text()
        m = re.search(r'\n(\d+) [–-] (\d+)\n(\d{2}\.\d{2}\.\d{4})', t)
        if not m or ("MATCHES" not in t and "M AT C H E S" not in t):
            continue
        dmy = m.group(3)
        iso = f"{dmy[6:10]}-{dmy[3:5]}-{dmy[0:2]}"
        out[iso] = _extract_goals_from_page(doc[i], roster, surnames)
    doc.close()
    return out


def update_goalscorers(html, dd, roster):
    print("3) Goleadoras (lendo os gols de cada jogo nos PDFs)...")
    surnames = {n.split(" ", 1)[1]: n for n in roster}
    valid_dates = {str(r["Date"])[:10] for r in dd.to_dict("records")}

    ledger = json.loads(LEDGER.read_text(encoding="utf-8")) if LEDGER.exists() else {}
    # lê os dois PDFs (novo sobrescreve, mas a detecção é determinística)
    for path in (PDF_OLD, PDF_NEW):
        for iso, goals in _match_pages(path, roster, surnames).items():
            if iso in valid_dates:            # só jogos da temporada atual
                ledger[iso] = goals            # keyed por data => nunca duplica
    LEDGER.write_text(json.dumps(ledger, ensure_ascii=False, indent=2), encoding="utf-8")

    # soma o livro-caixa
    tally = defaultdict(int)
    for iso, goals in ledger.items():
        for pl, n in goals.items():
            tally[pl] += n
    player_total = sum(tally.values())
    team_total = sum(int(r["Goals"]) for r in dd.to_dict("records") if r["Team"] == "Perth")
    own_goals = team_total - player_total   # gols contra (adversário) = diferença

    scorers = sorted(tally.items(), key=lambda kv: (-kv[1], kv[0]))
    js = "const SCORERS = [\n  " + ",\n  ".join(
        "{name:'%s', goals:%d}" % (n, g) for n, g in scorers) + "\n];"
    html = re.sub(r'const SCORERS = \[.*?\];', js, html, flags=re.S)

    n_rounds = dd[dd["Team"] == "Perth"].shape[0]
    html = re.sub(r'>Full season · Rounds 1–\d+<', f'>Full season · Rounds 1–{n_rounds}<', html)
    og_txt = f" ({player_total} by players + {own_goals} own goal{'s' if own_goals != 1 else ''})" if own_goals > 0 else ""
    html = re.sub(r'Goals by player across the full season \(Rounds 1–\d+\)\. \d+ goals total[^<]*\.',
                  f'Goals by player across the full season (Rounds 1–{n_rounds}). {team_total} goals total{og_txt}.', html)
    print(f"   {len(ledger)} jogos no livro-caixa · {player_total} gols de jogadoras + {own_goals} contra = {team_total}")
    return html


# ─────────────────────────────────────────────────────────────────────────────
# 4) JOGADORAS (linha = janela do PDF) + GOLEIRAS (temporada completa)
# ─────────────────────────────────────────────────────────────────────────────
def _gk_season(path_report, roster, valid_dates, durations):
    """Jogos+minutos da temporada p/ goleiras, do relatório individual delas."""
    pass  # (implementado inline abaixo)


def rebuild_players(html, dd, roster):
    print("4) Jogadoras (últimos 10 jogos) + goleiras (temporada completa)...")
    base = {p["name"]: p for p in json.loads(re.search(r'const PLAYERS = (\[.*?\]);', html, re.S).group(1))}
    names = set(base)
    mbn = ew.parse_perth_sc_overview(str(PDF_NEW), names)
    pdf = ew.parse_perth_sc_pdf(str(PDF_NEW), names)

    valid_dates = {str(r["Date"])[:10] for r in dd.to_dict("records")}
    durations = {str(r["Date"])[:10]: int(r["Duration"])
                 for r in dd.to_dict("records") if r["Team"] == "Perth"}

    # ── goleiras: aparições da temporada (união das escalações dos 2 PDFs) ──
    surnames = {n.split(" ", 1)[1]: n for n in roster}
    appear = defaultdict(set)   # player -> set(datas)
    for path in (PDF_OLD, PDF_NEW):
        if not path.exists():
            continue
        doc = fitz.open(str(path))
        for i in range(len(doc)):
            t = doc[i].get_text()
            m = re.search(r'\n\d+ [–-] \d+\n(\d{2}\.\d{2}\.\d{4})', t)
            if not m or ("MATCHES" not in t and "M AT C H E S" not in t):
                continue
            dmy = m.group(1)
            iso = f"{dmy[6:10]}-{dmy[3:5]}-{dmy[0:2]}"
            if iso not in valid_dates:
                continue
            for n in roster:
                if n in t:
                    appear[n].add(iso)
        doc.close()

    def gk_minutes(name, report_pdf):
        """Minutos da temporada somando o relatório individual da goleira (se houver)
        + durações dos jogos que ela jogou mas não estão no relatório."""
        played = appear.get(name, set())
        mins = {}
        p = BASE / report_pdf
        if p.exists():
            doc = fitz.open(str(p))
            for i in range(len(doc)):
                t = doc[i].get_text()
                if not t.strip().startswith("Match\n"):
                    continue
                for mm in re.finditer(r'(.+?) (\d{2}\.\d{2}\.\d{4})\n(\d+)\n', t):
                    dmy = mm.group(2)
                    iso = f"{dmy[6:10]}-{dmy[3:5]}-{dmy[0:2]}"
                    if iso in played:
                        mins[iso] = int(mm.group(3))
                break
            doc.close()
        # jogos jogados sem minuto no relatório -> usa a duração da partida
        for iso in played:
            mins.setdefault(iso, durations.get(iso, 90))
        total = sum(mins.values())
        return len(played), total

    players = []
    for name, b in base.items():
        pl = dict(b)
        if b.get("is_goalkeeper"):
            report = "D. Schroeder.pdf" if name == "D. Schroeder" else \
                     "E. Ingrey.pdf" if name == "E. Ingrey" else None
            if report:
                m, tot = gk_minutes(name, report)
                if m:
                    pl["matches"] = m
                    pl["total_min"] = tot
                    pl["avg_min"] = round(tot / m)
            players.append(pl)
            continue
        # jogadora de linha: janela do PDF novo
        rec = pdf.get(name)
        if name in mbn:
            pl["matches"] = mbn[name]
        if rec:
            num = ap._numeric_record(rec)
            pl["total_min"] = int(num.get("total_min", pl.get("total_min", 0)))
            if pl.get("matches"):
                pl["avg_min"] = round(pl["total_min"] / pl["matches"])
            pl["goals"] = num.get("goals", 0); pl["xg"] = round(num.get("xg", 0.0), 2)
            pl["assists"] = num.get("assists", 0); pl["xa"] = round(num.get("xa", 0.0), 2)
            for f in ap.RATIO_FIELDS_A:
                n, d = num.get(f + "_n", 0), num.get(f + "_d", 0)
                pl[f] = f"{n}/{d}"; pl[f + "_pct"] = round(d / n * 100, 1) if n else None
            pl["losses_total"], pl["losses_own"] = num.get("losses_n", 0), num.get("losses_d", 0)
            pl["recoveries"], pl["recoveries_opp"] = num.get("recoveries_n", 0), num.get("recoveries_d", 0)
            pl["touches_pa"] = num.get("touches_pa", 0)
            pl["yc"], pl["rc"] = num.get("cards_n", 0), num.get("cards_d", 0)
            dd2 = pl.setdefault("duels_detail", {})
            for f in ap.RATIO_FIELDS_B:
                n, d = num.get(f + "_n", 0), num.get(f + "_d", 0)
                dd2[f] = f"{n} / {d}"
                k = "def_pct" if f == "def_duels" else "off_pct" if f == "off_duels" else f + "_pct"
                dd2[k] = f"{round(d / n * 100)}%" if n else "-"
            pp = pl.setdefault("passing", {})
            for f in ap.RATIO_FIELDS_C:
                n, d = num.get(f + "_n", 0), num.get(f + "_d", 0)
                pp[f] = f"{n} / {d}"; pp[f + "_pct"] = round(d / n * 100, 1) if n else 0.0
            pp["deep_completions"] = num.get("deep_completions", 0)
            pp["key_passes"] = num.get("key_passes", 0); pp["shot_assists"] = num.get("shot_assists", 0)
        players.append(pl)

    html = re.sub(r'const PLAYERS = \[.*?\];',
                  'const PLAYERS = ' + json.dumps(players, ensure_ascii=False) + ';', html, flags=re.S)
    n_rounds = dd[dd["Team"] == "Perth"].shape[0]
    win_lo = max(1, n_rounds - 9)
    # rótulo condicional: goleira = temporada; linha = janela
    html = re.sub(r"Rounds \d+–\d+ · últimos 10 jogos do relatório",
                  f"Rounds {win_lo}–{n_rounds} · últimos 10 jogos do relatório", html)
    gk = [p for p in players if p.get("is_goalkeeper")]
    print("   " + " | ".join(f"{p['name']} {p['matches']}j" for p in gk))
    return html


# ─────────────────────────────────────────────────────────────────────────────
# 5) SHOOTING — figura (crop pág.19) + tabela por tipo
# ─────────────────────────────────────────────────────────────────────────────
def update_shooting(html, dd):
    print("5) Shooting (figura + tabela do PDF novo)...")
    img1, img2 = ap.step2b_update_shots_images()
    if img1:
        html = re.sub(r'const FINISHING_IMG1 = "[^"]*";', f'const FINISHING_IMG1 = "{img1}";', html)
    if img2:
        html = re.sub(r'const FINISHING_IMG2 = "[^"]*";', f'const FINISHING_IMG2 = "{img2}";', html)

    if PDF_NEW.exists():
        doc = fitz.open(str(PDF_NEW))
        page_txt = None
        for i in range(len(doc)):
            if doc[i].get_text().startswith("Shots\nGoal\nOn target\nMiss"):
                page_txt = doc[i].get_text(); break
        doc.close()
        if page_txt:
            keys = [("Total", "total"), ("Foot shots", "foot"), ("Head shots", "head"),
                    ("Inside penalty area", "inside"), ("Outside penalty area", "outside"),
                    ("After crosses", "crosses"), ("After set pieces", "setpieces"),
                    ("DFKs and penalties", "dfk")]
            sd = {}
            for label, k in keys:
                m = re.search(re.escape(label) + r"\n(\d+) / (\d+) ([\d.]+)%\n([\d.]+)\n(\d+)", page_txt)
                if m:
                    sd[k] = (int(m.group(1)), int(m.group(2)), float(m.group(3)),
                             float(m.group(4)), int(m.group(5)))
            if len(sd) == 8:
                body = ",\n  ".join(
                    "%s:{shots:%d, on_target:%d, pct:%s, xg:%s, goals:%d}" % (k, *sd[k])
                    for _, k in keys)
                html = re.sub(r'const SD = \{.*?\};', "const SD = {\n  " + body + "\n};", html, count=1, flags=re.S)
                print(f"   Shooting summary: {sd['total'][0]} chutes, {sd['total'][4]} gols")
    return html


# ─────────────────────────────────────────────────────────────────────────────
# 6) NPLW — classificação + rankings (reaproveita extract_wyscout)
# ─────────────────────────────────────────────────────────────────────────────
def update_npl(html, dd):
    print("6) NPLW Comparison / rankings...")
    if not NPL_PDF.exists():
        print("   NPL_Comparison.pdf ausente — pulando.")
        return html
    ap.step2_update_npl_json(dd)
    npl = json.loads(NPL_JSON.read_text(encoding="utf-8"))
    html = re.sub(r"const REPORT_DATA = \{.*?\};\n",
                  "const REPORT_DATA = " + json.dumps(npl, ensure_ascii=False) + ";\n", html, flags=re.S)
    rk = npl.get("playerRankings")
    if rk:
        html = re.sub(r"const NPL_RANKINGS = \{.*?\n(?=\s*const playerRankings = NPL_RANKINGS\[player\.name\];)",
                      "const NPL_RANKINGS = " + json.dumps(rk, ensure_ascii=False) + ";\n  ", html, flags=re.S)
    return html


# ─────────────────────────────────────────────────────────────────────────────
# 7) FIGURAS DA GOLEIRA (dos prints salvos na pasta)
# ─────────────────────────────────────────────────────────────────────────────
def update_gk_images(html):
    print("7) Figuras da goleira (dos prints, se houver)...")
    done = []
    if GK_IMG_SHOTS.exists():
        b = base64.b64encode(GK_IMG_SHOTS.read_bytes()).decode()
        html = re.sub(r'const SCH_SHOTS_AGAINST_IMG = "[^"]*";',
                      f'const SCH_SHOTS_AGAINST_IMG = "{b}";', html)
        done.append(GK_IMG_SHOTS.name)
    if GK_IMG_CROSSES.exists():
        b = base64.b64encode(GK_IMG_CROSSES.read_bytes()).decode()
        html = re.sub(r'const SCH_CONSTRUCTION_IMG = "[^"]*";',
                      f'const SCH_CONSTRUCTION_IMG = "{b}";', html)
        done.append(GK_IMG_CROSSES.name)
    print("   " + (", ".join(done) if done else "nenhum print novo — mantidas as atuais"))
    return html


# ─────────────────────────────────────────────────────────────────────────────
# 8) valida JSON e salva; 9) publica no GitHub
# ─────────────────────────────────────────────────────────────────────────────
def save_and_validate(html):
    for name, pat in [("DATA", r"const DATA = (\[.*?\]);"),
                      ("SD_TOTAL", r"const SD_TOTAL = (\{.*?\});"),
                      ("PLAYERS", r"const PLAYERS = (\[.*?\]);"),
                      ("REPORT_DATA", r"const REPORT_DATA = (\{.*?\});"),
                      ("NPL_RANKINGS", r"const NPL_RANKINGS = (\{.*?\});")]:
        m = re.search(pat, html, re.S)
        if not m:
            raise SystemExit(f"ERRO: constante {name} não encontrada — nada foi salvo.")
        json.loads(m.group(1))
    shutil.copy(HTML, HTML.with_suffix(".bak.html"))
    HTML.write_text(html, encoding="utf-8")
    print("8) HTML validado e salvo (backup em perth_azzurri_painel.bak.html).")


def publish(rnd):
    print("9) Publicando no GitHub...")
    files = ["perth_azzurri_painel.html", "team_stats_perth.xlsx", "npl_comparison_data.json",
             "goalscorers_season.json", "Perth_SC.pdf", "perth_sc_old.pdf", "NPL_Comparison.pdf",
             "D. Schroeder.pdf", "E. Ingrey.pdf", "sch_shots_against.png", "sch_crosses_setpieces.png",
             "atualizar_tudo.py"]
    existing = [f for f in files if (BASE / f).exists()]
    subprocess.run(["git", "add"] + existing, cwd=BASE)
    if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=BASE).returncode == 0:
        print("   Nada novo pra publicar."); return
    subprocess.run(["git", "commit", "-m", f"Atualização semanal — Round {rnd}"], cwd=BASE)
    r = subprocess.run(["git", "push"], cwd=BASE, capture_output=True, text=True)
    print("   Publicado! Site atualiza em 1-2 min." if r.returncode == 0
          else "   AVISO: falha no push — rode 'git push' manualmente.\n   " + r.stderr.strip())


def main():
    if not EXCEL.exists():
        raise SystemExit("team_stats_perth.xlsx não encontrado.")
    dd = rebuild_dashboard_data()
    html = HTML.read_text(encoding="utf-8")
    roster = {p["name"] for p in json.loads(re.search(r'const PLAYERS = (\[.*?\]);', html, re.S).group(1))}

    html, rnd, ld, goals = inject_team_data(html, dd)
    html = update_goalscorers(html, dd, roster)
    html = rebuild_players(html, dd, roster)
    html = update_shooting(html, dd)
    html = update_npl(html, dd)
    html = update_gk_images(html)
    save_and_validate(html)
    publish(rnd)
    print("\n✅ Pronto! Abra perth_azzurri_painel.html pra conferir.")


if __name__ == "__main__":
    main()
