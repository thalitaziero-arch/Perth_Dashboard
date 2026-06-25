#!/usr/bin/env python3
"""
extract_wyscout.py
===================

Extrai os dados do PDF "NPL Comparison" (Season Report da Hudl Wyscout,
Western Australia NPL Women) e gera/atualiza o arquivo `npl_comparison_data.json`
usado pelo painel Perth Azzurri.

USO:
    python extract_wyscout.py caminho/para/o_novo_relatorio.pdf [-o npl_comparison_data.json] [--round "Round 11 · Junho 2026"]

Requisitos:
    pip install pdfplumber --break-system-packages

O script:
  1. Lê o PDF e extrai standings, goals scored/conceded (typology + dynamics),
     formations, attack flanks e todas as métricas de teamStats.
  2. Mostra um diff resumido contra o JSON existente (se houver) para você
     conferir rapidamente o que mudou.
  3. Salva o resultado em npl_comparison_data.json (faz backup do anterior
     como npl_comparison_data.json.bak).

Observação: o parser depende do layout padrão do relatório Wyscout
"Western Australia NPL Women — Season Report". Se a Wyscout mudar o layout
do PDF, alguns valores podem precisar de ajuste manual — o script avisa
quando uma seção não bate com os 8 times esperados.
"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit(
        "Erro: a biblioteca 'pdfplumber' não está instalada.\n"
        "Instale com:  pip install pdfplumber --break-system-packages"
    )

TEAMS = [
    "Perth", "Fremantle City", "Balcatta", "Perth RedStar",
    "West NTC", "Subiaco", "Sorrento", "UWA Nedlands",
]


# ──────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────

def words_by_row(words, y_tol=2.5):
    """Group extracted words into rows based on their vertical position."""
    rows = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        placed = False
        for row in rows:
            if abs(row[0]["top"] - w["top"]) <= y_tol:
                row.append(w)
                placed = True
                break
        if not placed:
            rows.append([w])
    for row in rows:
        row.sort(key=lambda w: w["x0"])
    rows.sort(key=lambda row: row[0]["top"])
    return rows


def split_val_pct(token):
    """
    Split a merged 'value%percent%' token like '382.6974%' into (382.69, 74).
    Wyscout PDFs often merge a decimal value with a following integer
    percentage with no space. Percentages are 1-3 digits; we prefer 2.
    """
    token = token.rstrip("%")
    for plen in (2, 1, 3):
        if len(token) > plen:
            pct_str = token[-plen:]
            val_str = token[:-plen]
            if re.match(r"^\d+\.\d+$", val_str) and re.match(r"^\d+$", pct_str):
                return float(val_str), int(pct_str)
    # fallback: just a plain number, no percent
    try:
        return float(token), None
    except ValueError:
        return None, None


def team_name_from_words(ws, idx, max_words=2):
    """Try to combine up to `max_words` consecutive word tokens into a known team name."""
    for n in range(max_words, 0, -1):
        candidate = " ".join(w["text"] for w in ws[idx:idx + n])
        if candidate in TEAMS:
            return candidate, n
    return None, 1


# ──────────────────────────────────────────────────────────────────────────
# Page parsers
# ──────────────────────────────────────────────────────────────────────────

def parse_standings(page):
    """Page 2 — Standings."""
    text = page.extract_text()
    out = []
    for line in text.splitlines():
        m = re.match(
            r"^(\d+)\s+(.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(-?\d+)\s+"
            r"([\d.]+)[+-][\d.]+\s+([\d.]+)[+-][\d.]+",
            line,
        )
        if not m:
            continue
        pos, team, pts, p, w, d, l, gf, ga, gd, xg, xga = m.groups()
        team = team.strip()
        # Bold rendering in the PDF can double/glue digits (e.g. "30" -> "3030",
        # or "7" -> "77"). Points are deterministic from the standard scoring
        # rule (3 per win, 1 per draw), so recompute instead of trusting the
        # raw extracted text — far more robust than pattern-matching the glitch.
        pts_clean = 3 * int(w) + int(d)
        out.append({
            "pos": int(pos),
            "team": team,
            "pts": pts_clean,
            "p": int(p),
            "w": int(w),
            "d": int(d),
            "l": int(l),
            "gf": int(gf),
            "ga": int(ga),
            "gd": int(gd),
            "xg": float(xg),
            "xga": float(xga),
        })
    return out


def parse_goals_typology(page, scored=True):
    """Pages 3/4 — Goals scored / conceded typology (rows above the 'match dynamics' header)."""
    words = page.extract_words()
    # find the y-position of the dynamics section header to cut off typology rows
    cutoff = None
    for w in words:
        if w["text"] in ("dynamics",):
            cutoff = w["top"]
            break
    rows = words_by_row(words)
    out = []
    for row in rows:
        if cutoff is not None and row[0]["top"] >= cutoff:
            continue
        ws = row
        team, used = team_name_from_words(ws, 0)
        if not team:
            continue
        rest = ws[used:]
        nums = []
        i = 0
        while i < len(rest):
            tok = rest[i]["text"]
            m = re.match(r"^(\d+)$", tok)
            if m:
                # check if next token is a (%) value belonging to this number
                pct = None
                if i + 1 < len(rest) and re.match(r"^\(\d+%\)$", rest[i + 1]["text"]):
                    pct = int(re.match(r"^\((\d+)%\)$", rest[i + 1]["text"]).group(1))
                    i += 1
                nums.append((int(tok), pct))
            i += 1

        if len(nums) < 5:
            continue
        total = nums[0][0]
        long_shots, long_pct = nums[1]
        counter, counter_pct = nums[2]
        set_piece, set_pct = nums[3]
        direct_fk, fk_pct = nums[4]

        entry = {
            "team": team,
            "total": total,
            "longShots": long_shots, "longShotsPct": long_pct or 0,
            "counter": counter, "counterPct": counter_pct or 0,
            "setPiece": set_piece, "setPiecePct": set_pct or 0,
            "directFK": direct_fk, "directFKPct": fk_pct or 0,
        }
        out.append(entry)
    return out


def parse_goals_dynamics(page, xg_key="xg"):
    """Pages 3/4 — Goals scored/conceded in match dynamics.
    xg_key: 'xg' for goals scored page, 'xga' for goals conceded page."""
    words = page.extract_words()
    rows = words_by_row(words)
    out = []
    i = 0
    while i < len(rows):
        row = rows[i]
        team, used = team_name_from_words(row, 0)
        if not team:
            i += 1
            continue
        nums_row = [w["text"] for w in row]
        ints = [t for t in nums_row if re.match(r"^\d+$", t)]
        if len(ints) < 9:
            i += 1
            continue
        total, h1t, h2t, t1, t2, t3, t4, t5, t6 = (int(x) for x in ints[-9:])

        # xG breakdown is 2 rows below: e.g. ['4.27', 'xG', '1.95', '2.32', ...]
        xg_total = 0.0
        if i + 2 < len(rows):
            tokens = [w["text"] for w in rows[i + 2]]
            floats = [t for t in tokens if re.match(r"^\d+\.\d+$", t)]
            if floats:
                xg_total = float(floats[0])

        entry = {
            "team": team,
            "total": total,
            "h1": h1t, "h2": h2t,
            "t1": t1, "t2": t2, "t3": t3, "t4": t4, "t5": t5, "t6": t6,
        }
        entry[xg_key] = xg_total
        out.append(entry)
        i += 3
    return out


def parse_formations(page):
    """Page 5 — Formations table. Each row contains two teams (left & right column)."""
    words = page.extract_words()
    rows = words_by_row(words)
    out = []
    for row in rows:
        toks = [w["text"] for w in row]
        # find all team-name matches (1 or 2 word tokens) at the start of a segment
        i = 0
        while i < len(toks):
            team, used = team_name_from_words(row, i)
            if team:
                rest_tokens = toks[i + used:]
                # collect up to 3 (formation, pct) pairs immediately following
                pairs = []
                j = 0
                while j < len(rest_tokens) - 1 and len(pairs) < 3:
                    f, p = rest_tokens[j], rest_tokens[j + 1]
                    if re.match(r"^\d", f) and re.match(r"^[\d.]+%$", p):
                        pairs.append((f, float(p.rstrip("%"))))
                        j += 2
                    else:
                        break
                if len(pairs) >= 2:
                    out.append({
                        "team": team,
                        "f1": pairs[0][0], "p1": pairs[0][1],
                        "f2": pairs[1][0], "p2": pairs[1][1],
                    })
                i += used + j
            else:
                i += 1
    return out


def parse_attack_flanks(page):
    """Page 5 — Attacks by flanks. Two bands of 12 percentages = 4 teams x 3 (L/C/R) per band."""
    words = page.extract_words()
    pct_words = [w for w in words if re.match(r"^\d+%$", w["text"]) and w["top"] > 340]
    if len(pct_words) != 24:
        return []

    # split into two bands by top position (band1 ~350-370, band2 ~388-410)
    pct_words.sort(key=lambda w: w["top"])
    band1 = sorted(pct_words[:12], key=lambda w: w["x0"])
    band2 = sorted(pct_words[12:], key=lambda w: w["x0"])

    row1_teams = ["West NTC", "Balcatta", "UWA Nedlands", "Fremantle City"]
    row2_teams = ["Subiaco", "Sorrento", "Perth RedStar", "Perth"]

    out = []
    for teams, band in ((row1_teams, band1), (row2_teams, band2)):
        pcts = [int(w["text"].rstrip("%")) for w in band]
        for idx, team in enumerate(teams):
            left, centre, right = pcts[idx * 3:idx * 3 + 3]
            out.append({"team": team, "left": left, "centre": centre, "right": right})
    return out


def parse_page6(page):
    """Page 6 — Goals/Conceded goals (top) and Shots/Shots against (bottom)."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:1]] == ["Shots"])
    bottom_rows = rows[split_idx:]

    def two_cols(rows_slice, right_vals=3):
        """Each data row: '<rank> team... val [val2 ...] <rank> team... val [val2 ...]'."""
        left, right = [], []
        for row in rows_slice:
            toks = [w["text"] for w in row]
            if not toks or not re.match(r"^\d+$", toks[0]):
                continue
            rank = int(toks[0])
            if rank < 1 or rank > 8:
                continue
            team, used = team_name_from_words(row, 1)
            if not team:
                continue
            idx = 1 + used
            vals = []
            while idx < len(toks) and re.match(r"^-?\d+\.?\d*$", toks[idx]) and len(vals) < right_vals:
                vals.append(float(toks[idx]))
                idx += 1
            if vals:
                entry = {"team": team, "val": vals[0]}
                if len(vals) >= 2:
                    entry["xg" if right_vals == 1 else "distance"] = vals[1]
                if len(vals) >= 3:
                    entry["xgPerShot"] = vals[2]
                left.append(entry)
            # right column
            if idx < len(toks) and re.match(r"^\d+$", toks[idx]) and int(toks[idx]) == rank:
                team2, used2 = team_name_from_words(row, idx + 1)
                if team2:
                    idx2 = idx + 1 + used2
                    vals2 = []
                    while idx2 < len(toks) and re.match(r"^-?\d+\.?\d*$", toks[idx2]) and len(vals2) < right_vals:
                        vals2.append(float(toks[idx2]))
                        idx2 += 1
                    if vals2:
                        entry2 = {"team": team2, "val": vals2[0]}
                        if len(vals2) >= 2:
                            entry2["xg" if right_vals == 1 else "distance"] = vals2[1]
                        if len(vals2) >= 3:
                            entry2["xgPerShot"] = vals2[2]
                        right.append(entry2)
        return left, right

    top_rows = rows[:split_idx]
    goals, conceded = two_cols(top_rows, right_vals=1)
    shots, shots_against = two_cols(bottom_rows, right_vals=3)
    return goals, conceded, shots, shots_against


def parse_ranked_rows(rows, start_idx, end_idx, value_col_finder):
    """
    Parse data rows of the form ['<rank>', <team words...>, <value tokens...>].
    `value_col_finder(rest_tokens)` should return the numeric value (float)
    for this row's left-hand list, given the tokens after the team name.
    Returns a list of {"team":..., "val":...} for ranks 1-8.
    """
    out = []
    for row in rows[start_idx:end_idx]:
        toks = [w["text"] for w in row]
        if not toks or not re.match(r"^\d+$", toks[0]):
            continue
        rank = int(toks[0])
        if rank < 1 or rank > 8:
            continue
        team, used = team_name_from_words(row, 1)
        if not team:
            continue
        rest = toks[1 + used:]
        val = value_col_finder(rest)
        if val is not None:
            out.append({"team": team, "val": val})
    return out


def first_value(tok):
    """Parse a token as either a plain float or a merged 'val%pct%' -> returns val."""
    if tok.endswith("%"):
        val, pct = split_val_pct(tok)
        return val
    m = re.match(r"^-?\d+\.?\d*$", tok)
    return float(m.group(0)) if m else None


def parse_page8(page):
    """Page 8 — Passes (avg/90, merged %) and Ball possession (plain %)."""
    words = page.extract_words()
    rows = words_by_row(words)
    # find split point: row starting with 'Ball','possession'
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:2]] == ["Ball", "possession"])
    passes_rows = rows[:split_idx]
    possession_rows = rows[split_idx:]

    passes = parse_ranked_rows(passes_rows, 0, len(passes_rows), lambda rest: first_value(rest[0]) if rest else None)
    possession = parse_ranked_rows(
        possession_rows, 0, len(possession_rows),
        lambda rest: float(rest[0].rstrip("%")) if rest and rest[0].endswith("%") else None
    )
    return passes, possession


def parse_page9(page):
    """Page 9 — Passes to final third / Deep completions, Progressive passes / Through passes."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:2]] == ["Progressive", "passes"])
    top_rows = rows[:split_idx]
    bottom_rows = rows[split_idx:]

    def left_right(rows_slice):
        left, right = [], []
        for row in rows_slice:
            toks = [w["text"] for w in row]
            if not toks or not re.match(r"^\d+$", toks[0]):
                continue
            rank = int(toks[0])
            if rank < 1 or rank > 8:
                continue
            # left team starts at idx 1
            team, used = team_name_from_words(row, 1)
            if not team:
                continue
            idx_after_left_team = 1 + used
            val_left = first_value(toks[idx_after_left_team])
            if val_left is not None:
                left.append({"team": team, "val": val_left})
            # find next rank number (start of right column) after the left value
            j = idx_after_left_team + 1
            while j < len(toks) and not re.match(r"^\d+$", toks[j]):
                j += 1
            if j < len(toks) and int(toks[j]) == rank:
                team2, used2 = team_name_from_words(row, j + 1)
                if team2:
                    idx_after_right_team = j + 1 + used2
                    if idx_after_right_team < len(toks):
                        val_right = first_value(toks[idx_after_right_team])
                        if val_right is not None:
                            right.append({"team": team2, "val": val_right})
        return left, right

    passes_final_third, deep_completions = left_right(top_rows)
    progressive, through_passes = left_right(bottom_rows)
    return progressive, deep_completions


def parse_page10(page):
    """Page 10 — Crosses / Dribbles, Touches in penalty area / Fouls suffered."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:4]] == ["Touches", "in", "penalty", "area"])
    top_rows = rows[:split_idx]
    bottom_rows = rows[split_idx:]

    crosses, dribbles = [], []
    for row in top_rows:
        toks = [w["text"] for w in row]
        if not toks or not re.match(r"^\d+$", toks[0]):
            continue
        rank = int(toks[0])
        if rank < 1 or rank > 8:
            continue
        team, used = team_name_from_words(row, 1)
        if not team:
            continue
        idx = 1 + used
        # crosses: merged val%, then two more merged tokens (→ ←), then rank2 + team2 + dribbles val
        val_left = first_value(toks[idx])
        if val_left is not None:
            crosses.append({"team": team, "val": val_left})
        # advance past the → ← merged tokens to find rank2
        j = idx + 1
        while j < len(toks) and not re.match(r"^\d+$", toks[j]):
            j += 1
        if j < len(toks) and int(toks[j]) == rank:
            team2, used2 = team_name_from_words(row, j + 1)
            if team2:
                idx2 = j + 1 + used2
                if idx2 < len(toks):
                    val_right = first_value(toks[idx2])
                    if val_right is not None:
                        dribbles.append({"team": team2, "val": val_right})

    touches_pa = []
    for row in bottom_rows:
        toks = [w["text"] for w in row]
        if not toks or not re.match(r"^\d+$", toks[0]):
            continue
        rank = int(toks[0])
        if rank < 1 or rank > 8:
            continue
        team, used = team_name_from_words(row, 1)
        if not team:
            continue
        idx = 1 + used
        val = first_value(toks[idx]) if idx < len(toks) else None
        if val is not None:
            touches_pa.append({"team": team, "val": val})

    return crosses, dribbles, touches_pa


def parse_page11(page):
    """Page 11 — Offensive duels / Defensive duels, Aerial duels / Loose ball duels."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:2]] == ["Aerial", "duels"])
    top_rows = rows[:split_idx]
    bottom_rows = rows[split_idx:]

    def left_right(rows_slice):
        left, right = [], []
        for row in rows_slice:
            toks = [w["text"] for w in row]
            if not toks or not re.match(r"^\d+$", toks[0]):
                continue
            rank = int(toks[0])
            if rank < 1 or rank > 8:
                continue
            team, used = team_name_from_words(row, 1)
            if not team:
                continue
            idx = 1 + used
            val_left = first_value(toks[idx]) if idx < len(toks) else None
            if val_left is not None:
                left.append({"team": team, "val": val_left})
            j = idx + 1
            while j < len(toks) and not re.match(r"^\d+$", toks[j]):
                j += 1
            if j < len(toks) and int(toks[j]) == rank:
                team2, used2 = team_name_from_words(row, j + 1)
                if team2:
                    idx2 = j + 1 + used2
                    if idx2 < len(toks):
                        val_right = first_value(toks[idx2])
                        if val_right is not None:
                            right.append({"team": team2, "val": val_right})
        return left, right

    off_duels, def_duels = left_right(top_rows)
    aerial_duels, loose_ball = left_right(bottom_rows)
    return off_duels, def_duels, aerial_duels


def parse_page12(page):
    """Page 12 — Interceptions / Pressing intensity (PPDA)."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:2]] == ["Shots", "blocked"])
    top_rows = rows[:split_idx]

    interceptions, ppda = [], []
    for row in top_rows:
        toks = [w["text"] for w in row]
        if not toks or not re.match(r"^\d+$", toks[0]):
            continue
        rank = int(toks[0])
        if rank < 1 or rank > 8:
            continue
        team, used = team_name_from_words(row, 1)
        if not team:
            continue
        idx = 1 + used
        val_left = first_value(toks[idx]) if idx < len(toks) else None
        if val_left is not None:
            interceptions.append({"team": team, "val": val_left})
        j = idx + 1
        if j < len(toks) and re.match(r"^\d+$", toks[j]) and int(toks[j]) == rank:
            team2, used2 = team_name_from_words(row, j + 1)
            if team2:
                idx2 = j + 1 + used2
                if idx2 < len(toks):
                    val_right = first_value(toks[idx2])
                    if val_right is not None:
                        ppda.append({"team": team2, "val": val_right})
    return interceptions, ppda


def parse_page7(page):
    """Page 7 — Losses / Recoveries (avg/90, first value before zone breakdown)."""
    words = page.extract_words()
    rows = words_by_row(words)
    split_idx = next(i for i, row in enumerate(rows) if [w["text"] for w in row[:1]] == ["Recoveries"])
    losses_rows = rows[:split_idx]
    recoveries_rows = rows[split_idx:]

    def first_col(rows_slice):
        out = []
        for row in rows_slice:
            toks = [w["text"] for w in row]
            if not toks or not re.match(r"^\d+$", toks[0]):
                continue
            rank = int(toks[0])
            if rank < 1 or rank > 8:
                continue
            team, used = team_name_from_words(row, 1)
            if not team:
                continue
            idx = 1 + used
            val = first_value(toks[idx]) if idx < len(toks) else None
            if val is not None:
                out.append({"team": team, "val": val})
        return out

    return first_col(losses_rows), first_col(recoveries_rows)


# ──────────────────────────────────────────────────────────────────────────
# Perth_SC.pdf — per-player season stats (Team Report, pages 2-4)
# ──────────────────────────────────────────────────────────────────────────
#
# IMPORTANT: this report must be generated in Wyscout with the date/match
# filter set to "all matches" / full season — NOT "last 5 matches" (the
# default). The parser below replaces season totals outright; it does not
# accumulate, so if the PDF only covers a subset of games the numbers will
# be wrong.

TABLE_A_FIELDS = [
    "total_min", "goals_xg", "assists_xa", "shots", "passes", "crosses",
    "dribbles", "duels", "losses", "recoveries", "touches_pa", "cards",
]
TABLE_B_FIELDS = [
    "total_min", "def_duels", "off_duels", "aerial", "loose", "blocks",
    "interceptions_clear", "tackles", "fouls", "freekicks", "setpieces",
    "direct_fk_corners", "corners_served", "throwins",
]
TABLE_C_FIELDS = [
    "total_min", "forward", "back", "lateral", "short_med", "long",
    "progressive", "final_third", "through", "deep_completions",
    "key_passes", "second_third_assists", "shot_assists",
]


def _split_player_blocks(lines, player_names):
    """Split a page's lines into per-player value blocks, using known
    player names as block boundaries (more robust than guessing field
    counts, since Wyscout sometimes omits a trailing dash for empty cells)."""
    name_idx = [i for i, l in enumerate(lines) if l.strip() in player_names]
    blocks = {}
    for j, i in enumerate(name_idx):
        name = lines[i].strip()
        end = name_idx[j + 1] - 1 if j + 1 < len(name_idx) else len(lines)
        # the line right before the next name is that name's leading number;
        # don't include it in the value block
        blocks[name] = [l.strip() for l in lines[i + 1:end]]
    return blocks


def _map_fields(values, field_names):
    """Map a variable-length value list onto field_names. Wyscout sometimes
    drops a trailing empty cell instead of printing '-', so pad on the right."""
    values = list(values) + ["-"] * (len(field_names) - len(values))
    return dict(zip(field_names, values[: len(field_names)]))


def _table_segments(lines):
    """Wyscout repeats a 'Player\\n...column headers...' footer after each
    table on the page. Split the page into per-table line ranges using
    those footers as separators, so a player name that appears in two
    different tables on the same page isn't mixed up."""
    footer_starts = [i for i, l in enumerate(lines) if l.strip() == "Player"]
    bounds = []
    seg_start = 0
    for f in footer_starts:
        bounds.append((seg_start, f))
        # next table's data starts at the first purely-numeric line after the footer
        j = f + 1
        while j < len(lines) and not lines[j].strip().isdigit():
            j += 1
        seg_start = j
    if seg_start < len(lines):
        bounds.append((seg_start, len(lines)))
    return bounds


def parse_perth_sc_overview(pdf_path, player_names):
    """Page 2 of Perth_SC.pdf: one row per player with Position, Age, Foot,
    Matches(total), Minutes(total/avg), Goals, Assists, Cards, Subs. The
    leading squad number is sometimes merged onto the name line (e.g.
    '23 E. Ingrey') and sometimes on its own line — handle both forms.
    Returns {name: matches_total (int)}."""
    import fitz
    doc = fitz.open(str(pdf_path))
    lines = doc[1].get_text().splitlines()
    doc.close()

    idxs = []
    for i, l in enumerate(lines):
        l2 = l.strip()
        m = re.match(r"^\d+\s+(.+)$", l2)
        cand = m.group(1) if m else l2
        if cand in player_names:
            idxs.append((i, cand))

    out = {}
    for j, (i, name) in enumerate(idxs):
        end = idxs[j + 1][0] if j + 1 < len(idxs) else len(lines)
        vals = [l.strip() for l in lines[i + 1:end]]
        # matches(total) is the integer token right before the first token
        # ending in "'" (total minutes)
        for k, v in enumerate(vals):
            if v.endswith("'") and k > 0 and vals[k - 1].isdigit():
                out[name] = int(vals[k - 1])
                break
    return out


def parse_perth_sc_pdf(pdf_path, player_names):
    """Parses Perth_SC.pdf (Wyscout Team Report) pages 2-4 into per-player
    season-stat dicts. `player_names` should be the full current roster
    (e.g. names already known from the dashboard) used to find row boundaries."""
    import fitz
    doc = fitz.open(str(pdf_path))
    # Tables A (season summary), B (duel details) and C (passing details) are
    # always pages 3-4 of the report, in that order — but depending on roster
    # size they may all fit on page 3, or spill onto page 4. Concatenate both
    # pages' lines and split into table segments generically, so either
    # layout works.
    combined_lines = doc[2].get_text().splitlines() + doc[3].get_text().splitlines()
    doc.close()
    segs = _table_segments(combined_lines)
    table_a = _split_player_blocks(combined_lines[segs[0][0]:segs[0][1]], player_names) if len(segs) > 0 else {}
    table_b = _split_player_blocks(combined_lines[segs[1][0]:segs[1][1]], player_names) if len(segs) > 1 else {}
    table_c = _split_player_blocks(combined_lines[segs[2][0]:segs[2][1]], player_names) if len(segs) > 2 else {}

    out = {}
    for name in player_names:
        rec = {}
        if name in table_a:
            rec["A"] = _map_fields(table_a[name], TABLE_A_FIELDS)
        if name in table_b:
            rec["B"] = _map_fields(table_b[name], TABLE_B_FIELDS)
        if name in table_c:
            rec["C"] = _map_fields(table_c[name], TABLE_C_FIELDS)
        if rec:
            out[name] = rec
    return out


# ──────────────────────────────────────────────────────────────────────────
# Player rankings (pages 14-18: "Players: Overview / Attack / Defence / Set pieces")
# ──────────────────────────────────────────────────────────────────────────

THREE_VAL_CATEGORIES = {"Goals", "Assists", "Shots", "Second assists", "Fouls"}
TEAMS_BY_LEN = sorted(TEAMS, key=len, reverse=True)


def _split_name_team(combined):
    for t in TEAMS_BY_LEN:
        if combined.endswith(t):
            return combined[: -len(t)].strip(), t
    return combined, None


def parse_player_rankings(pages):
    """pages: list of page.extract_text() strings for pages 14-18 (idx 13-17)."""
    all_lines = []
    for text in pages:
        all_lines.extend(text.splitlines())

    # find category boundaries: a line is a category title if the *next*
    # line starts with '↓' (the column header marker).
    titles = []
    for i, line in enumerate(all_lines):
        if i + 1 < len(all_lines) and all_lines[i + 1].startswith("↓"):
            titles.append(i)

    records = []
    for ti, title_idx in enumerate(titles):
        category = all_lines[title_idx].strip()
        end = titles[ti + 1] if ti + 1 < len(titles) else len(all_lines)
        value_count = 3 if category in THREE_VAL_CATEGORIES else 1
        i = title_idx + 1
        rank = 0
        while i < end:
            line = all_lines[i].strip()
            if i + 1 < end and all_lines[i + 1].strip() in TEAMS:
                name, team = line, all_lines[i + 1].strip()
                vals = [all_lines[i + 2 + k].strip() for k in range(value_count)]
                rank += 1
                records.append({"category": category, "rank": rank, "name": name,
                                 "team": team, "values": vals})
                i += 2 + value_count
            elif line.isdigit() and 4 <= int(line) <= 10 and i + 1 < end:
                name, team = _split_name_team(all_lines[i + 1].strip())
                if team:
                    vals = [all_lines[i + 2 + k].strip() for k in range(value_count)]
                    rank = int(line)
                    records.append({"category": category, "rank": rank, "name": name,
                                     "team": team, "values": vals})
                    i += 2 + value_count
                    continue
                i += 1
            else:
                i += 1
    return records


CATEGORY_LABELS = {
    "Goals": "xG", "Assists": "xA", "Shots": "xG/Shot",
    "Second assists": "3rd assists",
}


TOTAL_UNIT_CATEGORIES = {"Direct free kicks", "Penalties"}


def build_npl_rankings(records):
    """Group Perth players' top-10 appearances into the NPL_RANKINGS shape
    used by the dashboard: {player_name: [{category, rank, value, unit}, ...]}."""
    out = {}
    for r in records:
        if r["team"] != "Perth":
            continue
        cat, vals = r["category"], r["values"]
        if cat == "Fouls":
            value, unit = f"{vals[0]} avg/90", f"YC {vals[1]} · RC {vals[2]}"
        elif cat in THREE_VAL_CATEGORIES:
            label = CATEGORY_LABELS.get(cat, "")
            value = f"{vals[0]} total"
            unit = f"{vals[1]} avg/90" + (f" · {label} {vals[2]}" if label else "")
        else:
            value = vals[0]
            unit = "total" if cat in TOTAL_UNIT_CATEGORIES else "avg/90"
        out.setdefault(r["name"], []).append({
            "category": cat, "rank": r["rank"], "value": value, "unit": unit,
        })
    return out


# ──────────────────────────────────────────────────────────────────────────
# Main extraction
# ──────────────────────────────────────────────────────────────────────────

def extract(pdf_path, round_label=None):
    data = {}
    with pdfplumber.open(pdf_path) as pdf:
        data["lastUpdated"] = round_label or "Round ? · " + "____"

        data["standings"] = parse_standings(pdf.pages[1])

        data["goalsScoredTypology"] = parse_goals_typology(pdf.pages[2], scored=True)
        data["goalsScoredDynamics"] = parse_goals_dynamics(pdf.pages[2], xg_key="xg")

        data["goalsConcededTypology"] = parse_goals_typology(pdf.pages[3], scored=False)
        data["goalsConcededDynamics"] = parse_goals_dynamics(pdf.pages[3], xg_key="xga")

        data["formations"] = parse_formations(pdf.pages[4])
        data["attackFlanks"] = parse_attack_flanks(pdf.pages[4])

        # Page 6 — Goals/Conceded goals (totals + xG), Shots/Shots against (avg + distance + xgPerShot)
        goals, conceded, shots, shots_against = parse_page6(pdf.pages[5])

        team_stats = {}
        team_stats["shots"] = [{"team": d["team"], "val": d["val"]} for d in shots]
        team_stats["shotsAgainst"] = [{"team": d["team"], "val": d["val"]} for d in shots_against]
        team_stats["shotsXG"] = [{"team": d["team"], "val": d["xgPerShot"]} for d in shots]

        # Page 8 — possession, passes
        passes, possession = parse_page8(pdf.pages[7])
        team_stats["passes"] = passes
        team_stats["possession"] = possession

        # Page 9 — progressive passes, deep completions
        progressive, deep_comp = parse_page9(pdf.pages[8])
        team_stats["progressive"] = progressive
        team_stats["deepComp"] = deep_comp

        # Page 10 — crosses, dribbles, touchesPA
        crosses, dribbles, touches_pa = parse_page10(pdf.pages[9])
        team_stats["crosses"] = crosses
        team_stats["dribbles"] = dribbles
        team_stats["touchesPA"] = touches_pa

        # Page 11 — offensive/defensive/aerial duels
        off_duels, def_duels, aerial_duels = parse_page11(pdf.pages[10])
        team_stats["offDuels"] = off_duels
        team_stats["defDuels"] = def_duels
        team_stats["aerialDuels"] = aerial_duels

        # Page 12 — interceptions, ppda
        interceptions, ppda = parse_page12(pdf.pages[11])
        team_stats["interceptions"] = interceptions
        team_stats["ppda"] = ppda

        # Page 7 — losses, recoveries
        losses, recoveries = parse_page7(pdf.pages[6])
        team_stats["losses"] = losses
        team_stats["recoveries"] = recoveries

        data["teamStats"] = team_stats

    # Pages 14-18 — player rankings. These pages lay out two ranking tables
    # side by side; pdfplumber's reading order interleaves them, so we use
    # pymupdf (fitz) here instead, which preserves the correct sequence.
    try:
        import fitz
        doc = fitz.open(pdf_path)
        if len(doc) >= 18:
            pages_text = [doc[i].get_text() for i in range(13, 18)]
            records = parse_player_rankings(pages_text)
            data["playerRankings"] = build_npl_rankings(records)
        doc.close()
    except ImportError:
        print("Aviso: pymupdf não instalado, pulando extração de player rankings "
              "(pip install pymupdf --break-system-packages).")

    return data


def summarize_section(name, value):
    if isinstance(value, list):
        return f"{name}: {len(value)} times"
    if isinstance(value, dict):
        parts = ", ".join(f"{k}={len(v)}" for k, v in value.items())
        return f"{name}: {{{parts}}}"
    return f"{name}: {value}"


def main():
    parser = argparse.ArgumentParser(description="Extrai dados do PDF NPL Comparison para JSON")
    parser.add_argument("pdf", help="Caminho do PDF 'NPL Comparison' (season report)")
    parser.add_argument("-o", "--output", default="npl_comparison_data.json",
                         help="Arquivo JSON de saída (padrão: npl_comparison_data.json)")
    parser.add_argument("--round", default=None,
                         help='Texto de "lastUpdated", ex: "Round 11 · Junho 2026"')
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"Arquivo não encontrado: {pdf_path}")

    print(f"Lendo {pdf_path} ...")
    data = extract(pdf_path, round_label=args.round)

    print("\nResumo do que foi extraído:")
    for k, v in data.items():
        print("  -", summarize_section(k, v))

    expected_8 = set(TEAMS)
    for key in ("standings", "goalsScoredTypology", "goalsConcededTypology", "formations", "attackFlanks"):
        teams_found = {row["team"] for row in data.get(key, [])}
        if teams_found != expected_8:
            missing = expected_8 - teams_found
            extra = teams_found - expected_8
            print(f"\n⚠️  Aviso: seção '{key}' não tem os 8 times esperados.")
            if missing:
                print(f"   Faltando: {sorted(missing)}")
            if extra:
                print(f"   Inesperado: {sorted(extra)}")

    for key, lst in data.get("teamStats", {}).items():
        if len(lst) != 8:
            print(f"\n⚠️  Aviso: teamStats.{key} tem {len(lst)} entradas (esperado 8). Verifique manualmente.")

    out_path = Path(args.output)
    if out_path.exists():
        backup = out_path.with_suffix(out_path.suffix + ".bak")
        shutil.copy(out_path, backup)
        print(f"\nBackup do JSON anterior salvo em: {backup}")

        with open(out_path, "r", encoding="utf-8") as f:
            old_data = json.load(f)
        print("\nDiferenças no 'lastUpdated':")
        print(f"  antes: {old_data.get('lastUpdated')}")
        print(f"  agora: {data.get('lastUpdated')}")

        old_standings = {r["team"]: r for r in old_data.get("standings", [])}
        print("\nDiferenças na classificação (pts):")
        for row in data.get("standings", []):
            old = old_standings.get(row["team"], {})
            if old.get("pts") != row["pts"]:
                print(f"  {row['team']}: {old.get('pts')} -> {row['pts']}")

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"\nSalvo em: {out_path}")
    print("\nIMPORTANTE: confira o arquivo gerado antes de subir para o servidor,")
    print("especialmente as seções de teamStats (possession, passes, duels, etc.),")
    print("pois alguns valores no PDF aparecem 'colados' e o parser usa heurísticas")
    print("para separá-los corretamente.")


if __name__ == "__main__":
    main()
