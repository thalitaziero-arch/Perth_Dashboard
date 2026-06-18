"""
Perth Azzurri — Atualização Semanal
Rode com: streamlit run update_weekly.py
"""

from pathlib import Path
import base64
import shutil

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Atualizar Dados — Perth Azzurri", layout="wide")
st.title("⚽ Perth Azzurri — Atualização Semanal")

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "team_stats_perth.xlsx"
IMAGES_DIR = BASE_DIR / "images"
IMAGES_DIR.mkdir(exist_ok=True)

# ─── carregar dados ────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    return pd.read_excel(EXCEL_FILE, sheet_name="DashboardData")

df = load_data()

tab1, tab2, tab3 = st.tabs(["➕ Novo Round", "🖼️ Imagens / Figuras", "📋 Ver / Apagar Dados"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — NOVO ROUND
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Adicionar estatísticas do novo round")
    st.info("Preencha os dados do Perth e do adversário separadamente. Os campos de % são calculados automaticamente.")

    next_round = int(df["Round"].max()) + 1 if not df.empty else 1

    col_a, col_b = st.columns(2)
    with col_a:
        round_num  = st.number_input("Round", value=next_round, min_value=1, step=1)
        match_date = st.date_input("Data do jogo")
        competition = st.text_input("Competição", value="Australia. Western Australia NPL Women")
        duration   = st.number_input("Duração (min)", value=90, min_value=1, step=1)
    with col_b:
        team_home  = st.text_input("Time 1 (geralmente Perth)", value="Perth")
        team_away  = st.text_input("Time 2 (adversário)")
        goals_home = st.number_input("Gols Time 1", value=0, min_value=0, step=1)
        goals_away = st.number_input("Gols Time 2", value=0, min_value=0, step=1)

    match_str = f"{team_home} - {team_away} {goals_home}:{goals_away}"
    st.markdown(f"**Jogo:** `{match_str}`")

    st.divider()

    def team_form(label: str, key: str):
        st.markdown(f"### {label}")
        c1, c2, c3 = st.columns(3)
        with c1:
            scheme  = st.text_input("Esquema", key=f"scheme_{key}", placeholder="4-3-3")
            shots   = st.number_input("Finalizações", key=f"shots_{key}", value=0, min_value=0)
            sot     = st.number_input("No alvo", key=f"sot_{key}", value=0, min_value=0)
            xg      = st.number_input("xG", key=f"xg_{key}", value=0.0, min_value=0.0, step=0.01, format="%.2f")
        with c2:
            passes  = st.number_input("Passes", key=f"passes_{key}", value=0, min_value=0)
            passes_acc = st.number_input("Passes precisos", key=f"passes_acc_{key}", value=0, min_value=0)
            possession = st.number_input("Posse (%)", key=f"poss_{key}", value=50.0, min_value=0.0, max_value=100.0, step=0.1, format="%.2f")
        with c3:
            losses_low  = st.number_input("Perdas baixo", key=f"ll_{key}", value=0, min_value=0)
            losses_mid  = st.number_input("Perdas médio", key=f"lm_{key}", value=0, min_value=0)
            losses_high = st.number_input("Perdas alto", key=f"lh_{key}", value=0, min_value=0)
            rec_low     = st.number_input("Recup. baixo", key=f"rl_{key}", value=0, min_value=0)
            rec_mid     = st.number_input("Recup. médio", key=f"rm_{key}", value=0, min_value=0)
            rec_high    = st.number_input("Recup. alto", key=f"rh_{key}", value=0, min_value=0)
            duels       = st.number_input("Duelos total", key=f"dt_{key}", value=0, min_value=0)
            duels_won   = st.number_input("Duelos ganhos", key=f"dw_{key}", value=0, min_value=0)

        losses_total = losses_low + losses_mid + losses_high
        rec_total    = rec_low + rec_mid + rec_high
        shots_pct    = round(sot / shots * 100, 2) if shots > 0 else 0.0
        passes_pct   = round(passes_acc / passes * 100, 2) if passes > 0 else 0.0
        duels_pct    = round(duels_won / duels * 100, 2) if duels > 0 else 0.0

        return {
            "Round": round_num,
            "Date": str(match_date),
            "Match": match_str,
            "Competition": competition,
            "Duration": duration,
            "Team": label,
            "Scheme": scheme,
            "Goals": goals_home if key == "home" else goals_away,
            "xG": xg,
            "Shots": shots,
            "Shots_on_target": sot,
            "Shots_acc_pct": shots_pct,
            "Passes": passes,
            "Passes_accurate": passes_acc,
            "Passes_acc_pct": passes_pct,
            "Possession_pct": possession,
            "Losses_total": losses_total,
            "Losses_low": losses_low,
            "Losses_medium": losses_mid,
            "Losses_high": losses_high,
            "Recoveries_total": rec_total,
            "Recoveries_low": rec_low,
            "Recoveries_medium": rec_mid,
            "Recoveries_high": rec_high,
            "Duels_total": duels,
            "Duels_won": duels_won,
            "Duels_won_pct": duels_pct,
        }

    col_home, col_away = st.columns(2)
    with col_home:
        row_home = team_form(team_home or "Time 1", "home")
    with col_away:
        row_away = team_form(team_away or "Time 2", "away")

    st.divider()
    if st.button("💾 Salvar Round no Excel", type="primary", use_container_width=True):
        if not team_away.strip():
            st.error("Preencha o nome do adversário.")
        else:
            new_rows = pd.DataFrame([row_home, row_away])
            # backup
            shutil.copy(EXCEL_FILE, EXCEL_FILE.with_suffix(".bak.xlsx"))

            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a",
                                if_sheet_exists="overlay") as writer:
                combined = pd.concat([df, new_rows], ignore_index=True)
                combined.to_excel(writer, sheet_name="DashboardData", index=False)

            st.success(f"Round {round_num} salvo com sucesso! Backup em team_stats_perth.bak.xlsx")
            st.cache_data.clear()
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — IMAGENS / FIGURAS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Upload de imagens e figuras do Wyscout")
    st.info("Arraste as imagens ou clique para selecionar. Elas são salvas na pasta `images/` com o nome que você escolher.")

    uploaded = st.file_uploader(
        "Selecione as imagens (PNG, JPG, PDF)",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
    )

    if uploaded:
        round_tag = st.text_input("Tag do round (ex: Round_11)", value=f"Round_{next_round}")
        if st.button("💾 Salvar imagens", type="primary"):
            for f in uploaded:
                save_path = IMAGES_DIR / f"{round_tag}_{f.name}"
                save_path.write_bytes(f.read())
                st.success(f"Salvo: {save_path.name}")

    st.divider()
    st.subheader("Imagens salvas")
    imgs = sorted(IMAGES_DIR.iterdir()) if IMAGES_DIR.exists() else []
    if not imgs:
        st.write("Nenhuma imagem ainda.")
    else:
        cols = st.columns(3)
        for i, img_path in enumerate(imgs):
            with cols[i % 3]:
                if img_path.suffix.lower() in {".png", ".jpg", ".jpeg"}:
                    st.image(str(img_path), caption=img_path.name, use_container_width=True)
                else:
                    st.write(f"📄 {img_path.name}")
                if st.button("🗑️ Apagar", key=f"del_img_{i}"):
                    img_path.unlink()
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — VER / APAGAR
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Dados atuais")
    st.dataframe(df, use_container_width=True)

    st.divider()
    st.subheader("Apagar round completo")
    rounds_available = sorted(df["Round"].unique().tolist())
    round_to_delete = st.selectbox("Selecione o round para apagar", rounds_available)
    preview = df[df["Round"] == round_to_delete][["Round", "Date", "Match", "Team"]]
    st.dataframe(preview, use_container_width=True)

    if st.button("🗑️ Apagar este round", type="secondary"):
        shutil.copy(EXCEL_FILE, EXCEL_FILE.with_suffix(".bak.xlsx"))
        updated = df[df["Round"] != round_to_delete]
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a",
                            if_sheet_exists="overlay") as writer:
            updated.to_excel(writer, sheet_name="DashboardData", index=False)
        st.success(f"Round {round_to_delete} apagado. Backup salvo.")
        st.cache_data.clear()
        st.rerun()
