import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Perth Dashboard", layout="wide")

with open("perth_azzurri_painel.html", "r", encoding="utf-8") as f:
    html = f.read()

components.html(html, height=5000, scrolling=True)
