import io
import re
import hashlib
from datetime import datetime

import pandas as pd
import streamlit as st

from processor import (
    read_table_any,
    suggest_column_mapping,
    clean_product_name,
    coerce_number,
)

# =========================
# Config
# =========================
st.set_page_config(page_title="Bolle App", page_icon="📦", layout="wide")

# =========================
# Utils - Preview & Download
# =========================
def _hash_text(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8", errors="ignore")).hexdigest()

def preview_and_download(generated_txt: str, default_filename: str = None, encoding: str = "utf-8"):
    """
    Mostra un'anteprima EDITABILE del TXT.
    Il file scaricato è SEMPRE il contenuto attuale dell'anteprima.
    """
    if not isinstance(generated_txt, str):
        st.error("preview_and_download: atteso testo (str).")
        return

    if not default_filename:
        default_filename = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    current_hash = _hash_text(generated_txt)
    if "preview_state" not in st.session_state:
        st.session_state.preview_state = {"orig_hash": None, "text": generated_txt}

    if st.session_state.preview_state.get("orig_hash") != current_hash:
        st.session_state.preview_state = {"orig_hash": current_hash, "text": generated_txt}

    st.subheader("📝 Anteprima TXT (modificabile)")
    st.caption("Modifica liberamente qui sotto: il file scaricato sarà esattamente questo contenuto.")

    st.session_state.preview_state["text"] = st.text_area(
        label="Contenuto TXT",
        value=st.session_state.preview_state["text"],
        key="txt_preview_area",
        height=420,
        help="Questo è il testo che verrà scaricato.",
    )

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("↩️ Ripristina originale", use_container_width=True):
            st.session_state.preview_state["text"] = generated_txt
            st.rerun()
    with col2:
        st.download_button(
            label="⬇️ Scarica TXT",
            data=st.session_state.preview_state["text"].encode(encoding, errors="ignore"),
            file_name=default_filename,
            mime="text/plain",
            use_container_width=True,
        )
    with col3:
        enc = st.selectbox(
            "Encoding",
            ["utf-8", "cp1252", "latin-1"],
            index=["utf-8", "cp1252", "latin-1"].index(encoding),
        )
        if enc != encoding:
            # Forza rerender mantenendo il testo
            st.session_state.preview_state["text"] = st.session_state.preview_state["text"]
            st.rerun()

# =========================
# Sidebar - Parametri
# =========================
st.sidebar.header("⚙️ Impostazioni Output")

spazi_qty_kg = st.sidebar.number_input(
    "Spazi tra quantità (pezzi) e kg", min_value=0, max_value=200, value=26, step=1,
    help="Numero esatto di spazi fissi tra il valore 'pezzi' e il valore 'kg'."
)

spazi_nome_qty = st.sidebar.number_input(
    "Spazi tra nome prodotto e quantità", min_value=1, max_value=200, value=1, step=1,
    help="Spazi tra il nome del prodotto e la quantità."
)

decimali_kg = st.sidebar.number_input(
    "Decimali per kg", min_value=0, max_value=6, value=3, step=1,
    help="Quanti decimali mostrare per i kg nel TXT."
)

usa_virgola_decimale = st.sidebar.checkbox(
    "Usa virgola come separatore decimale (kg)", value=False,
    help="Se attivo, i kg verranno formattati con la virgola (es. 12,345)."
)

pulizia_nome = st.sidebar.selectbox(
    "Pulizia nome prodotto",
    ["Nessuna", "Base (rimuovi codici/sku comuni)", "Aggressiva (solo lettere, numeri, spazi, -_/.)"],
    index=1,
    help="Controlla quanto 'pulire' i nomi (rimuovere codici, parentesi, SKU, ecc.)."
)

rimuovi_doppispazi = st.sidebar.checkbox(
    "Compatta spazi multipli nel nome prodotto", value=True,
    help="Sostituisce sequenze di spazi multipli con uno spazio singolo all'interno del nome."
)

nome_file_base = st.sidebar.text_input(
    "Nome file (senza estensione)", value="bolle_export",
    help="Il file scaricato sarà <nome>.txt"
)

st.sidebar.markdown("---")
st.sidebar.caption("Carica CSV o Excel. Poi mappa le colonne e genera l'anteprima.")

# =========================
# Upload
# =========================
st.title("📦 Bolle App — Generatore TXT con Anteprima Modificabile")

uploaded = st.file_uploader("Carica file (.csv, .xlsx, .xls)", type=["csv", "xlsx", "xls"])
if not uploaded:
    st.info("Carica un file per iniziare.")
    st.stop()

# Leggi tabella
try:
    df = read_table_any(uploaded)
except Exception as e:
    st.error(f"Errore lettura file: {e}")
    st.stop()

if df.empty:
    st.error("Il file è vuoto o non contiene righe valide.")
    st.stop()

st.success(f"File caricato: {uploaded.name} — {len(df):,} righe")
with st.expander("Anteprima tabella (prime 50 righe)"):
    st.dataframe(df.head(50), use_container_width=True)

# =========================
# Mappatura colonne
# =========================
st.subheader("🧭 Mappa le colonne")
suggest = suggest_column_mapping(df.columns)

col1, col2, col3 = st.columns(3)
with col1:
    col_nome = st.selectbox("Colonna: Nome Prodotto", options=df.columns.tolist(),
                            index=df.columns.get_indexer([suggest.get("name", df.columns[0])])[0])
with col2:
    col_pezzi = st.selectbox("Colonna: Quantità (pezzi)", options=df.columns.tolist(),
                             index=df.columns.get_indexer([suggest.get("qty", df.columns[min(1, len(df.columns)-1)])])[0])
with col3:
    col_kg = st.selectbox("Colonna: Kg", options=df.columns.tolist(),
                          index=df.columns.get_indexer([suggest.get("kg", df.columns[min(2, len(df.columns)-1)])])[0])

# =========================
# Trasformazioni
# =========================
def transform_name(x: str) -> str:
    s = str(x) if pd.notna(x) else ""
    if pulizia_nome.startswith("Base"):
        s = clean_product_name(s, mode="base")
    elif pulizia_nome.startswith("Aggressiva"):
        s = clean_product_name(s, mode="aggressive")
    if rimuovi_doppispazi:
        s = re.sub(r"\s{2,}", " ", s).strip()
    return s

work = pd.DataFrame({
    "nome": df[col_nome].map(transform_name),
    "pezzi": df[col_pezzi].map(lambda v: coerce_number(v, integer=True)),
    "kg": df[col_kg].map(lambda v: coerce_number(v, integer=False)),
})

# Avvisi qualità dati
bad_qty = work["pezzi"].isna().sum()
bad_kg = work["kg"].isna().sum()
if bad_qty or bad_kg:
    st.warning(f"⚠️ Valori non numerici: pezzi={bad_qty}, kg={bad_kg}. Saranno trattati come 0.")
    work["pezzi"] = work["pezzi"].fillna(0).astype(int)
    work["kg"] = work["kg"].fillna(0.0)

# =========================
# Generazione TXT (grezza)
# =========================
def fmt_kg(value: float) -> str:
    if value is None:
        value = 0.0
    if usa_virgola_decimale:
        # Usa virgola come separatore
        s = f"{value:.{decimali_kg}f}".replace(".", ",")
    else:
        s = f"{value:.{decimali_kg}f}"
    return s

def build_line(nome: str, pezzi: int, kg: float) -> str:
    left = f"{nome}{' ' * spazi_nome_qty}{int(pezzi)}"
    middle = " " * spazi_qty_kg
    right = fmt_kg(kg)
    return f"{left}{middle}{right}"

lines = [build_line(r.nome, r.pezzi, r.kg) for r in work.itertuples(index=False)]
generated_txt = "\n".join(lines)

# =========================
# Anteprima NON modificabile (tabellare) + differenze spazi
# =========================
with st.expander("🔎 Controllo formattazione (non modificabile)"):
    show = work.copy()
    show["__spazi_nome_qty__"] = spazi_nome_qty
    show["__spazi_qty_kg__"] = spazi_qty_kg
    st.dataframe(show.head(100), use_container_width=True)
    st.code("\n".join(lines[:10]), language="text")

# =========================
# Anteprima MODIFICABILE + Download
# =========================
final_name = (nome_file_base.strip() or "bolle_export") + ".txt"
preview_and_download(generated_txt, default_filename=final_name, encoding="utf-8")
