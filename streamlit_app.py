import io
from datetime import datetime

import pandas as pd
import streamlit as st

# =========================
# CONFIG APP
# =========================
st.set_page_config(page_title="Bolle App — Export TXT", page_icon="📦", layout="wide")
st.title("📦 Bolle App — Export TXT (Anteprima modificabile)")

# -------------------------
# FORMATO DI DEFAULT (puoi cambiarli nell'Expander "Impostazioni (opzionali)")
# -------------------------
DEFAULT_SPAZI_NOME_QTY = 1     # spazi tra NOME e PEZZI
DEFAULT_SPAZI_QTY_KG   = 26    # spazi tra PEZZI e KG
DEFAULT_DECIMALI_KG    = 3     # decimali per KG
DEFAULT_DEC_SEP        = "."   # "." oppure ","
DEFAULT_FILENAME_BASE  = "bolle_export"

# Stato per anteprima
if "txt_base" not in st.session_state:
    st.session_state.txt_base = ""
if "txt_preview" not in st.session_state:
    st.session_state.txt_preview = ""

# =========================
# FUNZIONI
# =========================
def read_excel_any(file) -> pd.DataFrame:
    """Legge il primo foglio dell'Excel; se non è Excel, prova CSV."""
    name = getattr(file, "name", "") or ""
    if name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    # Fallback: CSV
    content = file.read()
    if isinstance(content, bytes):
        text = content.decode("utf-8", errors="ignore")
    else:
        text = str(content)
    sep = ";" if text.count(";") > text.count(",") else ","
    return pd.read_csv(io.StringIO(text), sep=sep)

def guess_mapping(columns):
    """Tenta di trovare colonne (nome, pezzi, kg) in base al nome."""
    cols_low = [c.lower() for c in columns]

    def find(keys, default_idx=0):
        for i, c in enumerate(cols_low):
            if any(k in c for k in keys):
                return columns[i]
        return columns[min(default_idx, len(columns)-1)]

    col_nome = find(["nome", "prodotto", "descr", "articolo"], 0)
    col_qty  = find(["pezzi", "quant", "qta", "qty"], 1 if len(columns) > 1 else 0)
    col_kg   = find(["kg", "peso", "weight"], 2 if len(columns) > 2 else len(columns)-1)
    return {"nome": col_nome, "pezzi": col_qty, "kg": col_kg}

def to_number(v, integer=False):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0 if integer else 0.0
    if isinstance(v, (int, float)):
        return int(v) if integer else float(v)
    s = str(v).strip().replace(" ", "")
    # Converti "1.234,56" -> "1234.56"
    if s.count(",") > 0 and s.count(".") == 0:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        # Rimuovi separatori migliaia con virgola (es. "1,234.56" -> "1234.56")
        if s.count(",") > 1:
            s = s.replace(",", "")
    try:
        x = float(s)
    except Exception:
        x = 0.0
    return int(round(x)) if integer else float(x)

def fmt_kg(value: float, decimali: int, dec_sep: str) -> str:
    s = f"{float(value):.{decimali}f}"
    if dec_sep == ",":
        s = s.replace(".", ",")
    return s

def build_txt(df: pd.DataFrame,
              col_nome: str, col_pezzi: str, col_kg: str,
              spazi_nome_qty: int, spazi_qty_kg: int,
              decimali_kg: int, dec_sep: str) -> str:
    """Crea il TXT finale riga per riga, senza fermate intermedie."""
    lines = []
    for r in df.itertuples(index=False):
        nome  = str(getattr(r, col_nome))
        pezzi = to_number(getattr(r, col_pezzi), integer=True)
        kg    = to_number(getattr(r, col_kg), integer=False)

        left   = f"{nome}{' ' * spazi_nome_qty}{int(pezzi)}"
        middle = " " * spazi_qty_kg
        right  = fmt_kg(kg, decimali_kg, dec_sep)
        lines.append(f"{left}{middle}{right}")
    return "\n".join(lines)

# =========================
# UI — UPLOAD
# =========================
uploaded = st.file_uploader("Carica il tuo file Excel (.xlsx/.xls). Supporto anche .csv.", type=["xlsx", "xls", "csv"])
if not uploaded:
    st.info("Carica un file per generare il TXT.")
    st.stop()

# Leggi file
try:
    df = read_excel_any(uploaded)
except Exception as e:
    st.error(f"Errore lettura file: {e}")
    st.stop()

if df.empty:
    st.error("Il file sembra vuoto.")
    st.stop()

# Mappatura automatica colonne
mapping = guess_mapping(df.columns)

# Impostazioni opzionali (collassate: NON interrompono il tuo flusso)
with st.expander("Impostazioni (opzionali)"):
    colA, colB, colC = st.columns(3)
    with colA:
        spazi_nome_qty = st.number_input("Spazi NOME → PEZZI", min_value=0, max_value=200,
                                         value=DEFAULT_SPAZI_NOME_QTY, step=1)
    with colB:
        spazi_qty_kg   = st.number_input("Spazi PEZZI → KG", min_value=0, max_value=200,
                                         value=DEFAULT_SPAZI_QTY_KG, step=1)
    with colC:
        decimali_kg    = st.number_input("Decimali KG", min_value=0, max_value=6,
                                         value=DEFAULT_DECIMALI_KG, step=1)

    dec_sep = st.radio("Separatore decimale (KG)", options=[".", ","],
                       index=0 if DEFAULT_DEC_SEP == "." else 1, horizontal=True)

    filename_base = st.text_input("Nome file (senza .txt)", value=DEFAULT_FILENAME_BASE)

# Se l'utente non apre l'expander, usa i default
if "spazi_nome_qty" not in locals():
    spazi_nome_qty = DEFAULT_SPAZI_NOME_QTY
    spazi_qty_kg   = DEFAULT_SPAZI_QTY_KG
    decimali_kg    = DEFAULT_DECIMALI_KG
    dec_sep        = DEFAULT_DEC_SEP
    filename_base  = DEFAULT_FILENAME_BASE

# =========================
# 1) GENERA TXT FINALE
# =========================
try:
    txt = build_txt(
        df=df,
        col_nome=mapping["nome"],
        col_pezzi=mapping["pezzi"],
        col_kg=mapping["kg"],
        spazi_nome_qty=spazi_nome_qty,
        spazi_qty_kg=spazi_qty_kg,
        decimali_kg=decimali_kg,
        dec_sep=dec_sep,
    )
except Exception as e:
    st.error(f"Errore durante la generazione del TXT: {e}")
    st.stop()

# Aggiorna base e, se è un nuovo file, resetta l'anteprima
if txt != st.session_state.txt_base:
    st.session_state.txt_base = txt
    st.session_state.txt_preview = txt

# =========================
# 2) MOSTRA ANTEPRIMA MODIFICABILE
# =========================
st.subheader("📝 Anteprima TXT (modificabile)")
st.caption("Il download userà esattamente questo contenuto.")

st.session_state.txt_preview = st.text_area(
    label="Contenuto TXT",
    value=st.session_state.txt_preview,
    height=420,
    key="txt_preview_area",
)

c1, c2, c3 = st.columns([1,1,2])
with c1:
    if st.button("↩️ Ripristina TXT originale", use_container_width=True):
        st.session_state.txt_preview = st.session_state.txt_base
        st.rerun()
with c2:
    final_name = (filename_base.strip() or DEFAULT_FILENAME_BASE) + ".txt"
    st.download_button(
        label="⬇️ Scarica TXT",
        data=st.session_state.txt_preview.encode("utf-8", errors="ignore"),
        file_name=final_name,
        mime="text/plain",
        use_container_width=True,
    )
with c3:
    st.write("")  # spazio
    st.caption(f"Righe: {st.session_state.txt_preview.count('\\n') + 1:,} — File: **{final_name}**")

# =========================
# 3) (OPZIONALE) ANTEPRIMA TABELLA
# =========================
with st.expander("Anteprima dati (prime 50 righe) — solo per controllo, non obbligatoria"):
    st.dataframe(df.head(50), use_container_width=True)
