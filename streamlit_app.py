import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime

st.set_page_config(page_title="Excel → TXT Bolle", page_icon="📦", layout="wide")
st.title("Excel → TXT (record fissi 128)")

# ---------------- Normalizzazione nomi colonna ----------------
def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def normcol(s: str) -> str:
    s = (s or "").strip().lower()
    s = strip_accents(s)
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def pick_col(norm_map, candidates):
    for c in candidates:
        if c in norm_map:
            return norm_map[c]
    for real_norm, real_name in norm_map.items():
        if any(real_norm.startswith(c) for c in candidates):
            return real_name
    return None

# ---------------- Regole/Helper tracciato ----------------
HDR_VALUTA = "EUR"
HDR_DOC_ID = "DSV"
HDR_TIPO_CLI = "1"

HDR_RE = re.compile(
    r"(?:\*\*\s*)?Rif\.\s*Doc\.\s*di\s*trasporto\s*(\d+)\s*del\s*(\d{2}/\d{2}/\d{4})[:\s]*",
    re.IGNORECASE
)

# Varianti di "pacchetti" finali da rimuovere (2 B, 10 PZ, (6pz), x12, ecc.)
PACK_TAILS = [
    r"\s*\(\s*\d+\s*(?:pz|pzs?|b)\.?\s*\)\s*$",
    r"\s*-\s*\d+\s*(?:pz|pzs?|b)\.?\s*$",
    r"\s+x?\d+\s*(?:pz|pzs?|b)\.?\s*$",
    r"\s+\d+\s*(?:pz|pzs?|b)\.?\s*$",
    r"\s+\d+\s*$",
    r"\s*\d+(?:pz|pzs?|b)\.?\s*$",
]
PACK_TAIL = re.compile("(" + "|".join(PACK_TAILS) + ")", re.IGNORECASE)

def left_pad(v, n):
    s = "" if v is None else str(v)
    return (s + " " * n)[:n]

def right_pad(v, n):
    s = "" if v is None else str(v)
    return (" " * n + s)[-n:]

def qty_10_3(v) -> str:
    try:
        i = int(round(float(v) * 1000))
        return str(i).zfill(10)
    except:
        return "0" * 10

def build_fixed_line(fields, total=128):
    buf = [" "] * total
    for start, length, val in fields:
        s = "" if val is None else str(val)
        s = s[:length]
        buf[start - 1:start - 1 + length] = list(s.ljust(length, " "))
    return "".join(buf)

def clean_descr(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = re.sub(r"\s+", " ", s).strip()
    prev = None
    while prev != s:
        prev = s
        s = PACK_TAIL.sub("", s).strip()
    return s

def um_from_cols(um_val, descr_val) -> str:
    um = (um_val or "").strip().upper()
    if um in ("KG", "PZ"):
        return um
    return "PZ" if " PZ" in f" {(descr_val or '').upper()}" else "KG"

def pick_sheet(xls: pd.ExcelFile) -> str:
    for s in xls.sheet_names:
        nl = s.lower()
        if "righe" in nl and "doc" in nl:
            return s
    return xls.sheet_names[0]

# ---------------- Conversione ----------------
def convert_excel_to_records(excel_bytes, cod_forn="", cod_cli_ricev=""):
    xls = pd.ExcelFile(excel_bytes)
    sheet = pick_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sheet)

    norm_map = {normcol(c): c for c in df.columns}
    col_descr = pick_col(norm_map, ["descrizione", "descrizionearticolo", "desc", "articolo"])
    col_cod   = pick_col(norm_map, ["cod", "codice", "codicearticolo", "codarticolo", "codart"])
    col_qta   = pick_col(norm_map, ["qta", "quantita", "quantitaconsegnata", "quantitaordinata", "qta1", "qta2"])
    col_um    = pick_col(norm_map, ["um", "uom", "unitamisura", "unita", "unitadimisura"])

    missing = []
    if not col_descr: missing.append("Descrizione")
    if not col_cod:   missing.append("Cod.")
    if not col_qta:   missing.append("Q.tà/Quantità")
    if missing:
        detected = ", ".join([f"{c}→{normcol(c)}" for c in df.columns])
        raise ValueError(f"Mancano colonne: {', '.join(missing)}.\nColonne lette: {detected}")

    records = []
    progressivo = 0
    current_header = None

    for _, row in df.iterrows():
        descr_raw = str(row.get(col_descr, "") or "").strip()
        m = HDR_RE.search(descr_raw) if descr_raw else None

        # TESTATA (01)
        if m:
            num_bolla = m.group(1)
            try:
                d = datetime.strptime(m.group(2), "%d/%m/%Y")
                data_bolla = d.strftime("%y%m%d")
            except:
                data_bolla = "000000"
            progressivo += 1
            current_header = (num_bolla, data_bolla)

            fieldsH = [
                (1, 2, "01"),
                (3, 5, str(progressivo).zfill(5)),
                (8, 7, ""),
                (15, 6, ""),
                (21, 7, left_pad(num_bolla, 7)),
                (28, 6, data_bolla),
                (34,15, left_pad(cod_forn, 15)),
                (49, 1, " "),
                (50,15, ""),
                (65,15, ""),
                (80,15, right_pad(cod_cli_ricev, 15)),
                (95, 1, "1"),
                (96, 1, " "),
                (97, 3, "EUR"),
                (100,7, ""),
                (107,3, "DSV"),
                (110,10, left_pad(num_bolla, 10)),
                (120,9, ""),
            ]
            lineH = build_fixed_line(fieldsH, 128)
            if len(lineH) != 128:
                raise ValueError("Record 01 non lungo 128.")
            records.append(lineH)
            continue

        # DETTAGLIO (02)
        if current_header is None:
            continue

        cod_val = row.get(col_cod, None)
        qta_val = row.get(col_qta, None)
        if pd.isna(cod_val) or pd.isna(qta_val):
            continue

        try:
            codice_art = str(int(cod_val))
        except:
            codice_art = str(cod_val or "")

        descr_pulita = clean_descr(descr_raw)
        um_val = row.get(col_um, "") if col_um else ""
        um = um_from_cols(um_val, descr_raw)
        quantita = qty_10_3(qta_val)

        fieldsD = [
            (1, 2, "02"),
            (3, 5, str(progressivo).zfill(5)),
            (8, 15, left_pad(codice_art, 15)),
            (23,30, left_pad(descr_pulita, 30)),
            (53, 2, left_pad(um, 2)),
            (55,10, quantita),
            (65,12, ""),  # Prezzo BLANK
            (74,12, ""),  # Importo BLANK (start 74)
            (83, 4, " "), # Pezzi BLANK (start 83)
            (87, 1, ""),  # Ass. IVA BLANK
            (88, 2, " "), # Cod. IVA BLANK
            (90, 1, ""),  # Tipo movimento BLANK
            (91, 1, "1"), # Tipo cessione
            (92, 5, "00000"), # Colli
            (97,12, ""),  # Filler
            (109,1, ""),  # Tipo resa
            (110,19, ""), # Filler finale
        ]
        lineD = build_fixed_line(fieldsD, 128)
        if len(lineD) != 128:
            raise ValueError("Record 02 non lungo 128.")
        records.append(lineD)

    if not records:
        raise RuntimeError("Nessun record generato. Controlla intestazioni e colonne.")

    return records

# ---------------- UI ----------------
st.markdown("Carica l’Excel (foglio con intestazioni bolla + righe articolo) e scarica il TXT a 128 caratteri.")

encoding = st.radio("Encoding TXT", ["utf-8", "cp1252"], horizontal=True, index=0)

c1, c2 = st.columns(2)
with c1:
    cod_forn = st.text_input("Codice fornitore (opzionale, 15 char)", value="")
with c2:
    cod_cli = st.text_input("Codice cliente ricevente (opzionale, right-align 15)", value="")

uploaded = st.file_uploader("Carica Excel (.xlsx / .xls)", type=["xlsx", "xls"])

# --- Stato per preview/salvataggio ---
if "txt_base" not in st.session_state:
    st.session_state.txt_base = ""
if "txt_preview" not in st.session_state:
    st.session_state.txt_preview = ""
if "txt_saved" not in st.session_state:
    st.session_state.txt_saved = ""   # ultima versione SALVATA (per il download)
if "last_saved_at" not in st.session_state:
    st.session_state.last_saved_at = None

def make_grid_dataframe(text: str, width: int = 128, show_dots: bool = True) -> pd.DataFrame:
    """
    Converte il testo in griglia (riga x colonna) con 1 char per cella.
    - width: larghezza fissa (128 per il tuo tracciato)
    - show_dots: se True, visualizza gli spazi come '·' solo a video
    """
    rows = []
    lines = text.splitlines()
    for i, line in enumerate(lines, start=1):
        padded = (line[:width]).ljust(width)
        row = {}
        for pos, ch in enumerate(padded, start=1):
            if show_dots and ch == " ":
                row[pos] = "·"
            else:
                row[pos] = ch
        rows.append(row)
    df = pd.DataFrame(rows)
    df.index = range(1, len(df) + 1)  # numerazione righe 1..N
    df.columns = range(1, width + 1)  # numerazione colonne 1..width
    return df

# CSS per monospazio e preservare spazi nella tabella
st.markdown("""
<style>
/* forza monospace nella griglia */
div[data-testid="stDataFrame"] table {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 12px;
}
div[data-testid="stDataFrame"] td, div[data-testid="stDataFrame"] th {
  white-space: pre !important;
}
</style>
""", unsafe_allow_html=True)

if uploaded:
    try:
        records = convert_excel_to_records(uploaded, cod_forn, cod_cli)
        st.success(f"OK ✅ Records generati: {len(records)}")

        # TESTO FINALE
        txt = "\n".join(records) + "\n"

        # Se è cambiato il risultato, reset base, preview e saved
        if txt != st.session_state.txt_base:
            st.session_state.txt_base = txt
            st.session_state.txt_preview = txt
            st.session_state.txt_saved = txt
            st.session_state.last_saved_at = None

        # ---- Barra azioni (Ripristina / Salva) ----
        st.subheader("📝 Anteprima TXT (modificabile)")
        top_l, top_c, top_r = st.columns([1, 2, 1], vertical_alignment="center")
        with top_l:
            if st.button("↩️ Ripristina TXT originale", use_container_width=True):
                st.session_state.txt_preview = st.session_state.txt_base
                st.toast("Anteprima ripristinata")
        with top_r:
            if st.button("💾 Salva modifiche", use_container_width=True):
                st.session_state.txt_saved = st.session_state.txt_preview
                st.session_state.last_saved_at = datetime.now().strftime("%H:%M:%S")
                st.toast("Modifiche salvate")

        if st.session_state.last_saved_at:
            st.caption(f"Ultimo salvataggio: **{st.session_state.last_saved_at}**")

        # Editor
        st.session_state.txt_preview = st.text_area(
            label="Contenuto TXT",
            value=st.session_state.txt_preview,
            height=320,
            key="txt_preview_area",
        )

        # ---- Griglia caratteri (vista) ----
        st.markdown("#### 🔍 Griglia caratteri (anteprima visiva)")
        show_dots = st.checkbox("Mostra spazi come · (solo anteprima)", value=True)
        grid_df = make_grid_dataframe(st.session_state.txt_preview, width=128, show_dots=show_dots)
        # Nota: è una vista, non modifica il testo vero
        st.dataframe(grid_df, use_container_width=True, height=min(600, 36 + 24 * min(20, len(grid_df))))

        # ---- Download: usa SOLO la versione SALVATA ----
        st.markdown("---")
        st.caption("Il download utilizza l'ultima versione **SALVATA** (non quella in modifica).")
        data_to_download = st.session_state.txt_saved.encode(encoding, errors="strict")
        st.download_button(
            "⬇️ Scarica TXT (versione salvata)",
            data=data_to_download,
            file_name="export_bolle.txt",
            mime="text/plain",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Errore: {e}")
else:
    st.info("Carica il file Excel per iniziare.")
