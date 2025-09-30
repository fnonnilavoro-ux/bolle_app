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
HDR_RE = re.compile(
    r"(?:\*\*\s*)?Rif\.\s*Doc\.\s*di\s*trasporto\s*(\d+)\s*del\s*(\d{2}/\d{2}/\d{4})[:\s]*",
    re.IGNORECASE
)

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

# ---------- Helper per griglie editabili (separate 01 / 02) ----------
def split_lines_by_type(text: str, width: int = 128):
    """Ritorna (order, headers_idx, details_idx, headers_lines, details_lines)
       order: lista degli indici originali e tipo: [('01', idx01), ('02', idx02), ...]"""
    lines = text.splitlines()
    order = []
    headers_idx, details_idx = [], []
    headers_lines, details_lines = [], []
    for i, line in enumerate(lines):
        rec_type = line[:2]
        if rec_type == "01":
            headers_idx.append(i)
            headers_lines.append(line[:width].ljust(width))
            order.append(("01", i))
        else:
            details_idx.append(i)
            details_lines.append(line[:width].ljust(width))
            order.append(("02", i))
    return order, headers_idx, details_idx, headers_lines, details_lines

def df_from_lines(lines, show_dots=True):
    rows = []
    for line in lines:
        row = [("·" if (ch == " " and show_dots) else ch) for ch in line]
        rows.append(row)
    df = pd.DataFrame(rows, columns=[str(i) for i in range(1, 129)])
    # mantieni lunghezza 128 sempre
    return df

def lines_from_df(df: pd.DataFrame, show_dots=True):
    out = []
    for _, row in df.iterrows():
        chars = []
        for ch in row.tolist():
            c = "" if ch is None else str(ch)
            if c == "":
                c = " "
            if show_dots and c == "·":
                c = " "
            c = c[0]  # solo primo carattere
            chars.append(c)
        out.append("".join(chars))
    return out

def merge_back(order, headers_idx, details_idx, headers_lines, details_lines):
    """Ricompone nel loro ordine originale."""
    h_map = dict(zip(headers_idx, headers_lines))
    d_map = dict(zip(details_idx, details_lines))
    merged = []
    for typ, i in order:
        merged.append(h_map[i] if typ == "01" else d_map[i])
    return "\n".join(merged) + "\n"

# ---------------- Stato ----------------
if "txt_base" not in st.session_state:      # generato dall'Excel (originale)
    st.session_state.txt_base = ""
if "txt_saved" not in st.session_state:     # versione SALVATA (per download)
    st.session_state.txt_saved = ""
if "order" not in st.session_state:
    st.session_state.order = []
if "h_idx" not in st.session_state:
    st.session_state.h_idx = []
if "d_idx" not in st.session_state:
    st.session_state.d_idx = []
if "df_h" not in st.session_state:          # DataFrame editabile testate
    st.session_state.df_h = pd.DataFrame()
if "df_d" not in st.session_state:          # DataFrame editabile dettagli
    st.session_state.df_d = pd.DataFrame()
if "show_dots" not in st.session_state:
    st.session_state.show_dots = True
if "last_saved_at" not in st.session_state:
    st.session_state.last_saved_at = None

# ---------------- UI ----------------
st.markdown("Carica l’Excel (foglio con intestazioni bolla + righe articolo).")

encoding = st.radio("Encoding TXT", ["utf-8", "cp1252"], horizontal=True, index=0)

c1, c2, c3 = st.columns([1,1,1])
with c1:
    cod_forn = st.text_input("Codice fornitore (opzionale, 15 char)", value="")
with c2:
    cod_cli = st.text_input("Codice cliente ricevente (opzionale, right-align 15)", value="")
with c3:
    st.checkbox("Mostra spazi come · (solo anteprima)", key="show_dots", value=st.session_state.show_dots)

# Barra comandi (Reset / Ripristina / Salva)
b1, b2, b3 = st.columns([1,1,1])
with b1:
    if st.button("🧹 Reset (nuovo file)", use_container_width=True, type="secondary"):
        for k in ["txt_base", "txt_saved", "order", "h_idx", "d_idx", "df_h", "df_d", "show_dots", "last_saved_at"]:
            if k in st.session_state:
                del st.session_state[k]
        st.experimental_rerun()
with b2:
    if st.button("↩️ Ripristina TXT originale", use_container_width=True):
        if st.session_state.txt_base:
            order, h_idx, d_idx, h_lines, d_lines = split_lines_by_type(st.session_state.txt_base, 128)
            st.session_state.order = order
            st.session_state.h_idx = h_idx
            st.session_state.d_idx = d_idx
            st.session_state.df_h = df_from_lines(h_lines, st.session_state.show_dots)
            st.session_state.df_d = df_from_lines(d_lines, st.session_state.show_dots)
            st.toast("Anteprima riportata allo stato originale")
with b3:
    if st.button("💾 Salva modifiche", use_container_width=True):
        if not st.session_state.df_h.empty or not st.session_state.df_d.empty:
            # Ricostruisci linee da entrambi gli editor
            h_lines = lines_from_df(st.session_state.df_h, st.session_state.show_dots)
            d_lines = lines_from_df(st.session_state.df_d, st.session_state.show_dots)
            txt_preview = merge_back(st.session_state.order, st.session_state.h_idx, st.session_state.d_idx, h_lines, d_lines)
            st.session_state.txt_saved = txt_preview
            st.session_state.last_saved_at = datetime.now().strftime("%H:%M:%S")
            st.toast("Modifiche salvate")

uploaded = st.file_uploader("Carica Excel (.xlsx / .xls)", type=["xlsx", "xls"])

# CSS per monospace, celle strette e sezioni
st.markdown("""
<style>
/* Monospace in editor */
div[data-testid="stDataEditor"] table {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 12px;
}
div[data-testid="stDataEditor"] td, div[data-testid="stDataEditor"] th {
  white-space: pre !important;
  padding: 2px 6px !important;
}

/* Box testate più scuro */
.section-headers {
  background: #111827; /* grigio molto scuro */
  padding: 10px 12px;
  border-radius: 8px;
  border: 1px solid #1f2937;
  margin-bottom: 12px;
}
.section-headers h4{
  color: #e5e7eb;
}

/* Box dettagli normale/chiaro */
.section-details {
  background: #0b0f17;
  padding: 10px 12px;
  border-radius: 8px;
  border: 1px solid #111827;
}
.section-details h4{
  color: #e5e7eb;
}

/* intestazioni riga/colonna più compatte */
div[data-testid="stDataEditor"] thead th {
  padding: 4px 6px !important;
}
</style>
""", unsafe_allow_html=True)

if uploaded:
    try:
        records = convert_excel_to_records(uploaded, cod_forn, cod_cli)
        base_txt = "\n".join(records) + "\n"

        if base_txt != st.session_state.txt_base or st.session_state.df_h.empty:
            st.session_state.txt_base = base_txt
            order, h_idx, d_idx, h_lines, d_lines = split_lines_by_type(base_txt, 128)
            st.session_state.order = order
            st.session_state.h_idx = h_idx
            st.session_state.d_idx = d_idx
            st.session_state.df_h = df_from_lines(h_lines, st.session_state.show_dots)
            st.session_state.df_d = df_from_lines(d_lines, st.session_state.show_dots)
            st.session_state.txt_saved = base_txt
            st.session_state.last_saved_at = None

        st.success(f"OK ✅ Records generati: {len(records)}")

        # Se l'utente cambia la preferenza dei puntini, rigenera SOLO la vista (non i dati sottostanti)
        st.session_state.df_h = df_from_lines(
            lines_from_df(st.session_state.df_h, True),  # prima “leva” i puntini attuali
            st.session_state.show_dots
        )
        st.session_state.df_d = df_from_lines(
            lines_from_df(st.session_state.df_d, True),
            st.session_state.show_dots
        )

        # ====== GRIGLIA TESTATE (scura) ======
        st.markdown('<div class="section-headers"><h4>Righe TESTATA (01)</h4>', unsafe_allow_html=True)
        edited_h = st.data_editor(
            st.session_state.df_h,
            key="grid_headers",
            use_container_width=True,
            num_rows="fixed",   # le testate hanno cardinalità fissa derivata dal file
            disabled=False,
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # ====== GRIGLIA DETTAGLI (chiara) ======
        st.markdown('<div class="section-details"><h4>Righe DETTAGLIO (02)</h4>', unsafe_allow_html=True)
        edited_d = st.data_editor(
            st.session_state.df_d,
            key="grid_details",
            use_container_width=True,
            num_rows="dynamic",  # lasciamo libertà in più sui dettagli
            disabled=False,
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # Aggiorna stato con le versioni editate
        st.session_state.df_h = edited_h
        st.session_state.df_d = edited_d

        # Info salvataggio
        if st.session_state.last_saved_at:
            st.caption(f"Ultimo salvataggio: **{st.session_state.last_saved_at}**")
        else:
            st.caption("Non salvato: scaricherai l’ultima versione **salvata**. Premi “Salva modifiche” per fissarla.")

        st.markdown("---")
        st.caption("Il download usa la **versione salvata** (non quella in modifica).")
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
