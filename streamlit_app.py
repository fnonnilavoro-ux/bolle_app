import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime

# === NUOVO: AG-Grid super fluida ===
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

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
        if c in norm_map: return norm_map[c]
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
    if not isinstance(s, str): return ""
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
        if "righe" in nl and "doc" in nl: return s
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
            if len(lineH) != 128: raise ValueError("Record 01 non lungo 128.")
            records.append(lineH)
            continue

        # DETTAGLIO (02)
        if current_header is None: continue
        cod_val = row.get(col_cod, None)
        qta_val = row.get(col_qta, None)
        if pd.isna(cod_val) or pd.isna(qta_val): continue

        try: codice_art = str(int(cod_val))
        except: codice_art = str(cod_val or "")

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
            (65,12, ""),   # Prezzo BLANK
            (74,12, ""),   # Importo BLANK
            (83, 4, " "),  # Pezzi BLANK
            (87, 1, ""),   # Ass. IVA BLANK
            (88, 2, " "),  # Cod. IVA BLANK
            (90, 1, ""),   # Tipo movimento BLANK
            (91, 1, "1"),  # Tipo cessione
            (92, 5, "00000"), # Colli
            (97,12, ""),   # Filler
            (109,1, ""),   # Tipo resa
            (110,19, ""),  # Filler finale
        ]
        lineD = build_fixed_line(fieldsD, 128)
        if len(lineD) != 128: raise ValueError("Record 02 non lungo 128.")
        records.append(lineD)

    if not records:
        raise RuntimeError("Nessun record generato. Controlla intestazioni e colonne.")
    return records

# ---------- Griglia unica (AG-Grid) ----------
CHAR_COLS = [str(i) for i in range(1, 129)]  # 1..128

def text_to_df(text: str, width: int = 128) -> pd.DataFrame:
    lines = text.splitlines()
    rows = []
    for line in lines:
        padded = (line[:width]).ljust(width)
        rows.append(list(padded))
    df = pd.DataFrame(rows, columns=CHAR_COLS)
    df.index = range(1, len(df) + 1)
    return df

def df_to_text(df: pd.DataFrame) -> str:
    out_lines = []
    for _, row in df.iterrows():
        out_lines.append("".join([(str(x) if pd.notna(x) and str(x) != "" else " ")[0] for x in row.tolist()]))
    return "\n".join(out_lines) + "\n"

# --- Session state ---
if "txt_base" not in st.session_state:      # generato dall’Excel (originale)
    st.session_state.txt_base = ""
if "txt_saved" not in st.session_state:     # versione SALVATA (per download)
    st.session_state.txt_saved = ""
if "grid_df" not in st.session_state:       # DataFrame con 1 char per cella
    st.session_state.grid_df = pd.DataFrame()
if "last_saved_at" not in st.session_state:
    st.session_state.last_saved_at = None
if "show_dots" not in st.session_state:     # visualizza spazi come · (solo rendering)
    st.session_state.show_dots = True

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
        for k in ["txt_base","txt_saved","grid_df","last_saved_at"]:
            if k in st.session_state: del st.session_state[k]
        st.rerun()
with b2:
    if st.button("↩️ Ripristina TXT originale", use_container_width=True):
        if st.session_state.txt_base:
            st.session_state.grid_df = text_to_df(st.session_state.txt_base, 128)
            st.session_state.txt_saved = st.session_state.txt_base
            st.session_state.last_saved_at = None
            st.toast("Anteprima riportata allo stato originale")
with b3:
    if st.button("💾 Salva modifiche", use_container_width=True):
        if not st.session_state.grid_df.empty:
            st.session_state.txt_saved = df_to_text(st.session_state.grid_df)
            st.session_state.last_saved_at = datetime.now().strftime("%H:%M:%S")
            st.toast("Modifiche salvate")

uploaded = st.file_uploader("Carica Excel (.xlsx / .xls)", type=["xlsx", "xls"])

# --- CSS monospace per AG-Grid ---
st.markdown("""
<style>
.ag-theme-streamlit { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Courier New", monospace; }
.ag-theme-streamlit .ag-cell, 
.ag-theme-streamlit .ag-header-cell { font-size: 12px; padding-left:4px; padding-right:4px; }
</style>
""", unsafe_allow_html=True)

if uploaded:
    try:
        records = convert_excel_to_records(uploaded, cod_forn, cod_cli)
        base_txt = "\n".join(records) + "\n"

        # Prima inizializzazione (solo quando cambia il contenuto)
        if base_txt != st.session_state.txt_base or st.session_state.grid_df.empty:
            st.session_state.txt_base = base_txt
            st.session_state.grid_df = text_to_df(base_txt, 128)
            st.session_state.txt_saved = base_txt
            st.session_state.last_saved_at = None

        st.success(f"OK ✅ Records generati: {len(records)}")

        # --------- AG-Grid options (ultra fluido) ----------
        gb = GridOptionsBuilder.from_dataframe(st.session_state.grid_df)
        gb.configure_grid_options(
            rowHeight=24,
            suppressMovableColumns=True,
            enableRangeSelection=True,
            suppressCellSelection=False,
            singleClickEdit=True,
            editType="fullCell",
            ensureDomOrder=True,
            rowBuffer=20,
            domLayout="normal",
        )
        # colonne 1..128: editable + valueSetter (limita a 1 char; '·' -> spazio)
        one_char_setter = JsCode("""
            function(params){
                let v = params.newValue;
                if (v === null || v === undefined) v = " ";
                v = String(v);
                if (v === "·") v = " ";
                if (v.length === 0) v = " ";
                params.data[params.colDef.field] = v.substring(0,1);
                return true;
            }
        """)
        # formatter: mostra · per spazi (solo rendering)
        space_formatter_on  = JsCode("function(params){ return (params.value === ' ') ? '·' : params.value; }")
        space_formatter_off = JsCode("function(params){ return params.value; }")

        for c in CHAR_COLS:
            gb.configure_column(
                c,
                header_name=c,
                editable=True,
                width=26,
                resizable=False,
                sortable=False,
                filter=False,
                valueSetter=one_char_setter,
                valueFormatter=space_formatter_on if st.session_state.show_dots else space_formatter_off,
            )

        # Colonna TYPE (calcolata, non editabile) per capire al volo 01/02
        gb.configure_column(
            "TYPE",
            header_name="TYPE",
            valueGetter=JsCode("function(p){ return (p.data['1']||' ')+(p.data['2']||' '); }"),
            editable=False,
            pinned='left',
            width=60,
            resizable=False,
            sortable=False,
            filter=False,
        )
        # Inserisci TYPE come prima colonna
        col_order = ["TYPE"] + CHAR_COLS
        grid_options = gb.build()
        grid_options["columnDefs"].sort(key=lambda d: col_order.index(d["field"]))

        # Righe testata (01) più scure
        grid_options["getRowStyle"] = JsCode("""
            function(params){
                var c1 = params.data['1'] || ' ';
                var c2 = params.data['2'] || ' ';
                if (c1==='0' && c2==='1'){
                    return { backgroundColor: '#1f2937', color: '#e5e7eb' }; // grigio scuro
                }
                return null;
            }
        """)

        # Rendering
        grid_resp = AgGrid(
            st.session_state.grid_df,
            gridOptions=grid_options,
            theme="streamlit",  # light/dark si adatta
            height=min(600, 28 * max(10, len(st.session_state.grid_df) + 2)),
            fit_columns_on_grid_load=False,
            allow_unsafe_jscode=True,
            update_mode=GridUpdateMode.VALUE_CHANGED,      # commit cell -> rerun (veloce)
            data_return_mode=DataReturnMode.AS_INPUT,       # ritorna i valori correnti
        )

        # Aggiorna il DF in sessione SOLO quando arriva un delta dal grid
        if grid_resp and "data" in grid_resp:
            st.session_state.grid_df = pd.DataFrame(grid_resp["data"])

        # Info salvataggio
        if st.session_state.last_saved_at:
            st.caption(f"Ultimo salvataggio: **{st.session_state.last_saved_at}**")
        else:
            st.caption("Non salvato: scaricherai l’ultima versione **salvata**. Premi “Salva modifiche” per fissarla.")

        st.markdown("---")
        st.caption("Il download usa la **versione salvata** (non quella in modifica).")
        st.download_button(
            "⬇️ Scarica TXT (versione salvata)",
            data=st.session_state.txt_saved.encode(encoding, errors="strict"),
            file_name="export_bolle.txt",
            mime="text/plain",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Errore: {e}")
else:
    st.info("Carica il file Excel per iniziare.")
