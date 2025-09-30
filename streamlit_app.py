import streamlit as st
import pandas as pd
import re, unicodedata
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

# ----------------- CONFIG -----------------
st.set_page_config(page_title="Excel → TXT (record fissi 128)", page_icon="📦", layout="wide")
st.markdown("<h1 style='margin:0'>📦 Excel → TXT (record fissi 128)</h1>", unsafe_allow_html=True)

# ----------------- UTILS -----------------
def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def normcol(s: str) -> str:
    s = (s or "").strip().lower()
    s = strip_accents(s)
    return re.sub(r"[^a-z0-9]", "", s)

def pick_col(norm_map, candidates):
    for c in candidates:
        if c in norm_map: return norm_map[c]
    for real_norm, real_name in norm_map.items():
        if any(real_norm.startswith(c) for c in candidates): return real_name
    return None

HDR_RE = re.compile(
    r"(?:\*\*\s*)?Rif\.\s*Doc\.\s*di\s*trasporto\s*(\d+)\s*del\s*(\d{2}/\d{2}/\d{4})[:\s]*",
    re.IGNORECASE
)

# Pulizia descrizioni con regex semplici (no mega-pattern)
PACK_TAIL_PATTERNS = [
    re.compile(r"\s*\(\s*\d+\s*(?:pz|pzs?|b)\.?\s*\)\s*$", re.IGNORECASE),
    re.compile(r"\s*-\s*\d+\s*(?:pz|pzs?|b)\.?\s*$", re.IGNORECASE),
    re.compile(r"\s+x?\d+\s*(?:pz|pzs?|b)\.?\s*$", re.IGNORECASE),
    re.compile(r"\s+\d+\s*(?:pz|pzs?|b)\.?\s*$", re.IGNORECASE),
    re.compile(r"\s+\d+\s*$", re.IGNORECASE),
    re.compile(r"\s*\d+(?:pz|pzs?|b)\.?\s*$", re.IGNORECASE),
]

def left_pad(v, n):  s = "" if v is None else str(v); return (s + " " * n)[:n]
def right_pad(v, n): s = "" if v is None else str(v); return (" " * n + s)[-n:]

def qty_10_3(v) -> str:
    try: i = int(round(float(v) * 1000)); return str(i).zfill(10)
    except: return "0" * 10

def build_fixed_line(fields, total=128):
    buf = [" "] * total
    for start, length, val in fields:
        s = "" if val is None else str(val)
        s = s[:length]
        buf[start-1:start-1+length] = list(s.ljust(length, " "))
    return "".join(buf)

def clean_descr(s: str) -> str:
    if not isinstance(s, str): return ""
    s = re.sub(r"\s+", " ", s).strip()
    prev = None
    while prev != s:
        prev = s
        for pat in PACK_TAIL_PATTERNS:
            s = pat.sub("", s).strip()
    return s

def um_from_cols(um_val, descr_val) -> str:
    um = (um_val or "").strip().upper()
    if um in ("KG","PZ"): return um
    return "PZ" if " PZ" in f" {(descr_val or '').upper()}" else "KG"

def pick_sheet(xls: pd.ExcelFile) -> str:
    for s in xls.sheet_names:
        nl = s.lower()
        if "righe" in nl and "doc" in nl: return s
    return xls.sheet_names[0]

def convert_excel_to_records(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    df = pd.read_excel(xls, sheet_name=pick_sheet(xls))

    norm_map = {normcol(c): c for c in df.columns}
    col_descr = pick_col(norm_map, ["descrizione","descrizionearticolo","desc","articolo"])
    col_cod   = pick_col(norm_map, ["cod","codice","codicearticolo","codarticolo","codart"])
    col_qta   = pick_col(norm_map, ["qta","quantita","quantitaconsegnata","quantitaordinata","qta1","qta2"])
    col_um    = pick_col(norm_map, ["um","uom","unitamisura","unita","unitadimisura"])

    missing=[]
    if not col_descr: missing.append("Descrizione")
    if not col_cod:   missing.append("Cod.")
    if not col_qta:   missing.append("Q.tà/Quantità")
    if missing:
        detected = ", ".join([f"{c}→{normcol(c)}" for c in df.columns])
        raise ValueError(f"Mancano colonne: {', '.join(missing)}.\nColonne lette: {detected}")

    records=[]; progressivo=0; current_header=None
    for _,row in df.iterrows():
        descr_raw = str(row.get(col_descr, "") or "").strip()
        m = HDR_RE.search(descr_raw) if descr_raw else None

        if m:  # 01
            num_bolla = m.group(1)
            try: data_bolla = datetime.strptime(m.group(2), "%d/%m/%Y").strftime("%y%m%d")
            except: data_bolla="000000"
            progressivo += 1; current_header=(num_bolla,data_bolla)
            fieldsH = [
                (1,2,"01"), (3,5,str(progressivo).zfill(5)),
                (8,7,""), (15,6,""),
                (21,7,left_pad(num_bolla,7)), (28,6,data_bolla),
                (34,15,left_pad("",15)), (49,1," "), (50,15,""), (65,15,""),
                (80,15,right_pad("",15)), (95,1,"1"), (96,1," "), (97,3,"EUR"),
                (100,7,""), (107,3,"DSV"), (110,10,left_pad(num_bolla,10)), (120,9,""),
            ]
            lineH=build_fixed_line(fieldsH,128)
            if len(lineH)!=128: raise ValueError("Record 01 non lungo 128.")
            records.append(lineH); continue

        if current_header is None:  # 02 solo dopo 01
            continue

        cod_val=row.get(col_cod,None); qta_val=row.get(col_qta,None)
        if pd.isna(cod_val) or pd.isna(qta_val): continue

        try: codice_art=str(int(cod_val))
        except: codice_art=str(cod_val or "")
        descr_pulita=clean_descr(descr_raw)
        um=um_from_cols(row.get(col_um,"") if col_um else "", descr_raw)
        quantita=qty_10_3(qta_val)
        fieldsD=[
            (1,2,"02"), (3,5,str(progressivo).zfill(5)),
            (8,15,left_pad(codice_art,15)), (23,30,left_pad(descr_pulita,30)),
            (53,2,left_pad(um,2)), (55,10,quantita),
            (65,12,""), (74,12,""), (83,4," "), (87,1,""),
            (88,2," "), (90,1,""), (91,1,"1"), (92,5,"00000"),
            (97,12,""), (109,1,""), (110,19,""),
        ]
        lineD=build_fixed_line(fieldsD,128)
        if len(lineD)!=128: raise ValueError("Record 02 non lungo 128.")
        records.append(lineD)

    if not records: raise RuntimeError("Nessun record generato.")
    return records

# ---- TXT <-> DF (1 char per cella) ----
CHAR_COLS = [str(i) for i in range(1,129)]

def text_to_df(text:str)->pd.DataFrame:
    rows = [(line[:128]).ljust(128) for line in text.splitlines()]
    df = pd.DataFrame([list(r) for r in rows], columns=CHAR_COLS)
    df.index = range(1,len(df)+1)
    return df.astype(str)

def df_to_text(df:pd.DataFrame)->str:
    return "\n".join("".join((str(x) or " ")[0] for x in row) for row in df.values) + "\n"

# ----------------- SESSION -----------------
if "txt_base" not in st.session_state:   st.session_state.txt_base=""
if "txt_saved" not in st.session_state:  st.session_state.txt_saved=""
if "grid_df"  not in st.session_state:   st.session_state.grid_df=pd.DataFrame()
if "grid_opts" not in st.session_state:  st.session_state.grid_opts=None
if "ready"    not in st.session_state:   st.session_state.ready=False
if "last_saved_at" not in st.session_state: st.session_state.last_saved_at=None

# ----------------- STYLES -----------------
st.markdown("""
<style>
/* barra in alto fissa */
.toolbar { position: sticky; top: 0; z-index: 10; padding: 10px 0 8px; background: var(--background-color); }
.ag-theme-streamlit { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Courier New", monospace; }
.ag-theme-streamlit .ag-cell, .ag-theme-streamlit .ag-header-cell { font-size: 12px; }
.ag-row-hover { filter: brightness(1.05); }
.ag-row-odd { background-color: rgba(127,127,127,0.04); }
</style>
""", unsafe_allow_html=True)

# ----------------- UPLOAD + TOOLBAR -----------------
uploaded = st.file_uploader("Carica Excel (.xlsx/.xls)", type=["xlsx","xls"])

# Toolbar: SALVA / RIPRISTINA / RESET / SCARICA (in alto)
st.markdown('<div class="toolbar">', unsafe_allow_html=True)
t1, t2, t3, t4 = st.columns([1,1,1,2], vertical_alignment="center")
with t1:
    if st.button("💾 Salva modifiche", use_container_width=True):
        st.session_state.txt_saved = df_to_text(st.session_state.grid_df) if not st.session_state.grid_df.empty else ""
        st.session_state.last_saved_at = datetime.now().strftime("%H:%M:%S")
        st.toast("Salvato")
with t2:
    if st.button("↩️ Ripristina originale", use_container_width=True):
        if st.session_state.txt_base:
            st.session_state.grid_df = text_to_df(st.session_state.txt_base)
            st.session_state.txt_saved = st.session_state.txt_base
            st.session_state.last_saved_at = None
            st.toast("Ripristinato")
with t3:
    if st.button("🧹 Reset (nuovo file)", use_container_width=True):
        for k in ["txt_base","txt_saved","grid_df","grid_opts","ready","last_saved_at"]:
            if k in st.session_state: del st.session_state[k]
        st.rerun()
with t4:
    st.download_button(
        "⬇️ Scarica (versione salvata)",
        data=(st.session_state.txt_saved or "").encode("utf-8", errors="strict"),
        file_name="export_bolle.txt",
        mime="text/plain",
        use_container_width=True,
    )
st.markdown('</div>', unsafe_allow_html=True)

# ----------------- PIPELINE CARICAMENTO -----------------
if uploaded and not st.session_state.ready:
    try:
        txt = "\n".join(convert_excel_to_records(uploaded)) + "\n"
        st.session_state.txt_base  = txt
        st.session_state.txt_saved = txt
        st.session_state.grid_df   = text_to_df(txt)
        st.session_state.ready     = True
        st.session_state.grid_opts = None  # build options una volta
        st.toast("File caricato")
    except Exception as e:
        st.error(f"Errore: {e}")

if not st.session_state.ready:
    st.stop()

# ----------------- GRIGLIA (AG-Grid) -----------------
if st.session_state.grid_opts is None:
    df = st.session_state.grid_df
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        editable=True, cellEditor='agTextCellEditor',
        resizable=False, sortable=False, filter=False,
    )
    gb.configure_grid_options(
        rowHeight=24,
        singleClickEdit=True,
        stopEditingWhenCellsLoseFocus=False,
        enableCellTextSelection=True,
        ensureDomOrder=True,
        undoRedoCellEditing=True,
        undoRedoCellEditingLimit=500,
        rowBuffer=200,
        suppressMovableColumns=True,
        maintainColumnOrder=True,
        domLayout="normal",
        rowSelection="multiple",
    )

    # Parser input (1 char) + formatter (spazio → ·)
    one_char_parser = JsCode("""
        function(p){
            let v = p.newValue;
            if (v === null || v === undefined) v = " ";
            v = String(v);
            if (v === "·") v = " ";
            if (v.length === 0) v = " ";
            return v.substring(0,1);
        }
    """)
    space_formatter = JsCode("function(p){ return (p.value === ' ') ? '·' : p.value; }")

    # Colonna numeri di riga (non parte dei dati)
    gb.configure_column(
        "ROW",
        header_name="ROW",
        valueGetter=JsCode("function(p){ return (p.node.rowIndex + 1).toString(); }"),
        editable=False, pinned='left', width=58, suppressMenu=True
    )
    # Colonne 1..128
    for c in CHAR_COLS:
        gb.configure_column(c, header_name=c, width=26,
                            valueParser=one_char_parser, valueFormatter=space_formatter)

    st.session_state.grid_opts = gb.build()

    # Evidenziazione righe testata (01)
    st.session_state.grid_opts["getRowStyle"] = JsCode("""
        function(p){
            var c1 = p.data['1'] || ' ';
            var c2 = p.data['2'] || ' ';
            if (c1==='0' && c2==='1'){
                return { backgroundColor: '#1f2937', color: '#e5e7eb' };
            }
            return null;
        }
    """)

    # Ordine colonne: ROW poi 1..128
    defs = st.session_state.grid_opts["columnDefs"]
    defs.append({"field":"ROW"})
    order = ["ROW"] + CHAR_COLS
    st.session_state.grid_opts["columnDefs"] = sorted(defs, key=lambda d: order.index(d["field"]) if d["field"] in order else 999)

# Render: NO_UPDATE → nessun rerun mentre editi
grid_resp = AgGrid(
    st.session_state.grid_df,
    gridOptions=st.session_state.grid_opts,
    theme="streamlit",
    height=min(720, 28 * max(12, len(st.session_state.grid_df) + 3)),
    allow_unsafe_jscode=True,
    update_mode=GridUpdateMode.NO_UPDATE,
    data_return_mode=DataReturnMode.AS_INPUT,
    fit_columns_on_grid_load=False,
    reload_data=False,
)

# Aggiorna df in memoria quando Streamlit ricalcola (es. dopo Salva/Ripristina)
if grid_resp and "data" in grid_resp:
    new_df = pd.DataFrame(grid_resp["data"])
    # rimuovi eventuale colonna ROW
    if "ROW" in new_df.columns:
        new_df = new_df.drop(columns=["ROW"])
    # assicurati che ci siano solo 1..128 nell'ordine giusto
    st.session_state.grid_df = new_df[CHAR_COLS].astype(str)

# Footer info
if st.session_state.last_saved_at:
    st.caption(f"Ultimo salvataggio: **{st.session_state.last_saved_at}**")
