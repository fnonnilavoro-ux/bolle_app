import io
import re
import pandas as pd

# -------------------------
# Lettura file
# -------------------------
def read_table_any(file_obj) -> pd.DataFrame:
    name = getattr(file_obj, "name", "") or ""
    if name.lower().endswith(".csv"):
        # Prova a indovinare il separatore
        content = file_obj.read()
        if isinstance(content, bytes):
            sample = content[:4096].decode("utf-8", errors="ignore")
            buf = io.StringIO(content.decode("utf-8", errors="ignore"))
        else:
            sample = str(content)
            buf = io.StringIO(sample)
        sep = ";" if sample.count(";") > sample.count(",") else ","
        df = pd.read_csv(buf, sep=sep)
        return df
    elif name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(file_obj)
    else:
        # fallback: tenta CSV
        try:
            file_obj.seek(0)
        except Exception:
            pass
        try:
            return pd.read_csv(file_obj)
        except Exception:
            try:
                file_obj.seek(0)
            except Exception:
                pass
            return pd.read_excel(file_obj)

# -------------------------
# Mappatura colonne (suggest)
# -------------------------
def suggest_column_mapping(columns):
    cols = [c.lower() for c in columns]
    def find(*keys, default=None):
        for k in keys:
            for i, c in enumerate(cols):
                if k in c:
                    return columns[i]
        return default or columns[0]

    name = find("nome", "prodotto", "descr", "articolo", default=columns[0])
    qty  = find("pezzi", "quant", "qta", "qty", default=columns[1] if len(columns) > 1 else columns[0])
    kg   = find("kg", "peso", "weight", default=columns[2] if len(columns) > 2 else columns[-1])
    return {"name": name, "qty": qty, "kg": kg}

# -------------------------
# Pulizia nomi prodotto
# -------------------------
SKU_PATTERNS = [
    r"\bSKU[:\s]*[A-Za-z0-9\-_/\.]+\b",
    r"\bCOD(?:ICE)?[:\s]*[A-Za-z0-9\-_/\.]+\b",
    r"\bART(?:ICOLO)?[:\s]*[A-Za-z0-9\-_/\.]+\b",
    r"\([^\)]*\)$",            # parentesi finali
    r"\[[^\]]*\]$",
    r"[#@][A-Za-z0-9\-_/\.]+$",
    r"\b[A-Z]{2,}\d{2,}\b$",   # tipo ABC123 alla fine
]

def clean_product_name(s: str, mode: str = "base") -> str:
    """mode: 'base' o 'aggressive'"""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = s.strip()
    if mode == "base":
        for pat in SKU_PATTERNS:
            s = re.sub(pat, "", s, flags=re.IGNORECASE).strip()
        # Rimuovi trattini/virgole terminali
        s = re.sub(r"[\s,\-_/\.]+$", "", s).strip()
        return s
    elif mode == "aggressive":
        s = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ0-9 \-_/\.]", " ", s)
        s = re.sub(r"\s{2,}", " ", s).strip(" -_/.,")
        return s
    else:
        return s

# -------------------------
# Coercizione numeri
# -------------------------
def coerce_number(v, integer: bool):
    if v is None:
        return 0 if integer else 0.0
    if isinstance(v, (int, float)):
        return int(v) if integer else float(v)
    s = str(v).strip()
    # Normalizza separatori: prima togli i separatori migliaia comuni
    s = s.replace(" ", "")
    if s.count(",") > 0 and s.count(".") == 0:
        # probabile europeo: "1.234,56" -> "1234,56" già senza punti
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        # stile punto decimale: rimuovi eventuali separatori migliaia con virgola
        if s.count(",") > 1:
            s = s.replace(",", "")
    try:
        x = float(s)
    except Exception:
        x = 0.0
    return int(round(x)) if integer else float(x)
