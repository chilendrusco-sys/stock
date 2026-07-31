import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


def _needs_regeneration(path: Path) -> bool:
    return not path.exists() or path.stat().st_size == 0


def ensure_sample_files():
    DATA_DIR.mkdir(exist_ok=True)
    if _needs_regeneration(DEFAULT_BASE_FILE):
        sample_base = pd.DataFrame([
            {"Artículo": "Pollo", "Lote": "L001", "€/kg": 3.20},
            {"Artículo": "Pollo", "Lote": "L002", "€/kg": 3.40},
            {"Artículo": "Carne", "Lote": "L010", "€/kg": 5.10},
            {"Artículo": "Arroz", "Lote": "", "€/kg": 1.20},
        ])
        sample_base.to_excel(DEFAULT_BASE_FILE, index=False, engine="openpyxl")

    if _needs_regeneration(DEFAULT_WAREHOUSE_FILE):
        sample_warehouse = pd.DataFrame([
            {"Artículo": "Pollo", "Lote": "L001", "kg": 12.5},
            {"Artículo": "Pollo", "Lote": "L002", "kg": 8.0},
            {"Artículo": "Carne", "Lote": "L010", "kg": 4.5},
            {"Artículo": "Arroz", "Lote": "", "kg": 20.0},
        ])
        sample_warehouse.to_excel(DEFAULT_WAREHOUSE_FILE, index=False, engine="openpyxl")


st.set_page_config(page_title="Valorización de stock", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    [data-testid="stAppViewContainer"] {
        background: radial-gradient(circle at 15% 0%, #1c2333 0%, #11151f 55%, #0c0f17 100%);
    }
    [data-testid="stSidebar"] {
        background: #131722;
        border-right: 1px solid #262d3f;
    }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    .hero {
        background: linear-gradient(120deg, #1b2436 0%, #222c42 100%);
        border: 1px solid #34405c;
        border-radius: 16px;
        padding: 22px 26px;
        margin-bottom: 18px;
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        color: #f5f7fb;
        margin-bottom: 4px;
    }
    .hero-subtitle {
        color: #aab4c8;
        font-size: 0.98rem;
    }
    .kpi-card {
        background: #171c2a;
        border: 1px solid #2c3650;
        border-radius: 14px;
        padding: 16px 18px;
        height: 100%;
    }
    .kpi-label {
        color: #8b96b3;
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 6px;
    }
    .kpi-value {
        color: #f5f7fb;
        font-size: 1.6rem;
        font-weight: 800;
    }
    .kpi-value.warn { color: #f4b860; }
    .kpi-value.ok { color: #5fd68e; }
    .info-box {
        background: #16233a;
        border: 1px solid #2c4a75;
        border-radius: 12px;
        padding: 14px 16px;
        color: #d7e3f7;
        font-size: 0.9rem;
    }
    section[data-testid="stFileUploaderDropzone"] {
        background: #1a1f2e;
        border: 1px dashed #3a4560;
        border-radius: 12px;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid #2c3650;
        border-radius: 12px;
    }
    h3 { color: #f5f7fb !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_BASE_FILE = DATA_DIR / "base_stock.xlsx"
DEFAULT_WAREHOUSE_FILE = DATA_DIR / "warehouse_sample.xlsx"

ensure_sample_files()


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text)


def find_column(df: pd.DataFrame, candidates: list[str]):
    normalized_columns = {normalize_text(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in normalized_columns:
            return normalized_columns[key]

    for col in df.columns:
        col_norm = normalize_text(col)
        for candidate in candidates:
            cand_norm = normalize_text(candidate)
            if cand_norm in col_norm or col_norm in cand_norm:
                return col
    return None


def load_excel(path):
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception:
        # Reintenta con detección automática por si es un .xls antiguo u otro formato soportado.
        return pd.read_excel(path)


ITEM_CANDIDATES = ["articulo", "articuloid", "producto", "producto_id", "codigo", "codigoarticulo", "referencia", "ref", "item", "article", "sku"]
# "LoteInterno" es el lote real. El "Nº de Serie"/SSCC identifica el bulto/palet, no el lote: se excluye a propósito.
LOTE_CANDIDATES = ["loteinterno", "lote_interno", "lote interno", "lote", "lot", "batch", "numerodelote", "numlote"]
DESC_CANDIDATES = ["descripcion", "denominacion", "nombre", "detalle", "desc", "descripciónarticulo", "textobreve"]
# Candidatos de peso ordenados de más específico a más genérico para evitar confundir peso neto con peso bruto.
QTY_CANDIDATES = [
    "pesoneto", "peso_neto", "peso neto", "netweight", "netweightkg",
    "kgstock", "stockkg", "cantidadkg", "cantkg", "kilos", "kilogramos",
    "kg", "cantidad", "stock", "qty", "unidades",
]


def prepare_base_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    item_col = find_column(df, ITEM_CANDIDATES)
    lote_col = find_column(df, LOTE_CANDIDATES)
    price_col = find_column(df, ["€/kg", "eurokg", "euro_kg", "precio_kg", "coste_kg", "porkg", "pricekg", "priceperkg", "precio"])

    if not item_col or not price_col:
        raise ValueError("El Excel base debe tener al menos una columna de artículo y una de precio €/kg.")

    base = df[[item_col, lote_col, price_col]].copy() if lote_col else df[[item_col, price_col]].copy()
    base = base.rename(columns={item_col: "articulo", price_col: "€/kg"})
    if lote_col:
        base = base.rename(columns={lote_col: "lote"})
    else:
        base["lote"] = ""

    base["articulo"] = base["articulo"].fillna("")
    base["lote"] = base["lote"].fillna("")
    base["€/kg"] = pd.to_numeric(base["€/kg"], errors="coerce")
    base = base.dropna(subset=["€/kg"]).copy()
    base["articulo_key"] = base["articulo"].apply(normalize_text)
    base["lote_key"] = base["lote"].apply(normalize_text)
    base["match_key"] = base["articulo_key"] + "|" + base["lote_key"]
    base["match_key_item"] = base["articulo_key"]
    return base[["articulo", "lote", "€/kg", "match_key", "match_key_item"]].drop_duplicates(subset=["match_key"], keep="first")


def detect_warehouse_columns(df: pd.DataFrame) -> dict:
    return {
        "articulo": find_column(df, ITEM_CANDIDATES),
        "lote": find_column(df, LOTE_CANDIDATES),
        "descripcion": find_column(df, DESC_CANDIDATES),
        "kg_stock": find_column(df, QTY_CANDIDATES),
    }


def prepare_warehouse_dataframe(df: pd.DataFrame, columns: dict) -> pd.DataFrame:
    item_col = columns["articulo"]
    lote_col = columns["lote"]
    desc_col = columns["descripcion"]
    qty_col = columns["kg_stock"]

    if not item_col or not qty_col:
        raise ValueError("El Excel de almacén debe tener al menos una columna de artículo y una de peso/cantidad (kg, peso neto...).")

    keep_cols = [item_col, qty_col]
    if lote_col:
        keep_cols.append(lote_col)
    if desc_col:
        keep_cols.append(desc_col)

    warehouse = df[keep_cols].copy()
    rename_map = {item_col: "articulo", qty_col: "kg_stock"}
    if lote_col:
        rename_map[lote_col] = "lote"
    if desc_col:
        rename_map[desc_col] = "descripcion"
    warehouse = warehouse.rename(columns=rename_map)

    if "lote" not in warehouse.columns:
        warehouse["lote"] = ""
    if "descripcion" not in warehouse.columns:
        warehouse["descripcion"] = ""

    warehouse["articulo"] = warehouse["articulo"].fillna("")
    warehouse["lote"] = warehouse["lote"].fillna("")
    warehouse["descripcion"] = warehouse["descripcion"].fillna("")
    warehouse["kg_stock"] = pd.to_numeric(warehouse["kg_stock"], errors="coerce").fillna(0)
    warehouse["articulo_key"] = warehouse["articulo"].apply(normalize_text)
    warehouse["lote_key"] = warehouse["lote"].apply(normalize_text)
    warehouse["match_key"] = warehouse["articulo_key"] + "|" + warehouse["lote_key"]
    warehouse["match_key_item"] = warehouse["articulo_key"]
    return warehouse[["articulo", "descripcion", "lote", "kg_stock", "match_key", "match_key_item"]]


def valorate_stock(base_df: pd.DataFrame, warehouse_df: pd.DataFrame) -> pd.DataFrame:
    base_lookup = base_df.set_index("match_key")
    warehouse = warehouse_df.copy()

    warehouse["€/kg"] = None
    warehouse["importe"] = None

    for idx, row in warehouse.iterrows():
        if row["match_key"] in base_lookup.index:
            price = base_lookup.loc[row["match_key"], "€/kg"]
        elif row["match_key_item"] in base_lookup["match_key_item"].values:
            item_matches = base_lookup[base_lookup["match_key_item"] == row["match_key_item"]]
            price = item_matches["€/kg"].iloc[0]
        else:
            price = None

        warehouse.at[idx, "€/kg"] = price
        if pd.notna(price):
            warehouse.at[idx, "importe"] = float(row["kg_stock"]) * float(price)
        else:
            warehouse.at[idx, "importe"] = None

    warehouse["€/kg"] = pd.to_numeric(warehouse["€/kg"], errors="coerce")
    warehouse["importe"] = pd.to_numeric(warehouse["importe"], errors="coerce")
    return warehouse[["articulo", "descripcion", "lote", "kg_stock", "€/kg", "importe"]]


st.markdown(
    """
    <div class="hero">
        <div class="hero-title">📦 Valorización dinámica de stock</div>
        <div class="hero-subtitle">
            El Excel base aporta el precio €/kg por artículo/lote. El Excel de almacén
            no se modifica: solo se le añade el precio correcto y el importe calculado.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("¿Cómo funciona esta valorización?", expanded=False):
    st.markdown(
        """
        **1. Excel base (precios)** — Contiene, para cada `artículo` (y opcionalmente `lote`), el precio **€/kg** vigente.
        Se actualiza mes a mes; es la única fuente de precios.

        **2. Excel de almacén (stock)** — Contiene los `artículo`/`lote` con su cantidad en **kg**. Este archivo
        **no se modifica**: se lee tal cual, sin tocar sus cantidades ni sus filas.

        **3. Cruce automático** — Para cada fila del almacén, la app busca el precio:
        - primero por coincidencia exacta de **artículo + lote**,
        - si no existe ese lote en el base, usa el precio general del **artículo** (sin lote).

        **4. Cálculo del importe** — `importe = kg_stock × €/kg`. El resultado es una tabla nueva
        (almacén + precio + importe) que puedes descargar en Excel; el archivo de almacén original queda intacto.

        Si un artículo/lote del almacén no aparece en el Excel base, el precio queda vacío y se marca para que lo revises.

        > **Nota sobre Lote vs Nº de Serie/SSCC**: el `LoteInterno` es el lote real usado para el cruce de precios.
        > El Nº de Serie/SSCC identifica el bulto o palet físico, no el lote, así que se ignora en la valorización.
        """
    )

with st.sidebar:
    st.header("📂 Archivos")
    st.caption("Excel base → precios €/kg (se actualiza mes a mes)")
    base_file = st.file_uploader("Excel base", type=["xlsx", "xls"], label_visibility="collapsed")
    st.caption("Excel de almacén → artículos/lotes + stock en kg")
    warehouse_file = st.file_uploader("Excel de almacén", type=["xlsx", "xls"], label_visibility="collapsed")
    st.markdown(
        '<div class="info-box">Si no subes archivos, se usan los ejemplos incluidos en la carpeta <code>data</code>.</div>',
        unsafe_allow_html=True,
    )

if base_file is None and DEFAULT_BASE_FILE.exists():
    base_path = DEFAULT_BASE_FILE
else:
    base_path = base_file

if warehouse_file is None and DEFAULT_WAREHOUSE_FILE.exists():
    warehouse_path = DEFAULT_WAREHOUSE_FILE
else:
    warehouse_path = warehouse_file

try:
    if base_path is None:
        st.warning("No se ha encontrado un Excel base. Sube uno o crea el ejemplo en la carpeta data.")
        st.stop()
    if warehouse_path is None:
        st.warning("No se ha encontrado un Excel de almacén. Sube uno o crea el ejemplo en la carpeta data.")
        st.stop()

    base_df = load_excel(base_path)
    warehouse_df = load_excel(warehouse_path)

    warehouse_columns = detect_warehouse_columns(warehouse_df)
    with st.expander("Columnas detectadas en el Excel de almacén", expanded=not warehouse_columns["kg_stock"]):
        st.write(
            {
                "Artículo": warehouse_columns["articulo"] or "⚠️ no detectada",
                "Descripción": warehouse_columns["descripcion"] or "— no detectada (opcional)",
                "Lote": warehouse_columns["lote"] or "— no detectada (opcional)",
                "Peso/Cantidad (kg)": warehouse_columns["kg_stock"] or "⚠️ no detectada",
            }
        )
        st.caption("Columnas originales del archivo: " + ", ".join(str(c) for c in warehouse_df.columns))

    base_ready = prepare_base_dataframe(base_df)
    warehouse_ready = prepare_warehouse_dataframe(warehouse_df, warehouse_columns)
    result = valorate_stock(base_ready, warehouse_ready)

    total_importe = float(result["importe"].sum()) if "importe" in result.columns else 0.0
    total_kg = float(result["kg_stock"].sum()) if "kg_stock" in result.columns else 0.0
    sin_precio = int(result["€/kg"].isna().sum())

    st.markdown("### Resumen")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Stock total</div>'
            f'<div class="kpi-value">{total_kg:,.2f} kg</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Valor total</div>'
            f'<div class="kpi-value ok">{total_importe:,.2f} €</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        cls = "warn" if sin_precio else "ok"
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Líneas sin precio</div>'
            f'<div class="kpi-value {cls}">{sin_precio}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("### Resultado")
    if sin_precio:
        st.warning(f"{sin_precio} línea(s) del almacén no encontraron precio en el Excel base. Revísalas abajo (€/kg vacío).")

    def _highlight_missing(row):
        return ["background-color: rgba(244, 184, 96, 0.18)" if pd.isna(row["€/kg"]) else "" for _ in row]

    st.dataframe(result.style.apply(_highlight_missing, axis=1), use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="resultado")
    output.seek(0)

    st.download_button(
        label="Descargar resultado en Excel",
        data=output.getvalue(),
        file_name="resultado_valoracion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as exc:
    st.error(f"No se pudo procesar el archivo: {exc}")
