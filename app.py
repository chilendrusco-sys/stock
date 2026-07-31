import re
import unicodedata
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from fpdf import FPDF


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
        background: radial-gradient(circle at 10% 0%, #16233A 0%, #0F172A 45%, #0B1220 100%);
    }
    [data-testid="stSidebar"] {
        background: #0B1220;
        border-right: 1px solid #1E293B;
    }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    .hero {
        background: linear-gradient(120deg, #16233A 0%, #10231C 60%, #0F172A 100%);
        border: 1px solid #22334A;
        border-radius: 18px;
        padding: 22px 26px;
        margin-bottom: 18px;
        box-shadow: 0 6px 18px rgba(59, 130, 246, 0.15);
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        color: #E2E8F0;
        margin-bottom: 4px;
    }
    .hero-subtitle {
        color: #94A3B8;
        font-size: 0.98rem;
    }
    .kpi-card {
        background: #131C2E;
        border: 1px solid #22334A;
        border-left: 5px solid #3B82F6;
        border-radius: 14px;
        padding: 16px 18px;
        height: 100%;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }
    .kpi-label {
        color: #94A3B8;
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 6px;
    }
    .kpi-value {
        color: #E2E8F0;
        font-size: 1.6rem;
        font-weight: 800;
    }
    .kpi-value.warn { color: #F59E0B; }
    .kpi-value.ok { color: #34D399; }
    .info-box {
        background: #0F2A22;
        border: 1px solid #1E4B3B;
        border-radius: 12px;
        padding: 14px 16px;
        color: #6EE7B7;
        font-size: 0.9rem;
    }
    section[data-testid="stFileUploaderDropzone"] {
        background: #0F172A;
        border: 1px dashed #3B82F6;
        border-radius: 12px;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid #22334A;
        border-radius: 12px;
    }
    h3 { color: #E2E8F0 !important; }
    [data-testid="stSidebar"] .block-container {
        padding-top: 1.1rem;
    }
    [data-testid="stSidebar"] h2 {
        font-size: 1.05rem;
        margin-bottom: 2px;
        color: #E2E8F0;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        font-size: 0.78rem;
        margin-bottom: 2px;
    }
    [data-testid="stSidebar"] section[data-testid="stFileUploaderDropzone"] {
        padding: 6px;
        min-height: unset;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
        font-size: 0.75rem;
    }
    [data-testid="stSidebar"] .stButton button,
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] button {
        width: 100%;
        font-size: 0.85rem;
    }
    [data-testid="stSidebar"] .element-container {
        margin-bottom: 2px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_BASE_FILE = DATA_DIR / "base_stock.xlsx"
DEFAULT_WAREHOUSE_FILE = DATA_DIR / "warehouse_sample.xlsx"
CUSTOM_BASE_META = DATA_DIR / "custom_base_name.txt"

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

    # El fallback respeta el orden de prioridad de candidates, no el orden de columnas del Excel.
    for candidate in candidates:
        cand_norm = normalize_text(candidate)
        for col in df.columns:
            col_norm = normalize_text(col)
            if cand_norm in col_norm or col_norm in cand_norm:
                return col
    return None


def load_excel(path):
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception:
        # Reintenta con detección automática por si es un .xls antiguo u otro formato soportado.
        return pd.read_excel(path)


ALMACEN_CANDIDATES = ["almacen", "nombrealmacen", "almacen_nombre", "nombre almacen", "warehouse", "deposito", "centro", "planta", "ubicacion"]
ITEM_CANDIDATES = ["articulo", "articuloid", "producto", "producto_id", "codigo", "codigoarticulo", "referencia", "ref", "item", "article", "sku"]
# "LoteInterno" es el lote real. El "Nº de Serie"/SSCC identifica el bulto/palet, no el lote: se excluye a propósito.
LOTE_CANDIDATES = ["loteinterno", "lote_interno", "lote interno", "lote", "lot", "batch", "numerodelote", "numlote"]
# Variantes de "Nombre de Artículo" van antes que "descripcion" para no confundirlo con la descripción de especificación.
DESC_CANDIDATES = ["nombrearticulo", "nombre_articulo", "nombre articulo", "nombredearticulo", "nombre de articulo", "nombre_de_articulo", "nombredelarticulo", "nombre del articulo", "nombre_del_articulo", "descripcion", "denominacion", "nombre", "detalle", "desc", "descripciónarticulo", "textobreve"]
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
    return base[["articulo", "lote", "€/kg", "match_key", "match_key_item", "lote_key"]].drop_duplicates(subset=["match_key"], keep="first")


def detect_warehouse_columns(df: pd.DataFrame) -> dict:
    used_columns = set()

    def _pick(candidates):
        col = find_column(df, candidates)
        if col is None or col in used_columns:
            return None
        used_columns.add(col)
        return col

    return {
        "almacen": _pick(ALMACEN_CANDIDATES),
        "articulo": _pick(ITEM_CANDIDATES),
        "lote": _pick(LOTE_CANDIDATES),
        "descripcion": _pick(DESC_CANDIDATES),
        "kg_stock": _pick(QTY_CANDIDATES),
    }


def prepare_warehouse_dataframe(df: pd.DataFrame, columns: dict) -> pd.DataFrame:
    almacen_col = columns.get("almacen")
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
    if almacen_col:
        keep_cols.append(almacen_col)

    warehouse = df[keep_cols].copy()
    rename_map = {item_col: "articulo", qty_col: "kg_stock"}
    if lote_col:
        rename_map[lote_col] = "lote"
    if desc_col:
        rename_map[desc_col] = "descripcion"
    if almacen_col:
        rename_map[almacen_col] = "almacen"
    warehouse = warehouse.rename(columns=rename_map)

    if "lote" not in warehouse.columns:
        warehouse["lote"] = ""
    if "descripcion" not in warehouse.columns:
        warehouse["descripcion"] = ""
    if "almacen" not in warehouse.columns:
        warehouse["almacen"] = ""

    warehouse["articulo"] = warehouse["articulo"].fillna("")
    warehouse["lote"] = warehouse["lote"].fillna("")
    warehouse["descripcion"] = warehouse["descripcion"].fillna("")
    warehouse["almacen"] = warehouse["almacen"].fillna("")
    warehouse["kg_stock"] = pd.to_numeric(warehouse["kg_stock"], errors="coerce").fillna(0)
    warehouse["articulo_key"] = warehouse["articulo"].apply(normalize_text)
    warehouse["lote_key"] = warehouse["lote"].apply(normalize_text)
    warehouse["match_key"] = warehouse["articulo_key"] + "|" + warehouse["lote_key"]
    warehouse["match_key_item"] = warehouse["articulo_key"]
    return warehouse[["almacen", "articulo", "descripcion", "lote", "kg_stock", "match_key", "match_key_item"]]


def valorate_stock(base_df: pd.DataFrame, warehouse_df: pd.DataFrame) -> pd.DataFrame:
    base_lookup = base_df.set_index("match_key")
    # Artículos para los que el Excel base NO especifica lote: el precio se aplica a cualquier lote de ese artículo.
    # Si el base sí tiene lotes concretos para un artículo, un lote del almacén que no aparezca ahí queda sin precio (no se adivina).
    base_no_lote = (
        base_df[base_df["lote_key"] == ""]
        .drop_duplicates(subset=["match_key_item"], keep="first")
        .set_index("match_key_item")["€/kg"]
    )
    warehouse = warehouse_df.copy()

    warehouse["€/kg"] = None
    warehouse["importe"] = None

    for idx, row in warehouse.iterrows():
        if row["match_key"] in base_lookup.index:
            price = base_lookup.loc[row["match_key"], "€/kg"]
        elif row["match_key_item"] in base_no_lote.index:
            price = base_no_lote.loc[row["match_key_item"]]
        else:
            price = None

        warehouse.at[idx, "€/kg"] = price
        if pd.notna(price):
            warehouse.at[idx, "importe"] = float(row["kg_stock"]) * float(price)
        else:
            warehouse.at[idx, "importe"] = None

    warehouse["€/kg"] = pd.to_numeric(warehouse["€/kg"], errors="coerce")
    warehouse["importe"] = pd.to_numeric(warehouse["importe"], errors="coerce")
    return warehouse[["almacen", "articulo", "descripcion", "lote", "kg_stock", "€/kg", "importe"]]


def _pdf_safe(text) -> str:
    text = "" if text is None else str(text)
    text = text.replace("€", "EUR")
    return text.encode("latin-1", errors="replace").decode("latin-1")


def build_pdf_report(result_df: pd.DataFrame, almacen_name: str, total_kg: float, total_importe: float, sin_precio: int) -> bytes:
    col_widths = {
        "almacen": 28,
        "articulo": 38,
        "descripcion": 55,
        "lote": 25,
        "kg_stock": 22,
        "€/kg": 20,
        "importe": 24,
    }
    headers = ["Almacén", "Artículo", "Descripción", "Lote", "kg_stock", "€/kg", "Importe"]

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    def _print_header():
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, _pdf_safe("Informe de valorización de stock"), ln=1)
        pdf.set_font("Helvetica", "", 10)
        if almacen_name:
            pdf.cell(0, 6, _pdf_safe(f"Almacén: {almacen_name}"), ln=1)
        pdf.cell(0, 6, _pdf_safe(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=1)
        pdf.cell(
            0, 6,
            _pdf_safe(
                f"Stock total: {total_kg:,.2f} kg   |   Valor total: {total_importe:,.2f} EUR   |   Líneas sin precio: {sin_precio}"
            ),
            ln=1,
        )
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 9)
        for header, key in zip(headers, col_widths):
            pdf.cell(col_widths[key], 8, _pdf_safe(header), border=1)
        pdf.ln(8)
        pdf.set_font("Helvetica", "", 8)

    _print_header()

    for _, row in result_df.iterrows():
        if pdf.get_y() > 185:
            pdf.add_page()
            _print_header()
        pdf.cell(col_widths["almacen"], 7, _pdf_safe(row.get("almacen", ""))[:18], border=1)
        pdf.cell(col_widths["articulo"], 7, _pdf_safe(row.get("articulo", ""))[:24], border=1)
        pdf.cell(col_widths["descripcion"], 7, _pdf_safe(row.get("descripcion", ""))[:36], border=1)
        pdf.cell(col_widths["lote"], 7, _pdf_safe(row.get("lote", ""))[:16], border=1)
        kg_val = row.get("kg_stock")
        pdf.cell(col_widths["kg_stock"], 7, _pdf_safe(f"{kg_val:,.2f}" if pd.notna(kg_val) else ""), border=1, align="R")
        precio_val = row.get("€/kg")
        pdf.cell(col_widths["€/kg"], 7, _pdf_safe(f"{precio_val:,.2f}" if pd.notna(precio_val) else "—"), border=1, align="R")
        importe_val = row.get("importe")
        pdf.cell(col_widths["importe"], 7, _pdf_safe(f"{importe_val:,.2f}" if pd.notna(importe_val) else "—"), border=1, align="R")
        pdf.ln(7)

    return bytes(pdf.output())


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

        **3. Cruce automático** — Para cada fila del almacén, la app busca el precio por coincidencia exacta de
        **artículo + lote**. Solo usa el precio general del artículo (sin lote) cuando el Excel base **no tiene
        ningún lote definido para ese artículo**; si el base sí tiene lotes concretos y el lote del almacén no
        está entre ellos, el precio se deja vacío para que lo revises (no se asigna el precio de otro lote).

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

    if base_file is not None:
        suffix = Path(base_file.name).suffix or ".xlsx"
        for old in DATA_DIR.glob("custom_base.*"):
            old.unlink(missing_ok=True)
        (DATA_DIR / f"custom_base{suffix}").write_bytes(base_file.getvalue())
        CUSTOM_BASE_META.write_text(base_file.name, encoding="utf-8")

    saved_base_candidates = list(DATA_DIR.glob("custom_base.*"))
    saved_base_path = saved_base_candidates[0] if saved_base_candidates else None
    saved_base_name = CUSTOM_BASE_META.read_text(encoding="utf-8").strip() if CUSTOM_BASE_META.exists() else None

    if saved_base_path is not None:
        st.caption(f"✅ Base guardada: **{saved_base_name or saved_base_path.name}** (se usa aunque cierres y vuelvas a entrar)")
        if st.button("🗑️ Quitar base guardada", use_container_width=True):
            for old in DATA_DIR.glob("custom_base.*"):
                old.unlink(missing_ok=True)
            CUSTOM_BASE_META.unlink(missing_ok=True)
            st.rerun()

    st.caption("Excel de almacén → artículos/lotes + stock en kg")
    warehouse_file = st.file_uploader("Excel de almacén", type=["xlsx", "xls"], label_visibility="collapsed")
    st.markdown(
        '<div class="info-box">Si no subes archivos, se usan los ejemplos incluidos en la carpeta <code>data</code>.</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
    download_placeholder = st.empty()
    pdf_placeholder = st.empty()

if base_file is not None:
    base_path = base_file
elif saved_base_path is not None:
    base_path = saved_base_path
elif DEFAULT_BASE_FILE.exists():
    base_path = DEFAULT_BASE_FILE
else:
    base_path = None

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
    warehouse_file_id = getattr(warehouse_path, "name", str(warehouse_path))
    with st.expander("Columnas detectadas en el Excel de almacén", expanded=not warehouse_columns["kg_stock"]):
        st.caption("Si la detección automática se equivoca, elige manualmente la columna correcta.")
        column_options = ["— Ninguna —"] + list(warehouse_df.columns)

        def _column_selector(label, field, required=False):
            detected = warehouse_columns[field]
            default_index = column_options.index(detected) if detected in column_options else 0
            choice = st.selectbox(
                label + (" *" if required else ""),
                column_options,
                index=default_index,
                key=f"colsel_{field}_{warehouse_file_id}",
            )
            return None if choice == "— Ninguna —" else choice

        warehouse_columns = {
            "almacen": _column_selector("Almacén", "almacen"),
            "articulo": _column_selector("Artículo", "articulo", required=True),
            "descripcion": _column_selector("Descripción", "descripcion"),
            "lote": _column_selector("Lote", "lote"),
            "kg_stock": _column_selector("Peso/Cantidad (kg)", "kg_stock", required=True),
        }
        st.caption("Columnas originales del archivo: " + ", ".join(str(c) for c in warehouse_df.columns))

    with st.spinner("Procesando archivos y calculando valorización..."):
        base_ready = prepare_base_dataframe(base_df)
        warehouse_ready = prepare_warehouse_dataframe(warehouse_df, warehouse_columns)
        result = valorate_stock(base_ready, warehouse_ready)

    editor_key = f"editor_{getattr(base_path, 'name', str(base_path))}_{getattr(warehouse_path, 'name', str(warehouse_path))}"

    st.markdown("### Resultado (editable)")
    st.caption("Solo las filas sin precio (vacío o 0) son editables. Las que ya tienen precio quedan bloqueadas. El importe se recalcula al vuelo y no se toca el Excel de almacén original.")

    result = result.reset_index(drop=True)
    needs_review_mask = result["€/kg"].isna() | (result["€/kg"] == 0)

    if needs_review_mask.any():
        st.warning(f"{int(needs_review_mask.sum())} línea(s) sin precio válido (vacío o 0). Edítalas abajo en la columna €/kg.")

    locked_df = result[~needs_review_mask].copy()
    editable_df = result[needs_review_mask].copy()

    search_col, filter_col = st.columns([3, 2])
    with search_col:
        search_term = st.text_input("🔍 Buscar", placeholder="Filtra por artículo, descripción, lote o almacén...")
    with filter_col:
        only_missing = st.checkbox("Mostrar solo filas sin precio")
    if search_term or only_missing:
        st.caption("El filtro solo afecta lo que ves aquí; confirma tus ediciones antes de cambiarlo para no perder cambios sin guardar.")

    # Reservan el orden visual (resumen antes que la tabla) aunque se rellenen más abajo,
    # una vez que la tabla ya ha aplicado las ediciones del usuario.
    kpi_placeholder = st.container()
    table_placeholder = st.container()

    def _filter_view(df: pd.DataFrame) -> pd.DataFrame:
        if not search_term:
            return df
        term_norm = normalize_text(search_term)
        mask = pd.Series(False, index=df.index)
        for col in ["articulo", "descripcion", "lote", "almacen"]:
            if col in df.columns:
                mask = mask | df[col].apply(normalize_text).str.contains(term_norm, na=False)
        return df[mask]

    locked_view = _filter_view(locked_df) if not only_missing else locked_df.iloc[0:0]
    editable_view = _filter_view(editable_df)

    number_column_config = {
        "kg_stock": st.column_config.NumberColumn("kg_stock", format="%.2f"),
        "€/kg": st.column_config.NumberColumn("€/kg", format="%.2f", step=0.01),
        "importe": st.column_config.NumberColumn("importe", format="%.2f"),
    }

    with table_placeholder:
        if not editable_view.empty:
            edited_view = st.data_editor(
                editable_view,
                key=editor_key,
                use_container_width=True,
                disabled=["almacen", "articulo", "descripcion", "lote", "kg_stock", "importe"],
                column_config=number_column_config,
            )
            edited_view["€/kg"] = pd.to_numeric(edited_view["€/kg"], errors="coerce")
            edited_view["importe"] = edited_view["kg_stock"] * edited_view["€/kg"]
            editable_df.update(edited_view)
        elif search_term or only_missing:
            st.caption("Ninguna fila sin precio coincide con el filtro actual.")

        if not locked_view.empty:
            st.caption("Filas con precio ya asignado (bloqueadas):")
            st.dataframe(locked_view, use_container_width=True, column_config=number_column_config)
        elif not only_missing and search_term:
            st.caption("Ninguna fila con precio coincide con el filtro actual.")

    result = pd.concat([locked_df, editable_df]).sort_index()

    total_importe = float(result["importe"].sum()) if "importe" in result.columns else 0.0
    total_kg = float(result["kg_stock"].sum()) if "kg_stock" in result.columns else 0.0
    sin_precio = int((result["€/kg"].isna() | (result["€/kg"] == 0)).sum())

    with kpi_placeholder:
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

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="resultado")
    output.seek(0)

    almacen_values = result["almacen"].astype(str).str.strip()
    almacen_values = almacen_values[almacen_values != ""]
    almacen_name = almacen_values.iloc[0] if not almacen_values.empty else ""
    safe_almacen = re.sub(r'[<>:"/\\|?*]', "", almacen_name).strip()
    file_name = f"resultado_valoracion_{safe_almacen}.xlsx" if safe_almacen else "resultado_valoracion.xlsx"

    download_placeholder.download_button(
        label="⬇️ Descargar Excel",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    pdf_bytes = build_pdf_report(result, almacen_name, total_kg, total_importe, sin_precio)
    pdf_file_name = f"informe_valoracion_{safe_almacen}.pdf" if safe_almacen else "informe_valoracion.pdf"
    pdf_placeholder.download_button(
        label="🧾 Descargar informe PDF",
        data=pdf_bytes,
        file_name=pdf_file_name,
        mime="application/pdf",
    )
except Exception as exc:
    st.error(f"No se pudo procesar el archivo: {exc}")
