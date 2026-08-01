# Stock valuation app

Aplicación en Streamlit para valorar stock a partir de dos archivos Excel:

1. **Excel base** — artículos/lotes con su precio €/kg vigente (se actualiza mes a mes).
2. **Excel de almacén** — artículos/lotes con su cantidad en stock (kg).

La app cruza ambos archivos y calcula, para cada línea de almacén, el precio €/kg correcto y el importe (`kg_stock × €/kg`). El Excel de almacén nunca se modifica: se lee tal cual y solo se le añaden las columnas de precio e importe en el resultado.

## 🌐 App desplegada

La app corre en Streamlit Community Cloud, desplegada desde este repo (rama `main`):

https://bfuuqc6rnwh76t8ijlrzte.streamlit.app

Cada `git push` a `main` provoca un redeploy automático.

## 🚀 Ejecutar en local (desarrollo)

```bash
pip install -r requirements.txt
python run_app.py
```

o directamente:

```bash
streamlit run app.py
```

Acceso protegido por contraseña (ver `APP_PASSWORD` en `app.py`).

## 📁 Estructura del proyecto

```
app.py                    # App Streamlit: UI + lógica de cruce de precios
generate_sample_data.py   # Genera los Excel de ejemplo en data/
run_app.py / run_app.bat  # Lanzador local del servidor Streamlit
requirements.txt          # Dependencias
.streamlit/config.toml    # Tema visual de la app
data/                     # Excel de ejemplo (base_stock.xlsx, warehouse_sample.xlsx)
tests/                    # Tests de la lógica de valorización (pytest)
```

## 🧮 Modelo de valorización: cómo se asigna el precio

Para cada fila del Excel de almacén, la app busca su precio €/kg en el Excel base siguiendo estas reglas, en este orden:

1. **Artículo + lote (obligatorio, sin excepciones).**
   Si el Excel base define uno o más lotes concretos para un artículo, el cruce exige coincidencia exacta de `artículo + lote` (y de `almacén`, si aplica el punto 2). No se admite ambigüedad: nunca se asigna a un lote el precio de otro lote del mismo artículo, aunque sean parecidos.
   - Solo se usa el precio general del artículo (sin lote) cuando el Excel base **no define ningún lote** para ese artículo.
   - Si el lote del almacén no aparece entre los que sí tiene definidos el base, la línea se deja **sin precio** para revisión manual — nunca se adivina.
   - `LoteInterno` es el campo que se usa como lote real. El `Nº de Serie`/SSCC identifica el bulto o palet físico, no el lote, así que se ignora a propósito en el cruce.

2. **Almacén (opcional).**
   Si el Excel base incluye una columna de almacén, el cruce exige además que el almacén coincida exactamente. Esto permite que el mismo `artículo + lote` tenga precios distintos según el almacén, sin que se trate como un error.

3. **Negocio (opcional, para conflictos).**
   Si tras aplicar 1 y 2 una misma combinación `artículo + lote (+ almacén)` tiene precios distintos en el base, y esos precios se explican porque pertenecen a negocios distintos, la app intenta resolverlo automáticamente usando las reglas negocio↔almacén (editables desde "🏷️ Reglas: qué negocio corresponde a cada almacén"). Si las reglas no bastan para decidir un único precio, la fila se deja pendiente de un validador manual explícito.

4. **Conflictos reales (no explicados por negocio).**
   Si el Excel base tiene la misma combinación `artículo + lote (+ almacén)` con precios distintos que ninguna regla explica, se usa el primero que aparece en el archivo y se avisa en pantalla con las filas exactas a corregir.

### ✅ Validación manual de precios pendientes

Las líneas del almacén sin precio (por lote no encontrado, o por conflicto de negocio no resuelto) aparecen en una tabla editable separada de la de precios ya asignados. Cada fila pendiente tiene:

- Un campo `€/kg` editable.
- Un checkbox **✅ Validar**.

Al escribir un precio válido (> 0) y marcar **✅ Validar**, esa fila pasa automáticamente a la tabla de precios ya asignados (bloqueada) y entra en el cálculo del resumen y las descargas. La validación se guarda en la sesión del navegador por combinación de archivos (base + almacén) mientras dure la sesión.

## ✅ Tests

```bash
pip install pytest
pytest tests/test_price_matching.py
```

Cubre la regla estricta de artículo + lote: un lote no incluido en el base no puede recibir el precio de otro lote del mismo artículo.
