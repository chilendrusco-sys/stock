# Stock valuation app

Aplicación en Streamlit para valorar stock a partir de dos archivos Excel:

1. Un Excel base con artículos/lotes y su precio por kg.
2. Un Excel de almacén con artículos/lotes y cantidades.

La app devuelve el precio €/kg correcto para cada artículo/lote y calcula el importe del stock.

## Ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py
```
