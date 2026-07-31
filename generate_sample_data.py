from pathlib import Path
import pandas as pd

out_dir = Path(__file__).resolve().parent / 'data'
out_dir.mkdir(exist_ok=True)

base = pd.DataFrame([
    {'Articulo': 'Pollo', 'Lote': 'L001', '€/kg': 3.20},
    {'Articulo': 'Pollo', 'Lote': 'L002', '€/kg': 3.40},
    {'Articulo': 'Carne', 'Lote': 'L010', '€/kg': 5.10},
    {'Articulo': 'Arroz', 'Lote': '', '€/kg': 1.20},
])
base.to_excel(out_dir / 'base_stock.xlsx', index=False, engine='openpyxl')

warehouse = pd.DataFrame([
    {'Articulo': 'Pollo', 'Lote': 'L001', 'kg': 12.5},
    {'Articulo': 'Pollo', 'Lote': 'L002', 'kg': 8.0},
    {'Articulo': 'Carne', 'Lote': 'L010', 'kg': 4.5},
    {'Articulo': 'Arroz', 'Lote': '', 'kg': 20.0},
])
warehouse.to_excel(out_dir / 'warehouse_sample.xlsx', index=False, engine='openpyxl')
print('Sample Excel files created in', out_dir)
