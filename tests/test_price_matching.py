import sys
import types
from pathlib import Path

import pandas as pd


class _FakeSessionState(dict):
    def setdefault(self, key, default=None):
        if key not in self:
            self[key] = default
        return self[key]


class _FakeStreamlit(types.SimpleNamespace):
    session_state = _FakeSessionState(authenticated=True)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def set_page_config(self, *args, **kwargs):
        return None

    def markdown(self, *args, **kwargs):
        return None

    def stop(self):
        return None

    def warning(self, *args, **kwargs):
        return None

    def success(self, *args, **kwargs):
        return None

    def error(self, *args, **kwargs):
        return None

    def dataframe(self, *args, **kwargs):
        return None

    def button(self, *args, **kwargs):
        return False

    def caption(self, *args, **kwargs):
        return None

    def form(self, *args, **kwargs):
        return None

    def form_submit_button(self, *args, **kwargs):
        return False

    def columns(self, *args, **kwargs):
        return [None, None, None]

    def text_input(self, *args, **kwargs):
        return ""

    def checkbox(self, *args, **kwargs):
        return False

    def rerun(self):
        return None

    def expander(self, *args, **kwargs):
        return self

    def container(self, *args, **kwargs):
        return self

    def data_editor(self, *args, **kwargs):
        return None

    def empty(self, *args, **kwargs):
        return self

    def spinner(self, *args, **kwargs):
        return self

    def file_uploader(self, *args, **kwargs):
        return None

    def selectbox(self, *args, **kwargs):
        return None

    def header(self, *args, **kwargs):
        return None

    @property
    def sidebar(self):
        return self


class _FakeColumnConfig(types.SimpleNamespace):
    TextColumn = lambda *args, **kwargs: None
    NumberColumn = lambda *args, **kwargs: None


fake_streamlit = _FakeStreamlit()
fake_streamlit.column_config = _FakeColumnConfig()
sys.modules.setdefault("streamlit", fake_streamlit)

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app  # noqa: E402


def test_price_is_only_taken_for_exact_article_and_lot_match():
    base_df = pd.DataFrame(
        [
            {"Artículo": "Pollo", "Lote": "L001", "€/kg": 3.2},
            {"Artículo": "Pollo", "Lote": "", "€/kg": 2.2},
        ]
    )
    warehouse_df = pd.DataFrame(
        [
            {"Artículo": "Pollo", "Lote": "L001", "kg": 5.0},
            {"Artículo": "Pollo", "Lote": "L999", "kg": 3.0},
        ]
    )

    prepared_base = app.prepare_base_dataframe(base_df, base_file_id="test")
    warehouse_columns = app.detect_warehouse_columns(warehouse_df)
    prepared_warehouse = app.prepare_warehouse_dataframe(warehouse_df, warehouse_columns)
    result = app.valorate_stock(prepared_base, prepared_warehouse)

    assert result.loc[0, "€/kg"] == 3.2
    assert pd.isna(result.loc[1, "€/kg"])
