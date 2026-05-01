import io
from dataclasses import dataclass

import pandas as pd

from app.services import clean_excel_types, get_replace_params, export_statistics_to_csv


@dataclass
class DummyFile:
    filename: str
    data: io.BytesIO
    row_idx: int
    log: str
    replace_count: int


def test_clean_excel_types_normalizes_columns_and_values():
    df = pd.DataFrame({1: [' a ', None], 'b': [' x ', ' y ']})
    out = clean_excel_types(df)
    assert '1' in out.columns
    assert out.iloc[0]['1'] == 'a'
    assert out.iloc[1]['1'] == ''


def test_get_replace_params_contains_rule_hash_and_counts():
    df = pd.DataFrame({'a': [1, 2]})
    params = get_replace_params(None, df, 1, 2, 'a', 'P-', '', [('k', 'a')])
    assert params['excel_rows'] == 2
    assert params['rule_count'] == 1
    assert 'rule_hash' in params


def test_export_statistics_to_csv_has_headers():
    files = [DummyFile('a.docx', io.BytesIO(b'abc'), 0, 'ok', 2)]
    csv_text = export_statistics_to_csv(files)
    assert '文件名' in csv_text
    assert 'a.docx' in csv_text
