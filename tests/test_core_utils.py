import pandas as pd

from app.core_utils import (
    clean_filename,
    sanitize_cache_key,
    generate_safe_filename,
    get_replace_blockers,
    dedupe_filename,
)


def test_clean_filename_illegal_chars():
    assert clean_filename('a/b:c*?"<>|.docx').startswith('a_b_c')


def test_sanitize_cache_key_safe():
    value = sanitize_cache_key('../rule:name')
    assert '..' not in value
    assert '/' not in value


def test_generate_safe_filename_basic():
    row = pd.Series({'姓名': '张三'})
    out = generate_safe_filename(row, '姓名', file_prefix='NO-')
    assert out.endswith('.docx')
    assert 'NO-' in out


def test_get_replace_blockers():
    df = pd.DataFrame({'a': [1]})
    blockers = get_replace_blockers(None, df, [], 2, 1)
    assert '请先上传Word模板' in blockers
    assert '请至少添加1条替换规则' in blockers
    assert '起始行不能大于结束行' in blockers


def test_dedupe_filename():
    used = set()
    a = dedupe_filename('结果.docx', used)
    b = dedupe_filename('结果.docx', used)
    assert a == '结果.docx'
    assert b != a
    assert b.startswith('结果_')
