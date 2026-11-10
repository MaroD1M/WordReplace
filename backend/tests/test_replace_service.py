import io
import zipfile

import pandas as pd
from docx import Document

from app.services.replace_service import (
    ReplacedFile,
    _apply_filename_options,
    _build_filename,
    clean_filename,
    export_zip,
    replace_word_with_format,
    sign_export_token,
    verify_export_token,
)


def _build_sample_doc() -> bytes:
    doc = Document()
    paragraph = doc.add_paragraph()
    first = paragraph.add_run("Hello ")
    first.bold = True
    second = paragraph.add_run("World")
    second.italic = True

    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = "Cell 【姓名】"
    inner = cell.add_table(rows=1, cols=1)
    inner.cell(0, 0).text = "Inner 【姓名】"

    doc.sections[0].header.paragraphs[0].text = "Header 【姓名】"
    doc.sections[0].footer.paragraphs[0].text = "Footer 【姓名】"

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def test_replacement_preserves_run_formatting_and_spans_runs():
    row = pd.Series({"name": "Alice"})
    output, count = replace_word_with_format(
        _build_sample_doc(),
        row,
        [("Hello World", "name")],
    )

    assert count == 1
    result = Document(io.BytesIO(output.getvalue()))
    paragraph = result.paragraphs[0]
    assert paragraph.text == "Alice"
    assert paragraph.runs[0].bold is True


def test_replacement_covers_nested_tables_and_headers_footers():
    row = pd.Series({"姓名": "张三"})
    output, count = replace_word_with_format(
        _build_sample_doc(),
        row,
        [("【姓名】", "姓名")],
    )

    assert count == 4
    result = Document(io.BytesIO(output.getvalue()))
    assert result.tables[0].cell(0, 0).paragraphs[0].text == "Cell 张三"
    assert result.tables[0].cell(0, 0).tables[0].cell(0, 0).text == "Inner 张三"
    assert result.sections[0].header.paragraphs[0].text == "Header 张三"
    assert result.sections[0].footer.paragraphs[0].text == "Footer 张三"


def test_replacement_is_counted_for_each_occurrence():
    doc = Document()
    doc.add_paragraph("【姓名】 and 【姓名】")
    buffer = io.BytesIO()
    doc.save(buffer)

    output, count = replace_word_with_format(
        buffer.getvalue(),
        pd.Series({"姓名": "李四"}),
        [("【姓名】", "姓名")],
    )

    assert count == 2
    result = Document(io.BytesIO(output.getvalue()))
    assert result.paragraphs[0].text == "李四 and 李四"


def test_build_filename_does_not_double_docx_extension():
    row = pd.Series({"name": "报告.docx"})
    assert _build_filename(row, "name", 0) == "报告.docx"


def test_build_filename_falls_back_when_column_missing():
    row = pd.Series({"other": "x"})
    assert _build_filename(row, "name", 4) == "文件_5.docx"


def test_clean_filename_replaces_illegal_characters():
    assert clean_filename('a/b:c*?"<>|.docx') == "a_b_c______.docx"


def test_apply_filename_options_formats():
    assert _apply_filename_options("a.docx", 3, True, "01", "", "") == "03_a.docx"
    assert _apply_filename_options("a.docx", 3, True, "0001", "", "") == "0003_a.docx"
    assert _apply_filename_options("a.docx", 3, True, "1.", "", "") == "3._a.docx"
    assert _apply_filename_options("a.docx", 3, True, "一", "", "") == "三_a.docx"
    assert _apply_filename_options("a.docx", 12, True, "一", "", "") == "十二_a.docx"
    assert _apply_filename_options("a.docx", 3, False, "1", "前缀", "后缀") == "前缀a后缀.docx"


def test_apply_filename_options_prefers_seq_column_value():
    assert _apply_filename_options("a.docx", 3, True, "1", "", "", seq_value="A-7") == "A-7_a.docx"


def test_export_token_roundtrip_and_expiry():
    token = sign_export_token("run-1", ttl_seconds=60)
    assert verify_export_token("run-1", token)
    assert not verify_export_token("run-1", token + "x")
    assert not verify_export_token("other-run", token)
    assert not verify_export_token("run-1", "not-a-token")

    expired = sign_export_token("run-1", ttl_seconds=-1)
    assert not verify_export_token("run-1", expired)


def test_export_zip_deduplicates_filenames():
    files = [
        ReplacedFile(item_id="1", filename="a.docx", data=io.BytesIO(b"one"), row_idx=0, replace_count=0),
        ReplacedFile(item_id="2", filename="a.docx", data=io.BytesIO(b"two"), row_idx=1, replace_count=0),
    ]

    buffer = export_zip(files)
    with zipfile.ZipFile(buffer) as archive:
        assert archive.namelist() == ["a.docx", "a_1.docx"]
        assert archive.read("a.docx") == b"one"
        assert archive.read("a_1.docx") == b"two"
