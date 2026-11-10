import io

import pandas as pd
import pytest
from docx import Document

from app.services.replace_service import run_replace


def _word_bytes(text: str = "【姓名】") -> bytes:
    doc = Document()
    doc.add_paragraph(text)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _excel_bytes(rows: list[dict[str, str]], columns: list[str] | None = None) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=columns).to_excel(buffer, index=False)
    return buffer.getvalue()


RULES = [{"keyword": "【姓名】", "excel_column": "姓名"}]


def test_empty_excel_returns_details_list():
    excel = _excel_bytes([], columns=["姓名", "文件"])
    result = run_replace(_word_bytes(), excel, RULES, 1, 10, "文件")
    assert result["total"] == 0
    assert result["details"] == []


def test_start_row_beyond_data_returns_empty_result():
    excel = _excel_bytes([{"姓名": "张三", "文件": "a"}])
    result = run_replace(_word_bytes(), excel, RULES, 10, 20, "文件")
    assert result["total"] == 0
    assert result["details"] == []


def test_missing_column_raises_value_error():
    excel = _excel_bytes([{"姓名": "张三", "文件": "a"}])
    with pytest.raises(ValueError, match="不存在"):
        run_replace(_word_bytes(), excel, [{"keyword": "【姓名】", "excel_column": "不存在"}], 1, 1, "文件")


def test_row_number_includes_header_row():
    excel = _excel_bytes([{"姓名": "张三", "文件": "a"}])
    result = run_replace(_word_bytes(), excel, RULES, 1, 1, "文件")
    assert result["details"][0]["row_number"] == 2


def test_duplicate_filenames_are_deduplicated():
    excel = _excel_bytes(
        [
            {"姓名": "张三", "文件": "同名"},
            {"姓名": "李四", "文件": "同名"},
        ]
    )
    result = run_replace(_word_bytes(), excel, RULES, 1, 2, "文件")
    names = [detail["file_name"] for detail in result["details"]]
    assert names == ["同名.docx", "同名_1.docx"]


def test_empty_cell_replaces_with_blank_by_default():
    excel = _excel_bytes([{"姓名": "", "文件": "a"}])
    result = run_replace(_word_bytes(), excel, RULES, 1, 1, "文件")
    output = Document(io.BytesIO(result["files"][0].data.getvalue()))
    assert output.paragraphs[0].text == ""


def test_empty_cell_can_keep_placeholder():
    excel = _excel_bytes([{"姓名": "", "文件": "a"}])
    result = run_replace(
        _word_bytes(),
        excel,
        RULES,
        1,
        1,
        "文件",
        keep_placeholder_on_empty=True,
    )
    output = Document(io.BytesIO(result["files"][0].data.getvalue()))
    assert output.paragraphs[0].text == "【姓名】"


def test_sheet_selection_reads_named_sheet():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        pd.DataFrame([{"姓名": "第一页", "文件": "a"}]).to_excel(writer, sheet_name="Sheet1", index=False)
        pd.DataFrame([{"姓名": "第二页", "文件": "b"}]).to_excel(writer, sheet_name="Sheet2", index=False)
    result = run_replace(_word_bytes(), buffer.getvalue(), RULES, 1, 1, "文件", sheet_name="Sheet2")
    output = Document(io.BytesIO(result["files"][0].data.getvalue()))
    assert output.paragraphs[0].text == "第二页"


def test_filename_reserved_name_and_length():
    long_name = "A" * 300
    excel = _excel_bytes([{"姓名": "张三", "文件": "CON"}])
    result = run_replace(_word_bytes(), excel, RULES, 1, 1, "文件")
    assert result["details"][0]["file_name"] == "_CON.docx"

    excel_long = _excel_bytes([{"姓名": "张三", "文件": long_name}])
    result_long = run_replace(_word_bytes(), excel_long, RULES, 1, 1, "文件")
    stem = result_long["details"][0]["file_name"][: -len(".docx")]
    assert len(stem) <= 120
