import io
import json

import pandas as pd
from docx import Document
from fastapi.testclient import TestClient

from app.main import app

RULES_JSON = json.dumps([{"keyword": "【姓名】", "excel_column": "姓名"}])
WORD_MEDIA = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
EXCEL_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _word_bytes() -> bytes:
    doc = Document()
    doc.add_paragraph("【姓名】")
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _excel_bytes(rows: list[dict[str, str]], columns: list[str] | None = None) -> bytes:
    buffer = io.BytesIO()
    pd.DataFrame(rows, columns=columns).to_excel(buffer, index=False)
    return buffer.getvalue()


def _files(word: bytes | None = None, excel: bytes | None = None) -> dict:
    return {
        "word_file": ("template.docx", word or _word_bytes(), WORD_MEDIA),
        "excel_file": ("data.xlsx", excel or _excel_bytes([{"姓名": "张三", "文件": "a"}]), EXCEL_MEDIA),
    }


def test_health_and_ready():
    with TestClient(app) as client:
        assert client.get("/health").json() == {"status": "ok"}
        assert client.get("/health/ready").status_code == 200


def test_preview_returns_sheet_information():
    with TestClient(app) as client:
        response = client.post("/preview", files=_files())
    assert response.status_code == 200
    data = response.json()
    assert data["sheet_names"] == ["Sheet1"]
    assert data["selected_sheet"] == "Sheet1"
    assert data["excel_columns"] == ["姓名", "文件"]
    assert data["excel_total_rows"] == 1


def test_preview_rejects_fake_extension():
    with TestClient(app) as client:
        response = client.post("/preview", files=_files(word=b"not a docx"))
    assert response.status_code == 400


def test_execute_empty_excel_returns_zero_without_error():
    empty = _excel_bytes([], columns=["姓名", "文件"])
    with TestClient(app) as client:
        response = client.post(
            "/replace/execute",
            files=_files(excel=empty),
            data={
                "start_row": "1",
                "end_row": "5",
                "file_name_column": "文件",
                "rules_json": RULES_JSON,
            },
        )
    assert response.status_code == 200
    data = response.json()
    assert data["total"] == 0
    assert data["details"] == []


def test_execute_missing_column_returns_400():
    with TestClient(app) as client:
        response = client.post(
            "/replace/execute",
            files=_files(),
            data={
                "start_row": "1",
                "end_row": "1",
                "file_name_column": "文件",
                "rules_json": json.dumps([{"keyword": "【姓名】", "excel_column": "不存在"}]),
            },
        )
    assert response.status_code == 400
    assert "不存在" in response.json()["detail"]


def test_execute_and_export_zip_with_token():
    excel = _excel_bytes([{"姓名": "张三", "文件": "a"}, {"姓名": "李四", "文件": "b"}])
    with TestClient(app) as client:
        response = client.post(
            "/replace/execute",
            files=_files(excel=excel),
            data={
                "start_row": "1",
                "end_row": "2",
                "file_name_column": "文件",
                "rules_json": RULES_JSON,
            },
        )
        assert response.status_code == 200
        payload = response.json()
        assert payload["success"] == 2
        assert [detail["file_name"] for detail in payload["details"]] == ["a.docx", "b.docx"]

        zip_response = client.get(
            f"/export/zip/{payload['run_id']}",
            params={"token": payload["export_token"]},
        )
        assert zip_response.status_code == 200
        assert zip_response.content.startswith(b"PK")

        bad_token = client.get(
            f"/export/zip/{payload['run_id']}",
            params={"token": "bad-token"},
        )
        assert bad_token.status_code == 403


def test_preview_falls_back_when_sheet_is_missing():
    with TestClient(app) as client:
        response = client.post(
            "/preview",
            files=_files(),
            data={"sheet_name": "不存在的表"},
        )
    assert response.status_code == 200
    data = response.json()
    assert data["selected_sheet"] == "Sheet1"
    assert data["excel_columns"] == ["姓名", "文件"]
