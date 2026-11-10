import io
import json

import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from fastapi.testclient import TestClient

from app.main import app
from app.services.replace_service import (
    RUN_CACHE,
    ReplacedFile,
    cache_run,
    sign_export_token,
)

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
        "excel_file": (
            "data.xlsx",
            excel or _excel_bytes([{"姓名": "张三", "文件": "a"}]),
            EXCEL_MEDIA,
        ),
    }


def _execute(client: TestClient, rows: int) -> dict:
    excel = _excel_bytes([{"姓名": f"名字{i}", "文件": f"file{i}"} for i in range(rows)])
    response = client.post(
        "/replace/execute",
        files=_files(excel=excel),
        data={
            "start_row": "1",
            "end_row": str(rows),
            "file_name_column": "文件",
            "rules_json": RULES_JSON,
        },
    )
    assert response.status_code == 200
    return response.json()


def test_config_endpoint():
    with TestClient(app) as client:
        response = client.get("/config")
    assert response.status_code == 200
    data = response.json()
    for key in (
        "app_version",
        "max_upload_size_mb",
        "max_excel_rows",
        "max_replace_rules",
        "max_preview_rows",
        "run_cache_ttl_seconds",
    ):
        assert key in data


def test_paginated_details_endpoint():
    with TestClient(app) as client:
        payload = _execute(client, 3)
        run_id = payload["run_id"]
        token = payload["export_token"]

        page1 = client.get(
            f"/export/details/{run_id}",
            params={"token": token, "page": 1, "page_size": 2},
        )
        assert page1.status_code == 200
        page1_data = page1.json()
        assert page1_data["total"] == 3
        assert page1_data["page"] == 1
        assert page1_data["page_size"] == 2
        assert len(page1_data["details"]) == 2
        assert page1_data["details"][0]["file_name"] == "file0.docx"

        page2 = client.get(
            f"/export/details/{run_id}",
            params={"token": token, "page": 2, "page_size": 2},
        )
        assert page2.status_code == 200
        page2_data = page2.json()
        assert page2_data["total"] == 3
        assert len(page2_data["details"]) == 1
        assert page2_data["details"][0]["file_name"] == "file2.docx"

        bad_token = client.get(
            f"/export/details/{run_id}",
            params={"token": "bad-token", "page": 1, "page_size": 2},
        )
        assert bad_token.status_code == 403

        bad_page = client.get(
            f"/export/details/{run_id}",
            params={"token": token, "page": 0, "page_size": 2},
        )
        assert bad_page.status_code == 400

        bad_page_size = client.get(
            f"/export/details/{run_id}",
            params={"token": token, "page": 1, "page_size": 201},
        )
        assert bad_page_size.status_code == 400


def test_delete_failed_row():
    result = {
        "files": [
            ReplacedFile(
                item_id="ok-1",
                filename="ok.docx",
                data=io.BytesIO(b"doc"),
                row_idx=0,
                replace_count=1,
            )
        ],
        "total": 2,
        "success": 1,
        "failed": 1,
        "replacements": 1,
        "details": [
            {
                "item_id": "ok-1",
                "seq": 1,
                "row_number": 2,
                "file_name": "ok.docx",
                "status": "成功",
                "replace_count": 1,
                "message": "",
            },
            {
                "item_id": "failed-1",
                "seq": 2,
                "row_number": 3,
                "file_name": "failed.docx",
                "status": "失败",
                "replace_count": 0,
                "message": "boom",
            },
        ],
    }
    run_id = cache_run(result)
    try:
        with TestClient(app) as client:
            response = client.delete(
                f"/export/result/{run_id}/failed-1",
                params={"token": sign_export_token(run_id)},
            )
        assert response.status_code == 200
        data = response.json()
        assert data["deleted"] is True
        assert data["total"] == 1
        assert data["success"] == 1
        assert data["failed"] == 0
        assert [detail["item_id"] for detail in data["details"]] == ["ok-1"]
        assert data["details"][0]["seq"] == 1
    finally:
        RUN_CACHE.pop(run_id, None)


def test_preview_shows_content_control_text():
    doc = Document()
    paragraph = doc.add_paragraph("普通文本")
    sdt = parse_xml(
        f'<w:sdt {nsdecls("w")}>'
        "<w:sdtPr/>"
        "<w:sdtContent>"
        "<w:r><w:t>内容控件文字</w:t></w:r>"
        "</w:sdtContent>"
        "</w:sdt>"
    )
    paragraph._p.append(sdt)
    buffer = io.BytesIO()
    doc.save(buffer)

    with TestClient(app) as client:
        response = client.post("/preview", files=_files(word=buffer.getvalue()))
    assert response.status_code == 200
    assert "内容控件文字" in response.json()["word_text"]
