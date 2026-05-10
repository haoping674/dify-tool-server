from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)


def test_health():
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"


def test_create_read_update_download_delete_workbook():
    create_payload = {
        "filename": "test.xlsx",
        "sheet_name": "Alarms",
        "headers": ["Date", "Equipment", "Count"],
        "rows": [["2026-05-01", "CDA-01", 3]],
    }
    create_response = client.post("/api/v1/excel/workbooks", json=create_payload)
    assert create_response.status_code == 200
    workbook_id = create_response.json()["workbook_id"]

    rows_response = client.get(f"/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/rows?start_row=1&limit=10")
    assert rows_response.status_code == 200
    assert rows_response.json()["rows"][0] == ["Date", "Equipment", "Count"]

    append_response = client.post(
        f"/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/rows",
        json={"values": ["2026-05-02", "CDA-02", 5]},
    )
    assert append_response.status_code == 200
    assert append_response.json()["row_index"] == 3

    update_response = client.put(
        f"/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/cells/C2",
        json={"value": 10},
    )
    assert update_response.status_code == 200

    cell_response = client.get(f"/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/cells/C2")
    assert cell_response.status_code == 200
    assert cell_response.json()["value"] == 10

    download_response = client.get(f"/api/v1/excel/workbooks/{workbook_id}/download")
    assert download_response.status_code == 200
    assert download_response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    delete_response = client.delete(f"/api/v1/excel/workbooks/{workbook_id}")
    assert delete_response.status_code == 200
