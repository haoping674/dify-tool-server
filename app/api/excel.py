from fastapi import APIRouter, File, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.schemas.excel import (
    CellResponse,
    CellUpdateRequest,
    RangeReadResponse,
    RangeWriteRequest,
    RowAppendRequest,
    RowsResponse,
    RowUpdateRequest,
    SheetCreateRequest,
    SheetRenameRequest,
    WorkbookCreateRequest,
    WorkbookInfo,
    WorkbookResponse,
)
from app.services import excel_service as svc

router = APIRouter(prefix="/excel", tags=["excel"])


def download_url(workbook_id: str) -> str:
    return f"{settings.api_prefix}/excel/workbooks/{workbook_id}/download"


@router.post("/workbooks", response_model=WorkbookResponse)
def create_workbook(payload: WorkbookCreateRequest):
    workbook_id, filename = svc.create_workbook(
        filename=payload.filename,
        sheet_name=payload.sheet_name,
        headers=payload.headers,
        rows=payload.rows,
    )
    return {"workbook_id": workbook_id, "filename": filename, "download_url": download_url(workbook_id)}


@router.get("/workbooks", response_model=list[WorkbookInfo])
def list_workbooks():
    return svc.list_workbooks()


@router.post("/workbooks/upload", response_model=WorkbookResponse)
async def upload_workbook(file: UploadFile = File(...)):
    workbook_id, filename = await svc.upload_workbook(file)
    return {"workbook_id": workbook_id, "filename": filename, "download_url": download_url(workbook_id)}


@router.get("/workbooks/{workbook_id}/download")
def download_workbook(workbook_id: str):
    path = svc.require_workbook_path(workbook_id)
    filename = svc.read_filename(workbook_id)
    return FileResponse(
        path=path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.delete("/workbooks/{workbook_id}")
def delete_workbook(workbook_id: str):
    svc.delete_workbook(workbook_id)
    return {"ok": True}


@router.get("/workbooks/{workbook_id}/sheets")
def list_sheets(workbook_id: str):
    return {"workbook_id": workbook_id, "sheets": svc.list_sheets(workbook_id)}


@router.post("/workbooks/{workbook_id}/sheets")
def create_sheet(workbook_id: str, payload: SheetCreateRequest):
    return {"workbook_id": workbook_id, "sheets": svc.create_sheet(workbook_id, payload.sheet_name)}


@router.patch("/workbooks/{workbook_id}/sheets/{sheet_name}")
def rename_sheet(workbook_id: str, sheet_name: str, payload: SheetRenameRequest):
    return {"workbook_id": workbook_id, "sheets": svc.rename_sheet(workbook_id, sheet_name, payload.new_sheet_name)}


@router.delete("/workbooks/{workbook_id}/sheets/{sheet_name}")
def delete_sheet(workbook_id: str, sheet_name: str):
    return {"workbook_id": workbook_id, "sheets": svc.delete_sheet(workbook_id, sheet_name)}


@router.get("/workbooks/{workbook_id}/sheets/{sheet_name}/rows", response_model=RowsResponse)
def read_rows(workbook_id: str, sheet_name: str, start_row: int = 1, limit: int = 50):
    return {
        "workbook_id": workbook_id,
        "sheet_name": sheet_name,
        "start_row": start_row,
        "rows": svc.read_rows(workbook_id, sheet_name, start_row, limit),
    }


@router.post("/workbooks/{workbook_id}/sheets/{sheet_name}/rows")
def append_row(workbook_id: str, sheet_name: str, payload: RowAppendRequest):
    row_index = svc.append_row(workbook_id, sheet_name, payload.values)
    return {"ok": True, "workbook_id": workbook_id, "sheet_name": sheet_name, "row_index": row_index}


@router.put("/workbooks/{workbook_id}/sheets/{sheet_name}/rows/{row_index}")
def update_row(workbook_id: str, sheet_name: str, row_index: int, payload: RowUpdateRequest):
    svc.update_row(workbook_id, sheet_name, row_index, payload.values)
    return {"ok": True, "workbook_id": workbook_id, "sheet_name": sheet_name, "row_index": row_index}


@router.delete("/workbooks/{workbook_id}/sheets/{sheet_name}/rows/{row_index}")
def delete_row(workbook_id: str, sheet_name: str, row_index: int):
    svc.delete_row(workbook_id, sheet_name, row_index)
    return {"ok": True, "workbook_id": workbook_id, "sheet_name": sheet_name, "row_index": row_index}


@router.get("/workbooks/{workbook_id}/sheets/{sheet_name}/cells/{cell}", response_model=CellResponse)
def read_cell(workbook_id: str, sheet_name: str, cell: str):
    return {"workbook_id": workbook_id, "sheet_name": sheet_name, "cell": cell, "value": svc.read_cell(workbook_id, sheet_name, cell)}


@router.put("/workbooks/{workbook_id}/sheets/{sheet_name}/cells/{cell}")
def update_cell(workbook_id: str, sheet_name: str, cell: str, payload: CellUpdateRequest):
    svc.update_cell(workbook_id, sheet_name, cell, payload.value)
    return {"ok": True, "workbook_id": workbook_id, "sheet_name": sheet_name, "cell": cell}


@router.get("/workbooks/{workbook_id}/sheets/{sheet_name}/range", response_model=RangeReadResponse)
def read_range(workbook_id: str, sheet_name: str, range_ref: str):
    return {
        "workbook_id": workbook_id,
        "sheet_name": sheet_name,
        "range": range_ref,
        "values": svc.read_range(workbook_id, sheet_name, range_ref),
    }


@router.put("/workbooks/{workbook_id}/sheets/{sheet_name}/range")
def write_range(workbook_id: str, sheet_name: str, payload: RangeWriteRequest):
    svc.write_range(workbook_id, sheet_name, payload.start_cell, payload.values)
    return {"ok": True, "workbook_id": workbook_id, "sheet_name": sheet_name, "start_cell": payload.start_cell}
