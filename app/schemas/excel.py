from typing import Any
from pydantic import BaseModel, Field


class WorkbookCreateRequest(BaseModel):
    filename: str = Field(default="workbook.xlsx", description="Workbook filename. .xlsx will be appended if omitted.")
    sheet_name: str = Field(default="Sheet1")
    headers: list[str] | None = Field(default=None, description="Optional header row")
    rows: list[list[Any]] = Field(default_factory=list, description="Optional initial rows")


class WorkbookResponse(BaseModel):
    workbook_id: str
    filename: str
    download_url: str


class WorkbookInfo(BaseModel):
    workbook_id: str
    filename: str
    size_bytes: int
    modified_at: float


class SheetCreateRequest(BaseModel):
    sheet_name: str


class SheetRenameRequest(BaseModel):
    new_sheet_name: str


class CellUpdateRequest(BaseModel):
    value: Any


class CellResponse(BaseModel):
    workbook_id: str
    sheet_name: str
    cell: str
    value: Any


class RowAppendRequest(BaseModel):
    values: list[Any]


class RowUpdateRequest(BaseModel):
    values: list[Any]


class RowsResponse(BaseModel):
    workbook_id: str
    sheet_name: str
    start_row: int
    rows: list[list[Any]]


class RangeReadResponse(BaseModel):
    workbook_id: str
    sheet_name: str
    range: str
    values: list[list[Any]]


class RangeWriteRequest(BaseModel):
    start_cell: str = Field(..., examples=["A1"])
    values: list[list[Any]]
