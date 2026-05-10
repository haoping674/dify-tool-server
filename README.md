# Dify Excel Tool Server

這是一個基於 **FastAPI** 的 Python Tool Server，目標是讓 Dify 離線版 / 實驗室環境可以透過 HTTP Tool 操作 Excel。

它適合用在：

- Dify Workflow 生成 Excel 報表
- Dify Agent 讀取 Excel 後分析資料
- 將外部開源 Python 能力包成內網 Tool
- 企業內部 AI 助手產生固定格式 `.xlsx` 檔案
- Excel CRUD API Prototype

> 補充：Apple Numbers 可以直接開啟 `.xlsx`，所以建議先生成 Excel，再交給 Numbers 開啟。

---

## 功能

### Workbook CRUD

- 建立 Excel workbook
- 上傳現有 `.xlsx`
- 列出 workbook
- 下載 workbook
- 刪除 workbook

### Sheet CRUD

- 列出 sheets
- 新增 sheet
- 重新命名 sheet
- 刪除 sheet

### Data CRUD

- 讀取 rows
- 新增 row
- 更新 row
- 刪除 row
- 讀取 cell
- 更新 cell
- 讀取 range
- 寫入 range

---

## 專案架構

```text
.
├── app
│   ├── api
│   │   └── excel.py
│   ├── core
│   │   └── config.py
│   ├── schemas
│   │   └── excel.py
│   ├── services
│   │   └── excel_service.py
│   └── main.py
├── examples
│   ├── create_workbook.json
│   └── dify_http_node_examples.md
├── tests
│   └── test_excel_api.py
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── README.md
```

---

## 本機啟動

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Windows PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

啟動後檢查：

```text
http://localhost:8000/health
```

Swagger API 文件：

```text
http://localhost:8000/docs
```

OpenAPI JSON：

```text
http://localhost:8000/openapi.json
```

---

## Docker 啟動

```bash
docker compose up -d --build
```

檢查：

```bash
curl http://localhost:8000/health
```

---

## API 範例

### 1. 建立 Workbook

```bash
curl -X POST http://localhost:8000/api/v1/excel/workbooks \
  -H "Content-Type: application/json" \
  -d @examples/create_workbook.json
```

回傳：

```json
{
  "workbook_id": "...",
  "filename": "alarm_report.xlsx",
  "download_url": "/api/v1/excel/workbooks/.../download"
}
```

### 2. 讀取 rows

```bash
curl "http://localhost:8000/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/rows?start_row=1&limit=50"
```

### 3. 新增 row

```bash
curl -X POST http://localhost:8000/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/rows \
  -H "Content-Type: application/json" \
  -d '{"values":["2026-05-03","CDA-02","CDA","Low",1]}'
```

### 4. 更新 cell

```bash
curl -X PUT http://localhost:8000/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/cells/E2 \
  -H "Content-Type: application/json" \
  -d '{"value":10}'
```

### 5. 讀取 range

```bash
curl "http://localhost:8000/api/v1/excel/workbooks/{workbook_id}/sheets/Alarms/range?range_ref=A1:E10"
```

### 6. 下載 workbook

```bash
curl -L -o alarm_report.xlsx http://localhost:8000/api/v1/excel/workbooks/{workbook_id}/download
```

---

## Dify 離線版怎麼接

### 方法 A：Workflow 使用 HTTP Request 節點

1. Dify Workflow 新增 `HTTP Request` 節點
2. Method 選 `POST` / `GET` / `PUT` / `DELETE`
3. URL 填 Tool Server 內網位址
4. Body 填 JSON
5. 將上一步 LLM 輸出的結構化資料傳入 Body

例如：

```text
User
  ↓
LLM：整理成 Excel JSON
  ↓
HTTP Request：POST /api/v1/excel/workbooks
  ↓
Answer：回傳 download_url
```

### 方法 B：Agent 使用 Tool

如果你的 Dify 版本支援自訂 Tool / OpenAPI Tool，可以直接匯入：

```text
http://你的-tool-server:8000/openapi.json
```

然後讓 Agent 自己選擇：

- 建立 Excel
- 新增 row
- 修改 cell
- 下載檔案

---

## 實驗室離線部署建議架構

```text
[Dify]
  ↓ HTTP Tool / Agent Tool
[FastAPI Excel Tool Server]
  ↓
[openpyxl]
  ↓
[storage/*.xlsx]
```

如果 Dify 和 Tool Server 都是 Docker，建議放在同一個 docker network。

---

## 安全注意事項

目前這個專案是 Prototype，已經有基本防護：

- workbook_id 格式限制
- 檔名 sanitize
- 上傳大小限制
- 只支援 `.xlsx`
- 防止直接用路徑讀寫任意檔案

正式企業環境建議再加：

- API Key
- 使用者權限
- 檔案生命週期清理
- 操作審計 log
- 檔案病毒掃描
- Nginx 反向代理
- HTTPS

---

## 測試

```bash
pytest -q
```

---

## 後續可擴充

你可以繼續加：

- Excel 樣式：欄寬、顏色、字型、框線
- 圖表生成
- 樞紐分析表類似功能
- pandas 統計分析
- CSV 匯入/匯出
- PDF 報表輸出
- PowerPoint 報表輸出
- API Key 驗證
- Dify Plugin manifest
- MCP Server
