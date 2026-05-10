# Dify HTTP Request Node Examples

假設 Tool Server 在實驗室內網位址：

```text
http://excel-tool-server:8000
```

如果 Dify 和 Tool Server 在同一個 docker compose network，可以用 container service name。
如果不在同一個 network，請改成實驗室內可連到的 IP，例如：

```text
http://192.168.1.20:8000
```

## 1. 建立 Excel

Method: `POST`

URL:

```text
http://excel-tool-server:8000/api/v1/excel/workbooks
```

Headers:

```json
{
  "Content-Type": "application/json"
}
```

Body:

```json
{
  "filename": "alarm_report.xlsx",
  "sheet_name": "Alarms",
  "headers": ["Date", "Equipment", "System", "Severity", "Count"],
  "rows": [
    ["2026-05-01", "CDA-01", "CDA", "High", 3],
    ["2026-05-02", "Chiller-01", "Chiller", "Medium", 5]
  ]
}
```

## 2. 讀取前 50 列

Method: `GET`

URL:

```text
http://excel-tool-server:8000/api/v1/excel/workbooks/{{workbook_id}}/sheets/Alarms/rows?start_row=1&limit=50
```

## 3. 新增一列

Method: `POST`

URL:

```text
http://excel-tool-server:8000/api/v1/excel/workbooks/{{workbook_id}}/sheets/Alarms/rows
```

Body:

```json
{
  "values": ["2026-05-03", "CDA-02", "CDA", "Low", 1]
}
```

## 4. 更新儲存格

Method: `PUT`

URL:

```text
http://excel-tool-server:8000/api/v1/excel/workbooks/{{workbook_id}}/sheets/Alarms/cells/E2
```

Body:

```json
{
  "value": 10
}
```

## 5. 下載 Excel

Method: `GET`

URL:

```text
http://excel-tool-server:8000/api/v1/excel/workbooks/{{workbook_id}}/download
```

Dify 若要把檔案給使用者，通常做法是：

1. Tool Server 回傳 `download_url`
2. Dify 回覆使用者「報表已生成」並附上連結
3. 如果是內網系統，可由你的前端或檔案服務代理下載

