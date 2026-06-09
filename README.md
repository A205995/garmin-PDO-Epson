# Garmin PDO Epson 選型工具

離線單檔 Epson 機器人選型工具，包含型號查詢、條件推薦、工程確認、已購資訊與工程報告 PDF 匯出。

## 使用方式

直接開啟 `index.html` 即可使用。

## 資料匯出

目前型號資料已可匯出成 Excel：

- `outputs/garmin_pdo_epson_robot_database.xlsx`

重新產生 Excel：

```bash
python3 scripts/export_robot_database_excel.py
```

## 資料原則

下載連結一律導向 Epson 官方網頁或系列手冊，以避免型號與檔案不一致。
