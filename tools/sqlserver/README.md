# 匯入 SQL Server

把負面新聞每日結果增量匯入公司的 SQL Server。**在公司電腦執行，由公司電腦主動去 GitHub 拉資料**，不需要開放任何對內連線。

## 準備（一次性）

1. 安裝 Python 3.11 以上，以及 [Microsoft ODBC Driver 18 for SQL Server](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)（或 17）。
2. 取得這個資料夾（下載 repo 的 `tools/sqlserver/` 即可，或 `git clone`），安裝套件：
   ```
   pip install -r requirements.txt
   ```
3. 在 SQL Server 建一個資料庫（例如 `NegativeNews`），並讓執行帳號有建立資料表與寫入的權限。第一次執行會自動建立資料表。

## 執行

先確認會匯入哪些檔案（不連資料庫）：

```
python import_news.py --dry-run
```

正式匯入（Windows 驗證範例）：

```
python import_news.py --conn "DRIVER={ODBC Driver 18 for SQL Server};SERVER=主機名稱;DATABASE=NegativeNews;Trusted_Connection=yes;TrustServerCertificate=yes"
```

帳密驗證則把 `Trusted_Connection=yes` 換成 `UID=帳號;PWD=密碼`。**密碼不要寫進腳本或 repo**，改設定環境變數 `NEWS_DB_CONN` 放整串連線字串，執行時就不用帶 `--conn`。

第一次會匯入全部歷史（2026-08-09 起，約 9,500 列），之後每次只匯入有變動的檔案。

| 選項 | 說明 |
|---|---|
| `--since 2026-10-01` | 只匯入此日期之後的每日檔 |
| `--no-participants` | 不匯入 13F／Form 4 摘要 |
| `--force` | 忽略已匯入紀錄，全部重新匯入 |
| `--dry-run` | 只列出檔案，不連資料庫 |

## 排程

用 Windows 工作排程器，每天台北 10:00 之後執行一次（負面新聞約 09:15 跑完，機構內部人 06:00–07:00 已更新）。動作設定：

- 程式：`python`
- 引數：`C:\路徑\import_news.py`
- 起始位置：`C:\路徑\`

連線字串用環境變數 `NEWS_DB_CONN`。有任何檔案匯入失敗時，程式結束碼為 1，工作排程器會顯示失敗。

## 資料表

| 資料表 | 內容 |
|---|---|
| `dbo.NegativeNews` | 每日新聞，一列一則。第一欄 `SourceFile` 是來源每日檔，主鍵為 (`SourceFile`, `news_id`)。欄位沿用 CSV 欄名（含中文，查詢時用 `[ ]` 括起來）；空值存為 NULL |
| `dbo.NegativeNewsImportState` | 每個來源檔最後匯入的版本，用來判斷下次是否要重新匯入 |
| `dbo.Institutional13FSummary` | 13F 摘要，每次整張取代為最新版 |
| `dbo.InsiderForm4Summary` | Form 4 摘要，每次整張取代為最新版 |

型別：`日期` 為 DATE；`Published Time` 為 DATETIME2，**已轉成台北時間**；`Level`、`Rule_Level`、`input_tokens`、`output_tokens` 為 INT；`FinBERT`、`曝險金額` 為 FLOAT；其餘文字欄為 NVARCHAR。

CSV 之後若新增欄位，程式會自動在資料表加上對應欄位，不需手動改表。

## 使用上要知道的事

- **同一則新聞可能出現在兩天的檔案**（目前有 46 則，日期不同），所以 `news_id` 單獨不是唯一值。要一則只算一次，可建立視圖：
  ```sql
  CREATE OR ALTER VIEW dbo.vNegativeNewsDistinct AS
  SELECT * FROM (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY news_id ORDER BY [日期]) AS rn
    FROM dbo.NegativeNews) t
  WHERE rn = 1;
  ```
- **Level 0** 代表人工覆核判定無關而排除的新聞，要看風險事件請篩 `Level >= 3`（L3 以上為「主要負面事件」）。
- 較新的每日檔含 13F／Form 4 欄位（`13F 機構動向` 等），較舊的檔沒有，這些欄位在舊資料會是 NULL。
- 舊日期的檔案之後仍可能被修改（補分析、人工修改、離線覆核），程式以檔案版本判斷，改過就整檔取代，所以不會重複也不會漏更新。
- 沒有刪除機制：來源檔被整個移除時，資料庫裡的舊資料不會自動刪除。
