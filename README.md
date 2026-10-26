# 美股負面新聞整合中心

自動抓取多方新聞來源，以 FinBERT 評估情緒、依規則與 LLM 判讀事件嚴重度（Level 1～5），整理成可查詢、可人工覆核、可下載 Excel 的整合工作台。

先讀這份即可；更細的設計與限制見文末「相關文件」。

---

## 運作方式

每天台北時間 08:20，`.github/workflows/negative-news-ai.yml` 對完整 S&P 500 公司清單執行方法一、方法二，並加上全市場事件搜尋。新聞時間窗固定為前一天 08:20 到當天 08:20。GitHub 排程常有延遲，因此另由 Cloudflare Worker（`risk-scheduler`）在 08:22 準時觸發，GitHub 排程保留為備援；晚到的備援排程會略過。

處理流程：

新聞擷取 → 同市場、同公司相同網址／相同標題去重 → FinBERT 與規則留存 → 選用 LLM 統一事件判讀 → 人工覆核優先 → 依新聞發布日期合併寫入 `news_daily/`（CSV 與 Excel）。

### 新聞擷取

| 方法 | 說明 |
|---|---|
| 方法一｜多網站擷取 | 彙整 MoneyDJ、經濟日報、鉅亨網、CNBC、TradingView |
| 方法二｜Google News | 彙整 Google News RSS 英文新聞，並執行 FinBERT |
| 全市場事件搜尋 | 以中英文風險查詢搜尋候選事件；不保證涵蓋每一則新聞 |

公司清單：`SP500.xlsx`、`DowJones30.xlsx`、`config/extra_watchlist.csv`，或人工執行時上傳自訂清單（至少含 Ticker、Company，可附 Market、Exchange、Aliases）。港股使用自訂公司列表，代號標準化為如 `0700.HK`；原有美股 TradingView 逐公司擷取不套用港股。

### 判讀

- **FinBERT**：計算每則新聞的情緒分數（正面機率減負面機率）。原始中文標題會先翻譯再評分，原標題另外保留。它只是情緒訊號，不直接決定 Level。
- **規則**：依 `Parameter_Event.xlsx`（事件代碼、中文名稱、Level、Action、關鍵字、排除詞）比對，結果存於 `Rule_*` 欄位供稽核。
- **LLM**（Groq `openai/gpt-oss-20b`）：讀標題與摘要，回傳是否相關、辨識公司、事件代碼、證據與中文摘要，再由程式依規則表對應 Level 與 Action。「AI LLM」是開關；未開啟時使用規則結果。
- **人工覆核優先**：已人工覆核的新聞不會被排程或補跑覆蓋。
- **待補分析**：LLM 額度不足、429 或 API 失敗時，新聞保留為「待補分析」，Level 空白；它不等於無風險，也不能當作 Level 0。
- **違約防呆**：CDS 利差擴大不代表違約。未經確認的違約、L4／L5 重大事件、公司主體無法核對時，都會要求人工覆核。
- **Ticker 核對**：依公司清單、別名與每日更新的 Nasdaq 官方目錄核對；未確認或同名公司不猜代號。

### LLM 用量

| 項目 | 設定 |
|---|---|
| 每日排程每次判讀上限 | 100 筆（`DAILY_MAX_LLM`） |
| 手動執行預設 | 35 筆；網站人工執行可調 0–100 |
| 每日共享上限 | 120 筆（`LLM_DAILY_MAX`） |
| 請求間隔 | 20 秒 |

這些是程式設定，供應商實際限制可能更低。遇到 429 後本次停止新增請求並保留待補分析。成功結果依新聞內容、模型、提示版本與規則版本快取；人工覆核與快取不占新增請求筆數。大量積壓需多次補分析，或改走「離線覆核」。

---

## 網站操作

Streamlit 入口為 `負面新聞整合網站.py`。

- **查詢新聞期間**：設定開始、結束日期與市場，支援跨日。可直接編輯 Ticker、Level、事件中文、Action、覆核註記並儲存。Level 3 黃、4 橘、5 紅。
- **人工執行新聞期間**：指定日期與時間，可複選 AI LLM／方法一／方法二，範圍可選全市場或指定公司清單（Dow Jones 30、S&P 500、上傳清單可複選）。單次最多 31 天；歷史擷取受新聞來源保留範圍限制。
- **只補分析既有新聞**：不重新抓取，優先使用既有資料與快取，補跑寫入同一批 `news_daily` 並保留人工修改。
- **機構內部人交易**：13F 與 Form 4 整合頁，見 [INSTITUTIONAL_INSIDER.md](INSTITUTIONAL_INSIDER.md)。
- **執行紀錄**：查看新判讀、沿用、待補分析與新聞來源異常。
- **下載**：Excel 含「完整新聞、主要負面事件（L3–5）、待人工覆核、待補分析、執行紀錄」五張表。
- **曝險加權（選填）**：上傳含 `Ticker`／`Symbol`／股票代號，以及 `Exposure`／`Amount`／`曝險金額`／`曝險`／`部位金額`／`部位` 其中一欄的 Excel，系統會併入曝險並繪製曝險風險象限圖（X 曝險金額、Y 事件嚴重度、泡泡大小為則數、顏色為最高 Level，以中位數十字線劃分優先處理區）。
- **新聞頻率異常偵測與走勢圖**：以查詢期間內最後有資料的發布日對比前面有資料日的平均，達規則門檻才提示；沒有資料的日子不假設為零。基準天數越多越準，需至少 2 天資料。

---

## 資料夾與檔案

| 位置 | 說明 |
|---|---|
| `news_daily/YYYY-MM-DD.csv`、`.xlsx` | **正式結果**，依台北時間的新聞發布日歸檔；排程與人工補跑共用。CSV 為可再讀取的逐筆資料，Excel 為閱讀成果 |
| `news_runs/` | 每次執行的結果、錯誤與最近 100 筆索引 |
| `news_requests/` | 網站送出的人工執行請求（push 後觸發 Actions） |
| `news_cache/`、`news_usage/` | LLM 成功答案月度快取、每日呼叫用量 |
| `news_audits/` | 歷史 Ticker 核對報告 |
| `review_inbox/` | 離線覆核收件匣（見下節） |
| `history/` | 原始歷史統計與舊版人工覆核紀錄（`negative_news_manual_overrides.csv`），程式仍會讀取，請勿刪除 |
| `institutional_insider/` | 機構內部人交易成果 |
| `config/` | 公司別名、發行人登錄、美股代號參照 |
| `pipeline/` | 現行主流程與覆核匯入程式 |
| `legacy_engine.py` | 原始擷取、FinBERT、曝險等函式，由新版調用 |
| `BackupCode/` | 歷史程式備份，不是部署入口 |
| `Parameter_Event.xlsx`、`SP500.xlsx`、`DowJones30.xlsx` | 事件規則與公司名單；缺任何一份程式會停止。覆蓋檔案即可更新，不需改程式 |

前期的 `risk_news_daily/`（舊版每日 Excel）、`ai_news_daily/`、`ai_news_raw/` 內容已全部併入 `news_daily/` 並移除，歷史仍可在 Git 紀錄中找回。

---

## 離線總覆核

對積壓的「待人工覆核」「待補分析」與疑似歸錯的新聞，可不使用 Groq 額度做離線覆核：

1. 產生覆核 Excel，線下確認，另存為 `原檔名_已確認.xlsx`。
2. 上傳到 `review_inbox/`。
3. 「離線覆核匯入」流程（`.github/workflows/review-import.yml`）自動寫回 `news_daily`，結果在 `review_inbox/done/`。

欄位、略過規則與手動檢查方式見 [review_inbox/README.md](review_inbox/README.md)。

---

## 部署與設定

### 套件與版本

`requirements.txt` 列出套件，`runtime.txt` 指定 Python 3.11。

### 密鑰（不可加入 Git）

| 位置 | 名稱 | 用途 |
|---|---|---|
| GitHub Actions secret | `GROQ_API_KEY` | LLM 判讀；不需放到網站 |
| Streamlit secrets | `GITHUB_TOKEN` | 網站送出人工執行請求與儲存人工修改；需此 repo 的 Contents read/write |
| Streamlit secrets | `GITHUB_REPO = "RMSecfb/negative-news"` | 目標 repo |
| Streamlit secrets（選填） | `GITHUB_BRANCH`（預設 `main`）、`GITHUB_HISTORY_PATH`（預設 `history/negative_news_history.csv`） | 分支與歷史統計路徑 |

缺少有效 Token 時，網站仍可公開查詢，但不能送出請求與儲存修改。Streamlit Community Cloud 的本機儲存不是永久的，重開或休眠後會清空，所以所有結果都寫回 GitHub。

### 開發驗證

```
python -m unittest discover -s tests -v
```

離線測試可設定 `NEWS_LOCAL_REPO` 指向測試資料夾。

---

## 限制

- 新聞來源有保留範圍，補跑不保證回溯完整 31 天；全市場搜尋不保證涵蓋每一則新聞。
- 同一事件被多家媒體報導會增加則數，頻率統計不等於獨立事件數。
- 新聞頻率異常偵測只看使用者目前查詢期間內已保存的新聞。

---

## 相關文件

- [CURRENT_ARCHITECTURE.md](CURRENT_ARCHITECTURE.md)：程式分工、排程流程、FinBERT／規則／LLM 的實際分工與用量限制。
- [INSTITUTIONAL_INSIDER.md](INSTITUTIONAL_INSIDER.md)：機構持股與 Form 4 整合。
- [review_inbox/README.md](review_inbox/README.md)：離線覆核匯入規則。
