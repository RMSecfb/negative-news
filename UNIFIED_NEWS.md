# 新版整合新聞工作台

每天台北時間 08:20，由 `.github/workflows/negative-news-ai.yml` 執行全市場事件搜尋，以及完整 S&P 500 公司清單的方法一、方法二。GitHub 排程可能延遲；新聞時間窗固定為前一天 08:20 到當天 08:20。

## 判讀流程

新聞擷取 → 同市場、同公司相同網址／相同標題去重 → 原有 FinBERT 與規則留存 → 選用 LLM 統一事件判讀 → 人工覆核優先 → 依新聞發布日期合併 CSV／Excel。

「AI LLM」是判讀開關。未開啟時使用原有規則；開啟時，沒有完成 AI 的新聞保留「待補分析」，不將規則結果冒充 AI 結果或無風險。只選 AI 時，先執行候選事件／公司搜尋。CDS 上升不代表公司已違約；未經確認的違約會要求人工覆核。

LLM 新判讀預設每次最多 35 筆，人工執行可調整 0–100。每日共享的程式上限預設 120 次，每次請求間隔 20 秒。這些是程式設定，實際供應商限制可能更低；遇到 429 後當次停止新增請求並保留待補分析。成功結果按新聞內容、模型、提示版本和規則版本快取，人工覆核與快取不占新增請求筆數。大量新聞可能需要多次補分析；「全市場搜尋」不保證涵蓋每一則新聞。

## 網站操作

- 查詢新聞期間：設定開始、結束日期及市場，支援跨日結果。直接編輯 Ticker、Level、事件中文、Action、覆核註記並儲存。Level 3 黃色、4 橘色、5 紅色。
- 人工執行新聞期間：指定日期與時間，可複選 AI LLM／方法一／方法二；選全市場或指定公司列表。Dow Jones 30、S&P 500、上傳列表可複選，自訂 CSV／Excel 至少包含 Ticker、Company，也可附 Market、Exchange、Aliases。港股使用自訂公司列表，原有美股 TradingView 逐公司擷取不套用港股。
- 只補分析既有新聞：不重新抓取，優先利用既有資料和快取。補跑會寫入同一批 `news_daily` 結果，保留人工修改。
- 執行紀錄：查看新判讀、沿用、待補分析與新聞來源異常。歷史擷取能力受新聞來源保留範圍限制。

Ticker 依公司清單、別名及每天更新的 Nasdaq 官方目錄核對；未確認或同名公司不猜代號。港股代號標準化為例如 `0700.HK`。

## 已有設定

- GitHub Actions secret：`GROQ_API_KEY`。
- Streamlit secrets：`GITHUB_TOKEN`、`GITHUB_REPO = "RMSecfb/negative-news"`，可選 `GITHUB_BRANCH = "main"`。沿用原先人工編輯所用 Token；需此 repository 的 Contents read/write，無須將 Groq Key 放到網站。
- 網站送出請求至 `news_requests/`，push 觸發 Actions；執行完成寫回 `news_daily/`、`news_runs/`、`news_cache/`、`news_usage/`。GitHub Actions 自帶的 Token 負責寫結果。網站缺少有效 Token 時仍可公開查詢，但不能送出與儲存。
- Excel 提供完整新聞、主要負面事件、待人工覆核、待補分析、執行紀錄。原有 `ai_news_daily`／`risk_news_daily` 在查詢時向後相容，不刪除舊資料。

## 開發驗證

`python -m unittest discover -s tests -v`

Streamlit 入口為 `負面新聞整合網站.py`，原有來源擷取、FinBERT、曝險及 13F 支援函式保留於 `legacy_engine.py`。離線測試可設定 `NEWS_LOCAL_REPO` 指向測試資料夾。API Key 與 Token 不可加入 Git。
