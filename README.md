# 美股負面新聞整合中心

給風管部使用的美股負面新聞監控工具。自動抓取多方新聞來源、用 AI 判斷新聞情緒、依規則分類事件嚴重度，並整理成可下載的 Excel 報表與互動式圖表。

線上版本：以 Streamlit Community Cloud 部署，程式進入點為 `負面新聞整合網站.py`。

---

## 功能總覽

### 1. 新聞擷取（三種方法擇一）

| 方法 | 說明 |
|---|---|
| 方法一｜多網站擷取 | 彙整 MoneyDJ、經濟日報、鉅亨網、CNBC、TradingView |
| 方法二｜Google News | 彙整 Google News RSS 英文新聞，並執行 FinBERT 情緒分析 |
| 方法一＋方法二｜完整整合 | 依序執行兩種方法、合併去除重複新聞，再執行 FinBERT 分析 |

抓取範圍可選擇：
- **Dow Jones 30**：讀取 repo 內的 `DowJones30.xlsx`
- **S&P 500**：讀取 repo 內的 `SP500.xlsx`
- **上傳公司列表**：另外上傳一份自訂 Excel 名單（存成 `Company_List_Custom.xlsx`），與前兩者互不影響

### 2. FinBERT 情緒評分

- 用 FinBERT 模型計算每則新聞的情緒分數（positive − negative 機率）
- **原始標題含中文字元的新聞不會送進模型評分**，一律標記為 `N/A`——因為 FinBERT 只認得英文，硬評中文標題分數並不準確
- FinBERT ≤ 0 的新聞會被視為情緒偏負面

### 3. 負面事件分類

依 `Parameter_Event.xlsx` 內建的規則表，把新聞分類成事件類型（如破產、下市、重大訴訟、監管裁罰、資料外洩、天災、分析師降評…）並給出 Level 1～5 嚴重度，同時附上建議處理動作（Action）與事件中文說明。

### 4. 負面新聞總覽（KPI、圖表、明細）

- KPI：總新聞則數、負面新聞則數、重大事件（Level ≥ 4）則數、涉及公司數
- 事件類型總占比（甜甜圈圖）
- 公司負面新聞排行（長條圖，前 10 名）
- 可篩選、可排序的負面新聞明細表，可下載完整整合新聞 / 負面新聞 / FinBERT ≤ 0 三種 Excel

### 5. 曝險加權（選填功能）

上傳一份「Ticker + 曝險金額」的 Excel（格式需求見下方），系統會：
- 把曝險金額併入負面新聞明細
- 畫出**曝險風險象限圖**：X 軸＝曝險金額、Y 軸＝事件嚴重度（最高 Level × 負面新聞則數）、泡泡大小＝負面新聞則數、顏色＝最高 Level
- 用中位數畫出十字虛線，右上象限（曝險高＋事件嚴重）為優先處理對象

**曝險清單 Excel 格式**：需要 `Ticker`/`Symbol`（或「股票代號」）與 `Exposure`/`Amount`/`曝險金額`/`曝險`/`部位金額`/`部位` 其中一組欄位（大小寫、順序不拘）。

### 6. 新聞頻率異常偵測 ＋ 時間序列走勢圖

- 自動比較「今天 vs. 過去每日平均」，超過 2 倍、或平常沒有負面新聞卻突然出現 3 則以上，就會列入「🔺 異常放量預警」
- 下拉選單可選「全市場」或特定公司，畫出負面新聞則數隨時間變化的折線圖
- **依賴歷史資料累積**，天數越多，判斷越準確（見下方「歷史資料持久化」說明）

---

## Repo 檔案結構

```
negative-news/
├─ 負面新聞整合網站.py        # 主程式（Streamlit app）
├─ Parameter_Event.xlsx      # 負面事件分類規則（必要）
├─ DowJones30.xlsx           # 道瓊 30 公司名單（必要）
├─ SP500.xlsx                # S&P 500 公司名單（必要）
├─ requirements.txt          # Python 套件需求
├─ runtime.txt                # Python 版本設定
├─ history/
  └─ negative_news_history.csv   # 每日負面新聞歷史統計（程式自動寫入，見下方說明）
```

`Parameter_Event.xlsx`、`DowJones30.xlsx`、`SP500.xlsx` 這三份 Excel 檔**必須放在 repo 根目錄**，程式啟動時會自動讀取；缺任何一份都會在網頁上顯示錯誤並停止執行。更新公司名單或分類規則，只要覆蓋這幾個檔案、重新整理網站即可套用，不需要改程式碼。

---

## 部署設定（Streamlit Community Cloud）

### 1. 基本需求

- `requirements.txt` 需包含：`streamlit`、`pandas`、`requests`、`beautifulsoup4`、`openpyxl`、`plotly`、`torch`、`transformers`
- `runtime.txt` 指定 Python 版本

### 2. Secrets（選填，用於歷史資料保留）

Streamlit Community Cloud 的儲存空間**不是持久化的**：容器重開、休眠喚醒、或重新部署程式碼，都會清空檔案系統，讓「新聞頻率異常偵測」的歷史天數歸零重新累積。

為了解決這個問題，程式支援把每日負面新聞統計寫回 GitHub repo 本身（`history/negative_news_history.csv`），不需要額外的資料庫服務。要啟用這個功能，到 **App → Manage app → Settings → Secrets** 貼上：

```toml
GITHUB_TOKEN = "ghp_xxxxxxxxxxxxxxxxxxxx"
GITHUB_REPO = "chengyu1212/negative-news"
```

- `GITHUB_TOKEN`：GitHub Personal Access Token（Fine-grained token），權限只需要對這個 repo 開 **Contents → Read and write**
- `GITHUB_REPO`：格式為 `帳號/repo名稱`

選填（有預設值，通常不需要改）：
- `GITHUB_BRANCH`（預設 `main`）
- `GITHUB_HISTORY_PATH`（預設 `history/negative_news_history.csv`）

**沒有設定這兩個 Secrets 也完全可以正常使用**——異常偵測／趨勢圖會自動退回讀本機暫存資料夾，只是天數會因容器重開而歸零，其他功能不受影響。畫面上「新聞頻率異常偵測與時間序列走勢」那段文字會顯示目前讀取的資料來源，用來確認有沒有接上。

---

## 使用步驟

1. **選擇公司列表**：Dow Jones 30 / S&P 500 / 上傳自訂名單
2. **設定擷取期間**：起訖日期與時間（台北時間）
3. **選擇擷取方法**：方法一、方法二、或兩者合併
4. （選填）**上傳曝險清單**：套用後可看到曝險風險象限圖
5. **開始執行**：可即時查看進度、累計執行時間，執行中可安全停止
6. 完成後在「負面新聞總覽」查看 KPI、圖表、異常警示、趨勢圖，並可篩選、下載明細 Excel

---

## 已知限制

- FinBERT 僅支援英文，原始標題含中文字元的新聞一律標記 `N/A`，不會被誤判為負面
- 若未設定 GitHub Secrets，歷史統計資料會因 Streamlit Cloud 容器重開而不定期歸零
- 新聞頻率異常偵測至少需要 2 天歷史資料才會啟動比較邏輯

## 之後可能的擴充方向

- 覆核工作流（狀態追蹤）：記錄每則負面新聞「誰看過、處理到哪一步」，目前尚未實作
- 自動通知機制：Level ≥ 4 事件主動推播到 Email / Teams / Slack
