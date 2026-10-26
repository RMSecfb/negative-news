# 機構內部人交易整合

13F 與 Form 4 都在本 repo 的 `sec_filings/` 擷取與產出（`13F_FINAL.py`、`Form4_FINAL.py` 保留各自判讀架構，方法見 [sec_filings/METHODOLOGY.md](../sec_filings/METHODOLOGY.md)），新聞系統透過其成果 Excel 整合，不把內部人交易直接改成新聞 Level。

## 新聞與Excel

新聞表格最後兩欄為「13F 機構動向」「Form4 內部人動向」。查詢及後續排程／人工執行的Excel同樣包含這兩欄，並保留季度、申報日、分析起迄、品質、來源版本、同步時間等稽核欄位。原有人工作業、LLM判讀與L3–5顏色維持。

13F顯示季度、追蹤機構合計股數QoQ、增持／減持家數；新建倉、清倉、機構個別持股等見完整明細。這是指定機構與股票清單的季度資料，不代表全市場所有機構。

Form4依來源成果的Trade Date分析期間，使用P／S交易筆數，NPR=(Buy Count-Sell Count)/(Buy Count+Sell Count)。非P/S不加入NPR，10b5-1保留，Form4/A依原專案不直接併算而要求覆核。ERROR、PARTIAL、解析錯誤、主體無法核對或修正版待核對均顯示資料品質，不以零交易冒充成功。

「尚無已發布成果」「不在本次追蹤清單」「股票代號待核對」「非美股／不適用」「成功查詢但無P/S交易」是不同狀態。Form4截至日超過3個日曆日會標示資料較舊。13F始終保留季度，不稱為即時持倉。

這是最新已匯入成果的參考對照，不是依新聞當日可得資訊回測。新聞與不同期間的持倉／交易以市場加Ticker對照，不模糊配對公司名，也不把機構明細展開成多列新聞。

## 擷取與同步（全部在 RMSecfb/negative-news）

| 台北時間 | Workflow | 內容 |
|---|---|---|
| 週二至週六 06:00 | `sec-form4-daily.yml` | 執行 `sec_filings/run_form4.py`，分析最近 30 個日曆日，成果提交到 `sec_filings/insider_form4/YYYY-MM-DD/`（保留最近 30 天），接著同步 |
| 週一至週五 06:30 | `sec-13f.yml` | 執行 `sec_filings/run_13f.py`，成果提交到 `sec_filings/institutional_13f/YYYYQn/`，接著同步 |
| 每天 07:00 | `institutional-insider-sync.yml` | 保底同步，只讀本 repo 的成果，不連 SEC |
| 08:22 | 負面新聞 | 與上面分開，見 README |

每個擷取 workflow 都是「擷取 → 提交 → 同步」，所以即使 GitHub 排程延遲，網站資料也會是擷取後的版本。三者與負面新聞共用同一個 concurrency group，不會同時寫入。

### 必要設定

GitHub → Settings → Secrets and variables → Actions → New repository secret：

| 名稱 | 內容 |
|---|---|
| `SEC_IDENTITY` | SEC 要求的存取聲明，格式為「名稱 聯絡信箱」，例如 `公司或專案名稱 contact@公司網域` |

SEC 不需要申請帳號或 API key，只要求自動化程式在請求中聲明名稱與聯絡信箱，且請求不超過每秒 10 次（本專案設 4 次）。請使用公司可收信的信箱。**沒設定這個 Secret 時，兩個擷取 workflow 會顯示警告並略過，不算失敗。**

### 13F 的目標季度與發布門檻

原版程式的目標季度取自機構清單 `13F_Manager_List.xlsx` 的 Report Period，需要人工維護。自動執行時改由 `run_13f.py` 指定為「最近一個完整的日曆季度」，並加上兩個保護：

- 新季度要有 **80% 以上**納入彙總的機構已申報才會發布；不足時保留上一季成果。
- 已發布的季度在季末後 60 天內每天重跑（補入晚申報者），之後自動略過；重跑時成功機構數不得少於已發布的版本。

手動補跑：Actions → SEC 13F 季度擷取 → Run workflow，可指定季末日或強制重跑。機構清單 `13F_Manager_List.xlsx` 與股票清單 `13F_Ticker_Universe.xlsx` 仍需人工維護；新增機構或股票後，下次擷取才會納入。

### 同步規則

同步程式只讀 `sec_filings/` 內已提交的成果，**不執行 SEC 擷取**，不使用 LLM。可辨識路徑：

- `institutional_13f/YYYYQn/13F_YYYYQn_Full_Universe.xlsx`
- `insider_form4/YYYY-MM-DD/Form4_YYYY-MM-DD_30D_Full_Universe.xlsx`（也接受其他天數，實際依工作表起迄）

驗證工作表、股票唯一性、日期及品質後，原子保存到 `institutional_insider/`。`latest.json` 記錄兩類成果各自的來源 commit、路徑、摘要位置、列數與同步時間。前次成功資料不因新檔損壞或所有擷取失敗而被清空。

網站獨立「機構內部人交易」分頁顯示兩類摘要、機構／交易明細、Run Log，並提供來源完整檔及新聞整合 Excel。新聞查詢時直接讀共享摘要，不必先在分頁按載入。

既有歷史 Excel 不因只同步持倉而整批重寫；重新查詢下載可取得最新摘要，後續新聞排程與人工執行會將摘要寫入新生成的每日成果。已開啟的查詢結果需重新查詢才刷新資料。
