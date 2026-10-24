# negative-news 架構、判讀方式與 SEC13F 整合評估

核對日期：2026-10-24。新聞來源 RMSecfb/negative-news main；13F 來源 JUWangWang/SEC13F main（dae0d8ec9ec978bfc3e593bbff44327b607a80c8）。本文件區分既有執行行為、本次頁面功能，以及尚未建立的13F排程。

## 一、主程式與新舊檔案

現在是一套新版工作台調用舊的擷取與分析函式；不是新舊兩套排程同時跑。

|位置|用途與目前地位|
|---|---|
|負面新聞整合網站.py|Streamlit唯一入口，轉入pipeline.workbench.main|
|pipeline/workbench.py|期間查詢、人工編輯、人工執行排隊、13F分頁、執行紀錄|
|pipeline/integrated.py|目前啟用的新聞擷取、去重、FinBERT、規則、LLM、發布主流程|
|legacy_engine.py|保留原始多網站、TradingView、Google News、翻譯、FinBERT、曝險及舊13F相關函式；由新版調用|
|pipeline/market.py|全市場事件關鍵字候選搜尋及優先排序；不是列舉每一家上市公司所有新聞|
|pipeline/llm.py|Groq結構化事件判讀、JSON驗證、證據核對|
|pipeline/events.py|Parameter_Event讀取及事件、違約防呆|
|pipeline/issuers.py、symbols.py|公司與代號核對、共享交易所參照、補充歷史／OTC資料|
|pipeline/records.py、repository.py|跨日查詢、去重、人工覆核、Excel、GitHub原子提交|
|pipeline/ticker_audit.py|全歷史Ticker核對補修，不使用LLM|
|pipeline/analytics.py|本次新增：新聞頻率計算、SEC13F成果解析、跨市場限制與整合Excel|
|pipeline/run.py、legacy_bridge.py、scripts/patch_existing_app.py|先前AI版本及升級工具，現行workflow未直接啟用|
|BackupCode/|歷史程式備份，並非部署入口；不要用它判斷現行行為|
|Parameter_Event.xlsx|事件代碼、中文名稱、Level、Action、關鍵字、排除詞、優先順序與主要網站|
|SP500.xlsx、DowJones30.xlsx、config/extra_watchlist.csv|公司參照或指定擷取清單|
|config/us_symbol_reference.csv、issuer_registry.csv|共享美股代號、公司別名、歷史／OTC／非美股或非上市主體參照|

## 二、資料夾與Excel應該看哪一份

|位置|說明|
|---|---|
|news_daily/YYYY-MM-DD.csv、.xlsx|目前正式結果，依台北時間的新聞發布日歸檔；排程與人工補跑共用|
|risk_news_daily/|舊版每日Excel，留供相容讀取；不是新版每日持續更新的位置|
|ai_news_daily/|先前獨立AI的CSV／JSON，以執行日命名，內容可能跨發布日；新版有跨日遷移相容|
|ai_news_raw/|早期AI搜尋候選快照，並非現在主要報表|
|history/|原始歷史資料與人工覆核紀錄；新版讀取時保留相容性|
|news_cache/|月度LLM成功答案快取|
|news_usage/|台北日期的共用LLM呼叫預留及實用量|
|news_requests/|網站送出的持久化人工執行請求|
|news_runs/|執行結果、錯誤及最近100筆索引|
|news_audits/|歷史Ticker核對報告|

查詢時優先讀news_daily，缺少新版日期才相容讀舊成果，不將同日三個資料夾盲目相加。舊檔保留歷史不表示仍在雙軌產檔；建議先保留，待另外做歷史完整性與使用依賴確認後才考慮歸檔。

每日與期間Excel均包含：完整新聞、主要負面事件、待人工覆核、待補分析、執行紀錄。L3黃、L4橘、L5紅。CSV為可再讀取的逐筆資料，Excel是可閱讀成果。

## 三、排程與人工執行流程

唯一workflow為.github/workflows/negative-news-ai.yml，cron為UTC 00:20，即台北08:20。GitHub排程可能延遲，並非精準到秒。排程預設取得前一天08:20至當天08:20。

1. 安裝CPU FinBERT等依賴，執行回歸測試，刷新美股公司參照。
2. 同時涵蓋完整S&P500方法一／方法二及全市場風險事件搜尋；S&P500不是分日輪替子集合。
3. 方法一使用原始多網站與公司擷取；方法二使用Google News公司搜尋。全市場事件發現使用中英文風險查詢及RSS，受來源可取得範圍限制，不能保證涵蓋全市場每一篇新聞。
4. 依發布時間篩選、合併去重，再在有選方法二時執行FinBERT，保留原始標題與翻譯後評分標題。
5. 計算原始規則結果，保留Rule_*欄位，再依AI開關決定最終事件。
6. 人工決定優先、成功快取其次，尚未判讀的候選再依優先順序呼叫LLM。
7. 核對公司／Ticker與新聞文字證據，驗證事件與違約條件。
8. 依新聞發布日原子寫回news_daily的CSV與Excel，同步快取、用量及執行紀錄。提交前重讀最新資料，避免覆蓋執行期間的人工修改。

網站人工執行會把請求寫入news_requests，推送觸發同一workflow。可多選方法與公司清單，亦可選只補分析既有新聞。單次執行期間最多31天；查詢期間最多366天。相同公司相同URL或標題合併，不同公司的同篇新聞保留個別紀錄。補跑保留人工決定；不保證新聞源能回溯完整31天。

## 四、FinBERT、規則和LLM的實際分工

FinBERT使用ProsusAI/finbert，分析金融文字正負情緒。原始中文標題經翻譯後可評分，原標題另保留。它不是事件分類器，也不是直接把分數轉成L1至L5。現行整合流程中，它不構成LLM之前的負分硬門檻；方法二啟用時合併候選會評分，相關值留作稽核與規則備援。不能將現行流程解讀成「FinBERT先剔除所有非負面，LLM只看剩下資料」。

原始規則以關鍵字、排除詞及Primary_Priority選事件；無命中時依主要網站與情緒給系統類別。這套舊式字串匹配仍可能讓default字詞造成誤判，所以開AI時Rule_*僅為比較與稽核，不是直接最終結論。

AI開啟時，LLM讀取標題、摘要、已知公司及允許的事件代碼與中文意義，回傳是否相關、辨識公司、事件代碼、事件狀態、是否確認違約、原文證據及中文摘要。程式再用Parameter_Event把事件代碼對應到Level及Action。現在若回傳多個事件代碼，主事件採模型回傳清單中第一個有效代碼，並未重新依最高Level或優先順序排序；完整代碼仍保留於LLM事件代碼欄。

模型目前是Groq託管的openai/gpt-oss-20b，不是呼叫OpenAI API。LLM不直接產生可以信任的Ticker；程式把公司文字與共享公司參照、別名、交易所資料及新聞證據核對。歧義、非本市場、非上市或證據不足保留空白與原因，不強補代號。人工Ticker受保護，未能核實時另留稽核警示。

CDS利差擴大不等於已發生債務違約。最終DEFAULT必須通過confirmed_default及CONFIRMED狀態等防呆；缺乏證據、L4/L5重大事件、確認違約或無法核對主體等情況會要求人工覆核。

AI關閉時使用原始規則作最終分類；AI開啟但額度不足／429／API失敗時，保留為待補分析，最終Level空白，不能當作L0、沒有風險或已完成審查。主要負面事件頁為L3–5，完整新聞仍可查看其他結果。

## 五、調高LLM判讀量的限制

若「規則量」指判讀筆數：目前網站每次最多100、預設35；Actions預設同為35；共享每日上限LLM_DAILY_MAX=120；兩次新請求間隔GROQ_MIN_INTERVAL_SECONDS=20秒。失敗HTTP嘗試也算用量。人工結果及成功快取沿用不增加新呼叫。每一公司／新聞紀錄分開判讀，多公司新聞可能多次呼叫。更新模型、Prompt、規則或新聞文字可讓舊快取失效、增加用量。

調整每日上限需改workflow的LLM_DAILY_MAX；調整排程預設需改MAX_LLM備援值及相關預設。超過單次100還需同步修改workbench輸入限制與integrated.validate_request的上限，不能只把Actions輸入改成500。35次的20秒間隔約11分鐘、100次約33分鐘，另加擷取、模型、API與輸出耗時；目前job timeout為90分鐘。

Groq另有組織層RPM、RPD、TPM、TPD等配額，與本程式上限分開。官方頁目前列gpt-oss-20b免費基本值為30RPM、1000RPD、8000TPM、200000TPD，但實際應以使用者Groq Console的組織／模型Limits為準，其他共用應用亦會消耗配額。20秒間隔只能降低每分鐘請求壓力，不能保證不撞Token限制。若每次平均輸入加輸出為2000 Tokens，單看200000TPD約僅容100次；此為估算，不是實測單筆用量或保證可用筆數。429後本次停止新增呼叫，其餘候選留待補分析。

若指增加事件規則：Parameter_Event目前52列，其中3個系統類別，49個事件可提供LLM。llm.py只傳入前180個事件代碼；未來超過180必須先調整此設計，否則後面事件不會送給模型。事件越多Prompt越長，Token用量、分類難度與成本會增加。每筆目前標題最多600字元、摘要1600字元、輸出最多750 Tokens，不是讀取全文。增加新聞量不能彌補來源漏抓與摘要缺證據。

## 六、本次13F頁面整合

原始SEC13F專案目前只有程式、設定清單及2026Q2成果，沒有requirements或.github/workflows。13F_FINAL.py有固定Windows路徑，目標季度取自機構清單Report Period最常見值，不是自動判斷最新應申報季度；它抓各機構最新13F-HR再比對季度，因此新季度開始後，若清單未更新，機構可能被標為STALE。也尚需評估修正版13F-HR/A與前季可比性，不能直接貼進Actions就稱為完成自動更新。

現有成果：30檔股票、819筆機構明細；37家機構中32家OK／Aggregate Y、3家STALE、2家REVIEW排除。代表這份追蹤清單，不代表全市場所有機構。中文欄位和部分訊號已損壞；本次接入會先驗證固定格式的英文錨點，再修復已知20欄名稱並由數值重建方向訊號，原始成果下載維持原檔。

網站新增獨立13F分頁，可讀取季度清單、選擇成果並載入，顯示資料來源commit SHA、季度、持倉基準日、執行品質與機構明細，下載來源Excel。只按美股Ticker精確左連接已查詢新聞，不把港股代號、公司模糊名稱或空白主體串進美股13F；不改新聞Level，也不增加新聞筆數。未涵蓋股票表示無此專案參考資料，不等於零持股。

整合Excel保留原有5張新聞表，再加新聞與13F對照、13F季度摘要、13F機構明細、13F執行紀錄、資料來源與期間。新聞及季度資料分表保存，對照列帶有13F季度。這是所選季度參考對照，尚非歷史當時資訊回測；不能把後來公布的季度結果當成新聞發生時已可取得的訊號。

## 七、Excel與排程建議（評估，尚未建立SEC排程）

建議保留兩套資料產出：新聞每天維持news_daily；13F按季度放institutional_13f，另加每次抓取時間／來源版本快照與latest索引。同季度晚交、補交或修正申報不能只依季度名稱認定資料相同。網站在查詢時提供整合Excel，避免每天複製全部819筆持倉到每份新聞Excel，亦避免13F失敗拖累08:20新聞任務。

建議獨立SEC workflow，每日較早時段（例如台北06:30）檢查新申報，僅對有變更的機構重算；申報密集期持續每天檢查，平時也可維持低成本增量檢查與人工觸發。SEC官方說明13F一般應於季末後45天內申報，因此季度資料不會在季末當天齊備。設計須保留每家實際季度、申報日、accession、是否納入彙總、更新時間與失敗原因。

自動化前需把路徑改成專案相對位置、SEC Identity由環境設定、明確指定目標季度並驗證前季是相鄰季度、固定依賴版本、妥善處理修正版申報、更新及凍結原始清單版本、增加增量快取與失敗不覆蓋上次成功成果的保護。這次沒有新增SEC登入／憑證／排程，也沒有更改SEC13F來源repo。來源新增季度成果後，可在網站更新清單直接讀取。

Form4_FINAL.py雖已存在，但目前repo沒有Form4成果檔，不應顯示虛構內部人訊號。後續可獨立每日抓取最近交易窗口、保存交易日與申報日，另外做Form4/A修正版對帳；再以同一Ticker對照介面接入。

## 八、新聞頻率與時間序列

新版頁面加入原有規則式放量警示：比較查詢期間最後有資料發布日和前面有資料發布日平均。基準為零且當日至少3則，或基準大於零且當日至少2則、達平均兩倍，才提示。採已分類L1–5、按news_id去重、按市場與Ticker分公司。未核對Ticker不做個股放量警示；待補分析不當作已確認負面。

沒有資料的日子不假設零新聞，以免歷史漏跑造成假警示。可選全部市場或特定公司看折線圖，並下載統計CSV。這是使用者目前查詢期間內的已保存新聞觀察，不是隱含載入所有歷史或全市場完整資料；想用較長基準應擴大查詢期間。單日或資料不足會明示。不同來源範圍、缺漏、未完成分類與同一事件多媒體報導都可能影響則數，不能解讀成獨立事件數或統計顯著性。

## 參考

- 現行新聞主流程：https://github.com/RMSecfb/negative-news/blob/main/pipeline/integrated.py
- 現行workflow：https://github.com/RMSecfb/negative-news/blob/main/.github/workflows/negative-news-ai.yml
- SEC13F來源：https://github.com/JUWangWang/SEC13F
- Groq限制：https://console.groq.com/docs/rate-limits
- SEC 13F FAQ：https://www.sec.gov/rules-regulations/staff-guidance/frequently-asked-questions-about-form-13f
