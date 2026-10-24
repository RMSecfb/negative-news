# 機構內部人交易整合

來源為 [JUWangWang/SEC13F](https://github.com/JUWangWang/SEC13F)。來源的13F_FINAL.py及Form4_FINAL.py均保留各自判讀架構，新聞系統透過其成果Excel整合，不把內部人交易直接改成新聞Level。

## 新聞與Excel

新聞表格最後兩欄為「13F 機構動向」「Form4 內部人動向」。查詢及後續排程／人工執行的Excel同樣包含這兩欄，並保留季度、申報日、分析起迄、品質、來源版本、同步時間等稽核欄位。原有人工作業、LLM判讀與L3–5顏色維持。

13F顯示季度、追蹤機構合計股數QoQ、增持／減持家數；新建倉、清倉、機構個別持股等見完整明細。這是指定機構與股票清單的季度資料，不代表全市場所有機構。

Form4依來源成果的Trade Date分析期間，使用P／S交易筆數，NPR=(Buy Count-Sell Count)/(Buy Count+Sell Count)。非P/S不加入NPR，10b5-1保留，Form4/A依原專案不直接併算而要求覆核。ERROR、PARTIAL、解析錯誤、主體無法核對或修正版待核對均顯示資料品質，不以零交易冒充成功。

「尚無已發布成果」「不在本次追蹤清單」「股票代號待核對」「非美股／不適用」「成功查詢但無P/S交易」是不同狀態。Form4截至日超過3個日曆日會標示資料較舊。13F始終保留季度，不稱為即時持倉。

這是最新已匯入成果的參考對照，不是依新聞當日可得資訊回測。新聞與不同期間的持倉／交易以市場加Ticker對照，不模糊配對公司名，也不把機構明細展開成多列新聞。

## 成果同步

獨立workflow：.github/workflows/institutional-insider-sync.yml。每日台北07:00（UTC前日23:00）或手動Run workflow，與08:20新聞任務分開。GitHub排程可能延遲。

它只讀取來源已提交main的成果，**不執行SEC擷取程式**，不使用LLM、不需要新增SEC憑證。來源目前有兩支程式，但2026-10-24核對時只有2026Q2的13F成果，尚未提交Form4 Excel。

可自動辨識來源原有路徑：

- institutional_13f/YYYYQn/13F_YYYYQn_Full_Universe.xlsx
- insider_form4/YYYY-MM-DD/Form4_YYYY-MM-DD_30D_Full_Universe.xlsx（也接受其他天數，實際依工作表起迄）

每次以同一來源commit SHA讀目錄及檔案。驗證工作表、股票唯一性、日期及品質後，原子保存到新聞repo的institutional_insider/。latest.json記錄兩類成果各自來源SHA、路徑、摘要位置、列數與同步時間。前次成功資料不因新檔損壞或所有擷取失敗而被清空。

網站獨立「機構內部人交易」分頁顯示兩類摘要、機構／交易明細、Run Log，並提供來源完整檔及新聞整合Excel。新聞查詢時直接讀共享摘要，不必先在分頁按載入。

既有歷史Excel不因只同步持倉而整批重寫；重新查詢下載可取得最新摘要，後續新聞排程與人工執行會將摘要寫入新生成的每日成果。已開啟的查詢結果需重新查詢才刷新資料。

## SEC原始擷取仍為獨立階段

要讓Form4實際有數據，需執行來源的Form4_FINAL.py並提交成果Excel到上述路徑。現有腳本仍是Windows本機路徑與既有SEC Identity設定。若進一步自動化SEC擷取，需要另行移植路徑與日期、固定edgartools版本、處理SEC限流、增量與修正版對帳；不能把成果同步成功解讀成SEC擷取也成功。
