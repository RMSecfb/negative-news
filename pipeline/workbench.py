"""Streamlit UI: period queries, direct review and queued unified crawls."""
from __future__ import annotations
import hashlib
import io
import os
import uuid
from datetime import datetime, timedelta, time as dt_time
from pathlib import Path
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, DataReturnMode
from .common import truth
from .issuers import company_frame, load_companies
from .integrated import enqueue, validate_request
from .records import TPE, read_days, major, review, pending, excel_bytes, manual_save
from .repository import GitHubRepository, LocalRepository, read_json

ROOT = Path(__file__).resolve().parents[1]


def setting(name, default=''):
    try:
        return str(st.secrets.get(name, os.getenv(name, default)))
    except Exception:
        return os.getenv(name, default)


def repository():
    local = os.getenv('NEWS_LOCAL_REPO')
    if local:
        return LocalRepository(local)
    return GitHubRepository(setting('GITHUB_REPO', 'RMSecfb/negative-news'), setting('GITHUB_TOKEN'), setting('GITHUB_BRANCH', 'main'))


def render_results():
    frame = st.session_state.get('news_query_frame', pd.DataFrame())
    if frame.empty:
        st.info('這個期間尚無新聞結果；可到「人工執行新聞期間」補跑。')
        return
    frame = frame.fillna('').copy()
    for field in ['Ticker', 'Company', 'Level', '事件中文', 'Action', '覆核註記', '人工版本']:
        if field not in frame:
            frame[field] = ''
    metrics = st.columns(4)
    metrics[0].metric('合併新聞', len(frame))
    metrics[1].metric('主要負面事件 · L3–5', len(major(frame)))
    metrics[2].metric('待人工覆核', len(review(frame)))
    metrics[3].metric('待 AI 補分析', len(pending(frame)))
    done = frame.get('LLM狀態', pd.Series('', index=frame.index)).eq('OK') | frame.get('已人工覆核', pd.Series('', index=frame.index)).map(truth)
    st.caption(f'已完成 AI／人工判讀 {int(done.sum())} / {len(frame)} 筆。待補分析不等於無風險；待人工覆核表示已有判讀但仍須確認。')
    left, right = st.columns([2, 1])
    mode = left.radio('結果分類', ['主要負面事件', '完整新聞', '待人工覆核', '待補分析'], horizontal=True)
    right.download_button('下載查詢期間合併 Excel', excel_bytes(frame, st.session_state.get('news_query_runs', [])), file_name='Negative_News_' + st.session_state.get('news_query_period', 'period') + '.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    current = {'主要負面事件': major(frame), '完整新聞': frame, '待人工覆核': review(frame), '待補分析': pending(frame)}[mode].copy()
    st.markdown('<span class="level-chip level3">Level 3 · 黃色</span> <span class="level-chip level4">Level 4 · 橘色</span> <span class="level-chip level5">Level 5 · 紅色</span>', unsafe_allow_html=True)
    edit = st.toggle('開啟人工編輯', key='unified_edit_enabled')
    if edit:
        st.caption('可修改 Ticker、Level、事件中文、Action、覆核註記。調整篩選或重新查詢前，請先儲存。')
    if current.empty:
        st.info('目前分類沒有新聞。')
        return
    preferred = ['日期', 'Market', 'Ticker', 'Ticker狀態', 'Company', 'Published Time', 'Title', 'Level', '事件中文', 'Action', '分類狀態', '已人工覆核', '覆核註記', '最終判斷來源', '來源方法', 'FinBERT', 'Rule_Event_Code', 'Rule_Level', 'LLM事件代碼', 'LLM辨識公司', 'Ticker來源', 'Ticker候選', 'Ticker核對依據', 'Ticker參照日期', '上市狀態', '原始Ticker', '中文摘要', 'URL', 'Source']
    current = current[[c for c in preferred if c in current] + [c for c in current if c not in preferred]]
    if 'Ticker稽核警示' in current:
        preferred.append('Ticker稽核警示')
    current['Level'] = current['Level'].astype(str).str.replace(r'\.0$', '', regex=True)
    builder = GridOptionsBuilder.from_dataframe(current)
    builder.configure_default_column(editable=False, sortable=True, filter=True, resizable=True, minWidth=105)
    for col in current.columns:
        if col not in preferred:
            builder.configure_column(col, hide=True)
    editable = ['Ticker', 'Level', '事件中文', 'Action', '覆核註記']
    for col in editable:
        builder.configure_column(col, editable=edit)
    builder.configure_column('Level', width=88, cellEditor='agSelectCellEditor', cellEditorParams={'values': ['1', '2', '3', '4', '5']})
    builder.configure_column('Title', width=390, wrapText=True, autoHeight=True)
    builder.configure_column('Ticker狀態', width=210, wrapText=True, autoHeight=True)
    builder.configure_column('Action', width=260)
    builder.configure_column('事件中文', width=220)
    builder.configure_grid_options(getRowStyle=JsCode("""function(p) {
        const n=Number(p.data.Level||0); const colors={3:'#FFF9C4',4:'#FFE0B2',5:'#FFCDD2'};
        return colors[n] ? {backgroundColor:colors[n],color:'#17283D'} : null;
    }"""), rowHeight=46, stopEditingWhenCellsLoseFocus=True)
    key = hashlib.sha256((st.session_state.get('news_query_period', '') + mode + str(edit) + '|'.join(current['news_id'].astype(str))).encode()).hexdigest()[:14]
    response = AgGrid(current, gridOptions=builder.build(), allow_unsafe_jscode=True, data_return_mode=DataReturnMode.AS_INPUT, update_on=['cellValueChanged'], height=510, theme='streamlit', key='unified_grid_' + key)
    if edit:
        reviewer = st.text_input('覆核人（選填）', key='unified_reviewer')
        if st.button('儲存人工覆核', type='primary', key='save_unified_review'):
            try:
                edited = response['data']
                if not isinstance(edited, pd.DataFrame):
                    edited = pd.DataFrame(edited)
                count = manual_save(repository(), current, edited, reviewer)
                period = st.session_state['news_query_dates']
                st.session_state['news_query_frame'] = read_days(repository(), period[0], period[1], st.session_state['news_query_market'])
                st.session_state['news_saved_note'] = f'已儲存 {count} 筆人工覆核；合併 Excel 已同步更新，後續補跑保留人工決定。'
                st.rerun()
            except Exception as exc:
                st.error(f'未儲存：{exc}')
    with st.expander('保留原有曝險與 13F 分析功能'):
        import legacy_engine as engine
        upload = st.file_uploader('曝險清單 Excel（選填）', type=['xlsx'], key='unified_exposure')
        exposure = engine.load_exposure_list_from_upload(upload) if upload else pd.DataFrame(columns=['Ticker', 'Exposure'])
        if upload:
            st.dataframe(exposure, hide_index=True)
        engine.render_sec13f_panel(major(frame), exposure)
    with st.expander('個股價格查詢（原功能）'):
        ticker = st.text_input('股票代號', value='META', key='unified_price_ticker')
        if st.button('查詢股價', key='unified_price_load'):
            import legacy_engine as engine
            begin, end = st.session_state['news_query_dates']
            prices = engine.stock_prices(ticker, begin.isoformat(), (end + timedelta(days=1)).isoformat())
            st.dataframe(prices, hide_index=True)


def query_panel():
    st.subheader('查詢新聞期間')
    today = datetime.now(TPE).date()
    a, b, c = st.columns([1, 1, 1])
    start = a.date_input('查詢開始日期', value=today - timedelta(days=2), key='query_start')
    end = b.date_input('查詢結束日期', value=today, key='query_end')
    market_name = c.selectbox('查詢市場', ['全部市場', '美股 US', '港股 HK'], key='query_market')
    market = {'全部市場': 'ALL', '美股 US': 'US', '港股 HK': 'HK'}[market_name]
    if st.button('查詢合併結果', type='primary', key='query_news'):
        if start > end or (end-start).days > 365:
            st.error('請選擇有效期間；單次查詢最多 366 天。')
        else:
            try:
                with st.spinner('讀取排程與人工補跑合併結果…'):
                    frame = read_days(repository(), start, end, market)
                    st.session_state['news_query_runs'] = read_json(repository().read('news_runs/index.json'), {'runs': []})['runs']
                st.session_state.update(news_query_frame=frame, news_query_dates=(start, end), news_query_market=market, news_query_period=f'{start}_{end}_{market}')
            except Exception as exc:
                st.error(f'查詢失敗：{exc}')
    if st.session_state.get('news_saved_note'):
        st.success(st.session_state.pop('news_saved_note'))
    if 'news_query_frame' in st.session_state:
        render_results()


def manual_panel():
    st.subheader('人工執行新聞期間')
    st.caption('排程與人工補跑共用合併結果。相同新聞去重，原有 AI 判讀與人工覆核優先沿用。')
    now = datetime.now(TPE).replace(second=0, microsecond=0)
    begin = now - timedelta(days=1)
    a, b = st.columns(2)
    sd = a.date_input('執行開始日期', value=begin.date(), key='run_start_date')
    stime = a.time_input('執行開始時間', value=begin.time(), key='run_start_time')
    ed = b.date_input('執行結束日期', value=now.date(), key='run_end_date')
    etime = b.time_input('執行結束時間', value=now.time(), key='run_end_time')
    st.caption('輸入時區：台北 UTC+8；結果依新聞原始發布日期歸檔。歷史擷取仍受各新聞源保留範圍限制。')
    chosen_methods = st.multiselect('方法（可多選；AI LLM 是統一判讀開關）', ['AI LLM', '方法一', '方法二'], default=['AI LLM', '方法一', '方法二'], key='run_methods')
    st.caption('方法一：原有多網站擷取。方法二：Google News 與 FinBERT。AI LLM：合併去重後統一判讀；只選 AI 時會先執行候選事件／公司搜尋。')
    market_name = st.selectbox('執行市場', ['美股 US', '港股 HK'], key='run_market')
    market = 'US' if market_name.startswith('美股') else 'HK'
    scope_name = st.radio('標的範圍', ['全市場', '指定公司列表'], horizontal=True, key='run_scope')
    names, uploaded_rows = [], []
    if scope_name == '指定公司列表':
        options = ['Dow Jones 30', 'S&P 500', '上傳列表'] if market == 'US' else ['上傳列表']
        defaults = ['S&P 500'] if market == 'US' else ['上傳列表']
        lists = st.multiselect('指定公司列表（可多選）', options, default=defaults, key='run_lists_' + market)
        names = [dict([('Dow Jones 30', 'DowJones30'), ('S&P 500', 'SP500')])[x] for x in lists if x != '上傳列表']
        if '上傳列表' in lists:
            uploads = st.file_uploader('上傳公司列表 Excel／CSV（可多份）', type=['xlsx', 'csv'], accept_multiple_files=True, key='run_uploads_' + market)
            st.caption('至少需 Ticker、Company；可另附 Market、Exchange、Aliases。港股代號保留前導零。')
            for upload in uploads:
                try:
                    data = pd.read_csv(upload, dtype=str) if upload.name.lower().endswith('.csv') else pd.read_excel(upload, dtype=str)
                    uploaded_rows.extend(company_frame(data, market).to_dict('records'))
                except Exception as exc:
                    st.error(f'{upload.name}：{exc}')
        try:
            companies = load_companies(ROOT, names, uploaded_rows, market)
            st.caption(f'合併清單並依市場／代號去重後：{len(companies)} 家公司。')
        except Exception as exc:
            st.error(str(exc))
    llm = 'AI LLM' in chosen_methods
    cap = st.number_input('本次最多新增 LLM 判讀筆數', min_value=0, max_value=100, value=35, step=5, disabled=not llm)
    resume = st.checkbox('只補分析既有新聞（不重新擷取）', key='resume_news', disabled=not llm)
    st.info('LLM 以去重後的候選計算，成功判讀可跨排程／補跑沿用。額度不足或限流的新聞保留為「待補分析」；人工覆核不會被自動結果覆寫。')
    if st.button('開始執行並合併', type='primary', key='queue_news'):
        try:
            if scope_name == '指定公司列表' and not names and not uploaded_rows:
                raise ValueError('請選擇有效公司列表或上傳公司清單')
            request = {'id': uuid.uuid4().hex, 'type': 'manual', 'start': datetime.combine(sd, stime, TPE).isoformat(), 'end': datetime.combine(ed, etime, TPE).isoformat(),
                       'market': market, 'scope': 'all' if scope_name == '全市場' else 'list', 'lists': names, 'companies': uploaded_rows,
                       'methods': [m for m in chosen_methods if m != 'AI LLM'], 'ai': llm, 'max_llm': int(cap), 'resume': bool(resume and llm)}
            run_id = enqueue(repository(), request)
            st.session_state['news_active_run'] = run_id
            st.success('執行請求已儲存並送入 GitHub Actions。可在下方或「執行紀錄」查看進度。')
        except Exception as exc:
            st.error(f'未送出執行：{exc}')
    if st.session_state.get('news_active_run'):
        render_active_run()


@st.fragment(run_every=15)
def render_active_run():
    run_id = st.session_state.get('news_active_run')
    if not run_id:
        return
    try:
        report = read_json(repository().read(f'news_runs/{run_id}.json'))
        if not report:
            st.info('已排隊，等待 GitHub Actions 啟動。若長時間未啟動，請在 Actions 確認此 workflow 已啟用。')
        elif report.get('status') == 'running':
            st.info(f'執行中：{report.get("start", "")} → {report.get("end", "")}。新聞量大時可能需數十分鐘。')
        elif report.get('status') == 'failed':
            st.error('執行失敗：' + str(report.get('error', '未知錯誤')))
        else:
            st.success(f'完成：新增 {report.get("新增", 0)}、重複合併 {report.get("重複合併", 0)}、LLM 新判讀 {report.get("llm_calls", 0)}、沿用 {report.get("cache_reused", 0)}、待補分析 {report.get("pending", 0)}。請回到查詢期間讀取合併結果。')
            if report.get('errors'):
                st.warning(f'部分新聞來源回報 {len(report["errors"])} 項異常，詳見執行紀錄；未宣稱完整來源覆蓋。')
    except Exception as exc:
        st.warning(f'目前無法讀取執行狀態：{exc}')


def history_panel():
    st.subheader('執行紀錄')
    if st.button('讀取／重新整理執行紀錄', key='load_news_runs'):
        try:
            repo = repository()
            reports = read_json(repo.read('news_runs/index.json'), {'runs': []}).get('runs', [])
            st.session_state['news_run_reports'] = sorted(reports, key=lambda r: r.get('started_at', r.get('finished_at', '')), reverse=True)
        except Exception as exc:
            st.error(str(exc))
    reports = st.session_state.get('news_run_reports', [])
    if reports:
        st.dataframe(pd.DataFrame([{k: r.get(k, '') for k in ['id', 'type', 'status', 'start', 'end', 'methods', 'llm_calls', 'cache_reused', 'pending', 'review', '新增', '重複合併']} for r in reports]), hide_index=True, use_container_width=True)
        selected = st.selectbox('查看執行詳細紀錄', range(len(reports)), format_func=lambda i: str(reports[i].get('id')))
        st.json(reports[selected])


def main():
    st.set_page_config(page_title='負面新聞整合中心', page_icon='🛡️', layout='wide')
    st.markdown('''<style>
        .block-container{max-width:1450px;padding-top:1.7rem}.level-chip{display:inline-block;padding:4px 12px;border-radius:5px;margin:0 5px 12px 0;color:#17283D;font-size:13px}
        .level3{background:#FFF9C4}.level4{background:#FFE0B2}.level5{background:#FFCDD2}
        div[data-testid="stMetric"]{border:1px solid #dce2ea;border-radius:9px;padding:14px}
        </style>''', unsafe_allow_html=True)
    st.title('負面新聞整合中心')
    st.caption('跨市場新聞監控 · 原有 FinBERT／規則保留 · 合併後由 LLM 統一判讀 · 人工覆核優先')
    st.info('每日 08:20（台北時間）：AI 全市場事件搜尋 ＋ 完整 S&P 500 方法一／方法二 → 去重 → FinBERT／規則 → LLM → 合併 Excel。')
    query, manual, history = st.tabs(['查詢新聞期間', '人工執行新聞期間', '執行紀錄'])
    with query:
        query_panel()
    with manual:
        manual_panel()
    with history:
        history_panel()
