"""Safely patch the existing root app with two merge wrappers + AI read-only view.
Run from repository root: python scripts/patch_existing_app.py
"""
from __future__ import annotations
import ast
from pathlib import Path

APP=Path('負面新聞整合網站.py')
MARKER='# === NEGATIVE_NEWS_AI_ADDON_V1 ==='
WRAPPERS={
'save_risk_news_snapshot_to_github': '''def save_risk_news_snapshot_to_github(frame: pd.DataFrame, snapshot_date: date) -> bool:
    """Merge with existing same-date data; preserve earlier universe and manual reviews."""
    from pipeline.legacy_bridge import merge_daily_snapshot
    return merge_daily_snapshot(frame, snapshot_date,
        _github_config, _github_risk_news_path, _github_headers,
        _github_get_file_sha, _dataframe_to_excel_bytes,
        _github_push_risk_news_snapshot)
''',
'append_negative_history_to_github': '''def append_negative_history_to_github(event_frame: pd.DataFrame, event_date: datetime) -> bool:
    """Recompute same-date summary from old snapshot + current run, never discard other tickers."""
    from pipeline.legacy_bridge import merge_daily_history
    return merge_daily_history(event_frame, event_date,
        _github_config, _github_risk_news_path, _github_headers,
        _github_fetch_history, _github_push_history, load_negative_history.clear)
'''
}

VIEW='''
# === NEGATIVE_NEWS_AI_ADDON_V1 ===
# This independent table does not overwrite the original Streamlit dashboard.
try:
    import io as _ai_io
    st.divider()
    st.subheader("AI 事件辨識（GitHub Actions）")
    st.caption("此區顯示新增的批次 LLM 結果；原本的新聞搜尋、歷史統計與人工覆核功能保留。")
    _ai_day = st.date_input("AI 批次日期", value=datetime.now(TAIPEI).date(), key="ai_addon_day")
    try:
        _ai_repo = st.secrets.get("GITHUB_REPO", "RMSecfb/negative-news")
    except Exception:
        _ai_repo = "RMSecfb/negative-news"
    _ai_url = f"https://raw.githubusercontent.com/{_ai_repo}/main/ai_news_daily/{_ai_day.isoformat()}.csv"
    if st.button("讀取 AI 分析結果", key="load_ai_news"):
        _ai_response = requests.get(_ai_url, timeout=18)
        if _ai_response.status_code == 200:
            _ai_data = pd.read_csv(_ai_io.BytesIO(_ai_response.content), dtype=str).fillna('')
            st.session_state['ai_addon_frame'] = _ai_data
            st.session_state['ai_addon_loaded_date'] = _ai_day.isoformat()
        elif _ai_response.status_code == 404:
            st.info("這天尚未有 GitHub Actions 的 AI CSV，請選擇已執行日期。")
        else:
            st.warning(f"讀取 AI CSV 失敗：HTTP {_ai_response.status_code}")
    _ai_data = st.session_state.get('ai_addon_frame',pd.DataFrame())
    if not _ai_data.empty:
        st.caption(f"目前載入日期：{st.session_state.get('ai_addon_loaded_date','')}｜新聞 {_ai_data.shape[0]} 筆")
        _ai_filter = st.selectbox("AI 分類狀態", ["全部", "已分類", "待人工覆核", "無關新聞", "未入選初篩"], key="ai_addon_filter")
        _ai_show = _ai_data if _ai_filter == '全部' else _ai_data[_ai_data.get('分類狀態',pd.Series('',index=_ai_data.index)) == _ai_filter]
        _ai_cols = [x for x in ("Ticker","Company","Title","Event_Code","事件中文","Level","事件狀態","分類狀態","需要人工覆核","中文摘要","Evidence","URL","監控池") if x in _ai_show.columns]
        st.dataframe(_ai_show[_ai_cols],use_container_width=True,hide_index=True)
        st.download_button("下載 AI CSV",_ai_show.to_csv(index=False).encode('utf-8-sig'),file_name=f"AI_News_{st.session_state.get('ai_addon_loaded_date','')}.csv",mime='text/csv')
except Exception as _ai_error:
    st.warning(f"AI 模組讀取失敗（不影響原網站）：{_ai_error}")
'''

def main():
    if not APP.exists():raise SystemExit('請在負面新聞 repo 根目錄執行此指令')
    content=APP.read_text(encoding='utf-8-sig')
    if MARKER in content:
        print('AI 區塊已安裝；不重複修補。')
        return
    nodes=ast.parse(content).body
    lines=content.splitlines(keepends=True)
    changes=[]
    for node in nodes:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in WRAPPERS:
            changes.append((node.lineno-1,node.end_lineno,WRAPPERS[node.name]))
    if len(changes)!=len(WRAPPERS):
        raise SystemExit('找不到兩個原版函式；為安全起見停止自動修改。')
    for start,end,replace in sorted(changes,reverse=True):
        lines[start:end]=[replace+'\n']
    modified=''.join(lines)+'\n'+VIEW
    ast.parse(modified)
    APP.write_text(modified,encoding='utf-8')
    print('修補成功：同日合併、歷史統計合併、AI CSV 讀取區塊。')
    print('建議檢查 git diff，確認之後才 commit。')

if __name__=='__main__':main()
