"""Backwards-compatible GitHub merge for the current Streamlit application.

Designed to be called by two small AST-based function replacements. Preserves
existing API token / repository settings, manual edits, same-day universe runs.
"""
from __future__ import annotations
import base64
import io
from datetime import date, datetime
import pandas as pd
import requests
from .common import merge_records, news_key, truth


def _old_snapshot(config, headers_fn, github_path):
    if not config:return pd.DataFrame()
    url=f"https://api.github.com/repos/{config['repo']}/contents/{github_path}"
    try:
        r=requests.get(url,headers=headers_fn(config['token']),params={'ref':config['branch']},timeout=20)
        if r.status_code==404:return pd.DataFrame()
        r.raise_for_status()
        encoded=r.json().get('content','').replace('\n','')
        if not encoded:return pd.DataFrame()
        return pd.read_excel(io.BytesIO(base64.b64decode(encoded)),dtype=str).fillna('')
    except Exception as exc:
        # Critical: NEVER silently overwrite unreadable old files.
        raise RuntimeError(f'讀取 GitHub 既有新聞快照失敗，已禁止覆寫：{exc}') from exc


def merge_daily_snapshot(frame,date_obj,config_fn,path_fn,headers_fn,sha_fn,to_bytes_fn,push_fn):
    cfg=config_fn()
    if not cfg or frame is None or frame.empty:return False
    path=path_fn(date_obj)
    try:
        existing=_old_snapshot(cfg,headers_fn,path)
        combined=merge_records(existing,frame.copy())
        # preserve displayed legacy column ordering and discard internal key.
        result=combined.drop(columns=['news_id'],errors='ignore')
        original=list(frame.columns)
        cols=list(dict.fromkeys(original+list(result.columns)))
        result=result.reindex(columns=cols,fill_value='')
        sha=sha_fn(path)
        return push_fn(to_bytes_fn(result),date_obj,sha)
    except Exception as exc:
        print('GitHub merge snapshot ERROR:',exc)
        return False


def merge_daily_history(frame,event_date,config_fn,path_fn,headers_fn,fetch_history_fn,push_history_fn,cache_clear):
    cfg=config_fn()
    if not cfg or frame is None or frame.empty:return False
    try:
        # Existing snapshot may contain other universes; rebuild totals from union.
        existing=_old_snapshot(cfg,headers_fn,path_fn(event_date.date()))
        combined=merge_records(existing,frame)
        for key in ['Ticker','Company','Level']:
            if key not in combined: combined[key]=''
        combined['Ticker']=combined['Ticker'].fillna('').astype(str).str.strip()
        combined['Company']=combined['Company'].fillna('').astype(str).str.strip()
        combined['Level']=pd.to_numeric(combined['Level'],errors='coerce').fillna(0).astype(int)
        combined=combined[combined['Ticker']!='']
        if combined.empty:return False
        grouped=combined.groupby(['Ticker','Company']).agg(則數=('Ticker','size'),最高Level=('Level','max')).reset_index()
        grouped['日期']=pd.to_datetime(event_date.date())
        hist,sha=fetch_history_fn()
        if not hist.empty:
            hist=hist[pd.to_datetime(hist['日期'],errors='coerce').dt.date!=event_date.date()]
        merged=pd.concat([hist,grouped],ignore_index=True)
        ok=push_history_fn(merged,sha)
        if ok: cache_clear()
        return ok
    except Exception as exc:
        print('GitHub merge history ERROR:',exc)
        return False
