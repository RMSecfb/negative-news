"""Normalized news schema and idempotent CSV persistence."""
from __future__ import annotations
import hashlib
import re
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
import pandas as pd

DROP_QUERY = {'utm_source','utm_medium','utm_campaign','utm_term','utm_content','gclid','fbclid','ref','source','from'}

def clean_url(value: str) -> str:
    value = str(value or '').strip()
    if not value: return ''
    try:
        p = urlsplit(value)
        if not p.netloc: return value.lower().rstrip('/')
        q = [(k,v) for k,v in parse_qsl(p.query) if k.lower() not in DROP_QUERY and not k.lower().startswith('utm_')]
        return urlunsplit((p.scheme.lower(),p.netloc.lower(),p.path.rstrip('/'),urlencode(sorted(q)),''))
    except Exception:
        return value.lower().rstrip('/')

def text_key(text: str) -> str:
    return re.sub(r'\s+', ' ', str(text or '').casefold()).strip()

def news_key(item: dict) -> str:
    """Ticker scoped: same story about two issuers remains two rows."""
    url = clean_url(item.get('URL',''))
    material = (url or text_key(item.get('Title',''))) + '|' + str(item.get('Ticker','')).upper().strip()
    return hashlib.sha256(material.encode()).hexdigest()[:24]

def truth(v) -> bool:
    return str(v).strip().lower() in ('true','1','yes','y')

def merge_records(old: pd.DataFrame, new: pd.DataFrame, *, key='news_id', manual_col='已人工覆核') -> pd.DataFrame:
    """Keep manual decisions, new discovery tags, and latest nonempty fields on reruns."""
    if old is None or old.empty: old = pd.DataFrame()
    if new is None or new.empty: new = pd.DataFrame()
    cols = list(dict.fromkeys([*old.columns, *new.columns]))
    if not cols: return pd.DataFrame()
    rows = {}
    for f in (old, new):
        for r in f.reindex(columns=cols, fill_value='').fillna('').to_dict('records'):
            ident = str(r.get(key,'')).strip() or news_key(r)
            r[key] = ident
            if ident not in rows:
                rows[ident] = r
                continue
            prev=rows[ident]
            tags=set(filter(None, re.split(r'[|;]',str(prev.get('監控池',''))))) | set(filter(None, re.split(r'[|;]',str(r.get('監控池','')))))
            changed={**prev, **{k:v for k,v in r.items() if str(v).strip() not in ('','nan')}}
            changed['監控池']='|'.join(sorted(tags))
            if truth(prev.get(manual_col,False)) and not truth(r.get(manual_col,False)):
                for k in ('Event_Code','事件中文','Level','Action','人工覆核人','人工覆核時間',manual_col):
                    if k in prev: changed[k]=prev[k]
            rows[ident]=changed
    return pd.DataFrame(rows.values(), columns=cols)

def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists(): return pd.DataFrame()
    try: return pd.read_csv(path, dtype=str, keep_default_na=False)
    except pd.errors.EmptyDataError: return pd.DataFrame()

def save_csv(path: Path, frame: pd.DataFrame):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_name(path.name+'.tmp')
    frame.fillna('').to_csv(tmp,index=False,encoding='utf-8-sig')
    tmp.replace(path)
