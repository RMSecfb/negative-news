"""Batch entrypoint, designed for GitHub Actions; keeps the existing Streamlit app intact."""
from __future__ import annotations
import argparse
import json
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import pandas as pd
from .common import load_csv, save_csv, merge_records, news_key
from .events import load_rules, candidate_rules, CDS_PAT, ACTUAL_DEFAULT, pick_rule, SYSTEM_CODES
from .ingest import fetch_news
from .market import priority, recent_records, deduplicate_headlines, coverage
from .llm import analyze, MODEL, PROMPT_VERSION

TPE=ZoneInfo('Asia/Taipei')
BASE=Path(__file__).resolve().parents[1]
OUTPUT_FIELDS=['news_id','日期','Ticker','Company','Published Time','Title','Summary','Source','URL','Language','監控池',
 'FinBERT','LLM辨識公司','Keyword候選','Keyword事件代碼','LLM狀態','LLM事件代碼','事件狀態','Confirmed_Default','Evidence',
 '中文摘要','Event_Code','事件中文','Level','Action','分類狀態','需要人工覆核','已人工覆核',
 'LLM模型','Prompt版本','input_tokens','output_tokens','API錯誤']

def normalize_legacy(repo:Path,daystr:str):
    """Import existing date snapshot if present. No writes to old risk_news_daily."""
    path=repo/'risk_news_daily'/f'{daystr}.xlsx'
    if not path.exists():return pd.DataFrame()
    try:df=pd.read_excel(path,dtype=str).fillna('')
    except Exception as exc:
        print('WARN legacy import:',exc);return pd.DataFrame()
    for col in ['Ticker','Company','Published Time','Title','Source','URL','FinBERT']:
        if col not in df: df[col]=''
    df=df[(df['Title']!='') & (df['URL']!='')].copy()
    df['監控池']='LEGACY_SNAPSHOT'
    df['Language']=''
    df['Summary']=''
    df['news_id']=[news_key(r) for r in df.to_dict('records')]
    return df

def useful_candidate(r):
    title=str(r.get('Title',''))
    hits=str(r.get('Keyword候選',''))
    if CDS_PAT.search(title) or ACTUAL_DEFAULT.search(title):return True
    if hits:return True
    if any(pool in str(r.get('監控池','')) for pool in ('MARKET_DISCOVERY','CUSTOM')):return True
    try:return float(r.get('FinBERT','1'))<=0
    except (TypeError,ValueError):return False

def get_priority(r):
    return priority(r)

def classify(frame:pd.DataFrame,rules:list[dict],api_key:str,max_llm:int,daystr:str):
    if frame.empty:return pd.DataFrame(columns=OUTPUT_FIELDS),{'llm_calls':0,'pending':0}
    available_codes=[str(r.get('Event_Code','')) for r in rules if r.get('Event_Code') and str(r.get('Event_Code')) not in SYSTEM_CODES]
    code_labels={str(r.get('Event_Code','')):str(r.get('事件中文','') or '') for r in rules}
    rows=[]
    for r in frame.fillna('').to_dict('records'):
        title=str(r.get('Title','')); summary=str(r.get('Summary',''))
        matches=candidate_rules(title+' '+summary,rules)
        r['Keyword候選']='|'.join(dict.fromkeys(h for x in matches for h in x['hits']))[:450]
        r['Keyword事件代碼']='|'.join(x['Event_Code'] for x in matches[:6])
        r['日期']=daystr
        r.setdefault('LLM狀態','PENDING')
        r.setdefault('已人工覆核','False')
        if str(r.get('LLM狀態','')) not in ('OK','MANUAL') and str(r.get('已人工覆核','')).lower() not in ('true','1'):
            # Legacy Event_Code must not be treated as an AI-verified decision.
            for field in ('Event_Code','事件中文','Level','Action'):
                r[field]=''
        if str(r.get('已人工覆核','')).lower() in ('true','1'):
            r['LLM狀態']='MANUAL';r['分類狀態']='人工覆核'
        rows.append(r)
    # Small calls quota. Repeated runs reuse successful analyses; errors stay queued.
    todo=sorted([r for r in rows if str(r.get('LLM狀態','')) not in ('OK','MANUAL') and useful_candidate(r)],key=get_priority,reverse=True)
    calls=0
    for r in todo:
        if not api_key:
            r['LLM狀態']='NO_API_KEY'; r['分類狀態']='待人工覆核';r['需要人工覆核']='True'
            continue
        if calls>=max_llm:
            r['LLM狀態']='QUOTA_PENDING';r['分類狀態']='待人工覆核';r['需要人工覆核']='True'
            continue
        answer, usage, error=analyze(r,available_codes,api_key,code_labels)
        calls+=1
        r['LLM模型']=usage.get('model',MODEL)
        r['Prompt版本']=PROMPT_VERSION
        r['input_tokens']=usage.get('input_tokens','')
        r['output_tokens']=usage.get('output_tokens','')
        if answer is None:
            r['LLM狀態']='API_ERROR';r['API錯誤']=error;r['分類狀態']='待人工覆核';r['需要人工覆核']='True'
            if 'rate limit' in str(error).lower():
                break
            continue
        r['LLM狀態']='OK'
        r['LLM辨識公司']=answer['company_mentioned']
        r['LLM事件代碼']='|'.join(answer['event_codes'])
        r['事件狀態']=answer['event_status']
        r['Confirmed_Default']=str(answer['confirmed_default'])
        r['Evidence']=answer['evidence'];r['中文摘要']=answer['summary_zh']
        chosen=next((pick_rule(c,rules) for c in answer['event_codes'] if pick_rule(c,rules)),None)
        conflict=bool(CDS_PAT.search(str(r.get('Title',''))) and not ACTUAL_DEFAULT.search(str(r.get('Title',''))) and answer['confirmed_default'])
        if chosen and answer['relevant']:
            r['Event_Code']=str(chosen['Event_Code'])
            r['事件中文']=str(chosen.get('事件中文','') or '')
            r['Level']=str(chosen.get('Level','') or '')
            r['Action']=str(chosen.get('Action','') or '')
        # Never let an unsupported confirmed DEFAULT alone classify a CDS headline as actual default.
        if CDS_PAT.search(str(r.get('Title',''))) and not ACTUAL_DEFAULT.search(str(r.get('Title',''))):
            label=str(r.get('事件中文',''))
            is_payment_default=('DEFAULT' in str(r.get('Event_Code','')).upper() or ('違約' in label and not any(x in label for x in ('交換','互換','掉期','CDS'))))
            if is_payment_default:
                r.update({'Event_Code':'','事件中文':'','Level':'','Action':''});conflict=True
        try: major_level=int(float(r.get('Level') or 0))>=4
        except (ValueError,TypeError): major_level=False
        if answer['relevant']:
            needs=(answer['requires_review'] or conflict or not r.get('Event_Code') or
                   not answer['evidence'] or major_level or answer['confirmed_default'] or not r.get('Ticker'))
        else:
            needs=answer['requires_review'] or not answer['evidence'] or not r.get('Ticker')
        r['需要人工覆核']=str(needs)
        r['分類狀態']='待人工覆核' if needs else '已分類'
        if not answer['relevant']:
            r['分類狀態']='待人工覆核' if needs else '無關新聞'
    for r in rows:
        if str(r.get('Event_Code','')).upper() == 'DEFAULT' and (
            str(r.get('Confirmed_Default','')).lower() not in ('true','1') or
            str(r.get('事件狀態','')) != 'CONFIRMED'
        ) and str(r.get('已人工覆核','')).lower() not in ('true','1'):
            r.update({'Event_Code':'','事件中文':'疑似違約／尚未確認（待覆核）','Level':'','Action':'人工核實新聞證據','分類狀態':'待人工覆核','需要人工覆核':'True'})
        if not useful_candidate(r):
            r.setdefault('分類狀態','未入選初篩')
        elif str(r.get('LLM狀態','')) not in ('OK','MANUAL'):
            r['分類狀態']='待人工覆核';r['需要人工覆核']='True'
    return pd.DataFrame(rows).reindex(columns=OUTPUT_FIELDS,fill_value=''),{
        'llm_calls':calls,'pending':sum(str(r.get('分類狀態',''))=='待人工覆核' for r in rows),
        'processed':sum(str(r.get('LLM狀態',''))=='OK' for r in rows),
        'no_key':not bool(api_key)}

def run():
    p=argparse.ArgumentParser()
    p.add_argument('--date',default=datetime.now(TPE).date().isoformat())
    p.add_argument('--max-tickers',type=int,default=int(os.getenv('MAX_TICKERS','60')))
    p.add_argument('--max-llm',type=int,default=int(os.getenv('MAX_LLM_CALLS','35')))
    p.add_argument('--scope',choices=['market','hybrid'],default='market')
    p.add_argument('--no-fetch',action='store_true',help='Reprocess existing snapshot and AI raw news without RSS.')
    a=p.parse_args(); day=datetime.fromisoformat(a.date).replace(tzinfo=TPE)
    if not a.no_fetch and a.date != datetime.now(TPE).date().isoformat():
        raise SystemExit('歷史日期補跑請勾選 no_fetch，RSS 搜尋只支援即時新聞，不能當成歷史回測。')
    repo=BASE
    rules=load_rules(repo/'Parameter_Event.xlsx')
    rawfile=repo/'ai_news_raw'/f'{a.date}.csv'
    oldraw=load_csv(rawfile)
    oldresult=load_csv(repo/'ai_news_daily'/f'{a.date}.csv')
    if a.no_fetch: fresh=pd.DataFrame(); crawl_stats={'fetch':'skipped'}
    else: fresh,crawl_stats=fetch_news(repo,day,a.max_tickers,scope=a.scope)
    legacy=normalize_legacy(repo,a.date) if a.scope=='hybrid' else pd.DataFrame()
    merged=merge_records(oldraw,fresh)
    merged=merge_records(merged,legacy)
    reference_time=datetime.now(TPE) if a.date==datetime.now(TPE).date().isoformat() else day.replace(hour=23,minute=59,second=59)
    merged, excluded_dates=recent_records(merged,reference_time)
    if a.scope=='market' and not merged.empty:
        merged=merged[merged.get('監控池',pd.Series('',index=merged.index)).astype(str).str.contains('MARKET_DISCOVERY|CUSTOM',regex=True)]
    merged=deduplicate_headlines(merged)
    save_csv(rawfile,merged)
    # Preserve previously analyzed results: enrich fresh with old result before scoring.
    if not oldresult.empty:
        oldresult=oldresult[oldresult['news_id'].isin(set(merged.get('news_id',[])))]
    enriched=merge_records(merged,oldresult)
    output, cls_stats=classify(enriched,rules,os.getenv('GROQ_API_KEY',''),max(0,a.max_llm),a.date)
    # classify already includes cached/manual rows; empty corrected fields must stay empty.
    # Respect older human decisions from the existing Streamlit manual override file.
    overrides=load_csv(repo/'history'/'negative_news_manual_overrides.csv')
    if not overrides.empty and not output.empty and 'URL' in overrides:
        for _,override in overrides.iterrows():
            uri=str(override.get('URL','')).strip()
            ticker=str(override.get('Ticker','')).strip().upper()
            if not uri:continue
            mask=(output['URL'].astype(str).str.strip()==uri)
            if ticker:mask=mask & (output['Ticker'].astype(str).str.upper()==ticker)
            for col in ('事件中文','Level','Action'):
                if col in override and str(override[col]).strip():output.loc[mask,col]=str(override[col])
            output.loc[mask,'已人工覆核']='True'
            output.loc[mask,'LLM狀態']='MANUAL'
            output.loc[mask,'分類狀態']='人工覆核'
            output.loc[mask,'需要人工覆核']='False'
    save_csv(repo/'ai_news_daily'/f'{a.date}.csv',output)
    stats={'date':a.date,'scope':a.scope,'window_hours':48,'excluded_stored_dates':excluded_dates,**coverage(output),**crawl_stats,**cls_stats,'raw_distinct':len(merged),'final_rows':len(output),
      'note':'AI CSV is separate from legacy Streamlit snapshot; no replacement of old risk_news_daily.'}
    out=repo/'ai_news_daily'/f'{a.date}.json';out.write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(stats,ensure_ascii=False,indent=2))
    # An incomplete / failed web fetch is reported rather than falsely claiming full coverage.
    if crawl_stats.get('errors') and len(crawl_stats['errors'])>0:
        print('WARN RSS partial failure:',len(crawl_stats['errors']))

if __name__=='__main__':run()
