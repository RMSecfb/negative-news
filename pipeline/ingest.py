"""Google News RSS in zh/en + issuer-independent market discovery; does not require a paid news feed."""
from __future__ import annotations
import csv
import os
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
import pandas as pd
import requests
from .common import news_key

MARKET_QUERIES=[
 '"credit default swap" spreads surge OR spike stocks',
 '"debt default" OR "missed interest payment" company',
 '"credit rating downgrade" OR "going concern" company',
 '"liquidity crisis" OR "debt restructuring" stocks',
 '"SEC investigation" OR "securities fraud" company',
 '"bankruptcy filing" OR "chapter 11" company',
 '"major cyberattack" OR "data breach" company',
 '美股 公司 違約 信用違約交換 CDS 信用風險',
 '美股 監管 調查 重大 訴訟 債務 危機',
]

def company_file(path:Path):
    df=pd.read_excel(path,dtype=str).fillna('')
    names={str(c).strip().lower():c for c in df.columns}
    tc=next((names[x] for x in ('ticker','symbol','股票代號') if x in names),None)
    cc=next((names[x] for x in ('company','name','股票名稱') if x in names),None)
    if tc is None or cc is None: raise ValueError(f'{path} 需有 Ticker、Company 欄位')
    return [(str(r[tc]).strip().upper(),str(r[cc]).strip()) for _,r in df.iterrows() if str(r[tc]).strip() and str(r[cc]).strip()]

def optional_watchlist(path:Path):
    if not path.exists(): return []
    with path.open(encoding='utf-8-sig',newline='') as f:
        return [(r['Ticker'].strip().upper(),r['Company'].strip()) for r in csv.DictReader(f) if r.get('Ticker') and r.get('Company')]

def ticker_batch(companies:list[tuple[str,str]], size:int, day:datetime):
    """Rotating selection across 500; does not silently claim full 500 coverage."""
    n=len(companies)
    if size<=0 or size>=n: return companies
    batch_count=(n+size-1)//size
    current=(day.date().toordinal()%batch_count)
    return companies[current*size:(current+1)*size]

def rss_fetch(query:str,language:str,session:requests.Session,limit:int):
    loc={'en':('en-US','US','US:en'),'zh':('zh-TW','TW','TW:zh-Hant')}[language]
    u='https://news.google.com/rss/search'
    r=session.get(u,params={'q':query,'hl':loc[0],'gl':loc[1],'ceid':loc[2]},timeout=22,
        headers={'User-Agent':'Mozilla/5.0 (compatible; RiskNewsResearch/1.0)'})
    r.raise_for_status()
    try:
        root=ET.fromstring(r.content)
    except ET.ParseError as exc:
        raise ValueError('Invalid Google News RSS response') from exc
    rows=[]
    for item in root.findall('.//channel/item')[:limit]:
        def val(tag):
            el=item.find(tag)
            return ''.join(el.itertext()).strip() if el is not None else ''
        title=val('title')
        source=val('source') or 'Google News'
        if title.endswith(' - '+source):title=title[:-(len(source)+3)]
        body=re.sub('<[^>]*>',' ',val('description'))
        rows.append({'Title':title,'Summary':re.sub(r'\s+',' ',body)[:1000],
          'URL':val('link'),'Source':source,'Published Time':val('pubDate'),'Language':language})
    return rows

def find_issuers(title:str,known:list[tuple[str,str]]):
    """Conservative verification: only identify a ticker if a known exact alias occurs in title."""
    match=[]
    for t,c in known:
        company=re.sub(r'\s+(?:incorporated|inc\.?|corp\.?|corporation|ltd\.?|limited|holdings?|group|class\s+[ab])\b.*$','',c,flags=re.I).strip()
        # Skip overly short/generic corporate names to reduce wrong issuer associations.
        if len(company)>=5 and re.search(r'(?<![A-Za-z])'+re.escape(company)+r'(?![A-Za-z])',title,re.I):
            match.append((t,c))
        elif len(t)>=3 and re.search(r'(?<![A-Za-z])(?:NYSE:|NASDAQ:|\()\s*'+re.escape(t)+r'\b',title,re.I):
            match.append((t,c))
    return list(dict.fromkeys(match))[:3]

def fetch_news(repo:Path, day:datetime, count:int=60, per_feed:int=6, pause:float=.3):
    sp=company_file(repo/'SP500.xlsx')
    watch=optional_watchlist(repo/'config'/'extra_watchlist.csv')
    all_known=list(dict.fromkeys(watch+sp))
    tickers=list(dict.fromkeys(watch+ticker_batch(sp,count,day)))
    session=requests.Session()
    rows=[]; errors=[]; calls=0
    # Limit Yahoo-like scrape rate / RSS fetching; still subject to provider throttling.
    for ticker, company in tickers:
        for lang in ('en','zh'):
            q=f'"{company}" (investigation OR default OR downgrade OR lawsuit OR debt OR CDS OR 風險 OR 違約 OR 調查) when:2d'
            try:
                items=rss_fetch(q,lang,session,per_feed)
                for x in items: rows.append({**x,'Ticker':ticker,'Company':company,'監控池':'CUSTOM' if (ticker,company) in watch else 'SP500_ROTATION'})
            except Exception as exc: errors.append(f'{ticker}/{lang}:{str(exc)[:110]}')
            calls+=1
            time.sleep(pause)
    # Market-wide discovery is independent of ticker list.
    for q in MARKET_QUERIES:
        lang='zh' if re.search('[\u4e00-\u9fff]',q) else 'en'
        try:
            for x in rss_fetch(q,lang,session,per_feed):
                issuers=find_issuers(x['Title'],all_known)
                if not issuers: issuers=[('','')]
                for ticker,company in issuers:rows.append({**x,'Ticker':ticker,'Company':company,'監控池':'MARKET_DISCOVERY'})
        except Exception as exc: errors.append(f'MARKET/{lang}:{str(exc)[:110]}')
        calls+=1;time.sleep(pause)
    for r in rows:r['news_id']=news_key(r)
    df=pd.DataFrame(rows)
    if not df.empty:
        df=df.drop_duplicates(['news_id','監控池']).reset_index(drop=True)
    return df,{'rss_calls':calls,'ticker_count':len(tickers),'sp500_total':len(sp),'errors':errors,
               'scan_mode':'FULL_SP500' if count>=len(sp) else 'ROTATING_SUBSET', 'news_rows':len(df)}
