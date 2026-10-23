"""Market-wide event discovery, freshness filtering and transparent triage."""
from __future__ import annotations
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import re
import pandas as pd

# Separate event queries avoid requiring all risk keywords to occur together.
MARKET_QUERIES = [
 '"credit default swap" (surge OR widen OR spike) (company OR stock)',
 '("debt default" OR "missed payment" OR "covenant breach") (company OR stock)',
 '("bankruptcy" OR "chapter 11" OR "insolvency") (company OR stock)',
 '("liquidity crisis" OR "debt restructuring" OR "going concern") company',
 '("credit rating downgrade" OR "junk rating") company',
 '("SEC investigation" OR "DOJ investigation" OR "criminal charges") company',
 '("accounting fraud" OR "financial restatement" OR "auditor resigns") company',
 '("antitrust" OR "regulatory fine" OR "trading halt" OR "delisting") stock',
 '("profit warning" OR "cuts guidance" OR "withdraws guidance") company',
 '("earnings miss" OR "revenue plunge" OR "dividend suspended") stock',
 '("major recall" OR "fatal accident" OR "plant shutdown") company',
 '("cyberattack" OR "ransomware" OR "data breach") company',
 '("export ban" OR "sanctions" OR "FDA rejects") company stock',
 '("shares plunge" OR "stock plunges" OR "shares tumble") company',
 '美股 (CDS OR 信用風險 OR 債務危機 OR 信評降級)',
 '美股 (破產 OR 違約 OR 債務重整 OR 流動性危機)',
 '美股 (監管調查 OR 反壟斷 OR 重大訴訟 OR 裁罰)',
 '美股 (財報造假 OR 重編財報 OR 會計弊案)',
 '美股 (下修財測 OR 獲利預警 OR 財報不如預期 OR 暴跌)',
 '美股 (網路攻擊 OR 資料外洩 OR 重大召回 OR 停產)',
]

SEVERE = re.compile(r'bankrupt|chapter 11|insolven|missed.{0,20}payment|default on|liquidity crisis|accounting fraud|criminal charges|going concern|破產|債務違約|流動性危機|財報造假|會計弊案', re.I)
SIGNIFICANT = re.compile(r'cuts guidance|profit warning|credit rating downgrade|investigation|restatement|cyberattack|recall|plung|下修財測|調查|重大召回|暴跌|信評降級', re.I)
EDITORIAL_SOURCE = re.compile(r'reuters|bloomberg|associated press|financial times|wall street journal|cnbc|路透|彭博|中央社|經濟日報|鉅亨', re.I)

def publication_time(value):
    try:
        dt = parsedate_to_datetime(str(value))
        if dt.tzinfo is None:
            return None
        return dt.astimezone(timezone.utc)
    except (ValueError, TypeError, OverflowError):
        try:
            dt = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
            return dt.astimezone(timezone.utc) if dt.tzinfo else None
        except (ValueError, TypeError):
            return None

def recent_records(frame, now, hours=48):
    """Drop unverifiable, stale and future dates; apply to stored rows on reruns too."""
    if frame.empty:
        return frame.copy(), 0
    cutoff = now.astimezone(timezone.utc) - timedelta(hours=hours)
    upper = now.astimezone(timezone.utc) + timedelta(minutes=5)
    times = frame.get('Published Time', pd.Series('', index=frame.index)).map(publication_time)
    keep = times.map(lambda dt: dt is not None and cutoff <= dt <= upper)
    return frame.loc[keep].copy(), int((~keep).sum())

def deduplicate_headlines(frame):
    """Collapse exact normalized headlines per verified issuer, preserving pool tags."""
    if frame.empty:
        return frame.copy()
    groups = {}
    for row in frame.fillna('').to_dict('records'):
        key = (str(row.get('Ticker', '')).upper(), re.sub(r'\W+', '', str(row.get('Title','')).casefold()))
        if key not in groups:
            groups[key] = row
        else:
            prev = groups[key]
            tags = set(str(prev.get('監控池','')).split('|')) | set(str(row.get('監控池','')).split('|'))
            prev['監控池'] = '|'.join(sorted(filter(None, tags)))
    return pd.DataFrame(groups.values())

def priority(record):
    text = str(record.get('Title','')) + ' ' + str(record.get('Summary',''))
    score = 120 if SEVERE.search(text) else 70 if SIGNIFICANT.search(text) else 0
    if re.search(r'(?<![A-Za-z])CDS(?![A-Za-z])|credit default swap|信用違約交換', text, re.I):
        score = max(score, 100)
    if EDITORIAL_SOURCE.search(str(record.get('Source',''))):
        score += 20
    if record.get('Keyword候選'):
        score += 15
    return score

def coverage(frame):
    if frame.empty:
        return {'candidate_total':0, 'analyzed_total':0, 'remaining_candidates':0, 'api_errors':0}
    statuses = frame.get('LLM狀態', pd.Series('', index=frame.index)).fillna('')
    return {
        'candidate_total':len(frame),
        'analyzed_total':int(statuses.isin(['OK','MANUAL']).sum()),
        'remaining_candidates':int((~statuses.isin(['OK','MANUAL'])).sum()),
        'api_errors':int((statuses=='API_ERROR').sum()),
    }
