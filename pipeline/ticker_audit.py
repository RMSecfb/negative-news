"""Audit every stored news day without recrawling or calling an LLM."""
from __future__ import annotations
import io,json,os,re
from pathlib import Path
from datetime import datetime,timedelta
import pandas as pd
from .common import truth
from .issuers import load_index
from .records import TPE,csv_frame,normalize,read_days,merge_news,excel_bytes
from .repository import GitHubRepository,LocalRepository,read_json,json_bytes

ROOT=Path(__file__).resolve().parents[1]
TICKER_FIELDS=['Ticker','Company','Exchange','Ticker狀態','Ticker來源','Ticker候選','Ticker核對依據','Ticker參照日期','上市狀態','原始Ticker','Ticker稽核警示']

def audit(repository,root=ROOT):
    paths=set(repository.list(''))
    sources=sorted(p for p in paths if re.fullmatch(r'(?:news_daily|ai_news_daily)/\d{4}-\d{2}-\d{2}\.csv|risk_news_daily/\d{4}-\d{2}-\d{2}\.xlsx',p))
    cache={}
    dates=[]
    skipped=0
    for path in sources:
        print('Audit source: '+path,flush=True)
        raw=repository.read(path)
        cache[path]=raw
        frame=pd.read_excel(io.BytesIO(raw),dtype=str).fillna('') if path.endswith('.xlsx') else csv_frame(raw)
        normalized=normalize(frame)
        skipped+=len(frame)-len(normalized)
        dates.append(datetime.fromisoformat(Path(path).stem).date())
        if not normalized.empty:
            dates.extend(datetime.fromisoformat(d).date() for d in normalized['日期'].unique())
    if not dates:
        return {'rows':0,'sources':0}
    class Snapshot:
        def read(self,path,revision=None):
            if path not in paths:return None
            if path not in cache:cache[path]=repository.read(path)
            return cache[path]
    index=load_index(root)
    frame=read_days(Snapshot(),min(dates),max(dates),issuer_index=index)
    # Include the original persisted history if it contains additional stories.
    history=Snapshot().read('history/negative_news_history.csv')
    if history:
        extra=csv_frame(history)
        if 'Published Time' in extra:
            frame,_=merge_news(frame,extra)
    before=frame.copy()
    frame=index.enrich(frame)
    original={str(row['news_id']):row for path,raw in cache.items() if path.startswith(('news_daily/','ai_news_daily/')) and path.endswith('.csv') for row in normalize(csv_frame(raw)).to_dict('records')}
    rows=[]
    for row in frame.to_dict('records'):
        old=original.get(str(row['news_id']),{})
        rows.append({**row,'稽核前Ticker':old.get('Ticker',row.get('原始Ticker','')),'Ticker是否變更':str(old.get('Ticker',row.get('原始Ticker','')))!=str(row.get('Ticker',''))})
    audit_frame=pd.DataFrame(rows).fillna('')
    stamp=datetime.now(TPE).strftime('%Y%m%d-%H%M%S')
    report={'id':'ticker-audit-'+stamp,'type':'ticker-audit','status':'success','started_at':datetime.now(TPE).isoformat(),'finished_at':datetime.now(TPE).isoformat(),
            'sources':len(sources),'rows':len(frame),'changed':int(audit_frame['Ticker是否變更'].sum()),'missing_publication_time':skipped,
            'ticker_statuses':frame['Ticker狀態'].value_counts().to_dict(),'llm_calls':0}
    def build(read,revision):
        files={}
        runs=read_json(read('news_runs/index.json'),{'runs':[]})
        runs['runs']=[report]+[r for r in runs['runs'] if r.get('id')!=report['id']]
        for day,group in frame.groupby('日期'):
            latest=normalize(csv_frame(read(f'news_daily/{day}.csv')))
            current={str(r['news_id']):r for r in latest.to_dict('records')}
            updates={str(r['news_id']):r for r in group.to_dict('records')}
            result=[]
            identities=list(current)+[ident for ident in updates if ident not in current]
            for ident in identities:
                row=dict(updates.get(ident) or current[ident])
                existing=current.get(str(row['news_id']))
                update=updates.get(str(row['news_id']))
                if existing:
                    # Audit changes identity fields only; concurrent event decisions win.
                    row={**row,**existing}
                if update and not (existing and truth(existing.get('已人工覆核')) and (existing.get('Ticker') or 'Ticker' in str(existing.get('人工修改欄位','')).split('|'))):
                    row.update({k:update.get(k,'') for k in TICKER_FIELDS})
                    if update.get('Ticker狀態') not in ('已核對','人工確認') and not truth(row.get('已人工覆核')) and row.get('分類狀態')!='待補分析':
                        row.update(需要人工覆核='True',分類狀態='待人工覆核')
                elif update:
                    row['Ticker稽核警示']=update.get('Ticker稽核警示','')
                result.append(row)
            out=pd.DataFrame(result).fillna('')
            files[f'news_daily/{day}.csv']=out.to_csv(index=False).encode('utf-8-sig')
            files[f'news_daily/{day}.xlsx']=excel_bytes(out,runs['runs'])
        reference=root/'config'/'us_symbol_reference.csv'
        if reference.exists():files['config/us_symbol_reference.csv']=reference.read_bytes()
        files[f'news_audits/{stamp}.csv']=audit_frame.to_csv(index=False).encode('utf-8-sig')
        files[f'news_audits/{stamp}.json']=json_bytes(report)
        files[f'news_runs/{report["id"]}.json']=json_bytes(report)
        files['news_runs/index.json']=json_bytes(runs)
        return files,report
    return repository.transaction(build,'Audit and backfill verified tickers across all stored news dates')

def main():
    import argparse
    parser=argparse.ArgumentParser()
    parser.add_argument('--local')
    args=parser.parse_args()
    repo=LocalRepository(args.local) if args.local else GitHubRepository(os.environ['GITHUB_REPOSITORY'],os.environ['GITHUB_TOKEN'])
    print(json.dumps(audit(repo),ensure_ascii=False,indent=2),flush=True)

if __name__=='__main__':main()
