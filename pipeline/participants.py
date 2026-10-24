"""Institutional/insider result adapter for JUWangWang/SEC13F.

Consumes the source project's workbooks. Does not turn missing/failed filings
into zero holdings or no-trade signals, and never changes news decisions.
"""
from __future__ import annotations
import io
import json
import re
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
from .repository import GitHubRepository, read_json, json_bytes

SOURCE_REPO = 'JUWangWang/SEC13F'
PREFIX = 'institutional_insider/'
DISPLAY_FIELDS = ['13F 機構動向', 'Form4 內部人動向']
AUDIT_FIELDS = ['13F參考季度', '13F最近申報日', '13F資料狀態', 'Form4分析開始', 'Form4分析結束', 'Form4資料狀態', '機構內部人來源版本', '機構內部人同步時間']


def workbook_paths(paths, kind):
    pattern = (r'institutional_13f/(\d{4}Q[1-4])/13F_\1_Full_Universe\.xlsx' if kind == '13f'
               else r'insider_form4/(\d{4}-\d{2}-\d{2})/Form4_\1_\d+D_Full_Universe\.xlsx')
    return sorted([p for p in paths if re.fullmatch(pattern, p)], reverse=True)


def parse_form4(raw):
    sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype={'Ticker': str})
    if not {'Ticker Summary', 'Form4 Detail', 'Run Log'}.issubset(sheets):
        raise ValueError('Form4 缺少 Ticker Summary、Form4 Detail 或 Run Log')
    summary = sheets['Ticker Summary']
    needed = {'Ticker', 'Analysis Start', 'Analysis End', 'Buy Count', 'Sell Count'}
    if not needed.issubset(summary):
        raise ValueError('Form4 摘要缺少代號、期間或 P/S 筆數')
    summary['Ticker'] = summary['Ticker'].fillna('').str.strip().str.upper()
    if summary.empty or summary['Ticker'].eq('').any() or summary['Ticker'].duplicated().any():
        raise ValueError('Form4 摘要代號空白、重複或沒有資料')
    starts = pd.to_datetime(summary['Analysis Start'], errors='coerce')
    ends = pd.to_datetime(summary['Analysis End'], errors='coerce')
    if starts.isna().any() or ends.isna().any() or starts.gt(ends).any() or starts.nunique()!=1 or ends.nunique()!=1:
        raise ValueError('Form4 分析期間無效或混合期間')
    summary['Analysis Start'] = starts.dt.strftime('%Y-%m-%d')
    summary['Analysis End'] = ends.dt.strftime('%Y-%m-%d')
    for col in ['Buy Count', 'Sell Count']:
        values = pd.to_numeric(summary[col], errors='coerce')
        if values.isna().any() or values.lt(0).any() or values.mod(1).ne(0).any():
            raise ValueError('Form4 P/S 筆數無效')
        summary[col] = values.astype(int)
    total = summary['Buy Count'] + summary['Sell Count']
    summary['NPR'] = (summary['Buy Count'] - summary['Sell Count']).div(total.where(total.ne(0)))
    summary['內部人交易動向'] = summary['NPR'].map(lambda n: '無P/S交易' if pd.isna(n) else '近期淨買入' if n > 0 else '近期淨賣出' if n < 0 else '買賣中性')
    logs = sheets['Run Log']
    if not {'Ticker', 'Status'}.issubset(logs) or logs['Ticker'].duplicated().any():
        raise ValueError('Form4 執行紀錄缺少代號、狀態或有重複代號')
    logs['Ticker'] = logs['Ticker'].fillna('').astype(str).str.strip().str.upper()
    if logs['Ticker'].duplicated().any():
        raise ValueError('Form4 執行紀錄代號重複')
    statuses = logs.set_index('Ticker').to_dict('index')
    qualities = []
    for ticker in summary['Ticker']:
        record = statuses.get(ticker, {})
        status = str(record.get('Status', '')).upper()
        flags = []
        if status != 'OK':
            flags.append(status or '缺執行紀錄')
        for field, label in [('Parse Errors','解析失敗'), ('Issuer Unknown Filings','主體未核對'), ('Form 4/A Amendments','有修正版待核對')]:
            value = pd.to_numeric(record.get(field, 0), errors='coerce')
            if pd.notna(value) and value > 0:
                flags.append(label)
        qualities.append('；'.join(flags) if flags else 'OK')
    summary['資料品質'] = qualities
    return sheets


def read_snapshot(repository, revision=None):
    if revision is None and isinstance(repository, GitHubRepository):
        revision = repository.revision()
    manifest = read_json(repository.read(PREFIX+'latest.json', revision), {'datasets': {}})
    manifest['_snapshot_revision'] = revision
    summaries = {}
    for kind, meta in manifest.get('datasets', {}).items():
        path = meta.get('summary_path', '')
        if kind not in ('13f', 'form4') or not path.startswith(PREFIX) or '..' in path:
            continue
        raw = repository.read(path, revision)
        if raw:
            summaries[kind] = pd.read_csv(io.BytesIO(raw), dtype=str).fillna('')
    return manifest, summaries


def annotate(news, snapshot=None, today=None):
    data = news.copy()
    if data.empty:
        return data
    manifest, summaries = snapshot or ({'datasets': {}}, {})
    today = today or datetime.now(ZoneInfo('Asia/Taipei')).date()
    maps = {}
    for kind, summary in summaries.items():
        if 'Ticker' not in summary or summary['Ticker'].duplicated().any():
            raise ValueError('機構內部人摘要代號缺少或重複')
        maps[kind] = summary.set_index('Ticker').to_dict('index')
    rows = []
    for original in data.fillna('').to_dict('records'):
        row = {**original, **{field:'' for field in DISPLAY_FIELDS+AUDIT_FIELDS}}
        ticker = str(row.get('Ticker','')).strip().upper()
        row['機構內部人同步時間'] = manifest.get('synced_at','')
        versions = []
        for kind, display, state in [('13f', DISPLAY_FIELDS[0], '13F資料狀態'), ('form4', DISPLAY_FIELDS[1], 'Form4資料狀態')]:
            meta = manifest.get('datasets',{}).get(kind,{})
            if meta.get('source_sha'):
                versions.append(kind+':'+meta['source_sha'][:12])
            if row.get('Market') != 'US':
                row[display] = row[state] = '非美股／不適用'
                continue
            if not ticker:
                row[display] = row[state] = '股票代號待核對'
                continue
            item = maps.get(kind,{}).get(ticker)
            if kind not in summaries:
                row[display] = row[state] = '資料讀取失敗' if manifest.get('load_error') else '尚無已發布成果'
                continue
            if not item:
                row[display] = row[state] = '不在本次追蹤清單'
                continue
            if kind == '13f':
                quarter = str(item.get('Quarter',''))
                row['13F參考季度'] = quarter
                row['13F最近申報日'] = item.get('Latest Filing Date','')
                qoq = pd.to_numeric(item.get('機構持股QoQ',''),errors='coerce')
                change = '無前季比較' if pd.isna(qoq) else f'{qoq:+.1%}'
                counts = ''
                for field,label in [('增持家數','增'),('減持家數','減')]:
                    n = pd.to_numeric(item.get(field,''),errors='coerce')
                    counts += f'{int(n)}{label}／' if pd.notna(n) else ''
                row[display] = f'{quarter}｜持股{change}｜{counts.rstrip("／")}'
                row[state] = '季度參考（非即時）'
                if meta.get('source_status') != 'available':
                    row[state] += '；沿用上次成果'
            else:
                begin, end = item.get('Analysis Start',''), item.get('Analysis End','')
                row['Form4分析開始'], row['Form4分析結束'] = begin,end
                quality = item.get('資料品質','缺品質紀錄')
                if quality != 'OK':
                    row[display] = f'截至{end}｜資料待覆核：{quality}'
                    row[state] = quality
                else:
                    buy,sell = int(float(item['Buy Count'])),int(float(item['Sell Count']))
                    signal = '無P/S交易' if buy+sell == 0 else '淨買入' if buy>sell else '淨賣出' if sell>buy else '買賣中性'
                    row[display] = f'截至{end}｜{signal}｜{buy}買／{sell}賣'
                    row[state] = '已完成期間查詢'
                parsed = pd.to_datetime(end,errors='coerce')
                if pd.notna(parsed) and parsed.date() < today-timedelta(days=3):
                    row[display] += '（資料較舊）'
                    row[state] += '；資料較舊'
        row['機構內部人來源版本'] = '｜'.join(versions)
        rows.append(row)
    return pd.DataFrame(rows, index=data.index)


def annotate_from_repository(news, repository, revision=None):
    try:
        snapshot = read_snapshot(repository,revision)
    except Exception:
        snapshot = ({'load_error':True,'datasets':{}}, {})
    return annotate(news,snapshot)


def sync(target, source=None):
    """Publish one validated snapshot. Failed imports retain the last good dataset."""
    from .analytics import parse_sec
    source = source or GitHubRepository(SOURCE_REPO)
    revision = source.revision()
    paths = source.list('', revision)
    imported, errors = {}, {}
    for kind in ['13f','form4']:
        candidates = workbook_paths(paths,kind)
        if not candidates:
            errors[kind] = '來源尚未發布成果 Excel'
            continue
        path = candidates[0]
        try:
            raw = source.read(path,revision)
            sheets = parse_sec(raw) if kind=='13f' else parse_form4(raw)
            logs = sheets['Run Log']
            if kind=='13f' and (not {'Status','Aggregate'}.issubset(logs) or not (logs['Status'].eq('OK') & logs['Aggregate'].eq('Y')).any()):
                raise ValueError('13F 沒有成功納入彙總的機構；保留上次成果')
            if kind=='form4' and not logs['Status'].isin(['OK','PARTIAL']).any():
                raise ValueError('Form4 所有公司擷取失敗；保留上次成果')
            summary = sheets['Ticker Summary'].copy()
            if kind=='13f':
                detail = sheets['13F Detail']
                if 'Filing Date' in detail:
                    dates = pd.to_datetime(detail['Filing Date'],errors='coerce').dt.strftime('%Y-%m-%d')
                    latest = detail.assign(_date=dates).groupby('Ticker')['_date'].max()
                    summary['Latest Filing Date'] = summary['Ticker'].map(latest).fillna('')
            imported[kind] = (raw,summary,path)
        except Exception as exc:
            errors[kind] = str(exc)[:200]
    stamp = datetime.now(ZoneInfo('Asia/Taipei')).isoformat()
    def build(read, rev):
        manifest = read_json(read(PREFIX+'latest.json'),{'datasets':{}})
        files = {}
        for kind in ['13f','form4']:
            if kind in imported:
                raw, summary, path = imported[kind]
                summary_path = PREFIX+kind+'/summary.csv'
                workbook_path = PREFIX+kind+'/'+path.split('/')[-1]
                files[summary_path] = summary.to_csv(index=False).encode('utf-8-sig')
                files[workbook_path] = raw
                manifest['datasets'][kind] = {'status':'ready','source_status':'available','source_repo':SOURCE_REPO,
                    'source_sha':revision,'source_path':path,'summary_path':summary_path,'workbook_path':workbook_path,'rows':len(summary),'loaded_at':stamp}
            else:
                meta = manifest['datasets'].setdefault(kind,{'status':'not_published'})
                meta.update(source_status='unavailable',error=errors[kind])
        manifest.update(synced_at=stamp,source_repo=SOURCE_REPO,checked_source_sha=revision)
        files[PREFIX+'latest.json'] = json_bytes(manifest)
        return files, manifest
    return target.transaction(build,'Sync institutional and insider transaction results')


def main():
    import os
    target = GitHubRepository(os.environ['GITHUB_REPOSITORY'],os.environ['GITHUB_TOKEN'])
    result = sync(target)
    print(json.dumps(result,ensure_ascii=False))


if __name__ == '__main__':
    main()
