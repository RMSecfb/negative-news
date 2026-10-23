"""Resolve issuers against company lists; never treat a model-generated ticker as verified."""
from __future__ import annotations
import re
from pathlib import Path
import pandas as pd


def name_key(value):
    text = str(value or '').casefold()
    text = re.sub(r'\b(incorporated|inc|corporation|corp|limited|ltd|class [ab])\b', '', text)
    return re.sub(r'[^a-z0-9\u3400-\u9fff]+', '', text)


def symbol(value, market='US'):
    text = str(value or '').strip().upper()
    if market == 'HK':
        digits = text.removesuffix('.HK')
        if re.fullmatch(r'\d{1,5}', digits):
            return str(int(digits)).zfill(4) + '.HK'
    return text


def company_frame(frame, market='US'):
    frame = frame.fillna('').copy()
    cols = {str(c).strip().lower(): c for c in frame.columns}
    tc = next((cols[k] for k in ('ticker', 'symbol', '股票代號') if k in cols), None)
    cc = next((cols[k] for k in ('company', 'name', '股票名稱') if k in cols), None)
    if tc is None or cc is None:
        raise ValueError('公司清單需有 Ticker/Symbol 與 Company/Name 欄位')
    out = pd.DataFrame({'Ticker': frame[tc].astype(str), 'Company': frame[cc].astype(str)})
    for key, fallback in [('Market', market), ('Exchange', ''), ('Aliases', '')]:
        out[key] = frame[cols[key.lower()]].astype(str) if key.lower() in cols else fallback
    out['Market'] = out['Market'].replace('', market).str.upper()
    if not out['Market'].isin(['US', 'HK']).all():
        raise ValueError('Market 目前支援 US 或 HK')
    out['Ticker'] = [symbol(t, m) for t, m in zip(out['Ticker'], out['Market'])]
    out['Company'] = out['Company'].str.strip()
    return out[(out.Ticker != '') & (out.Company != '')].drop_duplicates(['Market', 'Ticker'])


def load_companies(repo: Path, names=('SP500',), uploads=(), market='US'):
    frames = []
    for name in names:
        if name not in ('SP500', 'DowJones30'):
            continue
        if market != 'US':
            raise ValueError('Dow Jones 30 與 S&P 500 僅支援美股；港股請上傳清單')
        frames.append(company_frame(pd.read_excel(repo / f'{name}.xlsx', dtype=str), 'US'))
    if uploads:
        frames.append(company_frame(pd.DataFrame(uploads), market))
    if not frames:
        return pd.DataFrame(columns=['Ticker', 'Company', 'Market', 'Exchange', 'Aliases'])
    return pd.concat(frames, ignore_index=True).drop_duplicates(['Market', 'Ticker']).reset_index(drop=True)


class IssuerIndex:
    def __init__(self, companies):
        self.exact = {}
        self.symbols = {}
        self.names = []
        for r in companies.fillna('').to_dict('records'):
            self.symbols[(r['Market'], symbol(r['Ticker'], r['Market']))] = r
            aliases = [r['Company'], *re.split(r'[;；|]', r.get('Aliases', ''))]
            # Punctuation identifies distinct securities: MET$A must never alias META.
            if re.fullmatch(r'[A-Za-z0-9]+', r['Ticker']):
                aliases.append(r['Ticker'])
            # Explicit aliases belong to a validated symbol from the supplied reference.
            if r['Market'] == 'US' and r['Ticker'] == 'META':
                aliases.extend(['Meta', 'Facebook', 'Meta Platforms'])
            if r['Market'] == 'US' and r['Ticker'] == 'CRWV':
                aliases.append('CoreWeave')
            for alias in aliases:
                key = name_key(alias)
                if not key:
                    continue
                item = (r['Market'], symbol(r['Ticker'], r['Market']))
                self.exact.setdefault((r['Market'], key), {})[item] = r
            self.names.append(r)

    def resolve(self, company, market='US'):
        direct = self.symbols.get((market, str(company).strip().upper()))
        if direct:
            return dict(direct), '股票代號精確核對'
        key = name_key(company)
        exact = self.exact.get((market, key), {})
        if len(exact) == 1:
            return dict(next(iter(exact.values()))), '公司清單／別名核對'
        if len(exact) > 1:
            return None, '同名公司，需人工確認'
        if len(key) >= 4:
            matches = {(r['Market'], r['Ticker']): r for r in self.names
                       if r['Market'] == market and name_key(r['Company']).startswith(key)}
            if len(matches) == 1:
                return dict(next(iter(matches.values()))), '公司名稱唯一比對'
        return None, '公司代號未確認'

    def enrich(self, frame):
        rows = []
        for r in frame.fillna('').to_dict('records'):
            market = str(r.get('Market', 'US') or 'US').upper()
            r['Market'] = market
            if not str(r.get('Ticker', '')).strip():
                company = r.get('LLM辨識公司') or r.get('Company', '')
                match, reason = self.resolve(company, market)
                if match:
                    r.update(Ticker=symbol(match['Ticker'], market), Company=match['Company'],
                             Exchange=match.get('Exchange', ''), Ticker來源=reason)
                else:
                    r['Ticker來源'] = reason
            else:
                r['Ticker'] = symbol(r['Ticker'], market)
                r.setdefault('Ticker來源', '公司清單／原始紀錄')
            rows.append(r)
        return pd.DataFrame(rows)
