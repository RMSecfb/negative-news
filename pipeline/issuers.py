"""Resolve issuers against company lists; never treat a model-generated ticker as verified."""
from __future__ import annotations
import re
import io
import unicodedata
from pathlib import Path
import pandas as pd


def name_key(value):
    text = unicodedata.normalize('NFKC', str(value or '')).casefold().replace('’', "'")
    # Remove security descriptions, not arbitrary business-name words.
    text = re.sub(r'\s+(?:(?:class|series)\s+[a-z0-9]+\s+)?(?:common (?:stock|shares)|ordinary shares|american depositary.*|depositary (?:shares|receipts)|ads\b|adr\b).*$', '', text)
    text = re.sub(r'\b(incorporated|inc|corporation|corp|limited|ltd|plc|class [ab])\b', '', text)
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


def load_index(root, repository=None, extra=None):
    """One reference for collection, cache reuse, historical queries and exports."""
    frames = [load_companies(Path(root), ['SP500', 'DowJones30'])]
    for filename in ('extra_watchlist.csv', 'us_symbol_reference.csv', 'issuer_registry.csv'):
        path = Path(root) / 'config' / filename
        raw = repository.read('config/' + filename) if repository is not None else None
        if raw:
            frame = pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False)
        elif path.exists():
            frame = pd.read_csv(path, dtype=str, keep_default_na=False)
        else:
            continue
        frames.append(frame)
    if extra is not None and not extra.empty:
        frames.append(extra)
    return IssuerIndex(pd.concat(frames, ignore_index=True).fillna(''))


class IssuerIndex:
    def __init__(self, companies):
        self.exact = {}
        self.symbols = {}
        self.names = []
        self.entities = {}
        self.symbol_aliases = {}
        for r in companies.fillna('').to_dict('records'):
            r['Market'] = str(r.get('Market') or 'US').upper()
            r['Ticker'] = symbol(r.get('Ticker', ''), r['Market'])
            r.setdefault('EntityType', 'LISTED')
            if not r['EntityType']:
                r['EntityType'] = 'LISTED'
            if r['Ticker']:
                self.symbols[(r['Market'], r['Ticker'])] = r
            aliases = [r['Company'], *re.split(r'[;；|]', r.get('Aliases', ''))]
            self.symbol_aliases.setdefault((r['Market'], r['Ticker']), set()).update(a for a in aliases if a)
            # Punctuation identifies distinct securities: MET$A must never alias META.
            if re.fullmatch(r'[A-Za-z0-9]+', r['Ticker']):
                aliases.append(r['Ticker'])
            for alias in aliases:
                key = name_key(alias)
                if not key:
                    continue
                item = (r['Market'], r['Ticker'] or name_key(r['Company']))
                self.entities.setdefault(key, {})[item] = r
                if r['Ticker'] and r['EntityType'] == 'LISTED':
                    self.exact.setdefault((r['Market'], key), {})[item] = r
            self.names.append(r)

    def resolve(self, company, market='US'):
        direct = self.symbols.get((market, str(company).strip().upper()))
        if direct and not re.fullmatch(r'[A-Za-z0-9]+', str(company).strip()):
            return dict(direct), '股票代號精確核對（保留符號）'
        key = name_key(company)
        exact = self.exact.get((market, key), {})
        if len(exact) == 1:
            return dict(next(iter(exact.values()))), '公司清單／別名核對'
        if len(exact) > 1:
            return None, '同名公司，需人工確認'
        direct = self.symbols.get((market, str(company).strip().upper()))
        if direct:
            return dict(direct), '股票代號精確核對'
        # A name containing a ticker still needs its name and ticker to agree.
        explicit = re.search(r'\(([A-Z][A-Z0-9.$-]{0,10})\)', str(company))
        if explicit:
            candidate = self.symbols.get((market, explicit.group(1)))
            bare = str(company)[:explicit.start()].strip()
            if candidate and name_key(bare) in {name_key(x) for x in [candidate['Company'], *re.split(r'[;；|]', candidate.get('Aliases', ''))]}:
                return dict(candidate), '名稱與括號代號交叉核對'
        return None, '公司代號未確認'

    def corroborated(self, row, match):
        text = unicodedata.normalize('NFKC', str(row.get('Title', '')) + '\n' + str(row.get('Summary', ''))).replace('’', "'")
        aliases = self.symbol_aliases.get((match['Market'], match.get('Ticker','')), set()) or [match['Company']]
        # Reference names contain legal/security suffixes; compare their cleaned
        # base name as a whole token sequence, never an arbitrary prefix.
        for alias in aliases:
            alias = unicodedata.normalize('NFKC', str(alias)).replace('’', "'")
            key = name_key(alias)
            if not key:
                continue
            words = re.sub(r'[^a-z0-9\u3400-\u9fff]+', '', text.casefold())
            if re.search(r'[\u3400-\u9fff]', str(alias)) and key in words:
                return True
            base = re.sub(r'\s+(?:(?:class|series)\s+[a-z0-9]+\s+)?(?:common (?:stock|shares)|ordinary shares|american depositary.*|depositary (?:shares|receipts)).*$', '', str(alias), flags=re.I)
            base = re.sub(r'\b(?:incorporated|inc|corporation|corp|limited|ltd|plc)\.?', '', base, flags=re.I).strip(' ,.')
            if base and re.search(r'(?<![A-Za-z0-9])'+re.escape(base).replace(r'\ ', r'\s+')+r'(?![A-Za-z0-9])', text, re.I):
                return True
        ticker = match.get('Ticker', '')
        if ticker and len(ticker) >= 2 and ticker not in {'ON','IT','OR','SO','AI','ALL','ARE','FOR','NOW','US','GO','BE','AN','AT'}:
            return bool(re.search(r'(?<![A-Za-z0-9])'+re.escape(ticker)+r'(?:\.US)?(?![A-Za-z0-9])', text))
        return bool(ticker and re.search(r'(?:\$|NASDAQ\s*:|NYSE\s*:)'+re.escape(ticker)+r'\b', text, re.I))

    def enrich(self, frame, audit_existing=True):
        from .common import truth
        rows = []
        for r in frame.fillna('').to_dict('records'):
            market = str(r.get('Market', 'US') or 'US').upper()
            r['Market'] = market
            supplied = symbol(r.get('Ticker', ''), market)
            if truth(r.get('已人工覆核')) and (supplied or 'Ticker' in str(r.get('人工修改欄位','')).split('|')):
                r.update(Ticker=supplied, Ticker狀態='人工確認', Ticker來源='人工覆核優先')
                checked = self.symbols.get((market, supplied))
                r['Ticker稽核警示'] = '' if (not supplied or (checked and self.corroborated(r, checked))) else '保留人工代號；目前新聞文字無法自動佐證'
                rows.append(r)
                continue
            if supplied and not r.get('原始Ticker'):
                r['原始Ticker'] = supplied
            subject = str(r.get('LLM辨識公司') or r.get('Company') or '').strip()
            if subject.casefold() in ('無', 'none', 'null', 'n/a', 'unknown'):
                subject = ''
            match, reason = self.resolve(subject, market)
            provided = self.symbols.get((market, supplied))
            if match is None and provided and name_key(subject) in {name_key(x) for x in [provided['Company'], *re.split(r'[;；|]', provided.get('Aliases', ''))]}:
                match, reason = provided, '原始代號與公司名稱交叉核對'
            if not subject and supplied:
                match = self.symbols.get((market, supplied))
            strict = audit_existing
            if match and (not strict or self.corroborated(r, match)) and (not supplied or supplied == match['Ticker']):
                r.update(Ticker=match['Ticker'], Company=match['Company'], Exchange=match.get('Exchange', ''), Ticker來源=reason,
                         Ticker狀態='已核對', Ticker核對依據=match.get('ReferenceSource') or match.get('Source') or '公司清單',
                         Ticker參照日期=match.get('VerifiedAt', ''), 上市狀態=match.get('ListingStatus') or '公司清單／交易所參照', Ticker候選='')
            elif supplied and not audit_existing:
                r.update(Ticker=supplied, Ticker狀態='來源代號／待主體核對')
            else:
                entities = self.entities.get(name_key(subject), {})
                entity = next(iter(entities.values())) if len(entities) == 1 else None
                status = '主體證據不足' if match else '公司代號待核對' if subject else '市場／多公司或主體待辨識'
                if '同名公司' in reason:
                    status = '公司名稱／股票代號歧義'
                if supplied and match and supplied != match['Ticker']:
                    status = '代號與主體衝突'
                if entity and entity.get('EntityType') != 'LISTED':
                    status = '非上市機構'
                elif entity and entity['Market'] != market:
                    status = '非本市場上市主體'
                r.update(Ticker='', Ticker狀態=status, Ticker來源=status, Ticker候選=(match or entity or {}).get('Ticker', supplied),
                         Ticker核對依據=(entity or {}).get('ReferenceSource', ''), Ticker參照日期=(entity or {}).get('VerifiedAt', ''))
                if not r.get('Company') and subject:
                    r['Company'] = (entity or {}).get('Company', subject)
                if r.get('分類狀態') != '待補分析' and not truth(r.get('已人工覆核')):
                    r.update(需要人工覆核='True', 分類狀態='待人工覆核')
            rows.append(r)
        return pd.DataFrame(rows)
