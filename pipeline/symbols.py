"""Download the official US symbol directory; model output alone is never a symbol source."""
from __future__ import annotations
import io
from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
SOURCES = [('https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqlisted.txt', 'Symbol', 'NASDAQ'),
           ('https://www.nasdaqtrader.com/dynamic/SymDir/otherlisted.txt', 'ACT Symbol', '')]


def refresh(repo=ROOT):
    frames = []
    for url, symbol_column, exchange in SOURCES:
        response = requests.get(url, timeout=35)
        response.raise_for_status()
        data = pd.read_csv(io.StringIO(response.text), sep='|', dtype=str).fillna('')
        if symbol_column not in data or 'Security Name' not in data:
            raise ValueError('官方股票代號目錄欄位異常，未採用此回應')
        data = data[~data[symbol_column].str.startswith('File Creation Time')]
        if 'Test Issue' in data:
            data = data[data['Test Issue'] == 'N']
        if 'ETF' in data:
            data = data[data['ETF'] == 'N']
        data = data[~data['Security Name'].str.contains(r'\b(?:warrants?|preferred|units|rights)\b', case=False, regex=True)]
        frames.append(pd.DataFrame({'Ticker': data[symbol_column], 'Company': data['Security Name'].str.split(' - ', n=1).str[0],
                                    'Market': 'US', 'Exchange': exchange if exchange else data.get('Exchange', '').map({'N':'NYSE','A':'NYSE American','P':'NYSE Arca','Z':'Cboe','V':'IEX'}).fillna(data.get('Exchange','')), 'Aliases': '',
                                    'ReferenceSource': url, 'VerifiedAt': datetime.now(timezone.utc).isoformat(), 'ListingStatus':'交易所目前名單'}))
    result = pd.concat(frames, ignore_index=True).drop_duplicates(['Market', 'Ticker'])
    if len(result) < 1000:
        raise ValueError('官方代號目錄筆數異常，未覆寫舊目錄')
    target = repo / 'config' / 'us_symbol_reference.csv'
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        previous = pd.read_csv(target, dtype=str, keep_default_na=False)
        missing = previous[~previous['Ticker'].isin(result['Ticker'])].copy()
        missing['ListingStatus'] = '不在最新交易所名單；歷史代號待核對'
        result = pd.concat([result, missing], ignore_index=True)
    result.to_csv(target, index=False, encoding='utf-8-sig')
    return len(result)


if __name__ == '__main__':
    try:
        print(f'Official US symbol reference: {refresh()} issuers', flush=True)
    except Exception as exc:
        print(f'WARNING: symbol refresh failed ({type(exc).__name__}); use validated repo company lists only.', flush=True)
