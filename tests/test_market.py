import sys
import types
import unittest
from datetime import datetime, timezone
from unittest.mock import patch
from pathlib import Path
import pandas as pd
try:
    import requests
except ModuleNotFoundError:
    sys.modules['requests']=types.ModuleType('requests')
from pipeline.market import recent_records, deduplicate_headlines, priority
from pipeline.ingest import fetch_news

class MarketTests(unittest.TestCase):
    def test_dates(self):
        df=pd.DataFrame({'Published Time':['Fri, 23 Oct 2026 12:00:00 GMT','Tue, 28 Jul 2026 07:00:00 GMT','invalid','Sat, 24 Oct 2026 12:00:00 GMT']})
        result, excluded=recent_records(df,datetime(2026,10,23,15,tzinfo=timezone.utc))
        self.assertEqual(len(result),1)
        self.assertEqual(excluded,3)
    def test_exact_headline_dedup(self):
        df=pd.DataFrame([dict(Ticker='CRWV',Title='Credit risk rises!',監控池='CUSTOM'),dict(Ticker='CRWV',Title='Credit risk rises',監控池='MARKET_DISCOVERY')])
        out=deduplicate_headlines(df)
        self.assertEqual(len(out),1)
        self.assertIn('MARKET_DISCOVERY',out.iloc[0]['監控池'])
    def test_major_events_rank_first(self):
        self.assertGreater(priority({'Title':'Company files bankruptcy','Source':'Reuters'}),priority({'Title':'Stock analyst changes price target','Keyword候選':'target'}))
    def test_market_mode_does_not_rotate_sp500(self):
        with patch('pipeline.ingest.company_file',return_value=[('AAPL','Apple')]), patch('pipeline.ingest.optional_watchlist',return_value=[]), patch('pipeline.ingest.rss_fetch',return_value=[]) as rss, patch('pipeline.ingest.requests.Session',create=True), patch('pipeline.ingest.time.sleep'):
            _, stats=fetch_news(Path('.'),datetime.now(timezone.utc),scope='market')
        self.assertEqual(stats['ticker_count'],0)
        self.assertEqual(stats['scan_mode'],'MARKET_EVENTS')
        self.assertGreaterEqual(rss.call_count,20)
        for call in rss.call_args_list:
            self.assertIn('when:2d',call.args[0])

if __name__=='__main__':
    unittest.main()
