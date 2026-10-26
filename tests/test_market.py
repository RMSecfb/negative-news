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
from pipeline.run import classify

class MarketTests(unittest.TestCase):
    def test_unconfirmed_default_is_not_classified_as_default(self):
        frame=pd.DataFrame([{'Ticker':'CRWV','Title':'Possible default','監控池':'CUSTOM','LLM狀態':'OK','Event_Code':'DEFAULT','Level':'5','Confirmed_Default':'False','事件狀態':'RUMORED'}])
        out,_=classify(frame,[], '',0,'2026-10-23')
        self.assertEqual(out.iloc[0]['Event_Code'],'')
        self.assertEqual(out.iloc[0]['分類狀態'],'待人工覆核')
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
    def test_routine_factset_survey_is_not_an_investigation(self):
        routine=priority({'Title':'鉅亨速報 - Factset 最新調查：洛克希德馬丁LMT-US的目標價調升至629.5元','Source':'鉅亨網'})
        probe=priority({'Title':'美國運通遭聯邦調查 涉反洗錢缺失','Source':'鉅亨網'})
        police=priority({'Title':'消息人士稱，FTC已展開調查','Source':'Test'})
        self.assertEqual(routine,0)
        self.assertGreaterEqual(probe,70)
        self.assertGreaterEqual(police,70)
        self.assertGreaterEqual(priority({'Title':'Factset 最新調查：某公司聲請破產'}),120)
    def test_llm_order_uses_finbert_within_same_priority(self):
        from pipeline.integrated import llm_order
        rows=[
            {'Title':'Quarterly update','FinBERT':'0.4'},
            {'Title':'Company files bankruptcy','FinBERT':'0.1'},
            {'Title':'Shares slip on weak outlook','FinBERT':'-0.8'},
            {'Title':'Shares slide after report','FinBERT':'-0.3'},
            {'Title':'Chinese headline','FinBERT':'N/A'},
        ]
        ordered=[r['Title'] for r in sorted(rows,key=llm_order)]
        self.assertEqual(ordered,['Company files bankruptcy','Shares slip on weak outlook','Shares slide after report','Quarterly update','Chinese headline'])
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
