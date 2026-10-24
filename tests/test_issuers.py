import unittest
import tempfile,io
from datetime import date
from pathlib import Path
import pandas as pd
from pipeline.issuers import IssuerIndex, load_index

ROOT=Path(__file__).resolve().parents[1]

class IssuerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index=load_index(ROOT)

    def row(self,name,title=None,ticker=''):
        return {'Market':'US','Ticker':ticker,'Company':'','LLM辨識公司':name,'Title':title or name+' faces a warning','Summary':'','LLM狀態':'OK','分類狀態':'已分類'}

    def test_reference_symbols_round_trip(self):
        # The complete directory, not only the three user examples.
        for (market,ticker),expected in self.index.symbols.items():
            if market=='US':
                found,_=self.index.resolve(ticker,market)
                if found is None:
                    from pipeline.issuers import name_key
                    self.assertGreater(len(self.index.exact.get((market,name_key(ticker)),{})),1,ticker)
                else:
                    self.assertEqual(found['Ticker'],ticker,ticker)

    def test_company_brand_cannot_be_confused_with_another_ticker(self):
        index=IssuerIndex(pd.DataFrame([{'Market':'US','Ticker':'HP','Company':'Helmerich & Payne Inc.'},{'Market':'US','Ticker':'HPQ','Company':'HP Inc.'}]))
        self.assertIsNone(index.resolve('HP')[0])
        unknown=index.enrich(pd.DataFrame([self.row('HP','HP reports earnings')])).iloc[0]
        self.assertEqual(unknown.Ticker,'')
        hinted=index.enrich(pd.DataFrame([self.row('HP','HP reports earnings','HPQ')])).iloc[0]
        self.assertEqual(hinted.Ticker,'HPQ')

    def test_general_security_descriptions_and_multiple_issuers(self):
        companies=pd.DataFrame([{'Market':'US','Ticker':'ALFA','Company':'Alpha Systems Inc. Class A Common Stock'}, {'Market':'US','Ticker':'BETA','Company':'Beta Digital Ltd. American Depositary Shares, each representing two shares'}, {'Market':'US','Ticker':'GAM','Company':'Gamma Group Corporation Common Stock'}])
        index=IssuerIndex(companies)
        for name,ticker in [('Alpha Systems','ALFA'),('Beta Digital Ltd. (BETA)','BETA'),('Gamma Group','GAM')]:
            self.assertEqual(index.resolve(name)[0]['Ticker'],ticker)
        self.assertIsNone(index.resolve('Alpha')[0])

    def test_verified_names_from_all_lists_remain_available(self):
        index=IssuerIndex(pd.DataFrame([{'Market':'US','Ticker':'ALFA','Company':'Alpha'}, {'Market':'US','Ticker':'ALFA','Company':'Alpha Technologies Inc. Common Stock'}]))
        row=self.row('Alpha Technologies Inc.','Alpha reports earnings','ALFA')
        out=index.enrich(pd.DataFrame([row])).iloc[0]
        self.assertEqual(out.Ticker,'ALFA')

    def test_all_observed_listed_subjects(self):
        for name,ticker in [('IBM','IBM'),('Getty Images','GETY'),('Eon Resources','EONR'),("Papa John's International, Inc.",'PZZA'),('Abbott','ABT'),('Sands','LVS'),('Meta','META'),('CoreWeave','CRWV'),('Karooooo','KARO'),('FuelCell Energy, Inc.','FCEL'),('PulteGroup','PHM'),('Netflix','NFLX'),('Nvidia','NVDA'),('Yiren Digital Ltd. (YRD)','YRD')]:
            out=self.index.enrich(pd.DataFrame([self.row(name)])).iloc[0]
            self.assertEqual(out['Ticker'],ticker,name)

    def test_partial_location_is_not_a_company(self):
        out=self.index.enrich(pd.DataFrame([self.row('Chesapeake','Data breach at Chesapeake business','CPK')])).iloc[0]
        self.assertEqual(out.Ticker,'')
        self.assertEqual(out['原始Ticker'],'CPK')

    def test_company_must_be_supported_by_news_text(self):
        out=self.index.enrich(pd.DataFrame([self.row('Meta Platforms','Greek parliament passes social media ban','META')])).iloc[0]
        self.assertEqual(out.Ticker,'')
        self.assertEqual(out['Ticker狀態'],'主體證據不足')

    def test_legacy_non_llm_rows_are_also_audited(self):
        row={**self.row('','Cost to battle wildfires rises','COST'),'Company':'Costco'}
        row.pop('LLM狀態')
        out=self.index.enrich(pd.DataFrame([row])).iloc[0]
        self.assertEqual(out.Ticker,'')
        self.assertEqual(out['原始Ticker'],'COST')

    def test_foreign_and_nonprofit_are_explained(self):
        for subject,status in [('Capcom','非本市場上市主體'),('Northern Light Health','非上市機構')]:
            out=self.index.enrich(pd.DataFrame([self.row(subject)])).iloc[0]
            self.assertEqual(out.Ticker,'')
            self.assertEqual(out['Ticker狀態'],status)

    def test_otc_reference_keeps_symbol_and_listing_status(self):
        out=self.index.enrich(pd.DataFrame([self.row('Getty Images')])).iloc[0]
        self.assertEqual(out.Ticker,'GETY')
        self.assertIn('OTC',out['上市狀態'])
        self.assertIn('sec.gov',out['Ticker核對依據'])

    def test_manual_ticker_is_not_changed(self):
        row={**self.row('IBM','Unclear headline','ZZZ'),'已人工覆核':'True'}
        out=self.index.enrich(pd.DataFrame([row])).iloc[0]
        self.assertEqual(out.Ticker,'ZZZ')
        self.assertEqual(out['Ticker狀態'],'人工確認')

    def test_deliberately_cleared_manual_ticker_is_not_refilled(self):
        row={**self.row('IBM'),'已人工覆核':'True','人工修改欄位':'Ticker|Level'}
        out=self.index.enrich(pd.DataFrame([row])).iloc[0]
        self.assertEqual(out.Ticker,'')
        self.assertEqual(out['Ticker狀態'],'人工確認')

    def test_conflicting_issuer_and_ticker_are_not_silently_kept(self):
        out=self.index.enrich(pd.DataFrame([self.row('IBM','IBM profit warning','META')])).iloc[0]
        self.assertEqual(out.Ticker,'')
        self.assertEqual(out['Ticker狀態'],'代號與主體衝突')

    def test_audit_includes_historical_files_and_preserves_manual_decisions(self):
        from pipeline.repository import LocalRepository
        from pipeline.ticker_audit import audit
        from pipeline.records import read_days
        with tempfile.TemporaryDirectory() as temp:
            repo=LocalRepository(temp)
            old=pd.DataFrame([{**self.row('IBM'),'Published Time':'2026-08-10T10:00:00+08:00','URL':'https://example.com/old','Level':'4','Event_Code':'EARNINGS_MISS'}])
            output=io.BytesIO()
            old.to_excel(output,index=False)
            human=pd.DataFrame([{**self.row('IBM',ticker='ZZZ'),'Published Time':'2026-10-23T10:00:00+08:00','URL':'https://example.com/manual','Level':'5','Event_Code':'MANUAL','已人工覆核':'True','人工版本':'human-v1'}])
            repo.transaction(lambda read,rev:({'risk_news_daily/2026-08-10.xlsx':output.getvalue(),'news_daily/2026-10-23.csv':human.to_csv(index=False)},None),'fixture')
            report=audit(repo)
            self.assertEqual(report['sources'],2)
            self.assertEqual(report['rows'],2)
            self.assertEqual(report['llm_calls'],0)
            found=read_days(repo,date(2026,8,10),date(2026,10,24))
            self.assertEqual(found[found.URL=='https://example.com/old'].iloc[0].Ticker,'IBM')
            manual=found[found.URL=='https://example.com/manual'].iloc[0]
            self.assertEqual((manual.Ticker,manual.Level,manual.Event_Code,manual['人工版本']),('ZZZ','5','MANUAL','human-v1'))

if __name__=='__main__':unittest.main()
