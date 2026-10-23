import os
import tempfile
import unittest
from datetime import date
from unittest.mock import patch
from pathlib import Path
import pandas as pd
from streamlit.testing.v1 import AppTest
from pipeline.records import normalize


class WorkbenchTests(unittest.TestCase):
    def test_query_period_multiple_lists_and_editable_results(self):
        with tempfile.TemporaryDirectory() as temp, patch.dict(os.environ, {'NEWS_LOCAL_REPO':temp}):
            for day in ['2026-10-22','2026-10-23']:
                df=normalize(pd.DataFrame([{'Published Time':day+'T10:00:00+08:00','Title':'Test story','Ticker':'META','Company':'Meta Platforms','Market':'US','URL':'https://example.com/'+day,'Level':'3','Event_Code':'CDS_SPREAD_WIDENING','事件中文':'CDS利差擴大','Action':'觀察','分類狀態':'已分類','LLM狀態':'OK'}]))
                path=Path(temp)/'news_daily'/f'{day}.csv'
                path.parent.mkdir(exist_ok=True)
                df.to_csv(path,index=False)
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'負面新聞整合網站.py'),default_timeout=30).run()
            self.assertEqual(len(app.exception),0)
            app.date_input(key='query_start').set_value(date(2026,10,22))
            app.date_input(key='query_end').set_value(date(2026,10,23))
            app.button(key='query_news').click().run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len(app.session_state['news_query_frame']),2)
            app.radio(key='run_scope').set_value('指定公司列表').run()
            app.multiselect(key='run_lists_US').set_value(['S&P 500','Dow Jones 30']).run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len(app.multiselect(key='run_lists_US').value),2)
            app.toggle(key='unified_edit_enabled').set_value(True).run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(len([b for b in app.button if b.key=='save_unified_review']),1)


if __name__=='__main__':
    unittest.main()
