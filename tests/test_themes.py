import unittest
import pandas as pd
from pipeline.themes import apply_themes, load_themes, match

RULES = [{'Event_Code': 'CDS_SPREAD_WIDENING', '事件中文': 'CDS利差擴大', 'Level': 3, 'Action': '觀察'},
         {'Event_Code': 'SECTOR_MOVE', '事件中文': '產業波動', 'Level': 1, 'Action': '參考'}]


class ThemeTests(unittest.TestCase):
    def test_matches_ai_debt_headlines_only(self):
        themes = load_themes()
        self.assertEqual(match({'Title': '標普示警美國資料中心信用降評風險'}, themes), ('AI 債務／資料中心信用', 'CDS_SPREAD_WIDENING'))
        self.assertEqual(match({'Title': 'Hyperscalers likely carry more AI debt than the market realizes'}, themes)[1], 'SECTOR_MOVE')
        self.assertEqual(match({'Title': 'American Express adds $250 Airbnb credit, but there is a catch'}, themes), ('', ''))
        self.assertEqual(match({'Title': 'Top CD rates from major banks'}, themes), ('', ''))

    def test_only_company_less_unreviewed_rows_become_theme_events(self):
        base = {'Title': 'AI Infrastructure Debt Meets Rising Yields', 'Title_ZH': '', 'Company': '', 'Event_Code': 'NOT_RELEVANT', 'Level': '0', '已人工覆核': 'False'}
        frame = pd.DataFrame([{**base, 'Ticker': ''}, {**base, 'Ticker': 'MSFT'}, {**base, 'Ticker': '', '已人工覆核': 'True'},
                              {**base, 'Ticker': '', 'Title': 'Walmart opens new store'}])
        out = apply_themes(frame, RULES)
        self.assertEqual(list(out['風險主題']), ['AI 債務／資料中心信用'] * 3 + [''])
        self.assertEqual(list(out['Event_Code']), ['CDS_SPREAD_WIDENING', 'NOT_RELEVANT', 'NOT_RELEVANT', 'NOT_RELEVANT'])
        self.assertEqual((out.loc[0, 'Level'], out.loc[0, 'Company'], out.loc[0, '最終判斷來源']), ('3', '產業主題', '主題規則'))


if __name__ == '__main__':
    unittest.main()
