import io
import shutil
import tempfile
import unittest
from datetime import date
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from pipeline.integrated import ROOT
from pipeline.records import normalize, read_days, manual_save
from pipeline.repository import LocalRepository
from pipeline import review_import as ri

DAY = '2026-10-23'


def story(n, ticker='META', company='Meta Platforms', code='MAJOR_LITIGATION', level='4'):
    return {'Ticker': ticker, 'Company': company, 'Market': 'US', 'Published Time': f'{DAY}T10:0{n}:00+08:00',
            'Title': f'Story {n}', 'URL': f'https://example.com/story{n}', 'Source': 'Test', '來源方法': '方法一',
            'Event_Code': code, '事件中文': '重大訴訟', 'Level': level, 'Action': '通知', '分類狀態': '待人工覆核', '需要人工覆核': 'True'}


def workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = '覆核'
    header = ['序號', '日期', '公司', 'news_id', '目前事件代碼', '目前Level', '最終決定', '新事件代碼', '更正公司', '更正Ticker', '覆核註記', '覆核人', '人工版本']
    ws.append(header)
    for i, r in enumerate(rows, 1):
        ws.append([i] + [r.get(h, '') for h in header[1:]])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


class ReviewImportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.repo = LocalRepository(self.tmp.name)
        shutil.copy(ROOT / 'Parameter_Event.xlsx', Path(self.tmp.name) / 'Parameter_Event.xlsx')
        frame = normalize(pd.DataFrame([story(1), story(2), story(3), story(4, ticker='', company='Apple Inc.')]))
        target = Path(self.tmp.name) / 'news_daily' / f'{DAY}.csv'
        target.parent.mkdir()
        frame.to_csv(target, index=False)
        self.ids = {t: i for t, i in zip(['Story 1', 'Story 2', 'Story 3', 'Story 4'], frame['news_id'])}

    def tearDown(self):
        self.tmp.cleanup()

    def row(self, n, **kw):
        base = {'日期': DAY, 'news_id': self.ids[f'Story {n}'], '目前事件代碼': 'MAJOR_LITIGATION', '目前Level': 4, '人工版本': ''}
        base.update(kw)
        return base

    def day(self):
        return read_days(self.repo, date(2026, 10, 23), date(2026, 10, 23)).set_index('Title')

    def run_import(self, rows):
        return ri.apply_review(self.repo, workbook(rows), 'batch_已確認.xlsx')

    def test_modify_exclude_keep_and_markers(self):
        out = self.run_import([
            self.row(1, 最終決定='修改', 新事件代碼='EARNINGS_WARNING', 覆核註記='財測下修', 覆核人='tester'),
            self.row(2, 最終決定='排除', 覆核人='tester'),
            self.row(3, 最終決定='維持', 覆核人='tester'),
        ])
        self.assertEqual((out['applied'], out['skipped']), (3, 0))
        d = self.day()
        a, b, c = d.loc['Story 1'], d.loc['Story 2'], d.loc['Story 3']
        self.assertEqual((a['Event_Code'], a['Level'], a['事件中文']), ('EARNINGS_WARNING', '3', '財測下修'))
        self.assertEqual(a['覆核註記'], '財測下修')
        self.assertEqual((b['Event_Code'], b['Level'], b['Action']), ('NOT_RELEVANT', '0', '排除、不納入風險事件'))
        self.assertEqual((c['Event_Code'], c['Level']), ('MAJOR_LITIGATION', '4'))
        self.assertEqual(c['人工修改欄位'], '')
        for r in (a, b, c):
            self.assertEqual((r['已人工覆核'], r['分類狀態'], r['需要人工覆核'], r['最終判斷來源'], r['人工覆核人']),
                             ('True', '人工覆核', 'False', '人工覆核', 'tester'))
            self.assertTrue(r['人工版本'] and r['人工覆核時間'])
        self.assertIn('Event_Code', a['人工修改欄位'])
        self.assertIsNotNone(self.repo.read('review_inbox/done/batch_已確認.xlsx'))
        self.assertIsNotNone(self.repo.read('review_inbox/done/batch_已確認_結果.csv'))
        self.assertIsNotNone(self.repo.read(f'news_daily/{DAY}.xlsx'))

    def test_company_and_ticker_correction(self):
        self.run_import([self.row(4, 最終決定='維持', 更正公司='Apple Inc.', 更正Ticker='aapl')])
        r = self.day().loc['Story 4']
        self.assertEqual((r['Ticker'], r['Company']), ('AAPL', 'Apple Inc.'))
        self.assertIn('Ticker', r['人工修改欄位'])

    def test_conflicts_and_invalid_rows_are_skipped_not_fatal(self):
        original = read_days(self.repo, date(2026, 10, 23), date(2026, 10, 23))
        edited = original.copy()
        edited.loc[edited['Title'] == 'Story 1', 'Level'] = '3'
        manual_save(self.repo, original, edited, 'someone')
        out = self.run_import([
            self.row(1, 最終決定='排除'),                                    # 已被別人覆核（人工版本不一致）
            self.row(2, 目前事件代碼='ANTITRUST_INVESTIGATION', 最終決定='排除'),  # 分類已改變
            self.row(3, 最終決定='修改', 新事件代碼='NO_SUCH_CODE'),              # 無效代碼
            self.row(3, 最終決定='修改'),                                         # 缺新事件代碼
            self.row(3, 最終決定='其他'),                                         # 無效決定
            {**self.row(3, 最終決定='維持'), 'news_id': 'missing'},                # 找不到
        ])
        self.assertEqual((out['applied'], out['skipped']), (0, 6))
        d = self.day()
        self.assertEqual(d.loc['Story 1']['Level'], '3')
        self.assertEqual(d.loc['Story 2']['已人工覆核'], 'False')

    def test_reimport_same_file_does_not_overwrite(self):
        rows = [self.row(1, 最終決定='排除')]
        self.assertEqual(self.run_import(rows)['applied'], 1)
        self.assertEqual(self.run_import(rows)['applied'], 0)

    def test_manual_save_allows_level_zero(self):
        original = read_days(self.repo, date(2026, 10, 23), date(2026, 10, 23))
        edited = original.copy()
        edited.loc[0, 'Level'] = '0'
        self.assertEqual(manual_save(self.repo, original, edited, 'x'), 1)

    def test_missing_columns_fail_before_writing(self):
        wb = Workbook()
        wb.active.title = '覆核'
        wb.active.append(['news_id'])
        out = io.BytesIO()
        wb.save(out)
        with self.assertRaisesRegex(ValueError, '缺少欄位'):
            ri.apply_review(self.repo, out.getvalue(), 'bad_已確認.xlsx')
        self.assertIsNone(self.repo.read('review_inbox/done/bad_已確認.xlsx'))


if __name__ == '__main__':
    unittest.main()
