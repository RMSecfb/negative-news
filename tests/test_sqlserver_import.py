import sys
import unittest
from datetime import date, datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools' / 'sqlserver'))
import import_news as m


class FakeCursor:
    def __init__(self, columns=(), state=()):
        self.log, self.columns, self.state, self.rows = [], list(columns), list(state), []
        self.last = ''

    def execute(self, sql, *params):
        self.log.append((sql, params))
        self.last = sql

    def executemany(self, sql, rows):
        self.log.append((sql, rows))
        self.rows.extend(rows)

    def fetchall(self):
        if 'sys.columns' in self.last:
            return [(c,) for c in self.columns]
        if 'FROM ' + m.STATE_TABLE in self.last:
            return list(self.state)
        return []

    def fetchone(self):
        return None


class FakeConnection:
    def __init__(self, cursor):
        self._cursor, self.commits, self.rollbacks = cursor, 0, 0

    def cursor(self):
        return self._cursor

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rollbacks += 1


def frame(**overrides):
    base = {'news_id': ['a', 'b'], 'Ticker': ['AAPL', ''], '日期': ['2026-10-26', '2026-10-26'],
            'Published Time': ['2026-10-26T08:30:00+08:00', '2026-10-25T20:00:00+00:00'],
            'Level': ['3', ''], 'FinBERT': ['-0.5', ''], 'Title': ['x', 'y']}
    base.update(overrides)
    return pd.DataFrame(base)


class HelperTests(unittest.TestCase):
    def test_git_blob_sha_matches_git(self):
        self.assertEqual(m.git_blob_sha(b''), 'e69de29bb2d1d6434b8b29ae775ad8c2e48c5391')
        self.assertEqual(m.git_blob_sha(b'hello\n'), 'ce013625030ba8dba906f756967f9e9ca394464a')

    def test_only_daily_news_and_summaries_are_wanted(self):
        self.assertTrue(m.wanted('news_daily/2026-10-26.csv'))
        self.assertFalse(m.wanted('news_daily/2026-10-26.xlsx'))
        self.assertFalse(m.wanted('news_daily/2026-10-26.csv', since='2026-10-27'))
        self.assertTrue(m.wanted('institutional_insider/13f/summary.csv'))
        self.assertFalse(m.wanted('news_runs/index.json'))

    def test_convert_types_and_empty_to_null(self):
        self.assertIsNone(m.convert('Level', ''))
        self.assertEqual(m.convert('Level', '3'), 3)
        self.assertEqual(m.convert('FinBERT', '-0.25'), -0.25)
        self.assertEqual(m.convert('日期', '2026-10-26'), date(2026, 10, 26))
        self.assertEqual(m.convert('Published Time', '2026-10-25T20:00:00+00:00'), datetime(2026, 10, 26, 4, 0))
        self.assertIsNone(m.convert('Level', 'abc'))
        self.assertEqual(m.convert('Title', 'text'), 'text')
        self.assertIsNone(m.convert('Title', ''))

    def test_identifiers_are_bracket_quoted(self):
        self.assertEqual(m.q('13F 機構動向'), '[13F 機構動向]')
        self.assertEqual(m.q('a]b'), '[a]]b]')


class ImportTests(unittest.TestCase):
    def test_new_table_replaces_the_file_before_inserting(self):
        cursor = FakeCursor()
        count = m.import_news_file(cursor, 'news_daily/2026-10-26.csv', frame())
        sql = [s for s, _ in cursor.log]
        create = next(s for s in sql if s.startswith('CREATE TABLE'))
        self.assertIn('[SourceFile] NVARCHAR(100) NOT NULL', create)
        self.assertIn('[news_id] NVARCHAR(64) NOT NULL', create)
        self.assertIn('[Level] INT', create)
        self.assertIn('[Title] NVARCHAR(MAX)', create)
        self.assertTrue(any('PRIMARY KEY' in s for s in sql))
        delete_at = next(i for i, s in enumerate(sql) if s.startswith('DELETE FROM'))
        insert_at = next(i for i, s in enumerate(sql) if s.startswith('INSERT INTO'))
        self.assertLess(delete_at, insert_at)
        self.assertEqual(count, 2)
        self.assertEqual(cursor.rows[0][0], 'news_daily/2026-10-26.csv')
        self.assertIsNone(cursor.rows[1][2])  # empty Ticker -> NULL

    def test_new_csv_column_is_added_to_existing_table(self):
        cursor = FakeCursor(columns=['SourceFile', 'news_id', 'Ticker', '日期', 'Published Time', 'Level', 'FinBERT'])
        m.import_news_file(cursor, 'news_daily/2026-10-26.csv', frame())
        alters = [s for s, _ in cursor.log if s.startswith('ALTER TABLE') and 'ADD [' in s]
        self.assertEqual(alters, ['ALTER TABLE dbo.NegativeNews ADD [Title] NVARCHAR(MAX) NULL'])

    def test_duplicate_or_missing_news_id_is_rejected_without_touching_the_table(self):
        for bad in (['a', 'a'], ['a', '']):
            cursor = FakeCursor()
            with self.assertRaises(ValueError):
                m.import_news_file(cursor, 'news_daily/2026-10-26.csv', frame(news_id=bad))
            self.assertEqual(cursor.log, [])

    def test_snapshot_is_replaced_whole(self):
        cursor = FakeCursor()
        m.import_snapshot(cursor, 'dbo.InsiderForm4Summary', pd.DataFrame({'Ticker': ['META'], 'Buy Count': ['1']}))
        sql = [s for s, _ in cursor.log]
        self.assertIn('DELETE FROM dbo.InsiderForm4Summary', sql)
        self.assertEqual(len(cursor.rows), 1)


class Resp:
    def __init__(self, payload=None, content=b''):
        self.payload, self.content = payload, content

    def raise_for_status(self):
        pass

    def json(self):
        return self.payload


class FakeSession:
    def __init__(self, files):
        self.files = files  # path -> bytes

    def get(self, url, **kwargs):
        if url.endswith('/commits/main'):
            return Resp({'sha': 'c' * 40})
        if '/git/trees/' in url:
            return Resp({'truncated': False, 'tree': [{'path': p, 'type': 'blob', 'sha': m.git_blob_sha(b)} for p, b in self.files.items()]})
        path = url.split('/' + 'c' * 40 + '/')[1]
        return Resp(content=self.files[path])


class RunTests(unittest.TestCase):
    def test_unchanged_files_are_skipped_and_bad_files_do_not_block_others(self):
        good = 'news_id,日期,Level\nn1,2026-10-26,3\n'.encode('utf-8-sig')
        bad = 'news_id,日期,Level\nn1,2026-10-27,3\nn1,2026-10-27,4\n'.encode('utf-8-sig')
        done = 'news_id,日期,Level\nn9,2026-10-25,1\n'.encode('utf-8-sig')
        files = {'news_daily/2026-10-26.csv': good, 'news_daily/2026-10-27.csv': bad,
                 'news_daily/2026-10-25.csv': done, 'README.md': b'x'}
        state = [('news_daily/2026-10-25.csv', m.git_blob_sha(done))]
        connection = FakeConnection(FakeCursor(state=state))
        logs = []
        result = m.run(connection, 'o/r', 'main', FakeSession(files), log=logs.append)
        self.assertEqual(result['imported'], ['news_daily/2026-10-26.csv'])
        self.assertEqual(list(result['failed']), ['news_daily/2026-10-27.csv'])
        self.assertEqual(connection.rollbacks, 1)
        self.assertNotIn('news_daily/2026-10-25.csv', result['imported'])

    def test_tampered_download_is_rejected(self):
        session = FakeSession({'news_daily/2026-10-26.csv': b'a'})
        orig = session.get

        def tampered(url, **kw):
            r = orig(url, **kw)
            if 'raw' in url:
                r.content = b'tampered'
            return r
        session.get = tampered
        with self.assertRaises(RuntimeError):
            m.download('o/r', 'c' * 40, 'news_daily/2026-10-26.csv', m.git_blob_sha(b'a'), session)


if __name__ == '__main__':
    unittest.main()
