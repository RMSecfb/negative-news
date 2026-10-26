"""Incrementally load negative-news results from GitHub into SQL Server.

Runs on the company PC (pull model: nothing is pushed into the company network).
Only files whose Git blob SHA changed since the last run are re-imported, and a
changed daily file is replaced as a whole (delete + insert in one transaction), so
later LLM back-fills, manual edits and offline reviews of old dates are picked up.

    python import_news.py --conn "DRIVER={ODBC Driver 18 for SQL Server};SERVER=host;DATABASE=db;Trusted_Connection=yes;TrustServerCertificate=yes"
    python import_news.py --dry-run          # list what would be imported; no database needed
"""
from __future__ import annotations
import argparse
import hashlib
import io
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

DEFAULT_REPO = 'RMSecfb/negative-news'
TPE = timezone(timedelta(hours=8))
NEWS_TABLE = 'dbo.NegativeNews'
STATE_TABLE = 'dbo.NegativeNewsImportState'
SNAPSHOT_TABLES = {
    'institutional_insider/13f/summary.csv': 'dbo.Institutional13FSummary',
    'institutional_insider/form4/summary.csv': 'dbo.InsiderForm4Summary',
}
NEWS_PATH = re.compile(r'news_daily/(\d{4}-\d{2}-\d{2})\.csv')

# Typed columns; everything else is NVARCHAR(MAX) (NULL when empty).
NEWS_TYPES = {
    'news_id': 'NVARCHAR(64) NOT NULL', 'Ticker': 'NVARCHAR(64)', 'Event_Code': 'NVARCHAR(100)',
    '日期': 'DATE', 'Published Time': 'DATETIME2(0)', 'Level': 'INT', 'Rule_Level': 'INT',
    'FinBERT': 'FLOAT', '曝險金額': 'FLOAT', 'input_tokens': 'INT', 'output_tokens': 'INT',
    'Market': 'NVARCHAR(16)',
}
SNAPSHOT_TYPES = {'Ticker': 'NVARCHAR(64)'}
INT_COLS = {'Level', 'Rule_Level', 'input_tokens', 'output_tokens'}
FLOAT_COLS = {'FinBERT', '曝險金額'}


def q(name: str) -> str:
    return '[' + name.replace(']', ']]') + ']'


def git_blob_sha(data: bytes) -> str:
    return hashlib.sha1(b'blob %d\x00' % len(data) + data).hexdigest()


def wanted(path: str, since: str | None = None) -> bool:
    match = NEWS_PATH.fullmatch(path)
    if match:
        return not since or match.group(1) >= since
    return path in SNAPSHOT_TABLES


def list_remote(repo: str, branch: str, session, token: str = '', since: str | None = None, participants: bool = True):
    """Return (commit_sha, {path: blob_sha}) for one immutable snapshot of the branch."""
    headers = {'Authorization': 'Bearer ' + token} if token else {}
    api = f'https://api.github.com/repos/{repo}'
    commit = session.get(f'{api}/commits/{branch}', headers=headers, timeout=60)
    commit.raise_for_status()
    sha = commit.json()['sha']
    tree = session.get(f'{api}/git/trees/{sha}', params={'recursive': '1'}, headers=headers, timeout=60)
    tree.raise_for_status()
    body = tree.json()
    if body.get('truncated'):
        raise RuntimeError('GitHub 檔案清單被截斷，無法確認完整檔案')
    files = {item['path']: item['sha'] for item in body['tree']
             if item['type'] == 'blob' and wanted(item['path'], since)
             and (participants or not item['path'].startswith('institutional_insider/'))}
    return sha, files


def download(repo: str, commit: str, path: str, blob_sha: str, session) -> bytes:
    # A raw URL addressed by commit SHA is immutable, so it can never be a stale cached copy.
    url = f'https://raw.githubusercontent.com/{repo}/{commit}/{path}'
    response = session.get(url, timeout=120)
    response.raise_for_status()
    if git_blob_sha(response.content) != blob_sha:
        raise RuntimeError(f'{path} 內容與 GitHub 檔案指紋不符')
    return response.content


def read_csv(raw: bytes) -> pd.DataFrame:
    return pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding='utf-8-sig')


def convert(column: str, value: str):
    value = (value or '').strip() if isinstance(value, str) else value
    if value in ('', None):
        return None
    try:
        if column == '日期':
            return date.fromisoformat(value[:10])
        if column == 'Published Time':
            return datetime.fromisoformat(value).astimezone(TPE).replace(tzinfo=None)
        if column in INT_COLS:
            return int(float(value))
        if column in FLOAT_COLS:
            return float(value)
    except ValueError:
        return None
    return value


def column_type(table: str, column: str) -> str:
    types = NEWS_TYPES if table == NEWS_TABLE else SNAPSHOT_TYPES
    return types.get(column, 'NVARCHAR(MAX)')


def ensure_table(cursor, table: str, columns: list[str], extra_first: list[str] | None = None):
    """Create the table if needed and add any column that a newer CSV introduced."""
    schema, name = table.split('.')
    fixed = extra_first or []
    cursor.execute(f"SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(?)", table)
    existing = {row[0] for row in cursor.fetchall()}
    if not existing:
        definitions = [f'{q(c)} {t}' for c, t in fixed] + [f'{q(c)} {column_type(table, c)}' for c in columns]
        cursor.execute(f'CREATE TABLE {table} ({", ".join(definitions)})')
        existing = {c for c, _ in fixed} | set(columns)
    for column in columns:
        if column not in existing:
            nullable = column_type(table, column).replace(' NOT NULL', '')
            cursor.execute(f'ALTER TABLE {table} ADD {q(column)} {nullable} NULL')
            existing.add(column)


def ensure_news_schema(cursor, columns: list[str]):
    ensure_table(cursor, NEWS_TABLE, columns, extra_first=[('SourceFile', 'NVARCHAR(100) NOT NULL')])
    cursor.execute("SELECT 1 FROM sys.indexes WHERE name = 'PK_NegativeNews' AND object_id = OBJECT_ID(?)", NEWS_TABLE)
    if not cursor.fetchone():
        cursor.execute(f'ALTER TABLE {NEWS_TABLE} ADD CONSTRAINT PK_NegativeNews PRIMARY KEY ({q("SourceFile")}, {q("news_id")})')
        cursor.execute(f'CREATE INDEX IX_NegativeNews_Date ON {NEWS_TABLE} ({q("日期")})')
        cursor.execute(f'CREATE INDEX IX_NegativeNews_Ticker ON {NEWS_TABLE} ({q("Ticker")})')
        cursor.execute(f'CREATE INDEX IX_NegativeNews_NewsId ON {NEWS_TABLE} ({q("news_id")})')


def ensure_state(cursor):
    cursor.execute(
        f"IF OBJECT_ID('{STATE_TABLE}') IS NULL CREATE TABLE {STATE_TABLE} "
        "(FilePath NVARCHAR(200) NOT NULL PRIMARY KEY, BlobSha CHAR(40) NOT NULL, "
        "[RowCount] INT NOT NULL, ImportedAt DATETIME2(0) NOT NULL)")


def known_state(cursor) -> dict[str, str]:
    cursor.execute(f'SELECT FilePath, BlobSha FROM {STATE_TABLE}')
    return {row[0]: row[1].strip() for row in cursor.fetchall()}


def remember(cursor, path: str, blob_sha: str, rows: int):
    cursor.execute(f'DELETE FROM {STATE_TABLE} WHERE FilePath = ?', path)
    cursor.execute(f'INSERT INTO {STATE_TABLE} (FilePath, BlobSha, [RowCount], ImportedAt) VALUES (?, ?, ?, SYSDATETIME())',
                   path, blob_sha, rows)


def insert_rows(cursor, table: str, frame: pd.DataFrame, lead: list[tuple[str, object]] = ()):
    columns = [c for c, _ in lead] + list(frame.columns)
    sql = f'INSERT INTO {table} ({", ".join(q(c) for c in columns)}) VALUES ({", ".join("?" for _ in columns)})'
    leading = [v for _, v in lead]
    rows = [leading + [convert(c, v) for c, v in zip(frame.columns, record)]
            for record in frame.itertuples(index=False, name=None)]
    if rows:
        cursor.executemany(sql, rows)
    return len(rows)


def import_news_file(cursor, path: str, frame: pd.DataFrame) -> int:
    if 'news_id' not in frame or frame['news_id'].eq('').any() or frame['news_id'].duplicated().any():
        raise ValueError(f'{path} 的 news_id 缺少或重複，未匯入')
    ensure_news_schema(cursor, list(frame.columns))
    cursor.execute(f'DELETE FROM {NEWS_TABLE} WHERE {q("SourceFile")} = ?', path)
    return insert_rows(cursor, NEWS_TABLE, frame, lead=[('SourceFile', path)])


def import_snapshot(cursor, table: str, frame: pd.DataFrame) -> int:
    imported_at = datetime.now(TPE).replace(tzinfo=None, microsecond=0)
    ensure_table(cursor, table, list(frame.columns) + ['ImportedAt'])
    cursor.execute(f'DELETE FROM {table}')
    frame = frame.copy()
    frame['ImportedAt'] = imported_at.isoformat(sep=' ')
    return insert_rows(cursor, table, frame)


def run(connection, repo: str, branch: str, session, token='', since=None, participants=True, force=False, log=print) -> dict:
    cursor = connection.cursor()
    ensure_state(cursor)
    connection.commit()
    commit, remote = list_remote(repo, branch, session, token, since, participants)
    state = {} if force else known_state(cursor)
    todo = sorted(p for p, sha in remote.items() if state.get(p) != sha)
    log(f'來源版本 {commit[:10]}；符合條件的檔案 {len(remote)} 個，需匯入 {len(todo)} 個')
    result = {'imported': [], 'failed': {}, 'rows': 0}
    for path in todo:
        try:
            frame = read_csv(download(repo, commit, path, remote[path], session))
            rows = (import_snapshot(cursor, SNAPSHOT_TABLES[path], frame) if path in SNAPSHOT_TABLES
                    else import_news_file(cursor, path, frame))
            remember(cursor, path, remote[path], rows)
            connection.commit()
            result['imported'].append(path)
            result['rows'] += rows
            log(f'  已匯入 {path}（{rows} 列）')
        except Exception as exc:  # one bad file must not block the others
            connection.rollback()
            result['failed'][path] = str(exc)[:300]
            log(f'  失敗 {path}：{exc}')
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description='把 negative-news 的每日結果增量匯入 SQL Server')
    parser.add_argument('--conn', default=os.getenv('NEWS_DB_CONN', ''), help='ODBC 連線字串（或設定環境變數 NEWS_DB_CONN）')
    parser.add_argument('--repo', default=DEFAULT_REPO)
    parser.add_argument('--branch', default='main')
    parser.add_argument('--since', help='只匯入此日期（含）之後的每日檔，格式 YYYY-MM-DD')
    parser.add_argument('--no-participants', action='store_true', help='不匯入 13F／Form 4 摘要')
    parser.add_argument('--force', action='store_true', help='忽略已匯入紀錄，全部重新匯入')
    parser.add_argument('--dry-run', action='store_true', help='只列出會匯入的檔案，不連資料庫')
    args = parser.parse_args(argv)
    token = os.getenv('GITHUB_TOKEN', '')
    session = requests.Session()
    if args.dry_run:
        commit, remote = list_remote(args.repo, args.branch, session, token, args.since, not args.no_participants)
        print(f'來源版本 {commit[:10]}，共 {len(remote)} 個檔案：')
        for path in sorted(remote):
            print('  ', path)
        return 0
    if not args.conn:
        parser.error('需要 --conn 或環境變數 NEWS_DB_CONN')
    import pyodbc  # imported late so --dry-run and the tests need no ODBC driver
    connection = pyodbc.connect(args.conn, autocommit=False)
    try:
        result = run(connection, args.repo, args.branch, session, token, args.since, not args.no_participants, args.force)
    finally:
        connection.close()
    print(f'完成：匯入 {len(result["imported"])} 個檔案、{result["rows"]} 列；失敗 {len(result["failed"])} 個')
    return 1 if result['failed'] else 0


if __name__ == '__main__':
    sys.exit(main())
