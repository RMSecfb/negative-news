"""Run 13F_FINAL.py with workspace paths, an explicit target quarter and a completeness gate.

The original script takes its target quarter from the (manually maintained)
"Report Period" column of 13F_Manager_List.xlsx. For unattended runs the quarter
is instead the last completed calendar quarter. A new quarter's workbook is only
published once enough aggregate managers have filed, so a half-filed quarter
never replaces the last complete one.
"""
from __future__ import annotations
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
import ast
import hashlib
import importlib.metadata
import json
import os
import shutil
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '13F_FINAL.py'
OUTPUT = ROOT / 'institutional_13f'
STAGING = ROOT / '.staging_13f'
MIN_COMPLETE_RATIO = float(os.getenv('MIN_COMPLETE_RATIO', '0.8'))
SKIP_AFTER_DAYS = 60  # 13F is due 45 days after quarter end; leave slack for late filers
TPE = ZoneInfo('Asia/Taipei')


def last_completed_quarter_end(today: date) -> date:
    ends = [date(y, m, d) for y in (today.year - 1, today.year) for m, d in ((3, 31), (6, 30), (9, 30), (12, 31))]
    return max(e for e in ends if e <= today)


def quarter_label(period: date) -> str:
    return f'{period.year}Q{(period.month - 1) // 3 + 1}'


def should_skip(today: date, period: date, existing_ok: int | None, force: bool = False) -> bool:
    """Skip once the window has closed and a published workbook already exists."""
    return (not force) and existing_ok is not None and (today - period).days > SKIP_AFTER_DAYS


def decide_publish(ratio: float, ok_count: int, existing_ok: int | None, min_ratio: float = MIN_COMPLETE_RATIO) -> tuple[bool, str]:
    if ratio < min_ratio:
        return False, f'納入彙總的機構僅 {ratio:.0%}，低於 {min_ratio:.0%}，不發布本季成果'
    if existing_ok is not None and ok_count < existing_ok:
        return False, f'本次成功機構 {ok_count} 少於已發布的 {existing_ok}，保留已發布成果'
    return True, '發布'


class RunSettings(ast.NodeTransformer):
    def __init__(self, period: str):
        self.period = period

    def visit_Assign(self, node):
        self.generic_visit(node)
        names = {t.id for t in node.targets if isinstance(t, ast.Name)}
        if 'BASE_DIR' in names:
            node.value = ast.parse('Path(' + repr(str(ROOT)) + ')', mode='eval').body
        elif 'OUTPUT_DIR' in names:
            node.value = ast.parse('Path(' + repr(str(STAGING)) + ')', mode='eval').body
        elif 'TARGET_REPORT_PERIOD' in names:
            node.value = ast.Constant(self.period)
        return node


def existing_ok_count(quarter: str) -> int | None:
    meta = OUTPUT / quarter / '13F_Run_Metadata.json'
    workbook = OUTPUT / quarter / f'13F_{quarter}_Full_Universe.xlsx'
    if not workbook.exists():
        return None
    try:
        return int(json.loads(meta.read_text(encoding='utf-8')).get('ok_aggregate_managers', 0))
    except Exception:
        return 0


def main():
    os.environ.setdefault('EDGAR_LOCAL_DATA_DIR', str(ROOT.parent / 'sec-cache'))
    os.environ.setdefault('EDGAR_RATE_LIMIT_PER_SEC', '4')
    started = datetime.now(TPE)
    today = started.date()
    forced = os.getenv('FORCE', '').lower() in ('1', 'true', 'yes')
    override = os.getenv('TARGET_REPORT_PERIOD', '').strip()
    period = date.fromisoformat(override) if override else last_completed_quarter_end(today)
    quarter = quarter_label(period)
    existing = existing_ok_count(quarter)
    if should_skip(today, period, existing, forced):
        print(f'{quarter} 已發布且超過申報期間 {SKIP_AFTER_DAYS} 天，略過（設定 FORCE=1 可強制重跑）')
        return 0
    print(f'目標季度 {quarter}（{period}），已發布成功機構數：{existing}')
    shutil.rmtree(STAGING, ignore_errors=True)
    tree = RunSettings(period.isoformat()).visit(ast.parse(SOURCE.read_text(encoding='utf-8-sig')))
    ast.fix_missing_locations(tree)
    namespace = {'__name__': '__main__', '__file__': str(SOURCE)}
    exec(compile(tree, str(SOURCE), 'exec'), namespace)
    logs: pd.DataFrame = namespace['run_log_df']
    eligible = namespace['eligible']
    aggregate_total = int(eligible['Include in Aggregate'].eq('Y').sum())
    ok_count = int((logs['Status'].eq('OK') & logs['Aggregate'].eq('Y')).sum())
    ratio = ok_count / aggregate_total if aggregate_total else 0.0
    publish, reason = decide_publish(ratio, ok_count, existing)
    produced = Path(namespace['FINAL_FILE'])
    manifest = {
        'source_repo': os.getenv('GITHUB_REPOSITORY', 'RMSecfb/negative-news'),
        'execution_commit': os.getenv('GITHUB_SHA', 'local'),
        'source_sha256': hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        'edgartools': importlib.metadata.version('edgartools'),
        'started_at': started.isoformat(),
        'finished_at': datetime.now(TPE).isoformat(),
        'quarter': quarter, 'target_period': period.isoformat(),
        'managers_eligible': int(len(eligible)), 'aggregate_managers': aggregate_total,
        'ok_aggregate_managers': ok_count, 'complete_ratio': round(ratio, 4),
        'status_counts': {str(k): int(v) for k, v in logs['Status'].value_counts().items()},
        'published': publish, 'reason': reason,
        'output_file': f'institutional_13f/{quarter}/{produced.name}',
    }
    if publish:
        target = OUTPUT / quarter
        target.mkdir(parents=True, exist_ok=True)
        shutil.copy2(produced, target / produced.name)
        (target / '13F_Run_Metadata.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    (ROOT / 'run_manifest_13f.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    shutil.rmtree(STAGING, ignore_errors=True)
    print(json.dumps(manifest, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
