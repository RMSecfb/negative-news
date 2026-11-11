"""Industry risk themes: news that matters as a sector signal even when no single company is the subject.

config/risk_themes.csv lists each theme. Every matching headline gets the theme name in 風險主題
(a tag only). A matching headline with no company Ticker, that no person has reviewed and that has
no LLM event, is additionally kept as a theme event instead of being dropped as irrelevant.
"""
from __future__ import annotations
import csv
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
COLUMN = '風險主題'
THEME_COMPANY = '產業主題'
SOURCE = '主題規則'
_NO_EVENT = ('', 'NOT_RELEVANT', 'UNKNOWN', 'NON_CORE_SOURCE', 'GENERAL_INFORMATION')


def load_themes(path=None):
    path = Path(path or ROOT / 'config' / 'risk_themes.csv')
    if not path.exists():
        return []
    with path.open(encoding='utf-8-sig', newline='') as handle:
        return [{'name': row['Theme'].strip(), 'pattern': re.compile(row['Pattern'], re.I),
                 'stress': re.compile(row['StressPattern'], re.I) if row.get('StressPattern') else None,
                 'stress_event': (row.get('StressEvent') or '').strip(), 'default_event': (row.get('DefaultEvent') or '').strip()}
                for row in csv.DictReader(handle) if (row.get('Theme') or '').strip()]


def match(record, themes):
    text = f"{record.get('Title', '') or ''} {record.get('Title_ZH', '') or ''}"
    for theme in themes:
        if theme['pattern'].search(text):
            code = theme['stress_event'] if theme['stress'] is not None and theme['stress'].search(text) else theme['default_event']
            return theme['name'], code
    return '', ''


def _truth(value):
    return str(value).strip().lower() in ('true', '1', 'yes')


def apply_themes(frame, rules, themes=None):
    """Tag 風險主題 on every match; turn company-less, unreviewed, event-less matches into theme events."""
    themes = load_themes() if themes is None else themes
    if frame.empty or not themes:
        return frame
    from .events import pick_rule
    frame = frame.copy()
    if COLUMN not in frame:
        frame[COLUMN] = ''
    for index, record in frame.iterrows():
        name, code = match(record, themes)
        if not name:
            continue
        frame.at[index, COLUMN] = name
        if _truth(record.get('已人工覆核')) or str(record.get('Ticker', '') or '').strip():
            continue
        if str(record.get('Event_Code', '') or '') not in _NO_EVENT:
            continue
        rule = pick_rule(code, rules) if code else None
        if not rule:
            continue
        for field in ('Event_Code', '事件中文', 'Level', 'Action'):
            frame.at[index, field] = str(rule.get(field, '') or '')
        frame.at[index, 'Level'] = re.sub(r'\.0$', '', str(frame.at[index, 'Level']))
        frame.at[index, 'Company'] = str(record.get('Company', '') or '').strip() or THEME_COMPANY
        frame.at[index, '最終判斷來源'] = SOURCE
        frame.at[index, '分類狀態'] = '已分類'
        frame.at[index, '需要人工覆核'] = 'False'
    return frame
