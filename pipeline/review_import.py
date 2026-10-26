"""Apply an offline review workbook (review_inbox/*_已確認.xlsx) to news_daily.

Same write path and markers as records.manual_save, with three additions:
  * 維持 / 修改 / 排除 decisions (排除 = NOT_RELEVANT, Level 0),
  * optional company / ticker correction,
  * rows that changed since the review file was produced are skipped and reported.
No LLM is called.
"""
from __future__ import annotations
import argparse
import io
import os
import sys
import uuid
from datetime import date, datetime
from pathlib import Path
import pandas as pd
from .records import TPE, csv_frame, excel_bytes, tags
from .repository import LocalRepository, read_json

INBOX = 'review_inbox/'
DONE = 'review_inbox/done/'
SHEET = '覆核'
REQUIRED = ['news_id', '日期', '人工版本', '最終決定', '新事件代碼']
OPTIONAL = ['目前事件代碼', '目前Level', '更正公司', '更正Ticker', '覆核註記', '覆核人', '標題（點擊開原文）', '公司', '序號']
DECISIONS = ('維持', '修改', '排除')
EXCLUDE_CODE = 'NOT_RELEVANT'
EXCLUDE_ACTION = '排除、不納入風險事件'
DEFAULT_REVIEWER = '離線覆核'
FINAL_FIELDS = ['Ticker', 'Company', 'Event_Code', '事件中文', 'Level', 'Action']


def cell(value):
    text = '' if value is None else str(value).strip()
    return '' if text.lower() in ('nan', 'none') else text


def level_text(value):
    text = cell(value)
    if not text:
        return ''
    try:
        return str(int(float(text)))
    except ValueError:
        return text


def load_rules(raw):
    frame = pd.read_excel(io.BytesIO(raw), sheet_name='Event', dtype=str).fillna('')
    rules = {}
    for row in frame.to_dict('records'):
        code = cell(row.get('Event_Code'))
        if code:
            rules[code] = {'Event_Code': code, '事件中文': cell(row.get('事件中文')), 'Level': level_text(row.get('Level')), 'Action': cell(row.get('Action'))}
    if EXCLUDE_CODE not in rules:
        raise ValueError('Parameter_Event.xlsx 缺少 NOT_RELEVANT，無法套用「排除」')
    return rules


def read_review(raw):
    frame = pd.read_excel(io.BytesIO(raw), sheet_name=SHEET, dtype=str).fillna('')
    missing = [c for c in REQUIRED if c not in frame.columns]
    if missing:
        raise ValueError('覆核檔缺少欄位：' + '、'.join(missing))
    frame.attrs['has_state'] = '目前事件代碼' in frame.columns and '目前Level' in frame.columns
    for c in OPTIONAL:
        if c not in frame.columns:
            frame[c] = ''
    return frame


def validate(row, rules):
    decision = cell(row['最終決定'])
    code = cell(row['新事件代碼'])
    if not cell(row['news_id']):
        return 'news_id 空白'
    if decision not in DECISIONS:
        return f'最終決定必須是 {"、".join(DECISIONS)}，目前為「{decision}」'
    if decision == '修改':
        if not code:
            return '「修改」必須填新事件代碼'
        if code not in rules:
            return f'新事件代碼 {code} 不在規則表內'
        if code == EXCLUDE_CODE:
            return '要排除請選「排除」，不要用「修改」填 NOT_RELEVANT'
    elif code:
        return '只有「修改」才能填新事件代碼'
    return ''


def target_values(decision, code, rules):
    if decision == '排除':
        rule = rules[EXCLUDE_CODE]
        return {'Event_Code': EXCLUDE_CODE, '事件中文': rule['事件中文'], 'Level': '0', 'Action': EXCLUDE_ACTION}
    if decision == '修改':
        rule = rules[code]
        return {'Event_Code': code, '事件中文': rule['事件中文'], 'Level': rule['Level'], 'Action': rule['Action']}
    return {}


def apply_review(repository, review_raw, name, rules_raw=None, now=None):
    """One transaction for the whole workbook. Returns {'results': [...], 'applied': n, 'skipped': n}."""
    review = read_review(review_raw)
    rules = load_rules(rules_raw if rules_raw is not None else repository.read('Parameter_Event.xlsx'))
    now = now or datetime.now(TPE)
    stem = Path(name).stem

    def build(read, revision):
        results, files, days = [], {}, {}
        changed_days = set()
        for position, row in enumerate(review.to_dict('records'), start=1):
            entry = {'序號': cell(row.get('序號')) or position, 'news_id': cell(row['news_id']), '日期': cell(row['日期']),
                     '公司': cell(row.get('公司')), '標題': cell(row.get('標題（點擊開原文）'))[:80],
                     '最終決定': cell(row['最終決定']), '結果': '略過', '原因': ''}
            results.append(entry)
            problem = validate(row, rules)
            if problem:
                entry['原因'] = problem
                continue
            day = cell(row['日期'])[:10]
            try:
                date.fromisoformat(day)
            except ValueError:
                entry['原因'] = f'日期格式錯誤：{day}'
                continue
            if day not in days:
                # Raw stored file: untouched rows stay byte-identical (no re-enrichment of other rows).
                days[day] = csv_frame(read(f'news_daily/{day}.csv'))
            frame = days[day]
            if frame.empty or 'news_id' not in frame:
                entry['原因'] = f'找不到 {day} 的資料檔'
                continue
            matches = frame.index[frame['news_id'].astype(str) == cell(row['news_id'])]
            if len(matches) != 1:
                entry['原因'] = '找不到這則新聞（news_id 不存在於該日期）'
                continue
            index = matches[0]
            current = frame.loc[index]
            if cell(current.get('人工版本')) != cell(row['人工版本']):
                entry['原因'] = '這筆新聞已有新的人工覆核或已匯入過，未覆蓋'
                continue
            expected_code, expected_level = cell(row['目前事件代碼']), level_text(row['目前Level'])
            if review.attrs.get('has_state') and \
                    (cell(current.get('Event_Code')) != expected_code or level_text(current.get('Level')) != expected_level):
                entry['原因'] = '分類在覆核檔產生後已改變（排程重判或他人修改），請重新產生覆核檔'
                continue
            new = target_values(entry['最終決定'], cell(row['新事件代碼']), rules)
            if cell(row['更正公司']):
                new['Company'] = cell(row['更正公司'])
            if cell(row['更正Ticker']):
                from .issuers import symbol
                new['Ticker'] = symbol(cell(row['更正Ticker']), str(current.get('Market') or 'US'))
            modified = [k for k in FINAL_FIELDS if k in new and cell(new[k]) != cell(current.get(k))]
            for field, value in new.items():
                frame.loc[index, field] = value
            frame.loc[index, '人工修改欄位'] = tags(cell(current.get('人工修改欄位')), '|'.join(modified))
            frame.loc[index, '覆核註記'] = cell(row['覆核註記'])
            for field, value in {'已人工覆核': 'True', '需要人工覆核': 'False', '分類狀態': '人工覆核', '最終判斷來源': '人工覆核',
                                 '人工版本': uuid.uuid4().hex, '人工覆核人': cell(row['覆核人']) or DEFAULT_REVIEWER,
                                 '人工覆核時間': now.isoformat()}.items():
                frame.loc[index, field] = value
            entry['結果'] = '已套用'
            entry['原因'] = ('更改：' + '、'.join(modified)) if modified else '維持原分類，僅標記已人工覆核'
            changed_days.add(day)
        for day in sorted(changed_days):
            frame = days[day]
            files[f'news_daily/{day}.csv'] = frame.to_csv(index=False).encode('utf-8-sig')
            files[f'news_daily/{day}.xlsx'] = excel_bytes(frame, read_json(read('news_runs/index.json'), {'runs': []})['runs'])
        applied = sum(1 for r in results if r['結果'] == '已套用')
        files[f'{DONE}{stem}_結果.csv'] = pd.DataFrame(results).to_csv(index=False).encode('utf-8-sig')
        files[f'{DONE}{name}'] = review_raw
        return files, {'results': results, 'applied': applied, 'skipped': len(results) - applied}

    return repository.transaction(build, f'離線覆核匯入：{name}')


def inbox_files(repository):
    return sorted(p for p in repository.list(INBOX)
                  if p.endswith('_已確認.xlsx') and p.count('/') == 1)


def main(argv=None):
    parser = argparse.ArgumentParser(description='Apply review_inbox/*_已確認.xlsx without any LLM call')
    parser.add_argument('--root', default='.', help='repository checkout (default: current directory)')
    parser.add_argument('--dry-run', action='store_true', help='validate and report only; write nothing')
    args = parser.parse_args(argv)
    root = Path(args.root).resolve()
    repository = LocalRepository(root)
    if args.dry_run:
        class DryRun(LocalRepository):
            def transaction(self, build, message):
                return build(self.read, 'local')[1]
        repository = DryRun(root)
    paths = inbox_files(repository)
    if not paths:
        print('review_inbox/ 沒有等待匯入的 *_已確認.xlsx')
        return 0
    failed = False
    summary = []
    for path in paths:
        name = path.split('/')[-1]
        try:
            outcome = apply_review(repository, repository.read(path), name)
        except Exception as error:
            print(f'::error::{name} 匯入失敗：{error}')
            failed = True
            continue
        line = f'{name}：已套用 {outcome["applied"]} 筆，略過 {outcome["skipped"]} 筆'
        print(line)
        summary.append(f'### {line}')
        for r in outcome['results']:
            if r['結果'] == '略過':
                detail = f'序號 {r["序號"]} {r["公司"]} {r["標題"]}：{r["原因"]}'
                print('  略過 ' + detail)
                summary.append('- 略過 ' + detail)
        if not args.dry_run:
            (root / path).unlink()
    target = os.environ.get('GITHUB_STEP_SUMMARY')
    if target and summary:
        with open(target, 'a', encoding='utf-8') as handle:
            handle.write('\n'.join(summary) + '\n')
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
