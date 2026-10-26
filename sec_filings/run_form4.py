"""Run the pinned user source with workspace paths and explicit Form4/A exclusion."""
from pathlib import Path
import ast
import hashlib
import importlib.metadata
import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd

ROOT = Path(__file__).resolve().parent
os.environ['EDGAR_LOCAL_DATA_DIR'] = str(ROOT.parent / 'sec-cache')
os.environ['EDGAR_RATE_LIMIT_PER_SEC'] = '4'
SOURCE = ROOT / 'Form4_FINAL.py'


class RunSettings(ast.NodeTransformer):
    def visit_Assign(self, node):
        self.generic_visit(node)
        names = {target.id for target in node.targets if isinstance(target,ast.Name)}
        if 'BASE_DIR' in names:
            node.value = ast.parse('Path('+repr(str(ROOT))+')',mode='eval').body
        elif 'TODAY' in names:
            today = datetime.now(ZoneInfo('Asia/Taipei')).date().isoformat()
            node.value = ast.parse('date.fromisoformat('+repr(today)+')',mode='eval').body
        elif 'SLEEP_SECONDS' in names:
            node.value = ast.Constant(0.35)
        return node

    def visit_Call(self,node):
        self.generic_visit(node)
        if isinstance(node.func,ast.Attribute) and node.func.attr=='to_dataframe' and isinstance(node.func.value,ast.Name) and node.func.value.id=='form4':
            return ast.copy_location(ast.Call(func=ast.Name(id='_raw_transactions',ctx=ast.Load()),args=[node.func.value],keywords=[]),node)
        if isinstance(node.func,ast.Attribute) and node.func.attr=='get_filings':
            form = next((k.value.value for k in node.keywords if k.arg=='form' and isinstance(k.value,ast.Constant)),None)
            if form in ('4','4/A'):
                # EdgarTools strips '/A' when amendments=False. Use True only
                # for the explicit '4/A' query so it counts amendments alone.
                node.keywords = [k for k in node.keywords if k.arg!='amendments']
                node.keywords.append(ast.keyword(arg='amendments',value=ast.Constant(form=='4/A')))
        return node


def raw_transactions(form4):
    frames = []
    for field in ['non_derivative_table','derivative_table']:
        table = getattr(form4,field,None)
        transactions = getattr(table,'transactions',None)
        data = getattr(transactions,'data',None)
        if data is not None and not data.empty:
            if not {'Date','Code'}.issubset(data):
                raise ValueError('Raw transaction table lacks Date/Code')
            frames.append(data.copy())
    return pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()


def main():
    started = datetime.now(ZoneInfo('Asia/Taipei')).isoformat()
    tree = RunSettings().visit(ast.parse(SOURCE.read_text(encoding='utf-8-sig')))
    ast.fix_missing_locations(tree)
    namespace = {'__name__':'__main__','__file__':str(SOURCE),'_raw_transactions':raw_transactions}
    exec(compile(tree,str(SOURCE),'exec'),namespace)
    methodology = namespace['methodology_df']
    methodology.loc[methodology['項目'].eq('Excluded Transactions'),'定義'] = '由non_derivative_table／derivative_table的原始逐筆交易保留非P/S，依交易日過濾，不使用彙總reporting_date。'
    with pd.ExcelWriter(namespace['FINAL_FILE'],engine='openpyxl') as writer:
        for name,key in [('Ticker Summary','summary_df'),('Dashboard Summary','dashboard_df'),('Form4 Detail','detail_df'),('Excluded Transactions','excluded_df'),('Run Log','run_log_df'),('Ticker Universe','universe_df'),('Methodology','methodology_df')]:
            namespace[key].to_excel(writer,sheet_name=name,index=False)
    logs = namespace['run_log_df']
    manifest = {
        'source_repo':os.getenv('GITHUB_REPOSITORY','RMSecfb/negative-news'),
        'execution_commit':os.getenv('GITHUB_SHA','local'),
        'source_sha256':hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        'edgartools':importlib.metadata.version('edgartools'),
        'pandas':importlib.metadata.version('pandas'),
        'started_at':started,
        'finished_at':datetime.now(ZoneInfo('Asia/Taipei')).isoformat(),
        'timezone':'Asia/Taipei',
        'analysis_start':namespace['ANALYSIS_START'].isoformat(),
        'analysis_end':namespace['ANALYSIS_END'].isoformat(),
        'tickers':len(namespace['universe_df']),
        'ps_transactions':len(namespace['detail_df']),
        'excluded_transactions':len(namespace['excluded_df']),
        'status_counts':{str(k):int(v) for k,v in logs['Status'].value_counts().items()},
        'runtime_changes':['workspace BASE_DIR','Taipei run date','4 requests/sec','amendments=False for Form 4; amendments=True for explicit Form 4/A query','raw per-transaction dates for non-P/S excluded records'],
        'output_file':namespace['FINAL_FILE'].relative_to(ROOT).as_posix()
    }
    (ROOT/'run_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    (namespace['FINAL_FILE'].parent/'Form4_Run_Metadata.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(manifest,ensure_ascii=False,indent=2),flush=True)
    if not logs['Status'].isin(['OK','PARTIAL']).any():
        raise RuntimeError('All ticker fetches failed; do not publish workbook')


if __name__=='__main__':
    main()
