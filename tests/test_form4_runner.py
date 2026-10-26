import ast
import sys
import unittest
from pathlib import Path
import pandas as pd
from types import SimpleNamespace
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'sec_filings'))
from run_form4 import RunSettings, raw_transactions


class RunnerTests(unittest.TestCase):
    def test_original_form_and_amendments_are_queried_separately(self):
        tree=RunSettings().visit(ast.parse('a=c.get_filings(form="4")\nb=c.get_filings(form="4/A")'))
        calls=[n for n in ast.walk(tree) if isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='get_filings']
        values=[{k.arg:ast.literal_eval(k.value) for k in call.keywords} for call in calls]
        self.assertEqual(values,[{'form':'4','amendments':False},{'form':'4/A','amendments':True}])

    def test_raw_excluded_transactions_keep_individual_trade_dates(self):
        raw=pd.DataFrame([{'Date':'2026-10-01','Code':'F','Shares':10},{'Date':'2026-10-02','Code':'M','Shares':20}])
        table=SimpleNamespace(transactions=SimpleNamespace(data=raw))
        form=SimpleNamespace(non_derivative_table=table,derivative_table=None)
        result=raw_transactions(form)
        self.assertEqual(result['Date'].tolist(),['2026-10-01','2026-10-02'])
        result.loc[0,'Shares']=0
        self.assertEqual(raw.loc[0,'Shares'],10)

    def test_raw_transaction_missing_dates_is_an_error(self):
        table=SimpleNamespace(transactions=SimpleNamespace(data=pd.DataFrame([{'Code':'F'}])))
        with self.assertRaises(ValueError):
            raw_transactions(SimpleNamespace(non_derivative_table=table))


if __name__=='__main__':
    unittest.main()
