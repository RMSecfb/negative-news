# AI integration regression tests
import ast
import unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from pipeline.common import merge_records
from pipeline.events import candidate_rules
from pipeline.run import classify

class RegressionTests(unittest.TestCase):
    def test_sources_compile(self):
        for path in Path(".").rglob("*.py"):
            ast.parse(path.read_text(encoding="utf-8-sig"))
    def test_same_day_union(self):
        a=pd.DataFrame([dict(Ticker="AAPL",URL="https://a")])
        b=pd.DataFrame([dict(Ticker="CRWV",URL="https://b")])
        self.assertEqual(len(merge_records(a,b)),2)
    def test_manual_decision_preserved(self):
        a=pd.DataFrame([dict(Ticker="AAPL",URL="https://a",Level="4",已人工覆核="True")])
        b=pd.DataFrame([dict(Ticker="AAPL",URL="https://a",Level="1",已人工覆核="False")])
        self.assertEqual(merge_records(a,b).iloc[0]["Level"],"4")
    def test_cds_is_not_default(self):
        rules=[dict(Event_Code="DEFAULT",Keyword=["default"],Exclude_Keyword=[],事件中文="債務違約",Primary_Priority=1)]
        self.assertEqual(candidate_rules("CoreWeave credit default swap spreads surge",rules),[])
        self.assertEqual(len(candidate_rules("CoreWeave default on debt",rules)),1)
    def test_llm_conflict_requires_review(self):
        rules=[dict(Event_Code="DEFAULT",Keyword=["default"],Exclude_Keyword=[],事件中文="債務違約",Level=5)]
        frame=pd.DataFrame([dict(Ticker="CRWV",Title="CoreWeave credit default swap spreads surge",Summary="",監控池="CUSTOM")])
        answer=dict(company_mentioned="CoreWeave",event_codes=["DEFAULT"],event_status="MARKET_IMPLIED",confirmed_default=True,evidence="credit default swap",summary_zh="信用風險升高",relevant=True,requires_review=False)
        with patch("pipeline.run.analyze",return_value=(answer,{},"")):
            result,_=classify(frame,rules,"test-key",1,"2026-10-23")
        self.assertEqual(result.iloc[0]["Event_Code"],"")
        self.assertEqual(result.iloc[0]["分類狀態"],"待人工覆核")

if __name__=="__main__":
    unittest.main()
