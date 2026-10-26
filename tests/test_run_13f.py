import ast
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'sec_filings'))
import run_13f as r


class QuarterTests(unittest.TestCase):
    def test_last_completed_quarter(self):
        cases = {date(2026, 10, 26): date(2026, 9, 30), date(2026, 1, 5): date(2025, 12, 31),
                 date(2026, 6, 30): date(2026, 6, 30), date(2026, 7, 1): date(2026, 6, 30),
                 date(2026, 3, 30): date(2025, 12, 31)}
        for today, expected in cases.items():
            self.assertEqual(r.last_completed_quarter_end(today), expected)

    def test_quarter_label(self):
        self.assertEqual(r.quarter_label(date(2026, 6, 30)), '2026Q2')
        self.assertEqual(r.quarter_label(date(2025, 12, 31)), '2025Q4')

    def test_skip_only_after_window_with_existing_result(self):
        period = date(2026, 6, 30)
        self.assertFalse(r.should_skip(date(2026, 8, 1), period, 30))
        self.assertTrue(r.should_skip(date(2026, 9, 1), period, 30))
        self.assertFalse(r.should_skip(date(2026, 9, 1), period, None))
        self.assertFalse(r.should_skip(date(2026, 9, 1), period, 30, force=True))


class GateTests(unittest.TestCase):
    def test_incomplete_quarter_is_not_published(self):
        self.assertFalse(r.decide_publish(0.5, 18, None)[0])

    def test_complete_quarter_is_published(self):
        self.assertTrue(r.decide_publish(0.86, 32, None)[0])

    def test_degraded_rerun_keeps_published_result(self):
        self.assertFalse(r.decide_publish(0.81, 30, 32)[0])
        self.assertTrue(r.decide_publish(0.9, 33, 32)[0])


class PatchTests(unittest.TestCase):
    def test_runner_overrides_paths_and_target_period_only(self):
        source = ('BASE_DIR = Path("D:/x")\nOUTPUT_DIR = BASE_DIR / "institutional_13f"\n'
                  'TARGET_REPORT_PERIOD = (Counter(periods).most_common(1)[0][0])\nOTHER = 1\n')
        tree = r.RunSettings('2026-09-30').visit(ast.parse(source))
        values = {n.targets[0].id: ast.unparse(n.value) for n in tree.body}
        self.assertEqual(values['TARGET_REPORT_PERIOD'], "'2026-09-30'")
        self.assertIn('.staging_13f', values['OUTPUT_DIR'])
        self.assertNotIn('D:/x', values['BASE_DIR'])
        self.assertEqual(values['OTHER'], '1')


if __name__ == '__main__':
    unittest.main()
