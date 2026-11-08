import unittest
from pipeline.prefilter import skip_llm


class PrefilterTests(unittest.TestCase):
    def test_skips_positive_or_routine_without_event_words(self):
        self.assertTrue(skip_llm({'Title': 'Apple shares rise on strong iPhone demand', 'FinBERT': '0.8'}))
        self.assertTrue(skip_llm({'Title': 'Analyst keeps price target on Microsoft', 'FinBERT': '-0.1'}))

    def test_keeps_event_words_and_unknown_sentiment(self):
        self.assertEqual(skip_llm({'Title': 'Company files for Chapter 11 bankruptcy', 'FinBERT': '0.9'}), '')
        self.assertEqual(skip_llm({'Title': 'Stock plunges 12% after probe', 'FinBERT': '0.5'}), '')
        self.assertEqual(skip_llm({'Title': 'Some neutral headline', 'FinBERT': ''}), '')
        self.assertEqual(skip_llm({'Title': '輝達CDS飆升逾3倍', 'FinBERT': '-0.4'}), '')


if __name__ == '__main__':
    unittest.main()
