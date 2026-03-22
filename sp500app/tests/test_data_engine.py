import unittest
from datetime import date

from data_engine import Bar, _stats, merge_bars, resample_bars


class ResampleTests(unittest.TestCase):
    def test_weekly_ohlc_resample(self):
        bars = [
            Bar(date(2026, 1, 5), 100, 102, 99, 101, "local"),
            Bar(date(2026, 1, 6), 101, 105, 100, 104, "local"),
            Bar(date(2026, 1, 7), 104, 106, 103, 105, "local"),
        ]
        weekly = resample_bars(bars, "weekly")
        self.assertEqual(len(weekly), 1)
        self.assertEqual(weekly[0].open, 100)
        self.assertEqual(weekly[0].high, 106)
        self.assertEqual(weekly[0].low, 99)
        self.assertEqual(weekly[0].close, 105)

    def test_merge_prefers_latest_source_on_overlap(self):
        local = [Bar(date(2026, 1, 5), 100, 101, 99, 100, "local")]
        yahoo = [Bar(date(2026, 1, 5), 101, 103, 100, 102, "yahoo")]
        merged = merge_bars(local, yahoo)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].source, "yahoo")
        self.assertEqual(merged[0].close, 102)


class StatsTests(unittest.TestCase):
    def test_stats_sample_size_and_up_probability(self):
        stats = _stats([1.0, -0.5, 2.0, 0.0])
        self.assertEqual(stats["sample_size"], 4)
        self.assertAlmostEqual(stats["up_probability"], 0.5)


if __name__ == "__main__":
    unittest.main()
