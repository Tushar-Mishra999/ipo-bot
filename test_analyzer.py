"""Test suite for fundamental IPO analyzer."""

import os
import unittest
from ipo_bot.analyzer import analyze_file, load_config, interpolate_score

SAMPLE_DIR = os.path.join(os.path.dirname(__file__), "sample_data")
CRYSTAL_FILE = os.path.join(SAMPLE_DIR, "Crystal_Crop_Template.xlsx")
MAHANADI_FILE = os.path.join(SAMPLE_DIR, "Mahanadi_Coalfields_Template.xlsx")


class TestIPOAnalyzer(unittest.TestCase):

    def test_config_loaded(self):
        config = load_config()
        self.assertIn("pillars", config)
        self.assertEqual(len(config["pillars"]), 5)
        self.assertIn("earnings_quality", config["pillars"])
        self.assertIn("forensic_risk", config["pillars"])
        self.assertIn("financial_strength", config["pillars"])
        self.assertIn("profitability_growth", config["pillars"])
        self.assertIn("cash_allocation", config["pillars"])

    def test_interpolate_score(self):
        # Higher is better: bad 0.5, good 1.0
        self.assertEqual(interpolate_score(0.5, 0.5, 1.0, True), 0.0)
        self.assertEqual(interpolate_score(1.0, 0.5, 1.0, True), 100.0)
        self.assertEqual(interpolate_score(0.75, 0.5, 1.0, True), 50.0)
        self.assertEqual(interpolate_score(0.2, 0.5, 1.0, True), 0.0)
        self.assertEqual(interpolate_score(1.5, 0.5, 1.0, True), 100.0)

        # Lower is better: bad 1.5, good 0.3
        self.assertEqual(interpolate_score(1.5, 1.5, 0.3, False), 0.0)
        self.assertEqual(interpolate_score(0.3, 1.5, 0.3, False), 100.0)
        self.assertEqual(interpolate_score(0.9, 1.5, 0.3, False), 50.0)
        self.assertEqual(interpolate_score(2.0, 1.5, 0.3, False), 0.0)
        self.assertEqual(interpolate_score(0.1, 1.5, 0.3, False), 100.0)

    def test_crystal_crop_analysis(self):
        if not os.path.exists(CRYSTAL_FILE):
            self.skipTest(f"{CRYSTAL_FILE} not found")
        res = analyze_file(CRYSTAL_FILE)
        print("\n--- Crystal Crop Protection Analysis ---")
        print(f"Company: {res['company_name']}")
        print(f"Composite Score: {res['composite_score']} / 100 (Coverage: {res['overall_coverage_pct']}%)")
        print(f"Verdict: {res['verdict']} (Base: {res['base_verdict']})")
        if res['override_note']:
            print(f"Override: {res['override_note']}")
        for p_key, p_val in res['pillars'].items():
            print(f"  {p_val['name']}: {p_val['score']} (Coverage: {p_val['coverage_pct']}%)")
        print("Top Strengths:")
        for s in res['top_strengths']:
            print(f"  + {s}")
        print("Top Red Flags:")
        for rf in res['top_red_flags']:
            print(f"  - {rf}")

        self.assertGreaterEqual(res["composite_score"], 0.0)
        self.assertLessEqual(res["composite_score"], 100.0)
        self.assertIn(res["verdict"], ["Invest", "Watchlist", "Avoid"])
        self.assertGreaterEqual(res["overall_coverage_pct"], 80.0)
        self.assertEqual(len(res["pillars"]), 5)

    def test_mahanadi_analysis(self):
        if not os.path.exists(MAHANADI_FILE):
            self.skipTest(f"{MAHANADI_FILE} not found")
        res = analyze_file(MAHANADI_FILE)
        print("\n--- Mahanadi Coalfields Analysis ---")
        print(f"Company: {res['company_name']}")
        print(f"Composite Score: {res['composite_score']} / 100 (Coverage: {res['overall_coverage_pct']}%)")
        print(f"Verdict: {res['verdict']} (Base: {res['base_verdict']})")
        if res['override_note']:
            print(f"Override: {res['override_note']}")
        for p_key, p_val in res['pillars'].items():
            print(f"  {p_val['name']}: {p_val['score']} (Coverage: {p_val['coverage_pct']}%)")
        print("Top Strengths:")
        for s in res['top_strengths']:
            print(f"  + {s}")
        print("Top Red Flags:")
        for rf in res['top_red_flags']:
            print(f"  - {rf}")

        self.assertGreaterEqual(res["composite_score"], 0.0)
        self.assertLessEqual(res["composite_score"], 100.0)
        self.assertIn(res["verdict"], ["Invest", "Watchlist", "Avoid"])
        self.assertGreaterEqual(res["overall_coverage_pct"], 80.0)
        self.assertEqual(len(res["pillars"]), 5)


if __name__ == "__main__":
    unittest.main()
