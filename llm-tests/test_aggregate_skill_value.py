import unittest

from aggregate_skill_value import aggregate


class AggregateTests(unittest.TestCase):
    def record(self, condition="without-skill", **extra):
        return {
            "task": "example", "transport": "mcp", "condition": condition,
            "repetition": 1, "passed": True, "tokens": 100, "skill_reads": [],
            "execution": {"private_path": "do not publish"}, **extra,
        }

    def test_interruption_counts_without_becoming_zero_usage(self):
        receipt = aggregate([
            self.record(), self.record("with-skill", tokens=150, skill_reads=[{"success": True}]),
            self.record(passed=False, tokens=None),
        ])
        self.assertEqual(receipt["reserved_attempts"], 3)
        self.assertEqual(receipt["verified_cases"], 2)
        self.assertEqual(receipt["token_increase_percent"], {"mcp": 50.0})
        self.assertNotIn("execution", receipt["cases"][0])

    def test_duplicate_completed_case_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Duplicate verified"):
            aggregate([self.record(), self.record()])

    def test_missing_usage_is_unknown(self):
        receipt = aggregate([self.record(tokens=None), self.record("with-skill")])
        self.assertEqual(receipt["token_increase_percent"], {})
        self.assertIsNone(receipt["groups"][1]["recorded_tokens"])
        self.assertIn("missing", receipt["token_comparison_exclusions"]["mcp"])

    def test_failed_or_unknown_read_attempts_are_not_successful_reads(self):
        for read in ({"success": False}, {"success": None}, {}):
            with self.subTest(read=read):
                receipt = aggregate([self.record("with-skill", skill_reads=[read])])
                self.assertFalse(receipt["cases"][0]["skill_read"])
                self.assertEqual(receipt["groups"][0]["skill_reads"], 0)

    def test_a_successful_read_is_counted_once_despite_other_read_attempts(self):
        receipt = aggregate([self.record("with-skill", skill_reads=[
            {"success": False}, {"success": True}, {"success": True},
        ])])
        self.assertTrue(receipt["cases"][0]["skill_read"])
        self.assertEqual(receipt["groups"][0]["skill_reads"], 1)

    def test_equal_counts_with_different_tasks_cannot_produce_a_percentage(self):
        receipt = aggregate([
            self.record(), self.record("with-skill", task="different", tokens=150),
        ])
        self.assertEqual(receipt["token_increase_percent"], {})
        self.assertIn("mcp", receipt["token_comparison_exclusions"])

    def test_equal_counts_with_different_repetitions_cannot_produce_a_percentage(self):
        receipt = aggregate([
            self.record(), self.record("with-skill", repetition=2, tokens=150),
        ])
        self.assertEqual(receipt["token_increase_percent"], {})
        self.assertIn("mcp", receipt["token_comparison_exclusions"])

    def test_incomplete_verified_pairs_cannot_produce_a_percentage(self):
        receipt = aggregate([
            self.record(), self.record("with-skill", tokens=150),
            self.record(task="unfinished"),
            self.record("with-skill", task="unfinished", passed=False, tokens=None),
        ])
        self.assertEqual(receipt["verified_cases"], 3)
        self.assertEqual(receipt["incomplete_or_unverified_attempts"], 1)
        self.assertEqual(receipt["token_increase_percent"], {})
        self.assertIn("mcp", receipt["token_comparison_exclusions"])

    def test_an_unmatched_transport_does_not_hide_a_matched_transport(self):
        receipt = aggregate([
            self.record(), self.record("with-skill", tokens=150),
            self.record(transport="cli"),
            self.record("with-skill", transport="cli", task="different", tokens=150),
        ])
        self.assertEqual(receipt["token_increase_percent"], {"mcp": 50.0})
        self.assertIn("cli", receipt["token_comparison_exclusions"])

    def test_zero_treatment_usage_is_not_missing_usage(self):
        receipt = aggregate([self.record(), self.record("with-skill", tokens=0)])
        self.assertEqual(receipt["token_increase_percent"], {"mcp": -100.0})

    def test_zero_baseline_usage_has_an_explicit_exclusion(self):
        receipt = aggregate([self.record(tokens=0), self.record("with-skill")])
        self.assertEqual(receipt["token_increase_percent"], {})
        self.assertIn("zero", receipt["token_comparison_exclusions"]["mcp"])
