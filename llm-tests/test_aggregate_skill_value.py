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
