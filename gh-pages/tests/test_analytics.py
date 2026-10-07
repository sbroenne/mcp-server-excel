"""Tests for the public usage analytics page. Run from ``gh-pages/``::

    python -m unittest discover -s tests
"""

from __future__ import annotations

import copy
import json
import re
import unittest
from pathlib import Path
from unittest import mock

import sitegen.analytics as analytics

FIXTURES = Path(__file__).parent / "fixtures"


def _load(name: str) -> dict[str, object]:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def _render(report: dict[str, object]) -> str:
    with mock.patch.object(analytics, "read", return_value=json.dumps(report)):
        return analytics.render_usage_analytics()


def _section(page: str, heading: str) -> str:
    start = page.index(heading)
    following = re.search(r"^#{2,3} ", page[start + len(heading):], re.MULTILINE)
    end = start + len(heading) + following.start() if following else len(page)
    return page[start:end]


def _widths(html: str) -> list[str]:
    return re.findall(r'style="width: ([\d.]+)%"', html)


class WeightedReportTests(unittest.TestCase):
    def setUp(self):
        self.report = _load("usage-analytics-v3.json")

    def with_both_entry_points(self) -> dict[str, object]:
        report = copy.deepcopy(self.report)
        report["entryPoints"][1] = {
            "name": "cli",
            "enoughData": True,
            "users": 10,
            "actions": 125,
            "workUnits": 300,
            "actionsPerUser": 12.5,
            "workUnitsPerUser": 30.0,
        }
        return report

    def test_renders_work_entry_points_habits_and_downloads(self):
        page = _render(self.report)
        for heading in (
            "## Command line and AI assistant",
            "## How people work",
            "## Where people get Excel MCP",
            "### npm downloads each week",
        ):
            self.assertIn(heading, page)

    def test_pie_compares_people_and_actions_when_both_groups_have_enough_data(self):
        page = _render(self.with_both_entry_points())
        section = _section(page, "### Share of people and actions")
        self.assertIn(
            'aria-label="Users: AI assistant (MCP Server) 80.0%, Command line (excelcli) 20.0%"',
            section,
        )
        self.assertIn(
            'aria-label="Actions: AI assistant (MCP Server) 80.0%, Command line (excelcli) 20.0%"',
            section,
        )
        self.assertIn("conic-gradient(#4051b5 0.00% 80.00%, #d97706 80.00% 100.00%)", section)

    def test_pie_is_hidden_when_a_group_lacks_enough_data(self):
        page = _render(self.report)
        self.assertNotIn("analytics-pie", page)
        self.assertIn("There is not enough data yet to show **Command line (excelcli)**", page)

    def test_percentage_pairs_use_a_fixed_scale_and_zero_has_no_bar(self):
        page = _render(self.report)
        sessions = _section(page, "### Size of AI assistant sessions")
        # 2-10 actions: 60% of sessions, 5% of actions; 11-50 actions: none.
        self.assertEqual(_widths(sessions)[:4], ["60.00", "5.00", "0.00", "0.00"])

    def test_percentage_bars_use_a_fixed_scale(self):
        page = _render(self.report)
        pairs = _section(page, "### Areas used together")
        self.assertEqual(_widths(pairs), ["40.00"])

    def test_work_share_bars_use_a_fixed_scale(self):
        page = _render(self.report)
        work = _section(page, "## Where most of the work goes")
        self.assertEqual(_widths(work), ["80.00", "20.00"])

    def test_small_values_stay_visible(self):
        self.assertEqual(analytics._analytics_bar_size(0, 100), 0)
        self.assertEqual(analytics._analytics_bar_size(0.5, 100), 2)
        self.assertEqual(analytics._analytics_bar_size(150, 100), 100)

    def test_session_summary_comes_from_the_data(self):
        page = _render(self.report)
        sessions = _section(page, "### Size of AI assistant sessions")
        self.assertIn(
            "Sessions with more than 200 actions were **20%** of sessions but "
            "**94.67%** of actions.",
            sessions,
        )
        self.assertNotIn("A small number of long sessions do much of the work", page)

    def test_session_sentence_is_omitted_without_long_sessions(self):
        report = copy.deepcopy(self.report)
        report["habits"]["assistantSessions"]["sizes"][-1] = {"size": "201+", "enoughData": False}
        page = _render(report)
        self.assertNotIn("more than 200 actions", page)

    def test_small_groups_are_not_published(self):
        page = _render(self.report)
        sessions = _section(page, "### Size of AI assistant sessions")
        self.assertNotIn(">1 action<", sessions)
        self.assertIn("Not shown because too few people had them: 1 action.", sessions)
        weekdays = _section(page, "### Weekdays and weekends")
        self.assertNotIn("<span>Saturday</span>", weekdays)
        self.assertIn("too few people used Excel MCP on them: Saturday.", weekdays)

    def test_whole_habit_blocks_hide_when_too_few_people(self):
        report = copy.deepcopy(self.report)
        report["habits"]["assistantSessions"] = {"enoughData": False}
        report["habits"]["returningUsers"] = {"enoughData": False}
        page = _render(report)
        self.assertIn("There is not enough data yet to describe AI assistant sessions.", page)
        self.assertIn("There are not enough new people yet to show whether they come back.", page)

    def test_first_advanced_use_states_its_window(self):
        page = _render(self.report)
        self.assertIn("started using Excel MCP in the last **60 days**", page)

    def test_long_gaps_between_reports_show_an_average_week(self):
        report = copy.deepcopy(self.report)
        report["downloads"]["weeklyGains"] = [
            {"week": "2026-09-16", "days": 7, "total": 70, "channels": {}},
            {"week": "2026-09-23", "days": 14, "total": 280, "channels": {}},
        ]
        page = _render(report)
        gains = _section(page, "### New downloads each week")
        self.assertIn("Bars marked * cover a gap between reports that was not one week", gains)
        self.assertIn('aria-label="Week of Sep 16: 70"', gains)
        self.assertIn('aria-label="Average week from Sep 23 to Oct 07: 140"', gains)
        self.assertIn(">Sep 23*</small>", gains)

    def test_short_gaps_between_reports_show_an_average_week(self):
        report = copy.deepcopy(self.report)
        report["downloads"]["weeklyGains"] = [
            {"week": "2026-09-30", "days": 7, "total": 70, "channels": {}},
            {"week": "2026-10-07", "days": 2, "total": 40, "channels": {}},
        ]
        page = _render(report)
        gains = _section(page, "### New downloads each week")
        self.assertIn("Bars marked *", gains)
        self.assertIn('aria-label="Average week from Oct 07 to Oct 09: 140"', gains)
        self.assertIn(">Oct 07*</small>", gains)

    def test_weekly_gaps_do_not_add_a_note(self):
        report = copy.deepcopy(self.report)
        report["downloads"]["weeklyGains"] = [
            {"week": "2026-09-30", "days": 7, "total": 70, "channels": {}}
        ]
        page = _render(report)
        self.assertNotIn("Bars marked *", page)


class LegacyReportTests(unittest.TestCase):
    def test_schema_2_report_still_renders(self):
        page = _render(_load("usage-analytics-v2.json"))
        self.assertIn("## Most common actions", page)
        self.assertNotIn("## How people work", page)
        self.assertNotIn("## Command line and AI assistant", page)


if __name__ == "__main__":
    unittest.main()
