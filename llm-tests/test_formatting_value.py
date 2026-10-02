import copy
import unittest

from skill_value import check_attempt_budget
from skill_value_tasks import check_snapshot
from test_skill_value import cases


def snapshots():
    values = [["Account", "Amount (USD)", "Fractional rate"], ["001", 135.25, 0.45],
              ["002", -210.5, 0.12], ["003", 0, 0], ["Total", -75.25, 0.19],
              [None, None, None], ["Keep this note", None, None]]
    formulas = copy.deepcopy(values)
    formulas[4][1:] = ["=SUM(B2:B4)", "=AVERAGE(C2:C4)"]
    before = {"calculationMode": "automatic", "sheets": [{
        "name": "Sheet1", "sourceValues": values, "sourceFormulas": formulas,
        "sourceFormats": [["General"] * 3 for _ in range(7)], "tables": [], "charts": [],
        "sourcePresentation": [[{"bold": False, "fontColor": 0, "text": str(value)}
                                for value in row] for row in values],
    }]}
    after = copy.deepcopy(before)
    sheet = after["sheets"][0]
    for cell in sheet["sourcePresentation"][0]:
        cell["bold"] = True
    for row in (1, 2, 3, 4):
        sheet["sourceFormats"][row][1:] = ['$#,##0.00;($#,##0.00);"-"', "0.0%"]
        sheet["sourcePresentation"][row][2]["text"] = f"{values[row][2] * 100:.1f}%"
    for row in (1, 2, 3):
        for column in (1, 2):
            sheet["sourcePresentation"][row][column]["fontColor"] = 16711680
    sheet["sourcePresentation"][2][1]["text"] = "($210.50)"
    sheet["sourcePresentation"][3][1]["text"] = "-"
    return before, after


class FormattingChecks(unittest.TestCase):
    def test_percent_symbol_without_scaling_is_not_a_correct_rate(self):
        before, after = snapshots()
        after["sheets"][0]["sourcePresentation"][1][2]["text"] = "0.45%"
        with self.assertRaises(AssertionError):
            check_snapshot("formatting-report", after, before)

    def test_matrix_has_sixteen_positive_and_eight_negative_cases(self):
        matrix = cases(None, "formatting")
        self.assertEqual(len(matrix), 24)
        self.assertEqual(sum(task in ("formatting-report", "financial-formatting")
                             for task, _, _, _ in matrix), 16)

    def test_checkers_accept_conventions_and_reject_each_wrong_state(self):
        before, after = snapshots()
        for task in ("formatting-report", "financial-formatting"):
            check_snapshot(task, after, before)
        for field, row, column, wrong in (
            ("sourceValues", 1, 2, 45),
            ("sourceFormulas", 4, 1, -75.25),
            ("sourceFormats", 1, 2, "0.0"),
        ):
            invalid = copy.deepcopy(after)
            invalid["sheets"][0][field][row][column] = wrong
            with self.assertRaises(AssertionError):
                check_snapshot("formatting-report", invalid, before)
        invalid = copy.deepcopy(after)
        invalid["sheets"][0]["sourcePresentation"][1][1]["fontColor"] = 0
        with self.assertRaises(AssertionError):
            check_snapshot("financial-formatting", invalid, before)

    def test_approved_ceiling_counts_prior_attempts(self):
        check_attempt_budget(24, 50, 100)
        for planned, prior, ceiling in ((51, 50, 100), (24, 50, 60), (1, -1, 100), (1, 0, 101)):
            with self.assertRaises(ValueError):
                check_attempt_budget(planned, prior, ceiling)
