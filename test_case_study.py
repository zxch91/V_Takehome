import hashlib
import math
import tempfile
import unittest
from contextlib import redirect_stderr
from datetime import date
from io import StringIO
from pathlib import Path

from openpyxl import Workbook, load_workbook
from pydantic import ValidationError

from calculations import analyse, npv, roll_dates, xirr
from case_study import main
from cashflows import Cashflow, load_cashflows, parse_date
from reporting import write_results

SOURCE = Path(__file__).with_name("CST_Case_Study_data__1_.xlsx")


class FinancialTests(unittest.TestCase):
    def test_known_xirr(self):
        flows = [
            (date(2008, 1, 1), -10000),
            (date(2008, 3, 1), 2750),
            (date(2008, 10, 30), 4250),
            (date(2009, 2, 15), 3250),
            (date(2009, 4, 1), 2750),
        ]
        rate = xirr(flows)
        self.assertAlmostEqual(rate, 0.373362535, places=8)
        self.assertAlmostEqual(npv(flows, rate, flows[0][0]), 0, places=7)

    def test_negative_irr_uses_actual_days(self):
        start, end = date(2023, 1, 1), date(2024, 7, 1)
        self.assertAlmostEqual(
            xirr([(start, -100), (end, 90)]), 0.9 ** (365 / (end - start).days) - 1, places=12
        )

    def test_same_day_netting_and_zero_return(self):
        self.assertEqual(
            xirr([(date(2025, 1, 1), 100), (date(2024, 1, 1), -120), (date(2024, 1, 1), 20)]), 0
        )

    def test_invalid_irr(self):
        for amounts in ([], [100], [-100, 230, -132], [-100, math.nan], [-100, math.inf]):
            with self.subTest(amounts=amounts), self.assertRaises(ValueError):
                xirr([(date(2025 + year, 1, 1), amount) for year, amount in enumerate(amounts)])

    def test_nav_before_and_after_cashflow(self):
        start, end = date(2025, 1, 1), date(2026, 1, 1)
        flows = [(start, -100), (end, 110)]
        rate = xirr(flows)
        self.assertAlmostEqual(npv(flows, rate, start), 0)
        self.assertAlmostEqual(npv(flows, rate, start, False), 100)
        self.assertEqual(npv(flows, rate, end), 110)
        self.assertEqual(npv(flows, rate, end, False), 0)
        self.assertAlmostEqual(npv(flows, rate, date(2025, 7, 1)), 100 * 1.1 ** (181 / 365))

    def test_invalid_discount_rates(self):
        for rate in (-1, -2, math.inf, math.nan):
            with self.subTest(rate=rate), self.assertRaises(ValueError):
                npv([], rate, date(2025, 1, 1))

    def test_rolls_preserve_month_end_and_leap_years(self):
        self.assertEqual(
            roll_dates(date(2025, 9, 30), date(2026, 3, 31)),
            [date(2025, 9, 30), date(2025, 12, 31), date(2026, 3, 31)],
        )
        self.assertEqual(roll_dates(date(2027, 11, 30), date(2028, 5, 31))[1], date(2028, 2, 29))

    def test_final_stub_and_invalid_term(self):
        self.assertEqual(
            roll_dates(date(2025, 1, 31), date(2025, 9, 15)),
            [date(2025, 1, 31), date(2025, 4, 30), date(2025, 7, 31), date(2025, 9, 15)],
        )
        with self.assertRaises(ValueError):
            roll_dates(date(2025, 1, 1), date(2025, 1, 1))


class InputTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "input.xlsx"

    def make_input(self, changes=None):
        book = Workbook()
        sheet = book.active
        sheet.append([field.alias for field in Cashflow.model_fields.values()])
        sheet.append([1, "Fund I", date(2025, 1, 1), "Investment", "GBP", -100, -120, "EUR"])
        sheet.append([2, "Fund I", date(2026, 1, 1), "Principal Repayment", "GBP", 110, 132, "EUR"])
        for coordinate, value in (changes or {}).items():
            sheet[coordinate] = value
        return book

    def load(self, changes=None):
        self.make_input(changes).save(self.path)
        return load_cashflows(self.path)

    def test_currency_and_date_corrections(self):
        records, corrections = self.load({"E2": " GPB ", "C3": "01/01/2026 00:00:00" + chr(96)})
        self.assertEqual(records[0].local_currency, "GBP")
        self.assertEqual(records[1].cashflow_date, date(2026, 1, 1))
        self.assertEqual({row["cell"] for row in corrections}, {"E2", "C3"})

    def test_invalid_rows(self):
        for changes in (
            {"C2": "31/02/2025"},
            {"E2": "GBX"},
            {"F3": -110},
            {"F2": True},
            {"F2": "NaN"},
            {"F2": "Infinity"},
            {"A3": 1},
            {"H3": "USD"},
            {"E2": "EUR"},
        ):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                self.load(changes)

    def test_rejects_invalid_model_amounts(self):
        records, _ = self.load()
        raw = records[0].model_dump()
        for value in (True, "NaN", "Infinity"):
            with self.subTest(value=value), self.assertRaises(ValidationError):
                Cashflow(**{**raw, "amount_local": value})

    def test_missing_formula_caches_including_formula_only_rows(self):
        for changes in ({"A3": "=A2+1"}, {"A4": "=A3+1"}):
            with (
                self.subTest(changes=changes),
                self.assertRaisesRegex(ValueError, "no cached result"),
            ):
                self.load(changes)

    def test_bad_headers_and_extra_data(self):
        for changes in ({"A1": "Incorrect ID"}, {"K2": 123}, {"F3": "#DIV/0!"}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                self.load(changes)

    def test_reordered_columns_and_blank_formatting(self):
        book = self.make_input()
        sheet = book.active
        order = [7, 4, 2, 0, 6, 1, 5, 3]
        for row in list(sheet.iter_rows()):
            values = [cell.value for cell in row]
            for cell, index in zip(row, order):
                cell.value = values[index]
        sheet["B1"] = " Local Currency "
        sheet["K20"].number_format = "0.00"
        book.save(self.path)
        records, corrections = load_cashflows(self.path)
        self.assertEqual(records[0].amount_base, -120)
        self.assertEqual([row["cell"] for row in corrections], ["B1"])

    def test_empty_or_ambiguous_sheets(self):
        book = self.make_input()
        book.active.delete_rows(2, 2)
        book.save(self.path)
        with self.assertRaises(ValueError):
            load_cashflows(self.path)
        book = self.make_input()
        book.copy_worksheet(book.active)
        book.save(self.path)
        with self.assertRaises(ValueError):
            load_cashflows(self.path)
        self.assertEqual(len(load_cashflows(self.path, "Sheet")[0]), 2)

    def test_date_formats(self):
        self.assertEqual(parse_date("2027-09-30"), date(2027, 9, 30))
        for invalid in ("09/30/2027", "30/09/2027 01:00:00", "2027-09-30 rubbish", 46000):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                parse_date(invalid)

    def test_base_currency_only_fund(self):
        records, corrections = self.load({"E2": "EUR", "E3": "EUR", "G2": -100, "G3": 110})
        results = analyse(records)
        self.assertEqual(results["FX Forwards"], [])
        write_results(Path(self.temp.name) / "results.xlsx", records, corrections, results)

    def test_cli_errors_preserve_input(self):
        self.make_input().save(self.path)
        before = self.path.read_bytes()
        for args in (
            [str(self.path), "--output", str(self.path)],
            [str(self.path.with_name("missing.xlsx"))],
        ):
            with redirect_stderr(StringIO()), self.assertRaises(SystemExit) as error:
                main(args)
            self.assertEqual(error.exception.code, 1)
        self.assertEqual(self.path.read_bytes(), before)


class WorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.records, cls.corrections = load_cashflows(SOURCE)
        cls.results = analyse(cls.records)

    def test_records_and_repairs(self):
        self.assertEqual(len(self.records), 126)
        self.assertEqual([record.record_id for record in self.records], list(range(1, 127)))
        self.assertEqual({row["cell"] for row in self.corrections}, {"E18", "C22", "C52"})

    def test_fund_and_currency_irrs(self):
        expected = [0.10047562818877, 0.07887049044048, 0.09950722477585, 0.12051546680617]
        for row, rate in zip(self.results["IRR"], expected):
            self.assertAlmostEqual(row["annual_irr"], rate, places=12)
        self.assertEqual(
            analyse([r for r in self.records if r.fund_name == "Fund I"]), self.results
        )
        self.assertAlmostEqual(
            analyse(self.records, "Fund II")["IRR"][0]["annual_irr"], expected[0], places=12
        )
        with self.assertRaises(ValueError):
            analyse(self.records, "Unknown")

    def test_nav_endpoints_and_rollforward(self):
        self.assertEqual(len(self.results["NAV Schedule"]), 63)
        for currency, principal in (("EUR", 150e6), ("GBP", 100e6), ("USD", 200e6)):
            rows = [row for row in self.results["NAV Schedule"] if row["currency"] == currency]
            rate = next(
                row["annual_irr"]
                for row in self.results["IRR"]
                if row["scope"] == "Currency" and row["currency"] == currency
            )
            self.assertEqual(rows[0]["nav_pre_cashflow"], 0)
            self.assertAlmostEqual(rows[0]["nav_post_cashflow"], principal, places=4)
            self.assertEqual(rows[-1]["nav_pre_cashflow"], principal)
            self.assertEqual(rows[-1]["nav_post_cashflow"], 0)
            for before, after in zip(rows, rows[1:]):
                expected = before["nav_post_cashflow"] * (1 + rate) ** (
                    (after["valuation_date"] - before["valuation_date"]).days / 365
                )
                self.assertAlmostEqual(after["nav_pre_cashflow"], expected, places=4)

    def test_forward_coverage_direction_and_continuity(self):
        self.assertEqual(len(self.results["FX Forwards"]), 40)
        for currency in ("GBP", "USD"):
            trades = [
                row for row in self.results["FX Forwards"] if row["notional_currency"] == currency
            ]
            self.assertEqual(len(trades), 20)
            self.assertEqual(trades[0]["trade_date"], date(2025, 9, 30))
            self.assertEqual(trades[-1]["delivery_date"], date(2030, 9, 30))
            for previous, following in zip(trades, trades[1:]):
                self.assertEqual(previous["delivery_date"], following["trade_date"])
            for trade in trades:
                exposure = next(
                    row["nav_post_cashflow"]
                    for row in self.results["NAV Schedule"]
                    if row["currency"] == currency and row["valuation_date"] == trade["trade_date"]
                )
                self.assertLessEqual(abs(float(trade["notional_amount"]) - exposure), 0.00501)
                self.assertEqual((trade["direction"], trade["buy_currency"]), ("SELL", "EUR"))
                self.assertIsNone(trade["forward_rate"])
        self.assertTrue(
            all(row["coverage_before_delivery"] < 1 for row in self.results["Hedge Coverage"])
        )

    def test_export_and_cleaned_input_roundtrip(self):
        before = hashlib.sha256(SOURCE.read_bytes()).digest()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            write_results(path, self.records, self.corrections, self.results)
            book = load_workbook(path)
            self.assertEqual(book["NAV Schedule"].max_row, 64)
            self.assertEqual(book["FX Forwards"].max_row, 41)
            self.assertEqual(book["Cleaned Input"].max_row, 127)
            book.close()
            records, corrections = load_cashflows(path, "Cleaned Input")
            self.assertEqual(corrections, [])
            self.assertAlmostEqual(
                analyse(records)["IRR"][0]["annual_irr"],
                self.results["IRR"][0]["annual_irr"],
                places=12,
            )
        self.assertEqual(hashlib.sha256(SOURCE.read_bytes()).digest(), before)


if __name__ == "__main__":
    unittest.main()
