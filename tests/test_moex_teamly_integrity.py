"""Portable CPU regression checks; all database access is explicitly stubbed."""
import ast
from collections import Counter
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TASKS = ROOT / "tasks" / "finalpool"


def load(task):
    spec = importlib.util.spec_from_file_location(task.replace("-", "_"), TASKS / task / "evaluation" / "main.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class Cursor:
    def __init__(self, rows):
        self.rows = rows
    def execute(self, sql, params=None):
        pass
    def fetchall(self):
        return self.rows
    def close(self):
        pass


class Connection:
    def __init__(self, rows):
        self.rows = rows
    def cursor(self):
        return Cursor(self.rows)
    def close(self):
        pass


class NewsTests(unittest.TestCase):
    def setUp(self):
        self.g = load("moex-market-news-teamly")
        tree = ast.parse((TASKS / "moex-market-news-teamly/preprocess/main.py").read_text())
        self.news = ast.literal_eval(next(n.value for n in tree.body if isinstance(n, ast.Assign)
                                         and any(isinstance(t, ast.Name) and t.id == "NEWS" for t in n.targets)))
        self.articles = [(t, title, pub, date[:10]) for t, title, pub, date in self.news]
        counts = Counter(r[2] for r in self.articles)
        self.expected = (len(self.articles), 6, counts.most_common(1)[0][0])
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def workbook(self, fake=False, reverse=False):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "News Articles"
        sheet.append(["Ticker", "Title", "Publisher", "Published_Date"])
        rows = sorted(sorted(self.articles, key=lambda r: r[3], reverse=True), key=lambda r: r[0])
        if reverse:
            rows = rows[::-1]
        for row in rows:
            sheet.append(["FAKE"] if fake else row)
        summary = wb.create_sheet("Summary")
        summary.append(["Metric", "Value"])
        for k, v in zip(["Total_Articles", "Unique_Tickers", "Top_Publisher"], self.expected):
            summary.append([k, v])
        wb.save(self.root / "Market_News_Digest.xlsx")

    def run_excel(self):
        with patch.object(self.g, "get_news_articles", return_value=self.articles), contextlib.redirect_stdout(io.StringIO()):
            self.g.check_excel(self.root, self.expected)

    def test_real_news_and_dates_pass(self):
        self.workbook()
        self.run_excel()
        self.assertEqual(self.g.CRITICAL_FAILED, [])
        self.assertEqual(self.g.FAIL_COUNT, 0)

    def test_fabricated_news_rejected(self):
        self.workbook(fake=True)
        self.run_excel()
        self.assertIn("News article metadata and ordering match source", self.g.CRITICAL_FAILED)

    def test_wrong_order_rejected(self):
        self.workbook(reverse=True)
        self.run_excel()
        self.assertIn("News article metadata and ordering match source", self.g.CRITICAL_FAILED)

    def test_wrong_source_metadata_rejected(self):
        self.workbook()
        path = self.root / "Market_News_Digest.xlsx"
        wb = openpyxl.load_workbook(path)
        wb["News Articles"]["C2"] = "Invented Publisher"
        wb.save(path)
        self.run_excel()
        self.assertIn("News article metadata and ordering match source", self.g.CRITICAL_FAILED)

    def test_date_and_word_count_prose_variants(self):
        self.assertEqual(self.g.normalized_date("26.05.2026"), "2026-05-26")
        self.assertTrue(self.g.mentions_count("Восемнадцать статей", 18))
        self.assertTrue(self.g.mentions_count("Всего три статьи", 3))
        self.assertFalse(self.g.mentions_count("Статей 118", 18))

    def test_missing_news_file_rejected(self):
        self.run_excel()
        self.assertIn("News article metadata and ordering match source", self.g.CRITICAL_FAILED)

    def test_teamly_real_per_ticker_summary_passes(self):
        lines = []
        for ticker in self.g.TICKERS:
            pubs = Counter(row[2] for row in self.articles if row[0] == ticker)
            lines.append(f"## {ticker}\nСтатей: {sum(pubs.values())}. Издатель: {pubs.most_common(1)[0][0]}")
        with patch.object(self.g.psycopg2, "connect", return_value=Connection([(4, "Центр рыночной аналитики", "\n".join(lines))])), patch.object(self.g, "get_news_articles", return_value=self.articles), contextlib.redirect_stdout(io.StringIO()):
            self.g.check_teamly()
        self.assertEqual(self.g.CRITICAL_FAILED, [])

    def test_teamly_tickers_alone_rejected(self):
        with patch.object(self.g.psycopg2, "connect", return_value=Connection([(4, "Центр рыночной аналитики", " ".join(self.g.TICKERS))])), patch.object(self.g, "get_news_articles", return_value=self.articles), contextlib.redirect_stdout(io.StringIO()):
            self.g.check_teamly()
        self.assertIn("Teamly per-ticker counts and publishers match source", self.g.CRITICAL_FAILED)

    def test_email_source_summary_passes(self):
        body = f"Статей {self.expected[0]}, топ-издатель {self.expected[2]}; " + " ".join(self.g.TICKERS)
        rows = [("Weekly Market News Digest", '["portfolio-team@investco.com"]', body, "research@investco.com")]
        with patch.object(self.g.psycopg2, "connect", return_value=Connection(rows)), contextlib.redirect_stdout(io.StringIO()):
            self.g.check_email(self.expected)
        self.assertEqual(self.g.CRITICAL_FAILED, [])

    def test_email_tokens_alone_rejected(self):
        rows = [("Weekly Market News Digest", '["portfolio-team@investco.com"]', "18 РБК", "wrong@example.com")]
        with patch.object(self.g.psycopg2, "connect", return_value=Connection(rows)), contextlib.redirect_stdout(io.StringIO()):
            self.g.check_email(self.expected)
        self.assertIn("Email sender and covered tickers match task", self.g.CRITICAL_FAILED)


class CompetitorTests(unittest.TestCase):
    def setUp(self):
        self.g = load("teamly-fetch-competitor")
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def workbook(self, fake=False):
        products, details = self.g.api_fixtures()
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Product Listing"
        sheet.append(["Product", "Price", "Rating", "Features"])
        for p in reversed(products):
            sheet.append([p["name"], p["price"], p["rating"], "FAKE" if fake else "; ".join(reversed(p["features"]))])
        sheet = wb.create_sheet("Detailed Analysis")
        sheet.append(["Product", "Description", "Pros", "Cons", "Target_Market"])
        for p in details:
            values = [p["description"], "; ".join(p["pros"]), "; ".join(p["cons"]), p["target_market"]]
            sheet.append([p["name"], *(["FAKE"] * 4 if fake else values)])
        wb.save(self.root / "Competitor_Analysis.xlsx")

    def test_source_fields_and_reordered_lists_pass(self):
        self.workbook()
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertTrue(self.g.check_excel(self.root))
        self.assertEqual(self.g.FAILED_NAMES, [])

    def test_fake_features_and_details_fail(self):
        self.workbook(fake=True)
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertFalse(self.g.check_excel(self.root))
        self.assertIn("Product features match API source", self.g.FAILED_NAMES)
        self.assertIn("Detailed Analysis content matches API source", self.g.FAILED_NAMES)

    def test_notes_array_with_analysis_facts_passes(self):
        (self.root / "memory").mkdir()
        (self.root / "memory/memory.json").write_text(json.dumps({"notes": ["Конкурентный анализ выполнен: 10 продуктов; Gamma Ultra имеет наивысший рейтинг 4.8."]}, ensure_ascii=False))
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertTrue(self.g.check_memory(self.root))

    def test_alternative_source_observation_in_memory_passes(self):
        (self.root / "memory").mkdir()
        (self.root / "memory/memory.json").write_text(json.dumps({"notes": ["Конкурентный анализ 10 продуктов завершён; Alpha Pro стоит 299."]}, ensure_ascii=False))
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertTrue(self.g.check_memory(self.root))

    def test_unrelated_json_does_not_count_as_memory(self):
        (self.root / "memory").mkdir()
        (self.root / "memory/memory.json").write_text(json.dumps({"unrelated": "competitor" * 20}))
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertFalse(self.g.check_memory(self.root))


class PortfolioTests(unittest.TestCase):
    def setUp(self):
        self.g = load("moex-portfolio-gsheet-pdf-gcal-email")

    def expected(self, latest=True):
        g = self.g
        class SeedCursor:
            def execute(self, sql, params):
                ticker = params[0]
                if "stock_info" in sql:
                    self.row = ({"currentPrice": 100, "longName": ticker, "sector": "Energy"},)
                elif "date <=" in sql:
                    self.row = (100,)
                else:
                    self.row = (150 if ticker == "GAZP.ME" and latest else 100,)
            def fetchone(self):
                return self.row
        return g.compute_expected(SeedCursor())[0]

    def test_current_price_matches_latest_close_not_quote_snapshot(self):
        expected = self.expected()
        self.assertEqual(expected["GAZP.ME"]["current"], 150)
        self.assertAlmostEqual(expected["GAZP.ME"]["alloc"], 150 / 550 * 100)

    def workbook(self, root, expected):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for name, headers, fields in [
            ("Holdings", ["Symbol", "Company", "Sector", "Current_Price", "Shares_Held", "Market_Value", "Allocation_Pct"], ["company", "sector", "current", "shares", "market_value", "alloc"]),
            ("Performance", ["Symbol", "Purchase_Price", "Current_Price", "Return_Pct", "Status"], ["purchase", "current", "ret", "status"]),
            ("Rebalancing", ["Symbol", "Current_Allocation", "Target_Allocation", "Drift_Pct", "Action"], ["alloc", None, "drift", "action"]),
        ]:
            ws = wb.create_sheet(name)
            ws.append(headers)
            for ticker, values in expected.items():
                ws.append([ticker, *[20 if key is None else values[key] for key in fields]])
        wb.save(root / "Portfolio_Dashboard_Reference.xlsx")

    def test_latest_price_dashboard_passes(self):
        with tempfile.TemporaryDirectory() as directory, contextlib.redirect_stdout(io.StringIO()):
            self.workbook(Path(directory), self.expected())
            self.g.check_xlsx(directory, self.expected())
        self.assertEqual(self.g.CRITICAL_FAILS, [])

    def test_stale_quote_dashboard_fails(self):
        with tempfile.TemporaryDirectory() as directory, contextlib.redirect_stdout(io.StringIO()):
            self.workbook(Path(directory), self.expected(latest=False))
            self.g.check_xlsx(directory, self.expected())
        self.assertIn("Portfolio numeric values match source", self.g.CRITICAL_FAILS)

    def test_nonfinite_dashboard_values_fail(self):
        self.assertIsNone(self.g.safe_float("NaN"))
        self.assertIsNone(self.g.safe_float("inf"))

    def test_missing_dashboard_fails(self):
        with tempfile.TemporaryDirectory() as directory, contextlib.redirect_stdout(io.StringIO()):
            self.g.check_xlsx(directory, self.expected())
        self.assertIn("Portfolio dashboard exists with all three sheets", self.g.CRITICAL_FAILS)


class DatabaseErrorTests(unittest.TestCase):
    def test_database_errors_propagate_instead_of_counting_as_model_failure(self):
        # Both connect failures and query failures must reach grade_process,
        # which classifies exceptions separately from a content-check FAIL.
        for task, function, args in [
            ("moex-market-news-teamly", "check_teamly", ()),
            ("moex-market-news-teamly", "check_email", ((18, 6, "РБК"),)),
            ("teamly-fetch-competitor", "check_teamly", ()),
        ]:
            with self.subTest(task=task, function=function):
                g = load(task)
                with patch.object(g.psycopg2, "connect", side_effect=RuntimeError("DB unavailable")), contextlib.redirect_stdout(io.StringIO()):
                    with self.assertRaisesRegex(RuntimeError, "DB unavailable"):
                        getattr(g, function)(*args)
                self.assertEqual(g.FAIL_COUNT, 0)
        g = load("moex-portfolio-gsheet-pdf-gcal-email")
        with patch.object(g, "get_conn", side_effect=RuntimeError("DB unavailable")), contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(RuntimeError, "DB unavailable"):
                g.check_db({}, [])
            with self.assertRaisesRegex(RuntimeError, "DB unavailable"):
                g.run_evaluation("unused", "unused", "unused", None)
        self.assertEqual(g.FAIL_COUNT, 0)
        g = load("moex-market-news-teamly")
        connection = Connection([])
        class BrokenCursor(Cursor):
            def execute(self, sql, params=None):
                raise RuntimeError("query failed")
        with patch.object(g.psycopg2, "connect", return_value=connection), patch.object(connection, "cursor", return_value=BrokenCursor([])), contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(RuntimeError, "query failed"):
                g.check_teamly()
        self.assertEqual(g.FAIL_COUNT, 0)


if __name__ == "__main__":
    unittest.main()
