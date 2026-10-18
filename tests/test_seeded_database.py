"""Opt-in integration checks against a fresh db/init.sql.gz PostgreSQL instance.

Set COWORK_TEST_PG=1 only for a disposable seeded database. Tests insert and remove
uniquely named spreadsheet fixtures. No model or mock SQL cursor is involved.
"""
import importlib.util
import os
from pathlib import Path
import unittest
import uuid
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]


def load_module(relative, name):
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@unittest.skipUnless(os.environ.get("COWORK_TEST_PG") == "1", "requires disposable seeded PostgreSQL")
class SeededDatabaseTests(unittest.TestCase):
    def connect(self):
        import psycopg2
        return psycopg2.connect(host=os.environ.get("PGHOST", "localhost"),
                                port=os.environ.get("PGPORT", "5432"),
                                dbname=os.environ.get("PGDATABASE", "cowork_gym"),
                                user="eigent", password="camel")

    def test_moex_expected_current_prices_match_actual_mcp_adapter(self):
        grader = load_module("tasks/finalpool/moex-portfolio-gsheet-pdf-gcal-email/evaluation/main.py",
                             "portfolio_seeded")
        adapter = load_module("local_servers/moex-finance-mcp/pg_adapter.py", "moex_seeded")
        connection = self.connect()
        self.addCleanup(connection.close)
        with connection.cursor() as cursor:
            expected, _ = grader.compute_expected(cursor)
        # The actual MCP launchers bridge libpq names to the adapter's PG_* names.
        # Mirror only that environment setup; both sides run real database queries.
        mcp_env = {"PG_HOST": os.environ.get("PGHOST", "localhost"),
                   "PG_PORT": os.environ.get("PGPORT", "5432"),
                   "PG_DATABASE": os.environ.get("PGDATABASE", "cowork_gym"),
                   "PG_USER": "eigent", "PG_PASSWORD": "camel"}
        with patch.dict(os.environ, mcp_env):
            for ticker, row in expected.items():
                with self.subTest(ticker=ticker):
                    self.assertIsNotNone(row["current"])
                    self.assertEqual(row["current"], adapter.PgTicker(ticker).info["currentPrice"])
        self.assertAlmostEqual(sum(row["alloc"] for row in expected.values()), 100.0)

    def test_hr_selects_five_unique_eligible_people_in_contract_order(self):
        grader = load_module("tasks/finalpool/rzd-hr1c-training-trip-kazan-strict-excel-email-gcal/evaluation/main.py",
                             "hr_seeded")
        connection = self.connect()
        self.addCleanup(connection.close)
        with connection.cursor() as cursor:
            cursor.execute('SELECT "EMPLOYEE_ID", "DEPARTMENT", "YEARS_EXPERIENCE" '
                           'FROM hr1c_data."HR__PUBLIC__EMPLOYEES"')
            employees = cursor.fetchall()
        eligible = [row for row in employees if row[1] in ("Продажи", "Маркетинг") and row[2] >= 3]
        reference = sorted(eligible, key=lambda row: (-row[2], row[0]))[:5]
        selected = grader.fetch_eligible_employees()
        self.assertEqual([row[0] for row in selected], [row[0] for row in reference])
        self.assertEqual(len({row[0] for row in selected}), 5)
        self.assertEqual([row[0] for row in selected], [10005, 10009, 10012, 10003, 10018])

    def test_canvas_grade_oracle_matches_seeded_submissions(self):
        import openpyxl
        connection = self.connect()
        self.addCleanup(connection.close)
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT c.name, COUNT(DISTINCT s.user_id), ROUND(AVG(s.score), 2),
                       MAX(s.score), MIN(s.score)
                FROM canvas.submissions s
                JOIN canvas.assignments a ON a.id = s.assignment_id
                JOIN canvas.courses c ON c.id = a.course_id
                WHERE s.score IS NOT NULL
                GROUP BY c.id, c.name ORDER BY AVG(s.score) DESC
            """)
            rows = cursor.fetchall()
        workbook = openpyxl.load_workbook(ROOT / "tasks/finalpool/canvas-grade-summary/groundtruth_workspace/Canvas_Grade_Summary.xlsx",
                                          data_only=True)
        self.addCleanup(workbook.close)
        expected = list(workbook["Grade Summary"].values)[1:]
        self.assertEqual(len(rows), len(expected))
        expected_by_course = {row[0]: row for row in expected}
        for row in rows:
            with self.subTest(course=row[0]):
                self.assertEqual(row[:2], expected_by_course[row[0]][:2])
                for actual, wanted in zip(row[2:], expected_by_course[row[0]][2:]):
                    self.assertAlmostEqual(float(actual), wanted, places=2)

    def test_required_shared_sheet_accepts_blank_default_tab_but_rejects_absence(self):
        from utils.evaluation.artifact_guards import require_gsheet
        identity = "cowork-regression-" + uuid.uuid4().hex
        connection = self.connect()
        self.addCleanup(connection.close)
        try:
            with connection.cursor() as cursor:
                cursor.execute("INSERT INTO gsheet.spreadsheets (id, title) VALUES (%s, %s)",
                               (identity, identity))
                cursor.executemany("INSERT INTO gsheet.sheets (spreadsheet_id, id, title) VALUES (%s, %s, %s)",
                                   [(identity, 0, "Sheet1"), (identity, 1, "Summary")])
                cursor.executemany("INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) "
                                   "VALUES (%s, 1, %s, 0, %s)",
                                   [(identity, 0, "Total_Courses"), (identity, 1, "22")])
            connection.commit()
            require_gsheet(identity, ())
            require_gsheet(identity, ("Summary",))
            with self.assertRaises(SystemExit):
                require_gsheet(identity, ("Missing",))
        finally:
            connection.rollback()
            with connection.cursor() as cursor:
                cursor.execute("DELETE FROM gsheet.cells WHERE spreadsheet_id = %s", (identity,))
                cursor.execute("DELETE FROM gsheet.sheets WHERE spreadsheet_id = %s", (identity,))
                cursor.execute("DELETE FROM gsheet.spreadsheets WHERE id = %s", (identity,))
            connection.commit()
        with self.assertRaises(SystemExit):
            require_gsheet(identity, ())


if __name__ == "__main__":
    unittest.main()
