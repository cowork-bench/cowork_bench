"""Real evaluator regressions; only PostgreSQL query results are adapted."""
import contextlib
import importlib.util
import io
from pathlib import Path
import shutil
import sys
import tempfile
import types
import unittest
from zipfile import ZipFile
from unittest.mock import patch

import openpyxl
from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
POOL = ROOT / "tasks/finalpool"


def evaluator(task):
    path = POOL / task / "evaluation/main.py"
    spec = importlib.util.spec_from_file_location("canvas_eval", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class QueryCursor:
    """PostgreSQL adapter representing an actual populated matching shared sheet."""
    def __init__(self, rows, present):
        self.rows, self.present, self.result = rows, present, []

    def execute(self, sql, params=None):
        if "COUNT(DISTINCT" in sql:
            self.result = [(len(self.rows),)]
        elif "col_index" in sql:
            self.result = [(ri, ci, value) for ri, row in enumerate(self.rows)
                           for ci, value in enumerate(row) if value is not None]
        elif "SELECT id, title" in sql:
            self.result = [(10, "Summary")] if self.present else []
        else:
            self.result = [(1,)] if self.present else []

    def fetchall(self):
        return self.result

    def fetchone(self):
        return self.result[0] if self.result else None

    def close(self):
        pass


class CanvasIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.workspace = Path(self.temp.name)
        self.addCleanup(self.temp.cleanup)

    def prepare(self, task):
        for source in (POOL / task / "groundtruth_workspace").iterdir():
            if source.suffix in {".xlsx", ".pptx"}:
                shutil.copyfile(source, self.workspace / source.name)

    def run_evaluator(self, task, present=True):
        if task == "canvas-enrollment-ppt-gsheet":
            book = openpyxl.load_workbook(POOL / task / "groundtruth_workspace/Enrollment_Overview.xlsx")
            rows = list(book["Enrollment Details"].values)
        else:
            rows = [("Metric", "Value"), ("Total_Courses", 22),
                    ("Overall_Avg_Score", 68.0)]
        cursor = QueryCursor(rows, present)
        connection = types.SimpleNamespace(cursor=lambda: cursor, close=lambda: None)
        database = types.SimpleNamespace(connect=lambda **kwargs: connection)
        module = evaluator(task)
        argv = ["evaluation/main.py", "--agent_workspace", str(self.workspace),
                "--groundtruth_workspace", str(POOL / task / "groundtruth_workspace")]
        with patch.dict(sys.modules, {"psycopg2": database}), patch.object(sys, "argv", argv), \
                contextlib.redirect_stdout(io.StringIO()):
            try:
                module.main()
            except SystemExit as exc:
                return exc.code
        return 0

    def test_grade_gold_and_headerless_summary_pass(self):
        task = "canvas-grade-summary"
        self.prepare(task)
        self.assertEqual(self.run_evaluator(task), 0)
        path = self.workspace / "Canvas_Grade_Summary.xlsx"
        book = openpyxl.load_workbook(path)
        book["Summary"].delete_rows(1)
        book.save(path)
        self.assertEqual(self.run_evaluator(task), 0)

    def test_missing_grade_shared_sheet_fails(self):
        task = "canvas-grade-summary"
        self.prepare(task)
        self.assertEqual(self.run_evaluator(task, present=False), 1)

    def test_enrollment_gold_passes_and_missing_shared_sheet_fails(self):
        task = "canvas-enrollment-ppt-gsheet"
        self.prepare(task)
        self.assertEqual(self.run_evaluator(task), 0)
        self.assertEqual(self.run_evaluator(task, present=False), 1)

    def test_enrollment_percentages_without_raw_counts_pass(self):
        task = "canvas-enrollment-ppt-gsheet"
        self.prepare(task)
        path = self.workspace / "Enrollment_Overview.pptx"
        presentation = Presentation(path)
        for shape in presentation.slides[3].shapes:
            if shape.has_text_frame and "Enrollment Distribution" not in shape.text:
                shape.text = "Students: 99.8%\nTeachers: 0.1%\nTAs: 0.1%\nActive: 69.2%\nCompleted: 30.8%"
        presentation.save(path)
        self.assertEqual(self.run_evaluator(task), 0)

    def test_enrollment_missing_required_columns_fails(self):
        task = "canvas-enrollment-ppt-gsheet"
        self.prepare(task)
        path = self.workspace / "Enrollment_Overview.xlsx"
        book = openpyxl.load_workbook(path)
        book["Enrollment Details"].delete_cols(3, 6)
        book.save(path)
        self.assertEqual(self.run_evaluator(task), 1)

    def test_enrollment_empty_slide_content_fails(self):
        task = "canvas-enrollment-ppt-gsheet"
        self.prepare(task)
        presentation = Presentation()
        for title in ["Course Enrollment Overview", "Enrollment Summary",
                      "Top 5 Courses by Enrollment", "Enrollment Distribution"]:
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1)).text = title
        presentation.save(self.workspace / "Enrollment_Overview.pptx")
        self.assertEqual(self.run_evaluator(task), 1)

    def test_external_workbook_reference_never_starts_calculation(self):
        from utils.evaluation.spreadsheets import load_calculated_workbook, UnsupportedExternalWorkbook
        path = self.workspace / "external.xlsx"
        book = openpyxl.Workbook()
        book.active["A1"] = "='[1]SecretSheet'!A1"
        book.save(path)
        with patch("utils.evaluation.spreadsheets.subprocess.run") as process:
            with self.assertRaises(UnsupportedExternalWorkbook):
                load_calculated_workbook(path)
            process.assert_not_called()

    def test_external_link_metadata_and_vba_never_start_calculation(self):
        from utils.evaluation.spreadsheets import load_calculated_workbook, UnsupportedExternalWorkbook
        for part in ("xl/externalLinks/externalLink1.xml", "xl/vbaProject.bin"):
            with self.subTest(part=part):
                path = self.workspace / "external.xlsx"
                book = openpyxl.Workbook()
                book.active["A1"] = "=1+1"
                book.save(path)
                with ZipFile(path, "a") as archive:
                    archive.writestr(part, b"untrusted embedded content")
                with patch("utils.evaluation.spreadsheets.subprocess.run") as process:
                    with self.assertRaises(UnsupportedExternalWorkbook):
                        load_calculated_workbook(path)
                    process.assert_not_called()

    def test_external_data_function_never_starts_calculation(self):
        from utils.evaluation.spreadsheets import load_calculated_workbook, UnsupportedExternalWorkbook
        path = self.workspace / "external.xlsx"
        book = openpyxl.Workbook()
        book.active["A1"] = '=WEBSERVICE("file:///grader/answer-key.xml")'
        book.save(path)
        with patch("utils.evaluation.spreadsheets.subprocess.run") as process:
            with self.assertRaises(UnsupportedExternalWorkbook):
                load_calculated_workbook(path)
            process.assert_not_called()

    @unittest.skipUnless(shutil.which("soffice") or shutil.which("libreoffice") or
                         Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").is_file(),
                         "requires real LibreOffice Calc")
    def test_native_formula_recalculates_without_changing_agent_file(self):
        task = "canvas-grade-summary"
        self.prepare(task)
        path = self.workspace / "Canvas_Grade_Summary.xlsx"
        book = openpyxl.load_workbook(path)
        formula = "=COUNTA('Grade Summary'!A2:A23)"
        book["Summary"]["B2"] = formula
        book.save(path)
        self.assertIsNone(openpyxl.load_workbook(path, data_only=True)["Summary"]["B2"].value)
        self.assertEqual(self.run_evaluator(task), 0)
        self.assertEqual(openpyxl.load_workbook(path)["Summary"]["B2"].value, formula)
        self.assertIsNone(openpyxl.load_workbook(path, data_only=True)["Summary"]["B2"].value)


if __name__ == "__main__":
    unittest.main()
