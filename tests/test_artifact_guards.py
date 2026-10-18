"""Regression tests for fail-closed validation of mandatory artifact formats."""
import ast
import builtins
import contextlib
import importlib.util
import io
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

from utils.evaluation.artifact_guards import REQUIRED_FILES, require_files

ROOT = Path(__file__).resolve().parents[1]


class ArtifactGuardsTests(unittest.TestCase):
    def test_all_reported_file_tasks_reject_absence_and_corruption(self):
        with tempfile.TemporaryDirectory() as directory:
            for task, names in REQUIRED_FILES.items():
                with self.subTest(task=task), contextlib.redirect_stdout(io.StringIO()):
                    with self.assertRaises(SystemExit):
                        require_files(directory, names)
                    for name in names:
                        (Path(directory) / name).write_bytes(b"not an office document")
                    with self.assertRaises(SystemExit):
                        require_files(directory, names)
                    for name in names:
                        (Path(directory) / name).unlink()

    def test_valid_office_formats_pass(self):
        import openpyxl
        from docx import Document
        from pptx import Presentation
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            book = openpyxl.Workbook()
            book.active.append(["valid data", 1])
            book.save(root / "valid.xlsx")
            doc = Document()
            doc.add_paragraph("valid data")
            doc.save(root / "valid.docx")
            slides = Presentation()
            slides.slides.add_slide(slides.slide_layouts[0])
            slides.save(root / "valid.pptx")
            require_files(root, ("valid.xlsx", "valid.docx", "valid.pptx"))

    def test_every_manifest_is_wired_into_direct_entrypoint(self):
        # Direct task CLIs cannot depend on the framework invoking a preflight.
        from utils.evaluation.artifact_guards import REQUIRED_GSHEETS
        for task in REQUIRED_FILES.keys() | REQUIRED_GSHEETS.keys():
            path = ROOT / "tasks/finalpool" / task / "evaluation/main.py"
            tree = ast.parse(path.read_text())
            main = next(node for node in tree.body
                        if isinstance(node, ast.FunctionDef) and node.name == "main")
            self.assertTrue(any(isinstance(node, ast.Call)
                                and isinstance(node.func, ast.Name)
                                and node.func.id == "require_task_artifacts"
                                for node in ast.walk(main)), task)

    def test_missing_parser_dependency_is_not_a_model_failure(self):
        original = builtins.__import__
        def missing(name, *args, **kwargs):
            if name == "openpyxl":
                raise ModuleNotFoundError("openpyxl unavailable")
            return original(name, *args, **kwargs)
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / "report.xlsx").write_bytes(b"placeholder")
            with patch.object(builtins, "__import__", missing), self.assertRaises(ModuleNotFoundError):
                require_files(directory, ("report.xlsx",))

    def test_shared_sheet_database_errors_are_not_model_failures(self):
        from utils.evaluation.artifact_guards import require_gsheet
        def unavailable(**kwargs):
            raise RuntimeError("database unavailable")
        with patch.dict(sys.modules, {"psycopg2": types.SimpleNamespace(connect=unavailable)}):
            with self.assertRaisesRegex(RuntimeError, "database unavailable"):
                require_gsheet("Required report", ())

    def test_every_required_shared_sheet_must_exist_and_be_populated(self):
        from utils.evaluation.artifact_guards import REQUIRED_GSHEETS, require_gsheet
        class Cursor:
            def execute(self, sql, params=None):
                self.sql = sql
            def fetchall(self):
                if "spreadsheets" in self.sql:
                    return [(1,)] if self.present else []
                return [(index, tab) for index, tab in enumerate(self.tabs or ("Sheet1",))]
            def fetchone(self):
                return (2 if self.populated else 0,)
            def close(self):
                pass
        for title, tabs in REQUIRED_GSHEETS.values():
            cursor = Cursor()
            cursor.tabs, cursor.present, cursor.populated = tabs, True, True
            connection = types.SimpleNamespace(cursor=lambda: cursor, close=lambda: None)
            with self.subTest(title=title), patch.dict(sys.modules, {
                    "psycopg2": types.SimpleNamespace(connect=lambda **kwargs: connection)}), \
                    contextlib.redirect_stdout(io.StringIO()):
                require_gsheet(title, tabs)
                cursor.populated = False
                with self.assertRaises(SystemExit):
                    require_gsheet(title, tabs)
                cursor.present = False
                with self.assertRaises(SystemExit):
                    require_gsheet(title, tabs)

    def test_tracker_without_excel_fails_before_secondary_checks(self):
        path = ROOT / "tasks/finalpool/arxiv-method-benchmark-tracker/evaluation/main.py"
        spec = importlib.util.spec_from_file_location("tracker_eval", path)
        module = importlib.util.module_from_spec(spec)
        with patch.dict(sys.modules, {"psycopg2": types.ModuleType("psycopg2")}):
            spec.loader.exec_module(module)
        with tempfile.TemporaryDirectory() as directory, patch.object(
                sys, "argv", [str(path), "--agent_workspace", directory]), patch.object(
                module, "check_teamly") as secondary, contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaises(SystemExit) as exc:
                module.main()
            self.assertEqual(exc.exception.code, 1)
            secondary.assert_not_called()


if __name__ == "__main__":
    unittest.main()
