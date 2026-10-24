"""Flat sibling imports must resolve under the ``-m`` grading path.

Canonical grading runs ``python -m tasks.finalpool.<task>.evaluation.main``
from /workspace, so the grader's own directory is never on sys.path. The
harbor port flattened two relative-import graders (canvas-exam-calendar-report,
canvas-pdf-grade-gsheet) without the authors' sys.path shim, which broke the
-m path with ModuleNotFoundError. Packaging must restore the shim.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from harbor_adapter.generate_harbor_canonical import (
    _ensure_sibling_import_shim,
    _first_flat_sibling_import,
)

ROOT = Path(__file__).resolve().parents[2]
TASKS = ROOT / "tasks" / "finalpool"

FLAT_MAIN = (
    "import argparse\n"
    "import json\n"
    "\n"
    "from check_demo import helper\n"
    "print('ok', helper())\n"
)
SIBLING = "def helper():\n    return 42\n"


def _make_task(root: Path, main_src: str) -> Path:
    evaluation = root / "tasks" / "finalpool" / "demo" / "evaluation"
    evaluation.mkdir(parents=True)
    (evaluation / "main.py").write_text(main_src, encoding="utf-8")
    (evaluation / "check_demo.py").write_text(SIBLING, encoding="utf-8")
    return evaluation / "main.py"


def _run_m(root: Path) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "tasks.finalpool.demo.evaluation.main"],
        cwd=root,
        capture_output=True,
        text=True,
        timeout=30,
    )


class SiblingImportShimTests(unittest.TestCase):
    def test_flat_import_resolves_under_m_after_shim(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main_py = _make_task(root, FLAT_MAIN)
            before = _run_m(root)
            self.assertIn("ModuleNotFoundError", before.stderr)
            self.assertTrue(_ensure_sibling_import_shim(main_py))
            after = _run_m(root)
            self.assertEqual(after.returncode, 0, after.stderr)
            self.assertIn("ok 42", after.stdout)
            # идемпотентность: второй проход ничего не вставляет
            self.assertFalse(_ensure_sibling_import_shim(main_py))

    def test_relative_imports_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            main_py = _make_task(root=Path(tmp), main_src="from .check_demo import helper\nprint('ok', helper())\n")
            self.assertFalse(_ensure_sibling_import_shim(main_py))
            self.assertNotIn("sys.path.insert", main_py.read_text(encoding="utf-8"))

    def test_authors_shim_already_present(self) -> None:
        source = (
            "import os, sys\n"
            "sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))\n"
            "from check_demo import helper\n"
            "print('ok', helper())\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            main_py = _make_task(root=Path(tmp), main_src=source)
            self.assertFalse(_ensure_sibling_import_shim(main_py))

    def test_non_sibling_import_untouched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            main_py = _make_task(root=Path(tmp), main_src="import json\nprint(json.dumps({'a': 1}))\n")
            self.assertFalse(_ensure_sibling_import_shim(main_py))

    def test_real_tree_detection(self) -> None:
        """Только два переписанных портом грейдера требуют shim."""
        need = [
            d.name
            for d in sorted(TASKS.iterdir())
            if (d / "evaluation").is_dir()
            and _first_flat_sibling_import(d / "evaluation" / "main.py") is not None
        ]
        self.assertEqual(
            need,
            ["canvas-exam-calendar-report", "canvas-pdf-grade-gsheet"],
        )


if __name__ == "__main__":
    unittest.main()
