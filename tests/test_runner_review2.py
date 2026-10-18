"""Second-review regressions: real evaluator subprocesses and setup classification."""
import asyncio
import ast
import csv
from enum import Enum
import importlib.util
import json
import os
from pathlib import Path
import shlex
import subprocess
import sys
import tempfile
from types import SimpleNamespace
from typing import Optional
import unittest
from unittest.mock import patch

from utils.evaluation.evaluator import TaskEvaluator
from utils.task_runner.errors import PreprocessError

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("attempt_review2", ROOT / "scripts/read_attempt_result.py")
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


class SecondRunnerReviewTests(unittest.TestCase):
    def test_bad_status_and_nonobject_verdict_are_unscored(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            log = root / "traj_log.json"
            for status in ("unexpected", "interrupted"):
                log.write_text(json.dumps({"config": {"task_dir": "fixture"}, "status": status}))
                self.assertEqual(reader.read_result(root, "fixture"), ("infra_error", "null"))
            log.write_text(json.dumps({"config": {"task_dir": "fixture"}, "status": "success"}))
            (root / "eval_res.json").write_text("[]")
            self.assertEqual(reader.read_result(root, "fixture"), ("evaluation_error", "null"))
            (root / "eval_res.json").write_text('{"pass":true}')
            for exit_code in (1, 125, 127, 137):
                self.assertEqual(reader.read_result(root, "fixture", agent_exit=exit_code),
                                 ("infra_error", "null"))
            self.assertEqual(reader.read_result(root, "fixture", timed_out=True, agent_exit=124),
                             ("wall_timeout", "null"))

    def test_calculation_timeout_remains_an_evaluator_error(self):
        import openpyxl
        from utils.evaluation.spreadsheets import load_calculated_workbook
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "formula.xlsx"
            workbook = openpyxl.Workbook()
            workbook.active["A1"] = "=1+1"
            workbook.save(path)
            for error in (subprocess.TimeoutExpired("soffice", 60), OSError("engine could not start")):
                with self.subTest(error=type(error).__name__), \
                     patch("utils.evaluation.spreadsheets.shutil.which", return_value="soffice"), \
                     patch("utils.evaluation.spreadsheets.subprocess.run", side_effect=error):
                    with self.assertRaisesRegex(RuntimeError, "calculation process failed"):
                        load_calculated_workbook(path)

    def test_partial_summary_cannot_publish_full_benchmark_percentage(self):
        source = (ROOT / "run_parallel.sh").read_text().split("<< 'PYEOF'\n", 1)[1].split("\nPYEOF", 1)[0]
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary) / "run"
            directory.mkdir()
            (directory / "run_meta.json").write_text('{"total_tasks":2}')
            summary = directory / "summary.csv"
            summary.write_text("task,status,eval_pass,duration_s\nfixture,success,True,1\n")
            result = subprocess.run([sys.executable, "-c", source, str(summary), str(directory)],
                                    text=True, capture_output=True, check=True)
            self.assertIn("Coverage: 1/2", result.stdout)
            self.assertIn("Benchmark score: unavailable", result.stdout)
            with (directory.parent / "index.csv").open() as stream:
                self.assertEqual(next(csv.DictReader(stream))["pass_pct"], "")

    def test_strands_preprocess_failure_is_infrastructure(self):
        # Execute the actual methods without importing a model SDK or calling it.
        tree = ast.parse((ROOT / "strands_runner/agent.py").read_text())
        keep = {"_find_cause", "_classify_error", "TaskStatus"}
        nodes = [node for node in tree.body if getattr(node, "name", None) in keep]
        cls = next(node for node in tree.body if getattr(node, "name", None) == "StrandsTaskAgent")
        nodes.append(next(node for node in cls.body if getattr(node, "name", None) == "_run_preprocess"))
        namespace = {"Optional": Optional, "Enum": Enum, "subprocess": subprocess,
                     "PreprocessError": PreprocessError, "_INFRA_ERROR_MARKERS": (),
                     "print_color": lambda *args: None}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), "strands_runner/agent.py", "exec"), namespace)
        config = SimpleNamespace(initialization=SimpleNamespace(
            process_command=shlex.join([sys.executable, "-c", "raise SystemExit(1)"])),
            agent_workspace="/tmp/fixture", launch_time="2026-10-18 12:00:00")
        with self.assertRaises(PreprocessError) as caught:
            namespace["_run_preprocess"](SimpleNamespace(task_config=config, debug=False))
        self.assertEqual(namespace["_classify_error"](caught.exception).value, "infra_error")

    def test_native_evaluator_pipeline_preserves_verdict_exception_and_log(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "workspace").mkdir()
            log = root / "traj_log.json"
            original = json.dumps({"config": {"task_dir": "fixture", "launch_time": "2026-10-18 12:00:00"},
                                   "status": "success"})
            log.write_text(original)
            module = root / "fixture_grader.py"
            command = shlex.join([sys.executable, "-m", "utils.evaluation.grade_process",
                                  "--module", "fixture_grader"])
            evaluation = SimpleNamespace(evaluation_command=command, groundtruth_workspace=str(root))
            for source, expected in [("raise SystemExit(0)", True), ("raise SystemExit(1)", False),
                                     ("raise ValueError('broken grader fixture')", None)]:
                with self.subTest(source=source):
                    module.write_text(source)
                    with patch.dict(os.environ, {"PYTHONPATH": os.pathsep.join([str(ROOT), str(root)]), "PYTHONDONTWRITEBYTECODE": "1"}), \
                         patch("utils.evaluation.evaluator.Evaluation.build", return_value=evaluation):
                        result = asyncio.run(TaskEvaluator.evaluate_from_log_file(str(log), expected_task="fixture"))
                    self.assertIs(result["pass"], expected)
                    self.assertIs(json.loads((root / "eval_res.json").read_text())["pass"], expected)
                    self.assertEqual(log.read_text(), original)
                    if expected is None:
                        self.assertEqual(result["returncode"], 2)
                        self.assertIn("broken grader fixture", result["stderr"])


if __name__ == "__main__":
    unittest.main()
