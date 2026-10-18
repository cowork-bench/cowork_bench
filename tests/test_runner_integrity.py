"""Model-free regression tests for verdict classification and isolated attempts."""
import asyncio
import ast
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from utils.evaluation.evaluator import TaskEvaluator, confined_workspace, write_result

spec = importlib.util.spec_from_file_location("attempt_result", ROOT / "scripts/read_attempt_result.py")
attempt_result = importlib.util.module_from_spec(spec)
spec.loader.exec_module(attempt_result)


class GraderProcessTests(unittest.TestCase):
    def run_grader(self, source):
        with tempfile.TemporaryDirectory(prefix="grader space ") as temp:
            path = Path(temp)
            (path / "probe_grader.py").write_text(source)
            env = dict(os.environ, PYTHONPATH=os.pathsep.join([temp, str(ROOT)]))
            return subprocess.run([sys.executable, "-m", "utils.evaluation.grade_process",
                                   "--module", "probe_grader"], cwd=ROOT, env=env,
                                  text=True, capture_output=True)

    def test_pass_and_explicit_failed_check(self):
        self.assertEqual(self.run_grader("raise SystemExit(0)").returncode, 0)
        self.assertEqual(self.run_grader("print('missing output'); raise SystemExit(1)").returncode, 1)

    def test_missing_dependency_is_not_model_failure(self):
        result = self.run_grader("import nonexistent_cowork_dependency")
        self.assertEqual(result.returncode, 2)
        self.assertIn("ModuleNotFoundError", result.stderr)

    def test_runtime_exception_is_not_model_failure(self):
        result = self.run_grader("raise ValueError('broken evaluator')")
        self.assertEqual(result.returncode, 2)
        self.assertIn("broken evaluator", result.stderr)

    def test_non_verdict_exit(self):
        self.assertEqual(self.run_grader("raise SystemExit('setup error')").returncode, 2)

    def test_evaluator_preserves_error_and_ignores_log_command(self):
        config = {"log_file": "/tmp/probe/traj_log.json", "agent_workspace": "/tmp/probe/workspace",
                  "evaluation": {"evaluation_command": "UNTRUSTED"},
                  "task_dir": "fixture", "launch_time": "2026-10-18 12:00:00 Sunday"}
        fresh = SimpleNamespace(evaluation_command="trusted", groundtruth_workspace="fixture/gold")
        for code, expected in [(0, True), (1, False), (2, None), (137, None)]:
            with self.subTest(code=code), patch("utils.evaluation.evaluator.Evaluation.build", return_value=fresh), \
                 patch("utils.evaluation.evaluator.run_command", return_value=("checks", "diagnostic", code)) as run:
                result = asyncio.run(TaskEvaluator.evaluate_one({"config": config, "status": "success"}))
            self.assertIs(result["pass"], expected)
            self.assertTrue(run.call_args.args[0].startswith("trusted "))
            if code:
                self.assertEqual(result["stderr"], "diagnostic")
                self.assertEqual(result["returncode"], code)

    def test_serialized_task_root_cannot_delete_another_attempt(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            victim = root / "outside" / "eval_res.json"
            victim.parent.mkdir()
            victim.write_text('{"pass":true}')
            attempt = root / "attempt"
            attempt.mkdir()
            (attempt / "workspace").mkdir()
            record = {"status": "success", "config": {
                "task_dir": "fixture", "task_root": str(victim.parent),
                "global_task_config": {"dump_path": str(victim.parent)},
                "launch_time": "2026-10-18 12:00:00"}}
            log = attempt / "traj_log.json"
            log.write_text(json.dumps(record))
            fresh = SimpleNamespace(evaluation_command="trusted", groundtruth_workspace="gold")
            with patch("utils.evaluation.evaluator.Evaluation.build", return_value=fresh), \
                    patch("utils.evaluation.evaluator.run_command", return_value=("", "", 0)):
                result = asyncio.run(TaskEvaluator.evaluate_from_log_file(str(log), expected_task="fixture"))
            self.assertTrue(result["pass"])
            self.assertEqual(victim.read_text(), '{"pass":true}')

    def test_preprocess_nonzero_raises_before_agent(self):
        tree = ast.parse((ROOT / "utils/roles/task_agent.py").read_text())
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "TaskAgent")
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == "_run_preprocess")
        namespace = {"subprocess": subprocess, "print_color": lambda *a: None, "PreprocessError": RuntimeError}
        exec(compile(ast.Module(body=[method], type_ignores=[]), "task_agent.py", "exec"), namespace)
        task = SimpleNamespace(initialization=SimpleNamespace(process_command=f"{sys.executable} -c 'raise SystemExit(1)'"),
                               agent_workspace="/tmp/workspace", launch_time="2026-10-18 12:00:00")
        with self.assertRaisesRegex(RuntimeError, "preprocess exited 1"):
            namespace["_run_preprocess"](SimpleNamespace(task_config=task, debug=False))


class AttemptResultTests(unittest.TestCase):
    def test_dangling_answer_key_link_rejected_but_internal_link_allowed(self):
        with tempfile.TemporaryDirectory() as temp:
            work = Path(temp) / "workspace"
            work.mkdir()
            (work / "answer.xlsx").symlink_to("/workspace/tasks/finalpool/task/groundtruth_workspace/answer.xlsx")
            self.assertFalse(confined_workspace(work))
            (work / "answer.xlsx").unlink()
            (work / "data.txt").write_text("agent output")
            (work / "alias.txt").symlink_to("data.txt")
            self.assertTrue(confined_workspace(work))
            alias = Path(temp) / "alias"
            alias.symlink_to(work, target_is_directory=True)
            self.assertFalse(confined_workspace(alias))

    def test_verdict_replaces_agent_link_without_modifying_target(self):
        with tempfile.TemporaryDirectory() as temp:
            base = Path(temp)
            victim = base / "untouched.json"
            victim.write_text('{"pass":true}')
            verdict = base / "eval_res.json"
            verdict.symlink_to(victim)
            write_result({"pass": False}, verdict)
            self.assertEqual(victim.read_text(), '{"pass":true}')
            self.assertFalse(verdict.is_symlink())
            self.assertFalse(json.loads(verdict.read_text())["pass"])

    def test_outer_budget_timeout_remains_a_budget_failure_without_a_log(self):
        with tempfile.TemporaryDirectory() as temp:
            self.assertEqual(attempt_result.read_result(temp, "a", timed_out=True), ("wall_timeout", "null"))

    def test_no_cross_attempt_search_or_text_fallback(self):
        with tempfile.TemporaryDirectory() as temp:
            base = Path(temp)
            old, current = base / "old", base / "current"
            old.mkdir(); current.mkdir()
            (old / "traj_log.json").write_text(json.dumps({"status": "success", "config": {"task_dir": "a"}}))
            (old / "eval_res.json").write_text('{"pass":true}')
            (current / "run.log").write_text("Pass: True\n====== Status: success ======\n")
            self.assertEqual(attempt_result.read_result(current, "a"), ("evaluation_error", "null"))
            (current / "traj_log.json").write_text((old / "traj_log.json").read_text())
            (current / "eval_res.json").write_text('{"pass":null,"status":"evaluation_error"}')
            self.assertEqual(attempt_result.read_result(current, "a"), ("evaluation_error", "null"))
            (current / "eval_res.json").write_text('{"pass":false,"status":"graded"}')
            self.assertEqual(attempt_result.read_result(current, "a"), ("success", "False"))
            self.assertEqual(attempt_result.read_result(current, "another-task"), ("evaluation_error", "null"))

    def test_ambiguous_attempt_is_unscored(self):
        with tempfile.TemporaryDirectory() as temp:
            for name in ("a", "b"):
                path = Path(temp) / name
                path.mkdir()
                (path / "traj_log.json").write_text("{}")
            self.assertEqual(attempt_result.read_result(temp, "a"), ("evaluation_error", "null"))


if __name__ == "__main__":
    unittest.main()
