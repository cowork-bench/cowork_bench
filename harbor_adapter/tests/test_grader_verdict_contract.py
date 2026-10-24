"""Static control-flow contract for all 496 grader verdict lines.

This does not execute the graders and is not a runtime proof of pass rate.
"""

from __future__ import annotations

import ast
import unittest
from pathlib import Path

from harbor_adapter.grader_verdict_contract import (
    PASS_FALSE,
    PASS_TRUE,
    analyze_source,
    analyze_task,
    apply_plan,
    is_applied,
    iter_task_dirs,
)
from utils.evaluation.evaluator import _content_outcome
from utils.evaluation.verdict import parse_verdict_lines

ROOT = Path(__file__).resolve().parents[2]
TASKS = ROOT / "tasks" / "finalpool"


def _call_name(node: ast.AST) -> str:
    func = getattr(node, "func", None)
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name):
        return f"{func.value.id}.{func.attr}"
    return ""


def _parents(tree: ast.AST) -> dict[ast.AST, ast.AST]:
    mapping: dict[ast.AST, ast.AST] = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            mapping[child] = node
    return mapping


def _pass_prints(source: str, tree: ast.AST) -> list[tuple[int, str, bool]]:
    """Return Pass prints as (lineno, statement, inside_except)."""
    parents = _parents(tree)
    found: list[tuple[int, str, bool]] = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and _call_name(node) == "print"):
            continue
        segment = (ast.get_source_segment(source, node) or "").strip()
        if "Pass:" not in segment:
            continue
        inside_except = False
        current: ast.AST | None = node
        while current in parents:
            current = parents[current]
            if isinstance(current, ast.ExceptHandler):
                inside_except = True
        found.append((node.lineno, segment, inside_except))
    return found


def _previous_statement(lines: list[str], lineno: int) -> tuple[int, str]:
    index = lineno - 2
    while index >= 0 and not lines[index].strip():
        index -= 1
    if index < 0:
        return 0, ""
    return index + 1, lines[index].strip()


class GraderVerdictStaticTests(unittest.TestCase):
    def test_all_graders_emit_one_line_per_content_path(self) -> None:
        """Coverage + application: prints live in the PACKAGED copy, not sources.

        The finalpool sources stay in the authors' original form; packaging
        (generate_harbor_canonical.convert_one) applies the plan to
        tests/evaluation/main.py. Here we assert the plan covers all 496
        graders and that applying it produces exactly the planned prints.
        """
        tasks = iter_task_dirs(TASKS)
        self.assertEqual(len(tasks), 496)
        mechanical = manual = 0
        for task_dir in tasks:
            source = (task_dir / "evaluation" / "main.py").read_text(encoding="utf-8")
            plan = analyze_task(task_dir)
            self.assertFalse(plan.unresolved, task_dir.name)
            self.assertFalse(is_applied(plan, source), task_dir.name)
            packaged = apply_plan(plan, source)
            self.assertTrue(is_applied(plan, packaged), task_dir.name)
            actions = {site.action for site in plan.sites}
            self.assertTrue(
                "pass_true" in actions or "conditional" in actions or plan.implicit_pass_lineno,
                task_dir.name,
            )
            self.assertTrue("pass_false" in actions or "conditional" in actions, task_dir.name)
            if plan.change_class == "manual":
                manual += 1
            else:
                mechanical += 1
        self.assertEqual(mechanical + manual, 496)
        self.assertEqual(manual, 17)

    def test_wrapper_does_not_synthesize_pass_line(self) -> None:
        source = (ROOT / "scripts" / "run_eval.py").read_text(encoding="utf-8")
        self.assertNotIn("Pass:", source)


class StrictParserTests(unittest.TestCase):
    def test_unique_pass_and_fail(self) -> None:
        self.assertEqual(parse_verdict_lines(["Pass: True"]).value, "True")
        self.assertEqual(parse_verdict_lines(["Pass: False"]).value, "False")
        passed = _content_outcome("checks\nPass: True\n", "", 0)
        failed = _content_outcome("checks\nPass: False\n", "", 1)
        self.assertTrue(passed["pass"])
        self.assertFalse(failed["pass"])

    def test_crash_duplicate_and_conflict_are_technical(self) -> None:
        crash = _content_outcome(
            "Traceback (most recent call last):\nValueError: boom\n",
            "",
            1,
        )
        self.assertIsNone(crash["pass"])
        missing = _content_outcome("Overall: FAIL\n", "", 1)
        self.assertIsNone(missing["pass"])
        duplicate = _content_outcome("Pass: False\nPass: False\n", "", 1)
        self.assertIsNone(duplicate["pass"])
        conflict = _content_outcome("Pass: True\nPass: False\n", "", 0)
        self.assertIsNone(conflict["pass"])
        mismatch = _content_outcome("Pass: True\n", "", 1)
        self.assertIsNone(mismatch["pass"])


class AgentOutputPathCheckTests(unittest.TestCase):
    """A missing agent deliverable is the agent's failure: Pass: False.

    A missing groundtruth and a missing workspace directory are not the
    agent's failure: those exits stay technical. Without the distinction
    every agent-fail turn became "missing verdict" → technical → three
    wasted retries per task.
    """

    SOURCE = (
        "import os\n"
        "import sys\n"
        "\n"
        "\n"
        "def main(args, gt_dir):\n"
        '    agent_file = os.path.join(args.agent_workspace, "Report.xlsx")\n'
        '    gt_file = os.path.join(gt_dir, "Report.xlsx")\n'
        "    if not os.path.exists(agent_file):\n"
        '        print(f"FAIL: Agent output not found: {agent_file}")\n'
        "        sys.exit(1)\n"
        "    if not os.path.exists(gt_file):\n"
        '        print(f"FAIL: Groundtruth not found: {gt_file}")\n'
        "        sys.exit(1)\n"
        "    if not os.path.isdir(args.agent_workspace):\n"
        '        print("FAIL: workspace missing")\n'
        "        sys.exit(1)\n"
        "    if accuracy < 70.0:\n"
        "        sys.exit(1)\n"
        "    sys.exit(0)\n"
    )

    def test_agent_output_exit_is_content_fail(self) -> None:
        plan = analyze_source("demo", Path("demo/evaluation/main.py"), self.SOURCE)
        self.assertFalse(plan.unresolved)
        agent = next(s for s in plan.sites if s.source == "path_check_agent_output")
        self.assertEqual(agent.action, "pass_false")
        self.assertFalse(agent.manual)
        gt = next(s for s in plan.sites if "exists(gt_file)" in s.source)
        self.assertEqual(gt.action, "technical")
        workspace = next(s for s in plan.sites if "isdir" in s.source)
        self.assertEqual(workspace.action, "technical")

    def test_groundtruth_only_graders_stay_technical(self) -> None:
        source = self.SOURCE.replace(
            '    if not os.path.exists(agent_file):\n'
            '        print(f"FAIL: Agent output not found: {agent_file}")\n'
            "        sys.exit(1)\n",
            "",
        )
        plan = analyze_source("demo", Path("demo/evaluation/main.py"), source)
        self.assertFalse(any(s.action == "pass_false" and s.source == "path_check_agent_output" for s in plan.sites))

    def test_packaged_print_lands_before_agent_exit(self) -> None:
        plan = analyze_source("demo", Path("demo/evaluation/main.py"), self.SOURCE)
        packaged = apply_plan(plan, self.SOURCE)
        self.assertTrue(is_applied(plan, packaged))
        lines = packaged.splitlines()
        fail_at = next(i for i, line in enumerate(lines) if "Agent output not found" in line)
        self.assertIn(PASS_FALSE, lines[fail_at + 1])
        gt_at = next(i for i, line in enumerate(lines) if "Groundtruth not found" in line)
        self.assertNotIn("Pass:", lines[gt_at + 1])


if __name__ == "__main__":
    unittest.main()
