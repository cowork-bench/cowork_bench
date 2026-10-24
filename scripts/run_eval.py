"""Standalone grader entrypoint — engine-independent.

This is the benchmark's integrity-critical half. It runs ONLY the evaluator
(utils.evaluation.evaluator.TaskEvaluator) against the agent's persisted log and
the groundtruth; it imports NO agent engine, so it can grade output produced by
any agent framework.

Usage (inside the eval container, full task tree + groundtruth mounted):
    python3 scripts/run_eval.py --dumps_dir /workspace/dumps

It locates the single traj_log.json the agent phase wrote under dumps_dir and
evaluates it. The evaluator re-resolves groundtruth/evaluation paths from the
log's task_dir against the current (full) tree — see evaluator.evaluate_one.
"""
import argparse
import asyncio
import glob
import os

from utils.evaluation.evaluator import TaskEvaluator


async def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dumps_dir", default="/workspace/dumps",
                        help="Directory the agent phase wrote traj_log.json into.")
    parser.add_argument("--log_file", default=None,
                        help="Explicit traj_log.json path (overrides --dumps_dir search).")
    args = parser.parse_args()

    log_file = args.log_file
    if not log_file:
        matches = sorted(glob.glob(os.path.join(args.dumps_dir, "**", "traj_log.json"),
                                   recursive=True))
        if not matches:
            print(f"[run_eval] No traj_log.json found under {args.dumps_dir}")
            return 2
        if len(matches) > 1:
            print(f"[run_eval] WARNING: multiple traj_log.json found, using first:\n  " +
                  "\n  ".join(matches))
        log_file = matches[0]

    print(f"[run_eval] Evaluating log: {log_file}")
    # TaskEvaluator is the single authority on the verdict: pass is True/False
    # only after the grader's own content decision with a matching exit code,
    # and None on any technical failure (traceback, missing log, crash).
    # This wrapper adds nothing to the decision — it maps the outcome to an
    # exit code and must not synthesize a Pass line.
    eval_res = await TaskEvaluator.evaluate_from_log_file(log_file)
    detail = eval_res.get("details", eval_res.get("failure", "N/A"))
    outcome = eval_res.get("pass")
    grader_rc = eval_res.get("grader_rc")
    if outcome is True:
        print(f"Details: {detail}")
        return 0
    if outcome is False:
        print(f"Details: {detail}")
        return 1
    print("evaluator technical failure")
    print(f"Details: {detail}")
    if isinstance(grader_rc, int) and grader_rc < 0:
        return 128 + abs(grader_rc)
    if isinstance(grader_rc, int) and grader_rc > 1:
        return grader_rc
    return 2


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
