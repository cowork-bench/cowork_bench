from typing import Dict, Any, List, Optional
from utils.data_structures.task_config import Evaluation
from utils.general.helper import run_command, read_json
import logging
import os
import shlex
import json
import tempfile
from pathlib import Path


def confined_workspace(workspace):
    """Reject links that become answer-key access when the eval tree is mounted."""
    workspace = Path(workspace).absolute()
    if workspace.is_symlink():
        return False
    workspace = workspace.resolve()
    try:
        for directory, directories, files in os.walk(workspace, followlinks=False):
            for name in directories + files:
                path = Path(directory) / name
                if path.is_symlink() and not path.resolve().is_relative_to(workspace):
                    return False
    except (OSError, RuntimeError):
        return False
    return True

def write_result(result, destination):
    """Replace an agent-writable verdict without following symlinks/hardlinks."""
    destination = Path(destination)
    with tempfile.NamedTemporaryFile(mode="w", dir=destination.parent, delete=False) as f:
        temporary = Path(f.name)
        try:
            json.dump(result, f, ensure_ascii=False, indent=2)
            f.close()
            os.replace(temporary, destination)
        finally:
            temporary.unlink(missing_ok=True)


class TaskEvaluator:
    """Task evaluator"""
    
    @staticmethod
    async def evaluate_one(dump_line: Dict[str, Any]) -> Dict[str, Any]:
        """
        Single task evaluation
        Expected content to be checked:
        - user response: check all outputs from user side
        - response: check all outputs from llm
        - tool calls: check all tool calls from llm
        - tool outputs: check all tool outputs
        ====== The following checks need to be started from config ======
        - local status: check files in specific workspace directory (e.g. saved some things, modified some things etc)
        - remote status: manually call MCP server to check if remote status is normally modified [not sure if possible]
        Use the above content to determine whether the task execution is successful or not
        """
        # TaskConfig.__post_init__ mutates serialized paths; never reconstruct it here.
        config = dump_line['config']
        task_dir = config['task_dir']
        pool = Path('tasks/finalpool').resolve()
        if not (pool / task_dir).resolve().is_relative_to(pool):
            raise ValueError('Task path escapes tasks/finalpool')
        task_status = dump_line['status']
        # Prepare information for evaluation
        # Per-task graders open --res_log_file with "w" and dump their own report
        # into it. Pointing that at the agent's traj_log.json destroyed the run
        # record (config + status + trajectory) on every graded task — 1565 of the
        # shipped logs are already unusable for regrading. All 496 graders only
        # write this file, none read it, so give them their own path.
        res_log_file = os.path.join(os.path.dirname(config['log_file']) or ".",
                                    "eval_report.json")
        agent_workspace = config['agent_workspace']
        # Strip weekday name (e.g. "Sunday") from launch_time for eval-script compatibility
        launch_time = " ".join((config.get('launch_time') or "").split()[:2])
        print(f"launch time in eval is {launch_time}")

        # Groundtruth isolation: in the split agent/eval flow the agent phase ran
        # with a sanitized task tree (no groundtruth_workspace/ or evaluation/), so
        # the config serialized into the log has these as None. Re-resolve them
        # against the CURRENT tree (the eval container mounts the full tasks dir).
        # For the legacy single-pass flow this is a no-op (paths already correct).
        # Never execute a command supplied in an agent-writable trajectory.
        fresh = Evaluation.build(task_dir,
                                 cn_mode=config.get("cn_mode", False))
        eval_command = fresh.evaluation_command
        groundtruth_workspace = fresh.groundtruth_workspace

        # First check task status: only SUCCESS is possible to pass; otherwise return pass = None
        if task_status != "success":
            return {
                "pass": None,
                "details": f"Task status: {task_status}, only SUCCESS counts as pass; pass is null"
            }

        # Fail closed. An unresolvable grader means we cannot judge this run, which
        # is never the same thing as the run being correct. Evaluation.build resolves
        # the grader through a repo-root-relative path (task_config.py), so it yields
        # None whenever the grader is invoked from another cwd or the task has since
        # left tasks/finalpool — both reachable, and both used to return pass=True
        # with zero checks executed.
        if eval_command is None:
            return {
                "pass": None,
                "status": "evaluation_error",
                "details": (f"Evaluation command could not be resolved for task "
                            f"'{task_dir}' — grader missing or cwd is not "
                            f"the repo root; refusing to report a pass."),
            }

        # Graders open this path themselves. Remove any agent-planted link first.
        Path(res_log_file).unlink(missing_ok=True)

        # Evaluate all content (only when task status is SUCCESS)
        if eval_command is not None:
            # try:
            args = shlex.join(["--res_log_file", res_log_file, "--agent_workspace", agent_workspace,
                               "--groundtruth_workspace", groundtruth_workspace or "",
                               "--launch_time", launch_time])
            command = f"{eval_command} {args}"
            output, error, returncode = await run_command(command, debug=True)
            print("== Evaluation STDOUT ==")
            print(output)
            print("== Evaluation STDERR ==")
            print(error)
            if returncode not in (0, 1):
                return {
                    "pass": None, "status": "evaluation_error",
                    "failure": output, "stderr": error, "returncode": returncode,
                }
            if returncode == 1:
                return {
                    "pass": False,
                    "status": "graded",
                    "failure": output,
                    "stderr": error,
                    "returncode": returncode,
                }
                
        # Finally, it's successful
        return {
            "pass": True,
            "status": "graded",
            "details": "All evaluation checks passed, and task status is success"
        }
    
    @staticmethod
    async def evaluate_from_log_file(log_file_path: str, allow_resume: bool = False,
                                     expected_task: Optional[str] = None) -> Dict[str, Any]:
        """Evaluate task from log file"""
        try:            
            if not os.path.exists(log_file_path):
                return {
                    "pass": None,
                    "status": "evaluation_error",
                    "failure": "log_file_not_found",
                    "details": f"Log file not found: {log_file_path}"
                }
            # if allow_resume AND we can load pre exist eval res, we just load it
            eval_file_path = os.path.join(os.path.dirname(log_file_path),"eval_res.json")
            # otherwise, we do real eval and store the eval result
            dump_line = read_json(log_file_path)
            if expected_task is not None and dump_line['config']['task_dir'] != expected_task:
                raise ValueError("Trajectory task ID does not match the requested task")
            # Serialized paths are agent-writable. Derive output paths from the
            # actual trajectory being graded, not from arbitrary paths in JSON.
            dump_line['config']['log_file'] = os.path.abspath(log_file_path)
            dump_line['config']['agent_workspace'] = os.path.join(os.path.dirname(os.path.abspath(log_file_path)), "workspace")
            if os.path.islink(log_file_path) or not confined_workspace(dump_line['config']['agent_workspace']):
                result = {"pass": False, "status": "graded",
                          "failure": "Output links escape the isolated agent workspace"}
                write_result(result, eval_file_path)
                return result
            if allow_resume and os.path.isfile(eval_file_path) and not os.path.islink(eval_file_path):
                return read_json(eval_file_path)
            eval_res = await TaskEvaluator.evaluate_one(dump_line)
            write_result(eval_res, eval_file_path)
            return eval_res
            
        except Exception as e:
            logging.error(f"Error evaluating from log file {log_file_path}: {e}")
            result = {
                "pass": None,
                "status": "evaluation_error",
                "failure": "evaluation_error",
                "details": str(e)
            }
            if os.path.isfile(log_file_path):
                write_result(result, os.path.join(os.path.dirname(log_file_path), "eval_res.json"))
            return result
    
    @staticmethod
    async def batch_evaluate(run_results: List[Dict[str, Any]], allow_resume: bool=False) -> List[Dict[str, Any]]:
        """Batch evaluate task results"""
        eval_results = []
        
        for run_result in run_results:
            eval_result = {
                "task_config_path": run_result["task_config_path"],
                "task_id": run_result.get("task_id", "unknown"),
            }
            
            if not run_result.get("success", False):
                eval_result["evaluation"] = {
                    "pass": False,
                    "failure": "task_execution_failed",
                    "details": run_result.get("error", "Unknown error")
                }
            else:
                log_file = run_result.get("log_file")
                if log_file:
                    eval_result["evaluation"] = await TaskEvaluator.evaluate_from_log_file(log_file, allow_resume = allow_resume)
                else:
                    eval_result["evaluation"] = {
                        "pass": False,
                        "failure": "no_log_file",
                        "details": "No log file generated"
                    }
            
            eval_result["pass"] = eval_result["evaluation"]["pass"]
            eval_results.append(eval_result)
        
        return eval_results
