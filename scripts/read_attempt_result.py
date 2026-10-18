"""Read only one isolated attempt's native artifacts; never infer scores from text."""
import argparse
import json
from pathlib import Path


def read_result(directory, task, timed_out=False, agent_exit=0):
    if timed_out:
        return "wall_timeout", "null"
    if agent_exit not in (0, 1):
        return "infra_error", "null"
    logs = list(Path(directory).rglob("traj_log.json"))
    if len(logs) != 1 or logs[0].is_symlink():
        return "evaluation_error", "null"
    try:
        record = json.loads(logs[0].read_text())
        if record["config"]["task_dir"] != task:
            return "evaluation_error", "null"
        status = record["status"]
        if status != "success":
            if status in {"failed", "max_turns_reached", "infra_error"}:
                return status, "null"
            return "infra_error", "null"
        if agent_exit != 0:
            return "infra_error", "null"
        result = json.loads(logs[0].with_name("eval_res.json").read_text())
        if not isinstance(result, dict):
            return "evaluation_error", "null"
        if result.get("status") == "evaluation_error" or type(result.get("pass")) is not bool:
            return "evaluation_error", "null"
        return "success", str(result["pass"])
    except (OSError, ValueError, KeyError, TypeError):
        return "evaluation_error", "null"


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("directory")
    parser.add_argument("task")
    parser.add_argument("--timed-out", action="store_true")
    parser.add_argument("--agent-exit", type=int, default=0)
    args = parser.parse_args()
    print(*read_result(args.directory, args.task, args.timed_out, args.agent_exit))
