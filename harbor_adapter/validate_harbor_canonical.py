#!/usr/bin/env python3
"""Static validation for Harbor-canonical generated dataset."""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

AGENT_LEAKS = re.compile(
    r"(?i)\b(strands|universal\s+agent|openhands|qwen[\s_-]?code|turn_finish|TURN_FINISHED|ask_user)\b"
)
SECRET_LEAKS = re.compile(
    r"(?i)(PGPASSWORD\s*[:=]\s*[\"']?[^\s\"']+|POSTGRES_PASSWORD\s*[:=]\s*[\"']?[^\s\"']+"
    r"|password[\"']?\s*:\s*[\"'][^\"']+[\"'])"
)
# Known historical demo password must never appear in generated packaging.
KNOWN_BAD = re.compile(r"(?i)\bcamel\b")

EXPECTED_NO_GT = {
    "canvas-enrollment-gsheet",
    "canvas-module-completion-teamly-gcal",
    "canvas-quiz-remediation-forms-gcal",
    "ch-support-csat-forms-gsheet-email",
    "ecommerce-commodity-impact",
    "insales-customer-order-gsheet-email",
    "insales-customer-survey-form",
    "insales-vip-customer-gsheet-gcal-email",
    "kulinar-diet-forms-teamly-excel",
    "kulinar-event-catering-excel-word",
    "kulinar-event-menu-planner",
    "kulinar-nutrition-benchmark",
    "kulinar-scholarly-health-study",
    "kulinar-weekly-gsheet-gcal",
    "kulinar-weekly-plan-gsheet-gcal",
    "market-competitive-intelligence-report",
    "moex-earnings-calendar-alert",
    "moex-market-news-digest-teamly-forms-email",
    "pw-scholarly-trend-analysis-excel-ppt",
    "rzd-kulinar-novgorod-trip-gcal-word",
    "sf-hr-attrition-gcal",
    "sf-hr-turnover-teamly",
    "sf-sales-discount-gsheet-email",
    "sf-support-agent-performance-gsheet-gcal",
    "yf-gold-performance-gsheet-notion",
}

# Usable Oracle payload missing (dir only/.gitkeep) — not the GitHub-25 list.
EXPECTED_EMPTY_GT = {
    "rzd-canvas-fieldtrip-novgorod-gcal-word-email",
    "rzd-hr1c-training-trip-kazan-excel-email-gcal",
    "rzd-kulinar-team-trip-spb-catering-excel-gcal",
    "scholarly-fetch-gsheet-citation",
}

EXPECTED_NO_ORACLE = EXPECTED_NO_GT | EXPECTED_EMPTY_GT


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


_PG_ENV_REQUIRED = (
    "POSTGRES_USER",
    "POSTGRES_PASSWORD",
    "POSTGRES_DB",
    "PGUSER",
    "PGPASSWORD",
    "PGDATABASE",
    "PGHOST",
    "PGPORT",
)


_SERVICE_BLOCK_RE = re.compile(r"^  ([A-Za-z_][\w.-]*):\s*$")


def _service_blocks(compose: str) -> dict[str, str]:
    """Split the services section into per-service YAML blocks.

    Service names contain hyphens (workspace-prep, mcp-gateway-db), so a
    \\w-only boundary regex would silently absorb them into the preceding
    service. A boundary is a line with exactly two spaces of indentation, a
    name and a colon; nested keys are indented four spaces or more, and
    top-level keys (networks:, volumes:) have no indentation at all.
    """
    blocks: dict[str, list[str]] = {}
    section: str | None = None
    current: str | None = None
    for line in compose.splitlines():
        top = re.match(r"^([A-Za-z_][\w.-]*):", line)
        if top:
            section = top.group(1)
            current = None
            continue
        if section != "services":
            continue
        boundary = _SERVICE_BLOCK_RE.match(line)
        if boundary:
            current = boundary.group(1)
            blocks.setdefault(current, [])
            continue
        if current is not None and (line.startswith("    ") or not line.strip()):
            blocks[current].append(line)
    return {name: "\n".join(lines) for name, lines in blocks.items()}


def _service_networks(block: str) -> set[str]:
    """Collect network names from a service block (list and mapping forms)."""
    nets: set[str] = set()
    in_networks = False
    for line in block.splitlines():
        if re.match(r"^    networks:\s*$", line):
            in_networks = True
            continue
        if not in_networks:
            continue
        item = re.match(r"^      (?:-\s+)?([A-Za-z_][\w.-]*):?(?:\s|$)", line)
        if item:
            nets.add(item.group(1))
        elif line.strip():
            in_networks = False
    return nets


def _pg_env_values(env_path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    for line in env_path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        data[key.strip()] = value.strip()
    return data


def _validate_pg_env(compose: str, env_path: Path) -> list[str]:
    """Validate the packaged env_file without creating or deleting files.

    Valid packaged state: environment/pg.env exists, is populated (postgres
    seeds its user from it; prep / mcp-gateway-db / grader connect with it)
    and is referenced by compose — while the agent service (main) has no way
    to reach the credentials from its container: no env_file or volume
    reference, no PG credentials in env, no mount of the package root or
    environment/, no docker socket, and no shared network with postgres.
    """
    references_pg_env = bool(re.search(r"(?m)^\s*-\s+\./pg\.env\s*$", compose))
    if references_pg_env and not env_path.exists():
        return ["compose references ./pg.env but environment/pg.env is missing"]
    if not references_pg_env:
        if env_path.exists():
            return ["environment/pg.env exists without a compose reference"]
        return []
    errs: list[str] = []
    data = _pg_env_values(env_path)
    for key in _PG_ENV_REQUIRED:
        if not data.get(key):
            errs.append(f"environment/pg.env is missing {key}")
    blocks = _service_blocks(compose)
    main = blocks.get("main")
    if main is None:
        errs.append("compose has no main service — agent blindness cannot be verified")
        return errs
    # Agent-blindness: the main container must have no path to the secret.
    if "pg.env" in main:
        errs.append("main service references pg.env — agent must not see credentials")
    if re.search(r"(?m)^\s+(PGPASSWORD|PG_PASSWORD|POSTGRES_PASSWORD):", main):
        errs.append("main service carries PG credentials in env")
    for volume in re.findall(r"(?m)^\s+-\s+(\.[^:\s]*):", main):
        if volume.rstrip("/") in (".", "./environment"):
            errs.append(f"main service mounts {volume} — would expose environment/pg.env")
    if "docker.sock" in main:
        errs.append("main service mounts docker.sock — agent could read postgres credentials")
    postgres = blocks.get("postgres")
    if postgres is not None:
        shared = sorted(_service_networks(main) & _service_networks(postgres))
        if shared:
            errs.append("main service shares network with postgres: " + ", ".join(shared))
    return errs


def _validate_agent_blindness(task_dir: Path) -> list[str]:
    """Scan agent-visible surfaces for the packaged password value.

    Compose-level blindness is checked by _validate_pg_env; this complements
    it on the files the agent receives directly: task.toml (MCP URLs and env),
    the public MCP manifest, the workspace-prep script, the instruction and
    the task payload that main unpacks at runtime.
    """
    pg_env = task_dir / "environment" / "pg.env"
    if not pg_env.is_file():
        return []
    password = _pg_env_values(pg_env).get("POSTGRES_PASSWORD", "")
    if not password:
        return []
    errs: list[str] = []
    surfaces = [
        "task.toml",
        "instruction.md",
        "environment/mcp_manifest_public.json",
        "environment/prep/prepare_workspace.py",
    ]
    for rel in surfaces:
        path = task_dir / rel
        if path.is_file() and password in path.read_text(encoding="utf-8", errors="replace"):
            errs.append(f"{rel} contains the pg password — agent-visible")
    payload = task_dir / "environment" / "task_payload"
    if payload.is_dir():
        for path in sorted(payload.rglob("*")):
            if path.is_file() and password.encode("utf-8") in path.read_bytes():
                errs.append(f"environment/task_payload/{path.name} contains the pg password")
    return errs


def validate_task(task_dir: Path) -> list[str]:
    errs: list[str] = []
    tid = task_dir.name
    required = [
        "instruction.md",
        "task.toml",
        "environment/docker-compose.yaml",
        "environment/mcp_manifest_public.json",
        "environment/mcp_runtime/mcp_gateway.py",
        "environment/prep/prepare_workspace.py",
        "tests/grader_private/task_contract.json",
        "tests/test.sh",
        "tests/verifier/workspace_lifecycle.py",
        "environment/lifecycle/workspace_lifecycle.py",
        "environment/lifecycle/completion.py",
        "metadata.json",
    ]
    for rel in required:
        if not (task_dir / rel).exists():
            errs.append(f"missing {rel}")

    instr = _read(task_dir / "instruction.md") if (task_dir / "instruction.md").exists() else ""
    if AGENT_LEAKS.search(instr):
        errs.append("agent-specific leak in instruction.md")

    toml = _read(task_dir / "task.toml") if (task_dir / "task.toml").exists() else ""
    if "mcp-gateway-public" not in toml:
        errs.append("task.toml missing mcp-gateway-public")
    if re.search(r"mcp-gateway-db|mcp-gateway-workspace", toml):
        errs.append("task.toml exposes internal gateway")
    if re.search(r"(?i)\bPG(PASSWORD|USER|HOST|DATABASE)\b", toml):
        errs.append("task.toml contains PG env")
    if "ask_user" in toml or AGENT_LEAKS.search(toml):
        errs.append("task.toml agent-specific leak")
    if 'transport = "streamable-http"' not in toml:
        errs.append("task.toml missing streamable-http")

    compose = (
        _read(task_dir / "environment" / "docker-compose.yaml")
        if (task_dir / "environment" / "docker-compose.yaml").exists()
        else ""
    )
    if "agent_net" not in compose or "mcp_workspace_net" not in compose:
        errs.append("compose missing required networks")
    if "mcp-gateway-public" not in compose:
        errs.append("compose missing mcp-gateway-public")
    # main must not get PG env
    main_block = re.search(r"(?ms)^  main:.*?(?=^  [a-z]|\Z)", compose)
    if main_block:
        mb = main_block.group(0)
        if re.search(r"(?i)PGPASSWORD|POSTGRES_PASSWORD|PGUSER:|PGHOST:", mb):
            errs.append("main service has PG env")
        if "env_file:" in mb and "pg.env" in mb:
            errs.append("main mounts pg.env")
        if "./lifecycle:/opt/cowork_lifecycle:ro" not in mb:
            errs.append("main missing lifecycle mount")
        if "../tests:" in mb:
            errs.append("main mounts tests tree")
    if "/opt/cowork_lifecycle/workspace_lifecycle.py finalize" not in toml:
        errs.append("task.toml main collect does not finalize via /opt/cowork_lifecycle")
    if (task_dir / "environment" / "prep" / "task_contract.json").exists():
        errs.append("prep still contains task_contract.json")
    if (task_dir / "solution" / "task_contract.json").exists():
        errs.append("solution contains a task contract")
    private_contract = task_dir / "tests" / "grader_private" / "task_contract.json"
    if private_contract.is_file():
        try:
            private_data = json.loads(private_contract.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            errs.append("grader contract is not valid JSON")
            private_data = None
        if isinstance(private_data, dict):
            if private_data.get("task_dir") != tid:
                errs.append("grader contract task_dir mismatch")
            evaluation = private_data.get("evaluation", None)
            if not isinstance(evaluation, dict):
                errs.append("grader contract missing evaluation object")
            elif (
                set(evaluation) != {"groundtruth_workspace", "evaluation_command"}
                or evaluation.get("groundtruth_workspace") is not None
                or evaluation.get("evaluation_command") is not None
            ):
                errs.append("grader contract evaluation is not null/null")
            if "evaluation.main" in json.dumps(private_data):
                errs.append("grader contract contains an evaluation module")
    script = (task_dir / "tests" / "test.sh").read_text(encoding="utf-8") if (task_dir / "tests" / "test.sh").exists() else ""
    if "--contract " not in script or "/tests/grader_private/task_contract.json" not in script:
        errs.append("verifier does not read the grader-only contract")
    # Graders with flat sibling imports must carry the sys.path shim: the -m
    # grading path never puts the grader's own directory on sys.path.
    packaged_main = task_dir / "tests" / "evaluation" / "main.py"
    if packaged_main.is_file():
        here = str(Path(__file__).resolve().parent)
        if here not in sys.path:
            sys.path.insert(0, here)
        from generate_harbor_canonical import _first_flat_sibling_import

        if _first_flat_sibling_import(packaged_main) is not None:
            errs.append("flat sibling imports without sys.path shim")
    for service in (
        "main",
        "workspace-prep",
        "mcp-gateway-workspace",
        "mcp-gateway-public",
        "mock-pages",
    ):
        block = re.search(
            rf"(?ms)^  {re.escape(service)}:.*?(?=^  [a-z0-9-]+:|\Z)",
            compose,
        )
        if not block:
            continue
        body = block.group(0)
        if "grader_private" in body or "task_contract" in body or "../tests:" in body:
            errs.append(f"{service} mounts private contract")
    if "evaluation_command" in instr or "groundtruth_workspace" in instr:
        errs.append("instruction contains grader contract")
    if "evaluation_command" in toml or "groundtruth_workspace" in toml:
        errs.append("task.toml contains grader contract")
    if "strands" in compose.lower():
        errs.append("compose mentions strands")
    # workspace gateway must not be on db_net
    ws_block = re.search(
        r"(?ms)^  mcp-gateway-workspace:.*?(?=^  [a-z0-9-]+:|\Z)", compose
    )
    if ws_block and re.search(r"(?m)^\s+- db_net\s*$", ws_block.group(0)):
        errs.append("mcp-gateway-workspace attached to db_net")

    # secrets / camel in packaging (allow prepare rewrite of source preprocess copies under task_payload)
    for path in task_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix in {".png", ".jpg", ".pdf", ".tar", ".gz", ".zip", ".pyc"}:
            continue
        rel = path.relative_to(task_dir).as_posix()
        if rel.startswith("environment/task_payload/"):
            continue  # source copies may still contain historical literals until prep rewrite
        if rel.startswith("tests/evaluation/") or rel.startswith("tests/groundtruth"):
            continue
        if rel.startswith("tests/task_data/"):
            continue  # source data copies (mock pages, initial workspace) — same rationale as task_payload
        if rel.startswith("solution/groundtruth_workspace/"):
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        if KNOWN_BAD.search(text) and "prepare_workspace.py" not in rel:
            # prepare_workspace may contain replacement needles for source preprocess
            if "password" in text.lower() or "POSTGRES" in text:
                errs.append(f"known password literal in {rel}")
        if AGENT_LEAKS.search(text) and rel.endswith((".md", ".toml", ".yaml", ".yml", ".json", ".sh")):
            if "metadata.json" in rel:
                continue
            errs.append(f"agent leak in {rel}")

    meta = json.loads(_read(task_dir / "metadata.json")) if (task_dir / "metadata.json").exists() else {}
    has_oracle = (task_dir / "solution" / "solve.sh").is_file()
    no_gt_marker = (task_dir / "solution" / "NO_GROUNDTRUTH").is_file()
    if tid in EXPECTED_NO_ORACLE:
        if has_oracle:
            errs.append("unexpected Oracle for no-GT/empty-GT task")
        if not no_gt_marker:
            errs.append("missing NO_GROUNDTRUTH marker")
    else:
        if not has_oracle:
            errs.append("missing Oracle solve.sh for GT task")
        if no_gt_marker:
            errs.append("NO_GROUNDTRUTH marker on GT task")

    # Docker Compose requires every referenced env_file to exist. The
    # packaged pg.env carries the per-dataset credentials: populated in the
    # dataset is the valid state, provided the agent container stays blind
    # to it (checked in _validate_pg_env and _validate_agent_blindness).
    compose_path = task_dir / "environment" / "docker-compose.yaml"
    if compose_path.exists():
        env_dir = task_dir / "environment"
        pg_env = env_dir / "pg.env"
        errs.extend(_validate_pg_env(compose, pg_env))
        errs.extend(_validate_agent_blindness(task_dir))
        try:
            subprocess.run(
                ["docker", "compose", "-f", str(compose_path), "config", "-q"],
                check=True,
                capture_output=True,
                text=True,
                timeout=60,
            )
        except FileNotFoundError:
            errs.append("docker not available for compose config")
        except subprocess.CalledProcessError as exc:
            errs.append(f"compose config failed: {(exc.stderr or exc.stdout or '')[:300]}")
        except Exception as exc:
            errs.append(f"compose config error: {exc}")

    _ = meta

    # Grader sidecar isolation for DB tasks.
    compose = task_dir / "environment" / "docker-compose.yaml"
    if compose.is_file():
        cy = compose.read_text(encoding="utf-8", errors="ignore")
        has_postgres = "\n  postgres:" in cy or "\n  postgres:\n" in cy
        if has_postgres:
            if "\n  grader:" not in cy and "\ngrader:" not in cy:
                errs.append("DB task missing grader sidecar")
            if "grader_out:/grader_out:ro" not in cy:
                errs.append("main must mount grader_out read-only")
            toml = (task_dir / "task.toml").read_text(encoding="utf-8", errors="ignore")
            if 'service = "grader"' not in toml:
                errs.append("task.toml missing verifier.collect for grader")
    return errs


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    tasks = sorted(
        p for p in root.iterdir() if p.is_dir() and (p / "task.toml").exists()
    )
    all_errs: dict[str, list[str]] = {}
    for t in tasks:
        e = validate_task(t)
        if e:
            all_errs[t.name] = e
            print(f"FAIL {t.name}: {e}")
        else:
            print(f"OK {t.name}")

    with_oracle = sum(1 for t in tasks if (t / "solution" / "solve.sh").is_file())
    no_gt = sorted(t.name for t in tasks if (t / "solution" / "NO_GROUNDTRUTH").is_file())
    report = {
        "n_tasks": len(tasks),
        "n_ok": len(tasks) - len(all_errs),
        "n_fail": len(all_errs),
        "with_oracle": with_oracle,
        "no_groundtruth": no_gt,
        "expected_no_gt": sorted(EXPECTED_NO_GT),
        "expected_empty_gt": sorted(EXPECTED_EMPTY_GT),
        "no_gt_match_expected": set(no_gt) == EXPECTED_NO_ORACLE,
        "failures": all_errs,
    }
    out = root / "STATIC_VALIDATION.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({k: report[k] for k in report if k != "failures"}, indent=2, ensure_ascii=False))
    return 1 if all_errs else 0


if __name__ == "__main__":
    raise SystemExit(main())
