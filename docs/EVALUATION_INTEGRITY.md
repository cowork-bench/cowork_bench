# Evaluation integrity fixes

The regression suite reproduces specific false passes, false failures and
infrastructure misclassification. It does not certify every task in the pool.

## Required artifacts and task contracts

- The 58 entrypoints listed in `KNOWN_ISSUES.md` now run a shared prerequisite
  check before their percentage-based score: 80 mandatory Office files across
  55 tasks and six Google Sheets configurations. Missing, empty or unreadable
  required files, and absent shared sheets, cannot disappear from the denominator.
  These prerequisites do not replace each task's content checks.
- Canvas enrollment checks the requested workbook columns, shared sheet and
  presentation values. Canvas grade summary accepts a headerless Summary sheet
  and checks its shared sheet. Formula workbooks are recalculated with LibreOffice
  Calc in a temporary profile without changing the submitted artifact.
- Kulinar price/cost uses the initial price list and fixed category contract,
  compares shared-sheet cells by coordinates, and checks ingredient identities.
  Terminal menu validation accepts a single JSONL entity and verifies document
  ingredients and recipe attributes, including table-based layouts.
- HR travel verifies the requested top-five employees, both journeys, personal
  email recipients and exact calendar times with timezone conversion. An
  unreadable catering workbook cannot pass.
- MOEX portfolio uses the latest source close consistently with its MCP tool.
  MOEX news and Teamly competitor tasks validate source content and metadata,
  rather than accepting keyword-only pages or a matching global item count.
- The academic review checks the five supplied papers, source numeric anchors,
  requested document structure and actual extracted word counts. These remain
  deterministic content checks: they cannot establish arbitrary scientific
  correctness or prove that a paper was read.

## Execution and verdicts

Both container runners now split the bundled CAMEL and Strands engines into
agent and evaluation phases. Custom engines must support `--phase agent` and set
`AGENT_PHASE_AWARE=1`. Direct `main.py --phase all` remains available for development;
it does not provide this container isolation.

Each parallel attempt has a fresh artifact directory. Only its native trajectory
and evaluator result are read; old global dumps and stdout strings cannot supply
a verdict. The evaluator re-resolves the trusted grading command, checks the
requested task ID, derives workspace paths from the actual log, rejects workspace
links escaping that directory, and replaces output verdicts without following
agent-planted links. The agent container is removed before grading starts.

`grade_process` distinguishes explicit task-check exits (0/1) from uncaught
exceptions or other exits (2). An unresolved evaluator or evaluation exception
produces `pass: null`, `status: evaluation_error`, with diagnostics retained.
Preprocess failure stops agent execution and records `infra_error`. A trusted
outer timeout remains a budget failure even if no trajectory was flushed.
Runs with infrastructure/evaluation errors have no aggregate `pass_pct` in the
central index; report their coverage and errors separately.

The image adds Calc, uses Python 3.12 for Python MCP environments and fails on
MCP build/sync errors. HR and ClickHouse pin their compatible MCP dependency;
MOEX and Yahoo declare the PostgreSQL driver they import.

## Reproducing the checks

Install Python 3.12, LibreOffice Calc, rsync, and these Python packages:

```sh
python -m pip install openpyxl python-docx python-pptx psycopg2-binary termcolor pypdf
python -m unittest discover -s tests -v
bash -n run_parallel.sh scripts/run_containerized.sh
```

Tests run real evaluator code on valid fixture artifacts and deliberately
corrupted variants. PostgreSQL-dependent tests adapt query results; they do not
start a database. Formula tests use real LibreOffice. The parallel-runner test
executes the shell script with a fake Docker CLI; it checks orchestration and
summary behavior, not Docker networking or image build success. CI runs the same
suite without model credentials.

A clean full image build and representative tasks against the live seeded
PostgreSQL database are still required before using the revised image for a
published benchmark. This patch does not rewrite historical scores. Regrade
preserved artifacts under a pinned version, record evaluation errors separately,
and rerun only cases whose original execution or source state was invalid.
