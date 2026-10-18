"""Execute the real parallel shell runner with a recording Docker test double.

The fake supplies container boundaries and native output fixtures, never models,
PostgreSQL or a real Docker daemon. Staging, rsync, semaphore concurrency,
verdict cleanup, artifact parsing and final CSV aggregation are real.
"""
import csv
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest

ROOT = Path(__file__).resolve().parents[1]

# Each invocation is a separate OS process, just like the Docker executable.
# Container mount records live in individual files; calls.jsonl uses O_APPEND.
FAKE_DOCKER = r'''
import json
import os
from pathlib import Path
import sys

root = Path(os.environ["FAKE_DOCKER_ROOT"])
state = root / "fake-state"
args = sys.argv[1:]

def record(event):
    with (state / "calls.jsonl").open("a") as stream:
        stream.write(json.dumps(event) + "\n")

record({"operation": "command", "args": args})
if not args:
    raise SystemExit(2)
if args[0] in ("network", "rm", "stop"):
    raise SystemExit(0)
if args[0] == "run":
    if "--name" not in args:
        # Image probe and one-shot FK repair: no model or DB process is started.
        raise SystemExit(0)
    name = args[args.index("--name") + 1]
    mounts = {}
    for index, arg in enumerate(args[:-1]):
        if arg == "-v":
            parts = args[index + 1].split(":")
            mounts[parts[1]] = {"source": parts[0], "mode": parts[2] if len(parts) > 2 else "rw"}
    info = {"name": name, "mounts": mounts}
    (state / (name + ".json")).write_text(json.dumps(info))
    if name.startswith("agent-"):
        tree = Path(mounts["/workspace/tasks"]["source"])
        info["stage_files"] = sorted(str(path.relative_to(tree)) for path in tree.rglob("*") if path.is_file())
        record({"operation": "agent_start", **info})
    elif name.startswith("eval-"):
        record({"operation": "eval_start", **info})
    print(name)
    raise SystemExit(0)
if args[0] == "exec":
    name = args[1]
    if "pg_isready" in args:
        raise SystemExit(0)
    info = json.loads((state / (name + ".json")).read_text())
    task = args[args.index("--task_dir") + 1]
    attempt = Path(info["mounts"]["/workspace/dumps"]["source"])
    run = attempt / "fixture-model" / ("SingleUserTurn-" + task)
    if name.startswith("agent-"):
        record({"operation": "agent_exec", "task": task, "args": args, "attempt": str(attempt)})
        if task == "timeout":
            # Trusted Docker/timeout exit status, with no trajectory written.
            raise SystemExit(124)
        run.mkdir(parents=True, exist_ok=True)
        (run / "workspace").mkdir()
        (run / "traj_log.json").write_text(json.dumps({"status": "success", "config": {"task_dir": task}}))
        # Must be removed after the agent is stopped and before grading.
        (run / "eval_res.json").write_text(json.dumps({"pass": True, "status": "graded"}))
        # Poison stdout too: it must never override a native failed/missing grade.
        print("====== Status: success ======")
        print("Pass: True")
        raise SystemExit(0)
    if name.startswith("eval-"):
        planted = list(attempt.rglob("eval_res.json"))
        record({"operation": "eval_exec", "task": task, "args": args,
                "attempt": str(attempt), "planted_verdicts": [str(p) for p in planted]})
        print("Pass: True")
        if task in ("no-grade", "timeout"):
            # Simulate grader setup/discovery failure; no fresh native verdict.
            raise SystemExit(2)
        result = ({"pass": None, "status": "evaluation_error", "stderr": "missing dependency"}
                  if task == "eval-error" else {"pass": task == "pass", "status": "graded"})
        (run / "eval_res.json").write_text(json.dumps(result))
        raise SystemExit(2 if result["pass"] is None else (0 if result["pass"] else 1))
raise AssertionError("Unsupported fake Docker invocation: " + repr(args))
'''


@unittest.skipUnless(os.name == 'posix' and shutil.which('bash') and shutil.which('rsync'),
                     'parallel runner requires POSIX bash and rsync')
class ParallelRunnerIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix='cowork-parallel-test-')
        cls.addClassCleanup(cls.tmp.cleanup)
        cls.fixture = Path(cls.tmp.name).resolve()
        (cls.fixture / 'scripts').mkdir()
        (cls.fixture / 'configs').mkdir()
        (cls.fixture / 'db').mkdir()
        (cls.fixture / 'fake-state').mkdir()
        shutil.copyfile(ROOT / 'run_parallel.sh', cls.fixture / 'run_parallel.sh')
        shutil.copyfile(ROOT / 'scripts/read_attempt_result.py', cls.fixture / 'scripts/read_attempt_result.py')
        # Not executed by fake Docker; keeping the actual entrypoint makes the
        # fixture reflect the command path contract checked below.
        shutil.copyfile(ROOT / 'scripts/run_eval.py', cls.fixture / 'scripts/run_eval.py')
        (cls.fixture / 'db/init.sql.gz').write_bytes(b'fixture seed: Docker is mocked')
        cls.tasks = ('pass', 'fail', 'eval-error', 'no-grade', 'timeout')
        for task in cls.tasks:
            directory = cls.fixture / 'tasks/finalpool' / task
            contents = {
                'docs/task.md': 'Fixture input',
                'preprocess/main.py': '# writable preprocessing source',
                'initial_workspace/input.txt': 'Input data',
                'evaluation/main.py': '# answer-checking code must be hidden',
                'groundtruth_workspace/answer.txt': 'answer key',
                'files/groundtruth_data.json': '{"answer": 42}',
            }
            for relative, text in contents.items():
                path = directory / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text)
            # A prior global dump with a true verdict must not rescue this run.
            old = cls.fixture / 'dumps/old-model' / ('SingleUserTurn-' + task)
            old.mkdir(parents=True)
            (old / 'traj_log.json').write_text(json.dumps({'status': 'success', 'config': {'task_dir': task}}))
            (old / 'eval_res.json').write_text('{"pass":true,"status":"graded"}')
        docker = cls.fixture / 'fake-docker'
        docker.write_text('#!' + sys.executable + '\n' + textwrap.dedent(FAKE_DOCKER))
        docker.chmod(0o755)
        # Do not forward the caller's credentials or benchmark options to logs.
        env = {key: os.environ[key] for key in ('PATH', 'HOME', 'TMPDIR', 'LANG') if key in os.environ}
        env.update(DOCKER=str(docker), FAKE_DOCKER_ROOT=str(cls.fixture),
                   MODEL_NAME='fixture-model', MODEL_PROVIDER='openai_compatible',
                   IMAGE='fixture-image', TASK_TIMEOUT='1')
        cls.process = subprocess.run(['bash', 'run_parallel.sh', '3', *cls.tasks],
                                     cwd=cls.fixture, env=env, text=True, capture_output=True, timeout=45)
        if cls.process.returncode:
            raise AssertionError(cls.process.stdout + cls.process.stderr)
        summaries = list((cls.fixture / 'benchmark_logs').glob('fully_parallel_*/summary.csv'))
        if len(summaries) != 1:
            raise AssertionError('No unique summary: ' + cls.process.stdout + cls.process.stderr)
        cls.log_directory = summaries[0].parent
        with summaries[0].open() as stream:
            cls.rows = {row['task']: row for row in csv.DictReader(stream)}
        cls.calls = [json.loads(line) for line in (cls.fixture / 'fake-state/calls.jsonl').read_text().splitlines()]

    def events(self, operation):
        return [entry for entry in self.calls if entry['operation'] == operation]

    def test_native_verdicts_and_errors_are_distinct(self):
        expected = {'pass': ('success', 'True'), 'fail': ('success', 'False'),
                    'eval-error': ('evaluation_error', 'null'),
                    'no-grade': ('evaluation_error', 'null'), 'timeout': ('wall_timeout', 'null')}
        self.assertEqual(set(self.rows), set(expected), self.process.stdout + self.process.stderr)
        for task, pair in expected.items():
            with self.subTest(task=task):
                row = self.rows[task]
                self.assertEqual((row['status'], row['eval_pass']), pair, self.process.stderr)
        self.assertIn('EVAL_ERROR', self.process.stdout)
        self.assertIn('EVAL_FAIL', self.process.stdout)
        with (self.fixture / 'benchmark_logs/index.csv').open() as stream:
            aggregate = next(csv.DictReader(stream))
        self.assertEqual(aggregate['pass'], '1')
        self.assertEqual(aggregate['eval_fail'], '1')
        self.assertEqual(aggregate['agent_fail'], '1')
        self.assertEqual(aggregate['other'], '2')
        self.assertEqual(aggregate['pass_pct'], '')

    def test_agent_mounts_are_sanitized_and_attempt_local(self):
        starts = self.events('agent_start')
        self.assertEqual(len(starts), len(self.tasks))
        outputs = set()
        for entry in starts:
            mounts = entry['mounts']
            outputs.add(mounts['/workspace/dumps']['source'])
            self.assertNotEqual(mounts['/workspace/dumps']['source'], str(self.fixture / 'dumps'))
            self.assertEqual(mounts['/workspace/tasks']['mode'], 'rw')
            self.assertEqual(mounts['/workspace/configs']['mode'], 'ro')
            self.assertTrue(any(path.endswith('/docs/task.md') for path in entry['stage_files']))
            self.assertTrue(any(path.endswith('/preprocess/main.py') for path in entry['stage_files']))
            self.assertFalse(any('groundtruth' in path or '/evaluation/' in path for path in entry['stage_files']))
            task_names = {Path(path).parts[1] for path in entry['stage_files']}
            self.assertEqual(len(task_names), 1)
        self.assertEqual(len(outputs), len(self.tasks))

    def test_camel_runs_only_agent_phase_and_evaluator_is_standalone(self):
        for entry in self.events('agent_exec'):
            args = entry['args']
            self.assertIn('/workspace/main.py', args)
            self.assertEqual(args[args.index('--phase') + 1], 'agent')
        evaluations = self.events('eval_exec')
        self.assertEqual(len(evaluations), len(self.tasks))
        for entry in evaluations:
            args = entry['args']
            self.assertIn('/workspace/scripts/run_eval.py', args)
            self.assertNotIn('/workspace/main.py', args)
            self.assertEqual(args[args.index('--dumps_dir') + 1], '/workspace/dumps')
            self.assertEqual(args[args.index('--task_dir') + 1], entry['task'])
        for entry in self.events('eval_start'):
            self.assertEqual(entry['mounts']['/workspace/tasks'],
                             {'source': str(self.fixture / 'tasks'), 'mode': 'ro'})

    def test_planted_old_and_streamed_verdicts_cannot_supply_a_pass(self):
        for entry in self.events('eval_exec'):
            self.assertEqual(entry['planted_verdicts'], [], entry)
        # Both current agent stdout and an old global dump say True. The current
        # attempt's grader wrote no verdict, so it must remain unscored.
        self.assertEqual(self.rows['no-grade']['status'], 'evaluation_error')
        self.assertEqual(self.rows['no-grade']['eval_pass'], 'null')
        self.assertIn('Pass: True', (self.log_directory / 'no-grade.log').read_text())
        old = self.fixture / 'dumps/old-model/SingleUserTurn-no-grade/eval_res.json'
        self.assertIs(json.loads(old.read_text())['pass'], True)


if __name__ == '__main__':
    unittest.main()
