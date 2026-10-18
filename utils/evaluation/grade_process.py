"""Run a Python grader with distinct verdict (0/1) and infrastructure (2) exits.

A normal ``sys.exit(1)`` is a failed check. An uncaught import/runtime exception
is an evaluation error, even though Python ordinarily gives both exit code 1.
"""
import argparse
import runpy
import sys
import traceback


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--module", required=True)
    args, grader_args = parser.parse_known_args()
    sys.argv = [args.module, *grader_args]
    try:
        runpy.run_module(args.module, run_name="__main__")
    except SystemExit as exc:
        if exc.code is None or exc.code == 0:
            return 0
        if isinstance(exc.code, int) and exc.code == 1:
            return 1
        print(f"Grader exited without a verdict: {exc.code!r}", file=sys.stderr)
        return 2
    except Exception:
        traceback.print_exc()
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
