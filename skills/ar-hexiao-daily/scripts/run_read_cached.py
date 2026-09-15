"""Load a task-pinned row cache before a permitted read-only analysis script."""
import argparse
from pathlib import Path
import runpy
import sys

from workbook_read_cache import configure


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cache", required=True)
    parser.add_argument("--fingerprint", required=True)
    parser.add_argument("--root", required=True)
    parser.add_argument("--script", required=True, choices=["classify_hexiao.py", "validate_plan.py", "build_execution_evidence.py", "build_flow_plan.py"])
    parser.add_argument("arguments", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    configure(Path(args.cache), args.fingerprint, Path(args.root))
    script = Path(__file__).resolve().parent / args.script
    arguments = args.arguments[1:] if args.arguments[:1] == ["--"] else args.arguments
    sys.argv = [str(script), *arguments]
    runpy.run_path(str(script), run_name="__main__")


if __name__ == "__main__":
    main()
