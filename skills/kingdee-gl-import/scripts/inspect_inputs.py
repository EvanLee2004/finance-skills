#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按表头认序时账 / 凭证列表，不靠文件名。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from openpyxl import load_workbook

from convert import inspect_dir  # noqa: E402


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="认序时账源表")
    parser.add_argument("--input-dir", default="")
    parser.add_argument("--input", default="")
    args = parser.parse_args(argv)
    report = inspect_dir(args.input_dir, args.input)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("ready") else 2


if __name__ == "__main__":
    raise SystemExit(main())
