from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


MODE_ONLY_MERGE = "仅合并"
MODE_MERGE_AND_SPLIT = "合并并按销售拆分"
MODE_SPLIT_EXISTING = "已有应收 all 直接拆分"
MODES = (MODE_ONLY_MERGE, MODE_MERGE_AND_SPLIT, MODE_SPLIT_EXISTING)


def run_stage(label: str, command: list[str]) -> None:
    completed = subprocess.run(
        command,
        cwd=Path(command[1]).resolve().parent,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode:
        lines = [line.strip() for line in (completed.stderr + "\n" + completed.stdout).splitlines()]
        detail = "；".join(line for line in lines if line)[-1200:]
        raise RuntimeError(f"{label}失败：{detail or '脚本返回非零退出码。'}")


def archive_directory(source: Path, target: Path) -> None:
    files = sorted(path for path in source.rglob("*") if path.is_file())
    if not files:
        raise RuntimeError("拆分完成，但没有生成销售人员工作簿。")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, path.relative_to(source))


def main() -> None:
    parser = argparse.ArgumentParser(description="应收账款合并与销售人员拆分")
    parser.add_argument("--mode", choices=MODES, default=MODE_MERGE_AND_SPLIT)
    parser.add_argument("--primary", required=True, help="源台账或已有应收 all")
    parser.add_argument("--reference", help="上一版应收 all")
    parser.add_argument("--base-month", help="账龄基准月，YYYYMM")
    parser.add_argument("--date", help="拆分结果日期标签，MMDD")
    parser.add_argument("--out", required=True, help="合并结果路径；同时用于确定任务输出目录")
    args = parser.parse_args()

    primary = Path(args.primary).resolve()
    if not primary.is_file():
        raise RuntimeError(f"输入文件不存在：{primary}")
    output_path = Path(args.out).resolve()
    output_dir = output_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    script_dir = Path(__file__).resolve().parent
    merge_script = script_dir / "merge.py"
    split_script = script_dir / "split.py"

    merged_path = output_dir / "应收账款合并结果.xlsx"
    if args.mode in (MODE_ONLY_MERGE, MODE_MERGE_AND_SPLIT):
        command = [sys.executable, str(merge_script), "--source", str(primary), "--out", str(merged_path)]
        if args.reference:
            command.extend(["--ref", str(Path(args.reference).resolve())])
        if args.base_month:
            command.extend(["--base-month", args.base_month])
        print("正在合并应收台账", flush=True)
        run_stage("应收账款合并", command)
        if not merged_path.is_file():
            raise RuntimeError("合并脚本未生成应收账款合并结果。")

    if args.mode in (MODE_MERGE_AND_SPLIT, MODE_SPLIT_EXISTING):
        split_input = merged_path if args.mode == MODE_MERGE_AND_SPLIT else primary
        split_dir = output_dir / "_拆分中间文件"
        shutil.rmtree(split_dir, ignore_errors=True)
        command = [sys.executable, str(split_script), "--input", str(split_input), "--out-dir", str(split_dir)]
        if args.date:
            command.extend(["--date", args.date])
        print("正在按销售人员拆分应收台账", flush=True)
        try:
            run_stage("应收按销售人员拆分", command)
            archive_directory(split_dir, output_dir / "应收按销售拆分结果.zip")
        finally:
            shutil.rmtree(split_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
