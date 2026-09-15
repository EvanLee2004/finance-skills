from __future__ import annotations

import argparse


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", required=True)
    parser.add_argument("--result", required=True)
    parser.parse_args()
    raise RuntimeError("此 Skill 尚未发布，不能创建执行任务。")


if __name__ == "__main__":
    main()
