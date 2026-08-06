"""兼容入口：正式实现已迁移到 jdy-cashflow-export Skill。"""

from pathlib import Path
import runpy


if __name__ == "__main__":
    target = (
        Path(__file__).resolve().parents[2]
        / "jdy-cashflow-export"
        / "scripts"
        / "jdy_selenium_rpa.py"
    )
    runpy.run_path(str(target), run_name="__main__")
