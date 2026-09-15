# 018 并入后 pytest 红项（不删金标）

跑于正式 clone，`.venv/bin/python`。

## receivables-merge

`2 failed, 6 passed`

- `test_merge_robustness.py::test_happy_path_main_table`
- `test_merge_robustness.py::test_renamed_columns_and_unknown_sheet_warn`（李尚 merge 不再出「认列告警」sheet）

## split-by-sales

`6 passed`

## ar-hexiao-daily

`68 failed, 501 passed, 1 skipped`

正本金标/原子写/「今日清单」sheet 等用例对的是并前脚本；极速包改了日清三表（任务范围 / 核销明细 / 流转表怎么填）和判定字段。失败用例未删。

## consolidated-statements

无 pytest。`entry.py --period 202608 --input-dir 空夹` EXIT 0，出部分底稿。
