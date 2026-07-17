# dept-expense-alloc · 部门费用归集分摊

总账会计斯佳姐每月「部门科目余额表 + 利润表」自动化技能（开发中）。

## 快测

```bash
cd skills/dept-expense-alloc
python3 tests/test_robustness.py
python3 scripts/inspect_inputs.py --input-dir 工作区/input
python3 scripts/allocate.py --input-dir 工作区/input --out 工作区/output/试跑.xlsx
```

## 真实跑

把当月材料按 `01_主体余额表` … `07_定稿对照` 丢进一个目录，opencode 触发本技能或手动：

```bash
python3 scripts/allocate.py --input-dir /绝对路径/当月包 --out /绝对路径/部门科目余额表_YYYYMM.xlsx
```

## 配置

见 `config/业务规则.md`。真实财务数据勿提交 git。
