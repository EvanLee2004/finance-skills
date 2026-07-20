# 部门费用归集分摊（dept-expense-alloc）v1.0.0

总账月度：主体余额 + 收入底稿 + 人员归属 + 用友按人 → 部门科目余额表 + 利润表 + 核对。

## 快测
```bash
python3 tests/test_robustness.py
```

## 跑
```bash
python3 scripts/allocate.py --input-dir /path/to/当月材料 --out /path/to/部门科目余额表.xlsx
```

依赖：`pip install openpyxl xlrd pandas`  
说明与安装提示词见部署包内《使用说明》。
