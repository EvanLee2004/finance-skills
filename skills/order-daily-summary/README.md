# order-daily-summary · 九点下单统计

自动登录智云抓「下单」表 → 组织架构归 **多语（不含运保）/数据/游戏/其他** → 输出「下单数据(万元)」xlsx。

## 快速跑

```bash
# 1. 凭据（二选一）
cp config/config.local.example.json config.local.json   # 再编辑填账号密码
# 或 export ZHIYUN_USER=… ZHIYUN_PASSWORD=…

# 2. 依赖
pip install playwright requests openpyxl
playwright install chromium

# 3. 运行（需内网）
python3 scripts/run.py --out ./工作区 --detail
```

## 测试

```bash
python -m pytest skills/order-daily-summary/tests -q
```

## 结构

- `scripts/run.py` — CLI 入口
- `scripts/date_window.py` / `summarize.py` / `fetch_orders.py` / `write_report.py`
- `config/销售组织架构.xlsx` + `业务规则.md`
- `tests/` — 纯函数单测（不连网）
