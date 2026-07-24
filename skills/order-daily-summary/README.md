# order-daily-summary · 九点下单统计

自动登录智云抓「下单」表 → 组织架构归 **多语（不含运保）/数据/游戏/其他** → 输出「下单数据(万元)」xlsx。

> **⚠ 这是「9 点快照」口径，别把晚跑的数字差当 bug。**
> 智云「下单日期」可被编辑：昨天的订单当天可能被改期/改单号（如某单从 7-23 挪到 7-24，SO 随之重排），
> 所以「昨日下单额」只在**某个时刻**确定。设计就是**每天 9:00 定时跑**取当日快照；晚跑（>09:15）会自动
> 在产物和 stdout 标红提示"数据截至 X 时、可能已改期"。同一天不同时刻跑得数不同 = 数据动了，**不是取数错**。
> 抓取行数与服务端声明不符会**直接报错**（宁可失败也不静默少算）。

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
