# 九点下单统计 · SOP

## 前置

1. 公司内网或 VPN，能打开 `http://192.168.10.167:18880`
2. Python ≥3.10；`pip install playwright requests openpyxl`；`playwright install chromium`
3. 本机 `config.local.json`（从 `config/config.local.example.json` 复制）或环境变量 `ZHIYUN_USER` / `ZHIYUN_PASSWORD`

## 运行

```bash
# <本skill目录> = 本文件所在 skill 根目录的绝对路径
python3 "<本skill目录>/scripts/run.py" \
  --out "<交付目录绝对路径>" \
  [--today YYYY-MM-DD] \
  [--config "<config.local.json 绝对路径>"] \
  [--detail]
```

## 交付核对

- 打开「下单数据」：日期行、分部门万元、总计
- 打开「处理日志」：窗口是否正确、未匹配是否为空
- 有未匹配 → 改 `config/销售组织架构.xlsx` 后重跑

## 定时（可选后续）

可在部署机用 cron/launchd 工作日 09:00 触发；v1 不自带定时配置。
