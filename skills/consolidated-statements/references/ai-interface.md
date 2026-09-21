# AI标准调用协议

Skill根目录由安装环境定位，记作SKILL_ROOT；不要把示例的/skill写死到其他机器。需要Python 3.11+和requirements.txt依赖。三个命令的stdout只输出一个JSON对象，进度在stderr，不含凭据。

```text
python -B <SKILL_ROOT>/scripts/skill_cli.py describe
python -B <SKILL_ROOT>/scripts/skill_cli.py check --request <本次请求.json> --evidence <本次抵销证据.json>
python -B <SKILL_ROOT>/scripts/skill_cli.py run --request <本次请求.json> --evidence <本次抵销证据.json>
```

先describe读取能力，再check，ready后run。check不写文件、不登录金蝶，只校验参数、证据结构、来源文件与执行条件；不代表财务验证通过。取数和证据核验按collection.md执行，JSON字段按execution.md准备。日期、路径由本次用户任务提供，不能沿用样例日期。源文件使用绝对路径；结果目录必须新建且与来源目录分离。

请求格式：
```json
{
  "parameters": {"period": "202608", "fetch_kingdee": false},
  "files": {"reports": [
    {"local_path": "/absolute/current-input.xlsx", "name": "原始文件名.xlsx",
     "file_id": "upload-id", "sha256": "来源文件真实SHA256"}
  ]},
  "output_dir": "/absolute/new-result",
  "source_metadata": {"confirmed_bs_hashes": ["已核实重分类口径的来源真实SHA256"]}
}
```

sha256可以省略，由脚本实算；不得使用示例占位值运行。confirmed_bs_hashes不能凭空填写。报告和抵销证据的完整字段见execution.md。

返回公共字段schema_version=1、skill、version、status。状态：
- ok：describe完成。
- ready：预检通过，尚未生成或验证报表。
- needs_input：preflight.issues列出缺资料或错误，退出码2；修复对应输入后才能执行。
- complete：result.summary.complete及elimination_complete同时通过；仍应报告实际范围和核验记录。
- partial：已有输出但来源缺失或检查未通过，退出码0；不能称完整财务结果，读取result.summary及warnings。
- failed：退出码1，error_code给出异常类别；不输出敏感异常文本。检查本次受控记录及输入；不自动重跑或覆盖结果。

每次执行独立进程，避免旧入口的模块替换在同进程并发。默认处理已取得的报表；fetch_kingdee仅在平台提供出站策略且已授权取数时使用，凭据通过stdin。抵销明细仍需Agent按取数操作获取并核验，不能仅凭本入口声称已自动采集完。三张单体报表及Agent采集的抵销明细均使用headless=True的Playwright浏览器，具体要求见collection.md。缺少无头浏览器或网络能力时只阻塞采集，不捏造金额。

现金流按测试底稿的公式结构：对应范围内部流出合计同时减收到其他与经营活动有关的现金、支付其他与经营活动有关的现金；本月及1月至目标月分别计算。金额以当期金蝶为准，样例300元差异不补差、不调整。此规则同样适用于母公司范围，不能把流入加入抵销额。

所有输出由run自动写入结果目录；保存stdout JSON作为正式任务执行回执时，使用新的文件路径，不能覆盖来源。核验JSON是正式业务依据，允许保留；临时下载副本、截图及中间脚本按任务清理。
