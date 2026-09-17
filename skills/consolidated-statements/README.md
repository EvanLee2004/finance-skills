# 合并报表与抵销

版本2026.09.17.2。执行流程见[SKILL.md](SKILL.md)，AI标准调用参数和返回状态见[调用协议](references/ai-interface.md)。

本版生成六表底稿及两个分册，包含按金蝶证据计算的内部抵销。先使用scripts/skill_cli.py check检查本次请求和抵销证据，再使用run执行；旧版无证据的调用参数不适用于本版。安装依赖见requirements.txt。

金蝶采集需要已授权浏览器或平台出站策略；凭据不存入仓库。默认可以处理已取得的报表，抵销明细按取数操作由Agent核验。
