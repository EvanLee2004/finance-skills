# compliance-spot-check · 合规文件抽查

从应收 all（+ 可选抽查历史）生成**本周抽查建议清单**（文字版，可粘进抽查表）。

- 入口：`scripts/recommend.py`
- 配置：`config/抽查规则.md`、`config/列名别名.json`
- 回归：`python3 tests/test_robustness.py`
- 对人说明：见 `SKILL.md`

设计依据：本地 `技能/合规文件抽查/方案与文档/20260623_合规文件抽查_技能方案.md`。
