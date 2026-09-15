# 平台 Skill 同步版本

此目录保存 2026-09-15 当前 Ubuntu 部署源码中的近期修改 Skill，来源提交和逐文件 SHA-256 见 [manifest.json](manifest.json)。

| 目录 | 工具 |
|---|---|
| ar-hexiao-daily | 应收核销标准版 |
| ar-hexiao-daily-lab | 应收核销极速版 |
| consolidated-statements | 合并报表 |
| receivables-merge-and-split | 应收合并与销售拆分 |

每个目录按平台原有结构完整同步，包括 tool.yaml、入口、业务脚本、配置和已有测试。核销仍依赖财务平台的编排服务，入口脚本自身不提供独立执行能力；合并报表依赖平台取数能力。使用时须配合来源提交对应的平台服务与依赖，凭据通过部署环境配置。

source-patches/ar-hexiao-daily 保存同一来源提交中 sources/finance-skills 内单独更新的智云取数脚本及两项回归测试，用于追踪源码副本差异，不是完整安装包。

仓库原有 skills/ 是同事独立安装版本，其功能和更新白名单继续沿用原来的流程。平台版集中在本目录，避免同名目录覆盖导致独立版丢失功能。此次同步不表示已完成独立版接口迁移，也不自动更新同事本机。

本次只同步源码和规则，不包含运行凭据、业务输入输出或数据库。同步校验覆盖文件哈希、Python/JSON 语法和 diff；不执行真实核销或财务写入。
