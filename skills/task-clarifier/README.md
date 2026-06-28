# task-clarifier · 理清需求（先问清再动手）

财务部**行为型技能**（纯提示词、无脚本）。给不太会表达需求的财务同事用：需求含糊 / 有多种理解时，agent **先问 1~5 个带选项的问题把需求问明白、再动手**，绝不靠猜；信息不全就要求补充。理清后直接接现成财务技能执行。

- 触发：用户说「帮我理清需求 / 我不知道咋说」，或要跑某技能但信息不全、含糊时；**也可在提示词开头粘一段固定"开场白"主动触发**（见使用手册）。
- 核心铁律见 `SKILL.md`。
- **改编自 [trailofbits/skills](https://github.com/trailofbits/skills) 的 `ask-questions-if-underspecified`（CC BY-SA 4.0, © Trail of Bits）；本技能文件同以 CC BY-SA 4.0 共享。**

维护：李明昊，2026-06。
