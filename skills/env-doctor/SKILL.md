---
name: env-doctor
description: >-
  财务技能包的「环境管家」：当任何技能跑不起来、报缺依赖时（ModuleNotFoundError、No module named、找不到 soffice/tesseract/poppler、pip 装不上、Python 版本太老），
  查本技能的《依赖与安装清单》拿到对应安装命令（**国内镜像优先、国外备选**）把环境装齐、再重试原技能。
  当用户说「配下环境 / 把依赖装齐 / 这个技能跑不起来缺东西 / 装一下需要的库 / 环境自检」，或任何技能脚本因缺库/缺系统工具报错时，主动使用本技能。
---

# 环境管家（env-doctor）

> **研发说明** · 甲骨易财务部自研基础设施技能（研发维护：李明昊，2026-06）。它不处理业务、不碰数据，只做一件事：**让别的技能能跑起来**——把它们需要的 Python 库和系统工具按"国内镜像优先"装齐。**纯提示词驱动，无脚本**：所有"缺什么→怎么装"的知识都在 `config/依赖与安装清单.md` 这张活表里，加了新技能就改这张表。

## 你（agent）该怎么干

**核心原则：用它的是财务同事，不是技术员。** 技能报了一串看不懂的英文报错（缺库/缺工具），不要把报错甩给她——**你自己**对照清单把缺的装上、重试，跑通了再用人话告诉她"刚才差个组件，我装好了，继续"。

### 两个触发时机

1. **用户主动说**「配下环境 / 把依赖装齐」→ 照 `config/依赖与安装清单.md` 的**「一把全装」命令**(国内镜像)把 Python 库装齐；系统工具(LibreOffice 等)按她的系统(Win/Mac)给的命令装。逐项确认装上了，汇报。

2. **任何技能报环境错时（最常见）** → 你**自动**接手，不用等她开口：
   - 从报错认出缺什么：`ModuleNotFoundError: No module named 'X'` → 缺 Python 库 X；`找不到 soffice / command not found: soffice` → 缺 LibreOffice；`Unable to get page count / poppler` → 缺 poppler；`tesseract is not installed` → 缺 tesseract。
   - **查 `config/依赖与安装清单.md`** 找到 X 对应的安装命令。
   - **国内镜像优先**装（清华 PyPI / 淘宝 npm / 国内镜像下载）；**装失败再退国外源**。
   - 装完**重试刚才那个技能**，跑通就继续，别再回头打扰她。

### 怎么装（铁律）

- **国内镜像优先、国外备选都在清单里**。Python 库一律先用 `pip install -i https://pypi.tuna.tsinghua.edu.cn/simple <包>`；清华不通再换阿里 `https://mirrors.aliyun.com/pypi/simple`；都不行最后退默认源 `pip install <包>`。
- **能一把装就别一个个装**：用清单里的「一把全装」命令把 Python 库一次装齐，省得来回。
- **系统工具(LibreOffice / poppler / tesseract / Node)** 按用户系统给命令(Win 用 winget/scoop 或国内镜像下载链接；Mac 用 brew)。这类是系统级安装，**能自动装就装，需要下载安装包/点同意的就把国内下载链接给用户、引导她装**——别卡死、别瞎来。
- **`import 名 ≠ pip 包名`**：报错里说缺 `docx` 要装的是 **python-docx**、缺 `pptx` 装 **python-pptx**、缺 `PIL` 装 **Pillow**。清单里有完整映射，照它来，别想当然 `pip install docx`(那是错的包)。

### 安全红线

- 只装/配环境，**绝不读、不碰任何业务数据**（应收表、发票、工资等）。
- 只装清单里列明的包/工具，**不随便装来路不明的东西**；清单里没有的报错，先看是不是本地模块(如 `validators`、`helpers`、`office` 是技能自带的本地代码，不是要 pip 装的包)，拿不准就停下问李明昊。

## 清单在哪

完整的「哪个技能要什么、怎么装(国内/国外)」见同目录 `config/依赖与安装清单.md`——**那是这张活表的本体**，本文件只讲"怎么用它"。加了新技能，把新依赖补进那张表即可，本文件不用改。
