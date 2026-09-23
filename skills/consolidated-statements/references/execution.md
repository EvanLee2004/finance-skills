# 执行与证据格式

依赖：Python 3.11+、openpyxl、xlrd；仅联网采集需要playwright和已配置的浏览器/出站策略。运行脚本前用--help确认入口。平台原生沙箱通常把本包挂在/skill、输入挂在/workspace/inputs，正式工作与证据放/workspace/outputs，单条命令时限以实际工具为准。不要写只读Skill目录或假定命令结束后/tmp会保留。

## 一次完整运行

```text
python /skill/scripts/entry.py --period 202608 --input-dir /workspace/outputs/run-inputs --out-dir /workspace/outputs/run-result --evidence /workspace/outputs/elimination-evidence.json --confirmed-bs-hashes /workspace/outputs/confirmed-bs-hashes.json
```

日期及路径必须替换为当前任务值。run-result必须是新的空目录；原始输入不能放入结果目录。输入文件夹只放用户本次选定的单体来源，不扫描混有历史底稿的下载目录。随机文件ID先依上传清单还原原始文件名。confirmed-bs-hashes.json为已验证重分类口径的原始报表SHA256数组；不能为了绕过校验填写未经确认的哈希。

也支持--request读取原工具协议：parameters.period、parameters.fetch_kingdee、files.reports列表、output_dir、source_metadata.confirmed_bs_hashes。每个文件对象提供local_path、name、file_id、sha256。请求不含凭据；--fetch-kingdee只在已提供网络策略及stdin凭据时启用，不能假定原生沙箱直接支持。脚本复用原工具采集器，但抵销明细由Agent按collection.md取齐后提供证据。

## 抵销证据JSON

根字段：
- period：YYYYMM，与报表一致。
- ledger：固定HEAD，表示抵销明细来自北京本部。
- collected_at：本次核查时间。
- companies：按公司代码分组，仅内部对手方，不包含HEAD。
- 每个公司可有bs、is、cf。缺任一适用项时相应范围抵销不会被标为完成。

bs对象：status为observed/blank/no_data/failed，refs为非空证据引用列表。observed另提供ending_debit十进制字符串。blank表示亲眼核实栏目空白，计算按0但保留状态；no_data、failed是缺证据，不默认为0。

is对象：status=observed，accounts_complete=true，refs非空，accounts数组包含核对过的不同科目。每项有account、has_receipt_month、has_receipt_ytd、receipt_month、credit_ytd、refs。金额是含税贷方原值，不预先除1.06；代码按公司换算。若没有收票业务，标false并仍如实记累计贷方；完全无对应数据用status=no_data。重复科目会拒绝，以防每月累计被重复求和。

cf对象：status=observed，refs非空，month_in、month_out、ytd_in、ytd_out都是明确的金额字符串。经过完整筛选确认没有匹配行才写0，并在refs注明查询与无匹配证据。不能将取数失败写0。

refs应指向本次受控导出/核验材料及sheet/行号，或浏览器核查记录中的期间、公司、供应商编码、科目、日期、凭证和分录定位。支持后续核查，不放账号密码、Cookie或令牌。

以下只演示一个公司的格式，实际必须补齐纳入范围的所有内部公司和适用报表；金额为合成示例，不可直接用于业务：

```json
{
  "period": "202608",
  "ledger": "HEAD",
  "collected_at": "2026-09-17T12:00:00+08:00",
  "companies": {
    "CULTURE": {
      "bs": {"status": "observed", "ending_debit": "100.00", "refs": ["当期核验记录#期末借方"]},
      "is": {
        "status": "observed", "accounts_complete": true, "refs": ["当期核验记录#全部科目"],
        "accounts": [
          {"account": "113311", "has_receipt_month": true, "has_receipt_ytd": true,
           "receipt_month": "53.00", "credit_ytd": "106.00", "refs": ["当期核验记录#科目113311"]}
        ]
      },
      "cf": {"status": "observed", "month_in": "0.00", "month_out": "10.00",
             "ytd_in": "5.00", "ytd_out": "20.00", "refs": ["当期核验记录#四个入口"]}
    }
  }
}
```

### 零抵销公司怎么写（2026-09-23 实测踩坑）

某家公司在余额表里**完全没有往来科目行**（也不涉及收票、现金流）时，
**不要写 `status: no_data`** —— `eliminations.py` 会把它计入"抵销证据缺失或口径未确认"，
使 `elimination_complete` 恒为 false（实测：济南子公司如此；改成下面写法后即通过）。

正确写法：**`status: observed` + 金额 `0` / `accounts: []`**，并**单独出一份"零抵销确认"
证据文件**，写清"为什么是 0"（余额表全量行数、筛查方式、四入口公司名匹配数为 0），
在 refs 里引用它。这符合本文件"经过完整筛选确认没有匹配行才写 0"的要求——
是"查过了、确实没有"，不是取数失败。

### 交付后必做：分册 vs 底稿逐格对比

把 `合并报表_*.xlsx`、`母公司报表_*.xlsx` 与底稿的同名工作表**逐格比对
「单元格值 + 公式文本」**（金额按 6 位小数归一化），预期**数值差异 0**。
可接受的"写法差异"只有三类，且应在报告里逐条列明：分册不复制标题格（如 A1）、
公式内常数写法（`-325429.3` vs `-325429.30`）、浮点显示位差
（`351608.079999999` vs `351608.08`）。

⚠️ **比对前不要用外部编辑器/预览器打开交付文件**：实测某次预览回写后，
合并分册 709 个公式全部被改写成 `<f>=...`（多一个前导等号；OOXML 规范不带等号，
配合工作簿里的 `fullCalcOnLoad` 有重算报错风险），而底稿与母公司分册是干净的——
即该改写与技能无关，是外部写回造成的。要比对公式写法，以技能刚产出的文件为准；
若发现 `<f>=`，用字符串级替换去掉前导等号即可复原（注意底稿用的是 `ns0:` 前缀 XML，
正则要兼容命名空间前缀）。

### 小企业准则「其中：利息费用」的口径（不要当异常）

小企业准则利润表第 19 行「其中：利息费用（收入以"-"号填列）」是**净利息**，
而「财务费用」= 净利息 **+ 手续费等**，两者**本就不相等**，属正常口径：
实例（202608）济南 财务费用 −179.97 = 净利息 −377.16 + 手续费 197.19；
湖南分公司 424.29 = −43.71 + 468.00。**按来源原值保留，不归零、不挂异常提示。**

另注意这类外部 `.xls` 的表头列序与金蝶标准导出相反：
`行次｜本年累计金额｜本月金额`（标准导出是 `行次｜本月金额｜本年累计金额`），
取数时按表头文字辨认列，不要按位置猜。

## 结果状态

输出包含完整底稿、两个分册和报表来源与核验记录.json。JSON保留原始来源、范围、抵销前后公式、逐项证据、检查结果。summary.complete同时要求原始24份报表和抵销核验完成；elimination_complete仅表示已纳入范围抵销证据与检查通过，不能替代coverage_complete。有缺失使用_部分，资料齐全但检查未通过使用_待核实。

失败后检查原因再生成新的结果目录；不要在已抵销文件上追加一次抵销。tests/test_eliminations.py使用合成报表验证范围、重复累计拒绝、换算、合计公式与分册，无需金蝶登录：
```text
python -B -m unittest discover -s tests -p test_eliminations.py
```
