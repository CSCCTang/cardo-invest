---
name: cardo-invest-02-匡算
description: 投资匡算：按财务口径把工程费、工程建设其他费、预备费、建设期利息汇总为总投资，并给出资金筹措（自有/借款/补助）与分年使用计划；Decimal 计算，缺数只报缺项不补数。
---

# 02 · 匡算

## 1. 职责

```text
总投资 = 工程费 + 工程建设其他费 + 预备费 + 建设期利息
预备费 = (工程费 + 其他费) × 预备费率
建设期利息 = Σ(当年提款 × 年利率 × 0.5)   # 口径：不计复利，只对当年提款半额计息
# 例外：输入给了 estimate.interest.amount（如可研附表1 的建设期融资费用）→ 直接采用，不按利率反推
资金筹措 = 资本金 + 借款 + 政策性补助（与总投资容差 ≤1 万元）
```

## 2. 输入 / 输出

| 方向 | 路径 | 契约 |
|---|---|---|
| 读 | `过程文件/00-总控/引擎项目/05_过程/投资决策/财务输入.json` | `cardo/finance-input/v1#estimate/finance` |
| 读 | `规则/财务口径.yaml`（本技能自带副本，sha256 `cc78e252…9b0b`） | 费率/计息/取整/容差唯一来源 |
| 写 | `过程文件/00-总控/引擎项目/05_过程/投资决策/投资估算.json` | `cardo/invest-estimate/v1` |
| 写 | `过程文件/02-匡算/投资估算.json` | 同上（总控复制，供 05/08 只读消费） |

产物字段：`engineering / otherFees / otherFeeTotal / contingency / constructionInterest / constructionInvestment /
totalInvestment / funding / fundingTotal / plan[] / tables{} / trace / gaps[] / status / note`；
金额单位万元、`Decimal` 两位；比率 % 两位。

`tables`（派生块，只由已算值派生）：`composition`（投资构成，含占总投资比例）、`yearPlan`（分年投资计划，
末年吸收尾差）、`fundSource`（资金来源与平衡差额）。

## 3. 命令

```powershell
python Skills/02-匡算/scripts/calc_invest.py --help
python Skills/02-匡算/scripts/calc_invest.py --project-dir 过程文件/00-总控/引擎项目 [--format json]
python Skills/02-匡算/scripts/calc_invest.py --input <财务输入.json> --out <输出目录> --format json --no-timestamp
```

退出码：`0` 算出；`1` 数值非法（负值/非数值/建设期年数≤0）；`2` 输入/口径缺失；`3` 待确认（gaps 非空，产物仍落盘）。

## 4. 闸门

1. `totalInvestment` 非空（算不出即 `note=缺项，非完整总投资`，总控判失败）；
2. 明细之和 = 大类合计 = 建设投资，逐级可加总（零容差）；
3. 资金筹措合计 − 总投资 ≤ 1 万元；
4. `trace` 含 `ruleVersion/engineVersion/inputFingerprint/rulesSha256/steps[]`。

## 5. 规则

1. **Decimal 全程**：禁用 float 累加，产物不得出现浮点尾差；
2. **缺数不补**：分项缺单价/工程量 → 该项入 gaps、大类合计置 null，不补 0；
3. **口径不从代码来**：费率/计息系数/取整只读 `规则/财务口径.yaml`；
4. **可追溯**：每个分项带公式与口径来源锚点（`ruleSource`）；
5. **不拍板**：取费标准争议 → gaps + exit 3，不自行选一套继续算。

## 6. 调优点

- 取费/费率/预备费：改 `规则/财务口径.yaml`（或财务参数 xlsx，改后者不用动规则）。
- 分年分摊：`tables.yearPlan` 按均匀分摊、末年吸收尾差；要按形象进度分摊需改本技能脚本（属口径变更）。
