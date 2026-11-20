---
name: cardo-invest-03-财务评价
description: 财务评价：算收入、经营成本、折旧摊销、增值税及附加、所得税、土地使用税、印花税、房产税，生成全部投资/资本金两套现金流，算 FIRR/FNPV/静态回收期/ROE/ROI/ICR/DSCR 并留计算轨迹；指标从成分现金流重算，不读模板结果单元格。
---

# 03 · 财务评价

## 1. 职责

两段脚本、一条流水线：

```text
calc_finance.py  → 收入 / 经营成本 / 折旧摊销 / 税费 → 两套现金流底表
calc_metrics.py  → 由现金流底表算 10 项指标 + 计算轨迹 + 派生表 + 敏感性
```

## 2. 输入 / 输出

| 方向 | 路径 | 契约 |
|---|---|---|
| 读 | `…/05_过程/投资决策/财务输入.json` | `cardo/finance-input/v1` |
| 读 | `…/05_过程/投资决策/投资估算.json` | `cardo/invest-estimate/v1` |
| 读 | `规则/财务口径.yaml`（本技能自带副本） | 公式/税率/取整/容差/对账锚点 |
| 写 | `…/05_过程/投资决策/财务测算.json` | `cardo/finance/v1` |
| 写 | `…/05_过程/投资决策/计算轨迹.json` | `cardo/finance-trace/v1` |
| 写（总控复制） | `过程文件/03-财务评价/财务测算.json`、`计算轨迹.json` | 供 05/08 只读 |

指标键（固定 10 项）：`firrAll / fnpvAll / ptAll / firrEquity / fnpvEquity / ptEquity / roe / roi / icr / dscr`。

派生块：`tables.costStructure / fullCostStructure / revenueStructure / cashflows(全部投资、资本金) / debtSchedule`、
`sensitivity`（投资/收入/经营成本 × −10%/0/+10% 单因素）。

## 3. 命令

```powershell
python Skills/03-财务评价/scripts/calc_finance.py --project-dir 过程文件/00-总控/引擎项目 --no-timestamp
python Skills/03-财务评价/scripts/calc_metrics.py --project-dir 过程文件/00-总控/引擎项目 --no-timestamp
```

退出码：`0` 算出；`1` 校验失败；`2` 输入/口径缺失；`3` 口径待确认（税率非法/除零/负现金流）。

## 4. 闸门

1. 10 项指标全部非空（缺项即失败，不拿 0 顶替）；
2. `tables.cashflows.allInvestment.rows` 33 行存在（表位 1/2 的数据源）；
3. 计算轨迹含输入指纹与每一步公式；
4. 折现系数 = 1/(1+r)^t（08 独立复算会抽检前 3 年）。

## 5. 规则

1. 金额 `Decimal`、比率 %，比率两位；IRR 容差 1e-8；
2. 缺数不补：收入项缺单价/数量 → gaps + 指标 status=gap；
3. 非法输入 exit 3：税率非法、除零、计算期不符，不静默修正；
4. **反作弊**：不读模板结果单元格；对账锚点只是校验真值，不是输入；
5. **幂等**：`--no-timestamp` 两次跑 JSON sha256 一致；
6. 口径唯一来源：`规则/财务口径.yaml`（改口径 = 版本 +1 + 全链回归）。

## 6. 对账锚点（模板回归，另案）

| 项 | 模板真值 | 容差 |
|---|---|---|
| 全部投资 FIRR / FNPV / PT | 7.13% / 11918 / 15.93 | <0.5% / ≤1 万元 / <0.5% |
| 资本金 FIRR / FNPV / PT | 13.72% / 13563 / 10.62 | 同上 |
| ROE / ROI | 12.7% / 8.0% | ±1pp（可复算 + 差异解释） |

> 本工作区基线以 `记忆与规则/验证基线-CardoInvest.json` 为准；本项目报告表位 1/2 采用可研口径
> （附表8：3.24%/1,397/12.31；附表9：8.19%/628/4.09），引擎按 13,133.78 万元、32 年、2.7% 复算得
> 2.88%/321.42/21.06，在 8.1.5.1 作对照；两套不混用，正式申报以可研批复为准。

## 7. 调优点

- 税率/折旧年限/折现率：改 `规则/财务口径.yaml`；
- 收入项/成本构成：改 `输入文件/财务参数.xlsx`；
- 敏感性区间：改 `calc_metrics.py` 的 `sensitivity` 段（口径变更需登记）。
