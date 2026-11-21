# cardo-invest

> 安装（二选一）：
> `dsh plugin --profile <profile> add github:CSCCTang/cardo-invest#v0.1.1`
> `dsh plugin --profile <profile> add <本地路径>\cardo-invest-0.1.1.tgz`

投资决策：给一张财务参数表，产出投资匡算、财务评价、三方案比选与《项目投资决策建议书》。

## 需要提供什么

| 输入 | 形态 | 最小样例 |
|---|---|---|
| 财务参数 | `输入文件\财务参数.xlsx`（「工程费」表至少 1 行：名称 + 金额(万元)） | `skills\cardo-invest\assets\inputs\财务参数-空白.xlsx` |
| 工作区参数 | `工作区参数.json`（至少含 `narrative.planningBasis`） | 首次运行会生成骨架，照着填即可 |
| 槽位表（可选） | `输入文件\槽位.xlsx` | `skills\cardo-invest\assets\inputs\槽位-空白.xlsx` |
| 上游产物（可选） | 现状指标/事实库/AHP 结果，放 `输入文件\` 任意层级 | — |

## 装完怎么调

在 dsh 对话里说「用 cardo-invest 跑，工作区是 `D:\proj\invest-01`」。

工作区是独立空文件夹：`输入文件\` 放料、`过程文件\` 留痕、`成果文件\` 出活。连跑第二次建议换新目录。

## 会得到什么

`成果文件\`：

- 《项目投资决策建议书》docx + PDF（10 章）
- 投资估算表（xlsx）
- 财务评价表（xlsx）

## 许可证与版本

- 版本 `0.1.1`；许可证 **AGPL-3.0-or-later**，见 [LICENSE](LICENSE) 与 [NOTICE](NOTICE)。
- 第三方组件与许可证见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。
- 运行环境：Windows；Python ≥ 3.11；dsh ≥ 0.2.2-rc.5。

## 文档

- [docs/01_使用流程.md](docs/01_使用流程.md)
- [docs/02_输入准备.md](docs/02_输入准备.md)
- [docs/03_常见失败.md](docs/03_常见失败.md)
