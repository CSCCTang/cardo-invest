# 第三方依赖与许可证清单（cardo-invest）

只列本包 `requirements.txt` 实际声明的依赖；未在此列的组件不由本包引入。

| 依赖 | 用途 | 许可证 |
|---|---|---|
| python-docx | 生成建议书 docx | MIT |
| openpyxl | 投资估算表 / 财务评价表 xlsx | MIT |
| Pillow | 图片与版面处理 | HPND（BSD 兼容） |

## 说明

- 本包**不依赖 PyMuPDF**；PyMuPDF 只出现在 `cardo-status`。
- 本包不包含任何第三方二进制、wheel 或 Python 运行时；依赖由使用者从 PyPI 安装。
- 四个包统一以 AGPL-3.0-or-later 发布；本包依赖全为宽松许可，如需闭源分发可自行改为 MIT 并保留上表声明。
