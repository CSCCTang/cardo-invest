"""01b · 可研/匡算财务提取 —— 从提资材料自动生成并**逐步补全** 财务参数.xlsx。

机制（与用户要求一致）：
  * `输入文件/财务参数.xlsx` **始终存在**：01b 先用技能包里的空白模板建出它，
    再把从材料里抽到的值**原地写回**（用户填过的格子不覆盖）。
  * 后续步骤若发现缺项，再次运行本脚本即可增量补全 —— 越跑越全。
  * 中间台账放 `过程文件/01b-可研财务提取/`，输入文件夹里只留财务参数.xlsx
    与它的空白模板。

用法（工作区根执行）：
    python <技能目录>/Skills/01b-可研财务提取/scripts/build_params.py --workspace .
    python <技能目录>/Skills/01b-可研财务提取/scripts/build_params.py --workspace . --dry-run
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from extract_feasibility import (  # noqa: E402
    clean_line, detect_project_direction, extract_contingency_rate,
    extract_economic_indicators, extract_funding,
    extract_headline_investment, extract_project_meta, find_pages, fmt,
    has_revenue_section, to_decimal,
)
from extract_appendix import extract_from_appendix_workbook  # noqa: E402
from extract_tables import extract_from_workbook  # noqa: E402
from params_table import (  # noqa: E402
    FILLED_NAME, SHEETS, ensure_blank_template, fix_sheet_rows, normalize_text,
    project_name_from_dir, validate_budget_balance,
)
from pdf_tables import (  # noqa: E402
    drop_parents, find_estimate_pages, parse_estimate_page,
)
# 写出逻辑见 params_writer.py（写表 + 保留用户已填值 + 幂等清表）
from params_writer import write_params  # noqa: E402

SCHEMA = "cardo/invest-params-prefill/v1"

# ---------------------------------------------------------------- 输入源发现

# 财务数据源：可研 PDF（原始设计文件）与投资估算类 xlsx
_FINANCE_PDF = ("*可研*.pdf", "*可行性研究报告*.pdf", "*项目建议书*.pdf")
_FINANCE_XLSX = ("*匡算*.xlsx", "*投资估算*.xlsx", "*附表*.xlsx")
# 成果文件里的匡算/财务评价工作簿（附表1/3/4/5/7）——比 PDF 好抽得多
_APPENDIX_XLSX = ("*附表*.xlsx", "*匡算*.xlsx", "*投资估算*.xlsx", "*财务评价*.xlsx")

# 上游成果（其他三条线的产出）—— 只登记，不解析
UPSTREAM_KINDS = {
    "现状分析报告": ("*现状分析报告*.docx",),
    "AHP权重表": ("*AHP权重表*.xlsx",),
    "AHP方案评分结论": ("*AHP方案评分结论*.xlsx",),
    "AHP权重与评分结论": ("*AHP权重与评分结论*.docx",),
    "现状指标": ("*现状指标*.xlsx",),
    "数据缺口清单": ("*数据缺口清单*.xlsx",),
    "估算开项数据采集要求": ("*估算开项数据采集要求*.xlsx",),
    "建议书模板": ("*建议书模板*.docx",),
}


def discover_inputs(workspace):
    """扫描 工作区，归类财务源、上游成果、附录工作簿。

    附录工作簿在**成果文件**里（匡算表），不在「输入文件」——
    实测各项目的成果目录结构不同（某企业在 成果/，某新区在项目子目录），
    因此整个工作区都扫。
    """
    root = Path(workspace) / "输入文件"
    finance = []
    appendix = []
    upstream = {}
    if not root.is_dir():
        return finance, upstream, appendix
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        for label, patterns in UPSTREAM_KINDS.items():
            if any(path.match(pattern) for pattern in patterns):
                upstream.setdefault(label, []).append(path)
        if path.suffix.lower() == ".pdf" and any(path.match(p) for p in _FINANCE_PDF):
            finance.append(path)
        elif path.suffix.lower() == ".xlsx" and any(path.match(p) for p in _FINANCE_XLSX):
            finance.append(path)
    # 成果目录里的匡算/附表工作簿（跳过我们自己生成的财务参数.xlsx）
    workspace_root = Path(workspace)
    for path in sorted(workspace_root.rglob("*.xlsx")):
        if path.name.startswith("财务参数"):
            continue
        if "成果" not in path.parts and path.parent.name != "输入文件":
            continue
        if any(path.match(p) for p in _APPENDIX_XLSX):
            appendix.append(path)
    return finance, upstream, appendix


# ---------------------------------------------------------------- 抽取

def extract_from_pdf(path):
    """从可研 PDF 抽：投资估算附表明细 + 结论句 + 资金来源 + 项目字段。"""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return None
    label = path.name
    try:
        doc = pymupdf.open(path)
    except Exception:
        return None
    try:
        pages = find_estimate_pages(doc)
        layout_categories = {"engineering": [], "otherFees": [], "contingency": [], "total": []}
        current = None
        layout = None
        for page_no in pages:
            parsed, current, layout = parse_estimate_page(
                doc[page_no - 1], label, page_no, current, layout)
            for key, values in parsed.items():
                layout_categories[key].extend(values)

        detail = {}
        category_rows = {}
        for key, values in layout_categories.items():
            items, cats = drop_parents(values)
            detail[key] = items
            category_rows[key] = cats

        invest_pages = find_pages(doc, ("投资估算", "投资概算", "工程费用", "总概算"))
        fund_pages = find_pages(doc, ("资金筹措", "资金来源", "融资方案"))
        headline = extract_headline_investment(doc, invest_pages, label)
        funding = extract_funding(doc, fund_pages, label)
        economic = extract_economic_indicators(doc, label)
        contingency_rate = extract_contingency_rate(doc, label)
        meta = extract_project_meta(doc, label)
        has_rev, has_cost = has_revenue_section(doc)
        return {
            "file": label, "kind": "pdf", "pages": doc.page_count,
            "projectDirName": "",
            "detail": detail, "categoryRows": category_rows,
            "headline": headline, "funding": funding, "meta": meta,
            "economic": economic, "contingencyRateText": contingency_rate,
            "direction": detect_project_direction(doc),
            "hasRevenue": has_rev, "hasCost": has_cost,
            "tablePages": pages[:12],
        }
    finally:
        doc.close()


def extract_from_xlsx(path):
    """从投资估算类 xlsx 抽（附表形态的）。"""
    data = extract_from_workbook(path)
    if not data:
        return None
    data["kind"] = "xlsx"
    detail = {}
    for key in ("engineering", "otherFees", "contingency"):
        items = data.get("detail", {}).get(key) or []
        detail[key] = items
    data["detail"] = detail
    return data


# ---------------------------------------------------------------- 组装 9 张表



def collect(sources):
    """把各来源汇总成 {sheet: {key: {value, source}}}。

    来源有两类：
      * pdf 源 —— 可研 PDF 抽出的附表明细与正文结论
      * xlsx 源 —— 成果文件里的匡算工作簿（附表1/3/4/5），
        **优先级高于 PDF**（Excel 本身就是结构化表格，不必再从排版 PDF 里猜）
    """
    table = {}
    provenance = []
    warnings = []

    def put(sheet, key, value, source):
        bucket = table.setdefault(sheet, {})
        if key in bucket:
            return
        bucket[key] = {"value": value, "source": source}
        provenance.append({"sheet": sheet, "key": key, "source": source})

    meta = {}
    direction = None
    for data in sources:
        if data.get("kind") == "pdf":
            for key, item in (data.get("meta") or {}).items():
                if not key.startswith("_"):
                    meta.setdefault(key, item)
            direction = direction or data.get("direction")

    # 项目名称：**项目目录名优先**（提资方自己的命名，最权威），
    # 其次 PDF 文件名，最后正文候选。
    # 填错比空着更糟 —— 错名会被当成事实引用。
    # （实测某企业片区目录里放着别条道路的可研，按 PDF 文件名取会得到
    #   「某路（A路-B路）市政道路工程项目」）
    for data in sources:
        dir_name = data.get("projectDirName")
        if not dir_name:
            continue
        name = project_name_from_dir(dir_name)
        if not name:
            continue
        source = {"file": dir_name, "sheet": "（项目目录名）", "row": None,
                  "quote": dir_name}
        put("项目", "项目名称", name, source)
        break

    # 项目表。**抽不到的字段整项不写**（而不是写 None）：
    # 下游 03 财务评价用 `constructionYears or DEFAULT` 取默认建设期，
    # 若这里写 None，下游 `int(None)` 直接崩（实测某企业可研无建设期表述）。
    # 写成"没有这个字段"→ 下游用默认值 → 链路不中断，且缺项已进缺项表。
    for key, label in (("projectName", "项目名称"), ("rebuildArea", "改造面积(㎡)"),
                       ("constructionYears", "建设期(年)"), ("operationYears", "运营期(年)")):
        item = meta.get(key)
        if item and item.get("value") is not None:
            put("项目", label, item["value"], item["source"])
    # 规模（面积）：匡算工作簿的「建设内容与规模表」/「对标分析」
    for data in sources:
        scope = data.get("scope") or {}
        if not scope:
            continue
        for key, label in (("siteArea", "改造面积(㎡)"), ("buildingArea", "改造面积(㎡)")):
            if key in scope and label not in table.setdefault("项目", {}):
                put("项目", label, scope[key], data.get("source") or {
                    "file": data.get("file", ""), "sheet": "建设内容与规模表",
                    "row": None, "quote": "{} {}".format(key, scope[key])})
    if direction:
        put("项目", "项目方向", direction, {
            "file": "（自动识别）", "sheet": None, "row": None,
            "quote": "按可研正文「老旧小区/老旧厂区/老旧街区」出现次数判定"})

    # 工程费 / 其他费。**xlsx 源优先**（匡算工作簿是结构化表，比 PDF 可靠）
    ordered = [d for d in sources if d.get("kind") == "xlsx"] + \
              [d for d in sources if d.get("kind") != "xlsx"]
    # 逐条补 source：匡算工作簿抽出的条目只有名称与金额，出处用工作簿级兜底
    for data in ordered:
        fallback = data.get("source") or {
            "file": data.get("file", ""), "sheet": "（工作簿）", "row": None,
            "quote": data.get("file", "")}
        for key in ("engineering", "otherFees", "contingency", "interest", "total"):
            for item in (data.get("detail") or {}).get(key) or []:
                item.setdefault("source", fallback)
        for item in data.get("revenue") or []:
            item.setdefault("source", fallback)
    for data in ordered:
        detail = data.get("detail") or {}
        for key, sheet in (("engineering", "工程费"), ("otherFees", "其他费")):
            bucket = table.setdefault(sheet, {})
            for item in detail.get(key) or []:
                name = item["name"]
                if name in bucket:
                    continue
                # 匡算工作簿来源的条目没有逐行 source → 用工作簿级来源兜底
                source = item.get("source") or data.get("source") or {
                    "file": data.get("file", ""), "sheet": "（工作簿）",
                    "row": None, "quote": name}
                bucket[name] = {"value": item["amount"], "source": source}
                provenance.append({"sheet": sheet, "key": name, "source": source})
    # 有些可研只给了「投资构成汇总表」（工程费用/工程建设其他费用/预备费/总投资
    # 四行），明细在别处甚至没有（实测某区）。这时**用大类值兜底**，
    # 保证下游能算总投资 —— 但出处要写明是汇总行，不是明细行。
    for data in sources:
        rows_by_cat = data.get("categoryRows") or {}
        for key, sheet in (("engineering", "工程费"), ("otherFees", "其他费")):
            bucket = table.setdefault(sheet, {})
            if bucket:
                continue
            for item in rows_by_cat.get(key) or []:
                name = item.get("name") or ""
                if not name:
                    continue
                bucket[name] = {"value": item["amount"], "source": item["source"]}
                provenance.append({"sheet": sheet, "key": name,
                                   "source": item["source"]})

    # 其他费会从两处合并：手工参数表里的汇总项（建设管理费 537.59…）与
    # 匡算附表的明细（建设单位管理费 277.72 + 监理 259.87 = 537.59）。
    # 用附表「二 工程建设其他费用」的权威合计做精确匹配，只留一组，避免重复计费
    # （本项目不处理会 2878.53 → 4342.42，总投资虚增 1610.28 万）。
    oth = table.get("其他费")
    if oth:
        from extract_tables import pick_by_total
        target = None
        for data in sources:
            for row in (data.get("categoryRows") or {}).get("otherFees") or []:
                target = row["amount"]
                break
            if target:
                break
        rows = [{"name": name, "amount": cell["value"], "source": cell.get("source")}
                for name, cell in oth.items()]
        table["其他费"] = {row["name"]: {"value": row["amount"],
                                       "source": row.get("source")}
                          for row in pick_by_total(rows, target)}

    # 预备费：取费率（金额由 02 按口径算）
    for data in sources:
        rate = data.get("contingencyRate")
        if rate and "预备费率" not in table.setdefault("预备费", {}):
            put("预备费", "预备费率", rate["rate"], rate["source"])
    # 预备费率：**优先用金额反算**（附表自洽，最可靠）：
    #   预备费率 = 预备费金额 ÷ (工程费用 + 工程建设其他费用)
    #   实测某企业 392.54/(4301.76+605.70)=8%、某项目 268.07/(2966.38+384.49)=8%、
    #        某项目 528.85/(5305.06+1305.60)=8%，三个都还原为 8%。
    # 不用「正文/附表里的百分数」：那里写的是**占比**（7.41%）不是费率，
    # 混用会把 8% 读成 7.41%。
    engineering_total = other_total = contingency_amount = None
    for data in sources:
        rows = data.get("categoryRows") or {}
        for item in rows.get("engineering") or []:
            name = item.get("name") or ""
            if ("工程费" in name or "建筑安装工程费" in name) and "其他" not in name:
                engineering_total = item["amount"]
                break
        for item in rows.get("otherFees") or []:
            # 大类名可能是「工程建设其他费用」「工程建设其它费」等变体，
            # 统一按「含其他费」判断（写死全名会漏，导致反算费率偏小）。
            if "其他费" in (item.get("name") or "") or "其它费" in (item.get("name") or ""):
                other_total = item["amount"]
                break
        for item in rows.get("contingency") or []:
            if "预备费" in (item.get("name") or ""):
                contingency_amount = item["amount"]
    if (engineering_total and other_total and contingency_amount
            and (engineering_total + other_total) > 0):
        rate = contingency_amount / (engineering_total + other_total)
        source = None
        for data in sources:
            for item in (data.get("categoryRows") or {}).get("contingency") or []:
                if "预备费" in (item.get("name") or ""):
                    source = item["source"]
                    break
            if source:
                break
        put("预备费", "预备费率", rate.quantize(Decimal("0.0001")), source or {
            "file": "（按附表自洽反算）", "sheet": None, "row": None,
            "quote": "预备费 {} ÷ (工程费用 {} + 工程建设其他费用 {})".format(
                contingency_amount, engineering_total, other_total)})

    # 正文/附表里写的费率（作为对照，不覆盖上面的反算值）
    for data in sources:
        if data.get("kind") != "pdf":
            continue
        if "预备费率" in table.setdefault("预备费", {}):
            continue
        # 1) 可研正文说明句：「基本预备费：按…计 8%」/「按…*8%估算」
        rate_hit = data.get("contingencyRateText")
        if rate_hit:
            put("预备费", "预备费率", rate_hit["rate"], rate_hit["source"])
            continue
        # 2) 附表备注列（大类行与明细行都试）
        rows = list((data.get("categoryRows") or {}).get("contingency") or [])
        rows += list((data.get("detail") or {}).get("contingency") or [])
        for item in rows:
            note = item.get("note") or ""
            match = re.search(r"(\d+(?:\.\d+)?)\s*%", note)
            if match:
                value = to_decimal(match.group(1))
                if value is not None:
                    put("预备费", "预备费率", value / Decimal(100), item["source"])
                    break

    # 资金结构（xlsx 源优先：匡算表里的资金筹措表最权威）
    for data in ordered:
        funding = data.get("funding") or {}
        mapping = (("capital", "资本金(万元)"), ("loan", "借款(万元)"),
                   ("subsidy", "补助(万元)"))
        for key, label in mapping:
            item = funding.get(key)
            if item:
                put("资金结构", label, item["amount"], item["source"])
        rate = funding.get("loanRate")
        if rate:
            put("资金结构", "借款年利率", rate["rate"], rate["source"])
        # 可研正文「经济指标」里的资金来源文字：
        #   财政资金 / 政府投资 / 补助 → 补助（subsidy）
        #   企业自筹 / 自筹资金 → 资本金（equity）
        #   银行贷款 / 专项债 → 借款（loan）
        # 匡算工作簿的资金筹措表：equity/subsidy/loan + 利率
        workbook_finance = data.get("finance") or {}
        workbook_amounts = workbook_finance.get("amounts") or {}
        for key, label in (("equity", "资本金(万元)"), ("subsidy", "补助(万元)"),
                           ("loan", "借款(万元)")):
            value = workbook_amounts.get(key)
            if value is not None and label not in table.setdefault("资金结构", {}):
                put("资金结构", label, value, data.get("source") or {
                    "file": data.get("file", ""), "sheet": "资金筹措表",
                    "row": None, "quote": "{} = {}".format(key, value)})
        if workbook_finance.get("rate") is not None:
            put("资金结构", "借款年利率", workbook_finance["rate"], data.get("source") or {
                "file": data.get("file", ""), "sheet": "资金筹措表",
                "row": None, "quote": "利率 {}".format(workbook_finance["rate"])})
        economic_funding = (data.get("economic") or {}).get("funding") or {}
        amounts = (data.get("economic") or {}).get("items") or {}
        total_amount = (amounts.get("资金筹措") or {}).get("value")
        for word, item in economic_funding.items():
            if not total_amount:
                break
            if re.search(r"财政|政府投资|补助", word):
                put("资金结构", "补助(万元)", total_amount, item["source"])
            elif re.search(r"自筹|自有资金", word):
                put("资金结构", "资本金(万元)", total_amount, item["source"])
            elif re.search(r"银行|贷款|专项债", word):
                put("资金结构", "借款(万元)", total_amount, item["source"])

    # 收入 / 成本（xlsx 源优先）
    for data in ordered:
        for item in data.get("revenue") or []:
            put("收入", item["name"], item["annualAmount"], item["source"])
        cost = data.get("cost") or {}
        for key, label in (("crew", "班组人数"), ("salaryWan", "人均年薪酬(万元)")):
            item = cost.get(key)
            if item:
                put("成本", label, item["value"], item["source"])

    # 建设期利息：可研给了「建设期融资费」就直接采用
    for data in sources:
        rows = (data.get("categoryRows") or {}).get("total") or []
        for item in rows:
            if "融资费" in (item.get("name") or "") or "利息" in (item.get("name") or ""):
                put("建设期利息", "建设期融资费用(万元)", item["amount"], item["source"])

    return table, provenance, warnings


# ---------------------------------------------------------------- 写 xlsx

def main():
    parser = argparse.ArgumentParser(
        description="从可研/匡算材料自动生成并逐步补全 财务参数.xlsx")
    parser.add_argument("--workspace", default=".")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args()

    workspace = Path(args.workspace).resolve()
    # scripts/ → 01b-可研财务提取/ → Skills/ → CardoInvest/
    package_root = Path(__file__).resolve().parents[3]

    # 1) 空白模板始终放到输入文件夹（用户可填的入口）
    blank, created = ensure_blank_template(workspace, package_root)
    print("[01b] 空白模板：{}".format(blank))
    if created:
        print("[01b]   → 已从技能包释放（用户可填写，填完另存为 财务参数.xlsx）")

    # 2) 扫描输入源
    finance_files, upstream, appendix_files = discover_inputs(workspace)
    # 项目目录名（项目名的权威来源）——由批量倒推写入；单项目运行时缺省用工作区名
    project_dir_name = ""
    marker = workspace / ".项目名"
    if marker.is_file():
        # utf-8-sig：PowerShell 的 Set-Content/记事本会写入 BOM，
        # 用 utf-8 读会把 BOM 留在项目名开头（污染全链项目名）。
        project_dir_name = marker.read_text(encoding="utf-8-sig").strip()
    if not project_dir_name:
        project_dir_name = Path(workspace).name
    print("[01b] 财务源候选 {} 个；上游成果 {}".format(
        len(finance_files), {k: len(v) for k, v in upstream.items()} or "无"))

    sources = []
    # 2a) 成果文件里的匡算工作簿（附表1/3/4/5/7）——优先级最高
    for path in appendix_files:
        data = extract_from_appendix_workbook(path)
        if not data:
            continue
        data["file"] = path.name
        data["source"] = {"file": path.name, "sheet": "（工作簿）", "row": None,
                          "quote": path.name}
        data["projectDirName"] = project_dir_name
        sources.append(data)
        detail = data.get("detail") or {}
        print("[01b] 匡算工作簿命中：{}".format(path.name))
        for key in ("engineering", "otherFees", "contingency"):
            items = detail.get(key) or []
            if items:
                print("      · {}：{} 项，合计 {}".format(
                    key, len(items), fmt(sum((i["amount"] for i in items), Decimal(0)))))
        if data.get("finance"):
            print("      · 资金：{} 利率 {}".format(
                data["finance"].get("amounts"), data["finance"].get("rate")))
    # 2b) 可研 PDF
    for path in finance_files:
        if path.suffix.lower() == ".pdf":
            data = extract_from_pdf(path)
        else:
            data = extract_from_xlsx(path)
        if not data:
            continue
        data["projectDirName"] = project_dir_name
        sources.append(data)
        if data.get("kind") == "pdf":
            detail = data.get("detail") or {}
            print("[01b] pdf 命中：{} ({} 页, 附表页 {})".format(
                path.name, data["pages"], data.get("tablePages")))
        else:
            print("[01b] xlsx 命中：{} → {}".format(path.name, list(data.get("sheets", {}).values())))
        for key in ("engineering", "otherFees", "contingency"):
            items = detail.get(key) if data.get("kind") == "pdf" else \
                (data.get("detail") or {}).get(key)
            if items:
                print("      · {}：{} 项，合计 {}".format(
                    key, len(items), fmt(sum((i["amount"] for i in items), Decimal(0)))))

    if not sources:
        print("[01b] 未找到可解析的财务源 → 保留空白模板，等用户补料")
        return 2

    table, provenance, warnings = collect(sources)

    # 3) 平衡自检：用附表**大类行**的权威值互校（比用自己拼的明细可靠）
    engineering = [item["value"] for item in (table.get("工程费") or {}).values()]
    other_fees = [item["value"] for item in (table.get("其他费") or {}).values()]
    total = None
    engineering_total = None
    other_fees_total = None
    contingency_total = None
    for data in sources:
        for key, values in (data.get("categoryRows") or {}).items():
            for item in values:
                name = item.get("name") or ""
                if "总投资" in name:
                    total = item["amount"]
                elif "工程费" in name and "其他" not in name:
                    engineering_total = item["amount"]
                elif "其他费" in name or "其它费" in name:
                    other_fees_total = item["amount"]
                elif "预备费" in name:
                    contingency_total = item["amount"]
    ok, note = validate_budget_balance(
        engineering, other_fees, [], [], total,
        engineering_total=engineering_total, other_fees_total=other_fees_total,
        contingency_total=contingency_total)
    print("[01b] 平衡自检：{}（{}）".format("通过" if ok else "不通过", note))
    if not ok:
        warnings.append({"code": "PREFILL-BALANCE", "what": note,
                         "need": "核对可研附表或由用户确认", "owner": "user"})

    # 4) 写出 输入文件/财务参数.xlsx（始终在输入文件夹）
    out_path = workspace / "输入文件" / FILLED_NAME
    notes = [
        "本表由 01b-可研财务提取 自动生成并逐步补全：数值全部来自提资材料原文，",
        "未使用任何行业经验值。留空的单元格代表「材料中尚未找到该数」。",
        "每个数字的出处见「出处台账」工作表（来源文件 / 工作表或页码 / 行号 / 原文片段）。",
        "你可以直接在本表里补填或改填 —— 下次运行 01b 时，已填的格子不会被覆盖。",
        "第 8 章「财务评价」若因缺收入/成本而算不出指标，属材料事实，不得用经验值顶替。",
    ]
    if not args.dry_run:
        write_params(out_path, table, provenance, notes)
        print("[01b] 已写出 {}".format(out_path))
    else:
        print("[01b] dry-run：未写文件")

    # 5) 过程台账
    process = workspace / "过程文件" / "01b-可研财务提取"
    if not args.dry_run:
        process.mkdir(parents=True, exist_ok=True)
        payload = {
            "schema": SCHEMA,
            "generatedAt": None if args.no_timestamp else datetime.now().isoformat(timespec="seconds"),
            "sources": [{"file": item.get("file"), "kind": item.get("kind"),
                         "sheets": item.get("sheets"),
                         "tablePages": item.get("tablePages")} for item in sources],
            "upstream": {k: [p.name for p in v] for k, v in upstream.items()},
            "filled": {k: len(v) for k, v in table.items()},
            "provenance": provenance,
            "warnings": warnings,
            "balanceCheck": {"passed": ok, "note": note},
        }
        (process / "预填台账.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[01b] 台账 → {}".format(process / "预填台账.json"))
    for item in warnings:
        print("   ! {}".format(item.get("what")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
