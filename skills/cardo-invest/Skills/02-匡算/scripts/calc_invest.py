"""投资匡算引擎（T5）：工程费 + 工程建设其他费 + 预备费 = 建设投资；+ 建设期利息 = 总投资。

输入：`<项目>/05_过程/投资决策/财务输入.json`（cardo/finance-input/v1）
口径：`skill/规则/财务口径.yaml`（investment 段；工程费/其他费/预备费/建设期利息公式；自包含副本）
输出：`<项目>/05_过程/投资决策/投资估算.json`（cardo/invest-estimate/v1）+ stdout 摘要

用法：
    python calc_invest.py --help
    python calc_invest.py --project-dir <项目目录>
    python calc_invest.py --input <财务输入.json> --out <输出目录> --format json --no-timestamp

退出码：0 算出；1 数值非法（负值/非数值）；2 输入或口径缺失/坏；3 待确认（缺数 gaps / 资金筹措不平）
金额单位：万元，Decimal 两位；缺数只报 gaps，不补数。

对账保证（逐级闭合）：
    建设投资 = 工程费 + 工程建设其他费 + 预备费
    总投资   = 建设投资 + 建设期利息
    资金筹措 = 资本金 + 专项融资 + 政策性补助（与总投资差 ≤ 1 万元）
    每一级都由「已定到两位的上一级」推出，故展示值逐级可加总复核，无浮点尾差。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent                      # -> skill/
sys.path.insert(0, str(SCRIPT_DIR))                 # _lib.py 与脚本同目录（自包含）

import _lib  # noqa: E402

ENGINE_VERSION = "calc_invest/0.3.0"
RULE_FILE = "财务口径.yaml"
RULES_DIR = SKILL_ROOT / "规则"                      # skill 自带口径（自包含，不依赖插件布局）
INPUT_REL = Path("05_过程") / "投资决策" / "财务输入.json"
OUTPUT_REL = Path("05_过程") / "投资决策"
TOLERANCE_WAN = Decimal("1.0")
HUNDRED = Decimal("100")
TABLES_SCHEMA = "cardo/invest-estimate-tables/v1"
TABLES_NOTE = (
    "派生块（v2/T19）：只由 estimate/finance 输入按财务口径 investment 段派生，不改原字段、不重定义口径。"
    "金额单位万元（Decimal 两位字符串）；比率单位 %（两位）。"
    "ratio = 占总投资比例；ratioOfParent = 占所属大类比例（如其他费明细占工程建设其他费）。"
    "yearPlan 的建设投资按建设期年数均匀分摊、末年吸收尾差，各年合计可逐项复算回总投资。"
    "fundSource 的 ratio 以资金筹措合计（fundingTotal）为分母（口径 YAML investment.funding）。"
    "任一分项缺数时为 null 并在 gaps 提示，不补数。"
)
NOTE_INCOMPLETE = "缺项，非完整总投资"


class InputError(Exception):
    """输入/口径缺失或结构错误 → exit 2。"""


class ValueError_(Exception):
    """数值非法 → exit 1。"""


def load_rules(rules_dir: Path):
    path = Path(rules_dir) / RULE_FILE
    if not path.is_file():
        raise InputError("缺少口径文件：{}".format(_lib.rel_posix(path, Path.cwd())))
    rules = _lib.load_yaml(path)
    if not rules or (rules.get("schema") != "cardo/invest-finance-rules/v1"):
        raise InputError("口径 schema 不正确：{}".format(rules.get("schema") if rules else None))
    return rules, path


def resolve_project(args) -> Path:
    if args.project_dir:
        project = Path(args.project_dir)
        if not project.is_dir():
            raise InputError("项目目录不存在：{}".format(args.project_dir))
        return project
    if args.input:
        # 显式给了输入文件：项目根只用于相对路径展示，取当前目录
        return Path.cwd()
    current = Path.cwd()
    if (current / "00_项目索引.md").is_file() or (current / ".cardo").is_dir():
        return current
    raise InputError("未传 --project-dir，且当前目录不像项目（缺 00_项目索引.md/.cardo）")


def load_input(args, project: Path):
    if args.input:
        path = Path(args.input)
        if not path.is_file():
            raise InputError("输入文件不存在：{}".format(args.input))
    else:
        path = project / INPUT_REL
        if not path.is_file():
            raise InputError("缺少财务输入：{}".format(_lib.rel_posix(path, project)))
    try:
        payload = _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("财务输入读取失败：{}".format(error))
    if not isinstance(payload, dict):
        raise InputError("财务输入必须是 JSON 对象")
    schema = payload.get("schema")
    if schema not in (None, "cardo/finance-input/v1"):
        raise InputError("财务输入 schema 不匹配：{}".format(schema))
    return payload, path


def dec_positive(value, field):
    try:
        number = _lib.dec(value)
    except (InvalidOperation, TypeError, ValueError):
        raise ValueError_("{} 不是合法数值：{!r}".format(field, value))
    if number < 0:
        raise ValueError_("{} 为负值：{}".format(field, number))
    return number


def amount2(value, field):
    """单项金额：万元两位（ROUND_HALF_UP，口径 rounding.amountDecimals=2）。"""
    return _lib.q2(dec_positive(value, field))


def sum_items(items, field_prefix, gaps):
    """合计单项工程费。

    单项先定到两位再累加 —— 明细之和恒等于合计（口径 YAML engineeringCost.formula=sum(单项工程工程费)）；
    任一项缺金额 → 记 gap 且本大类合计置 None（不补 0、不出半截合计）。
    """
    total = Decimal("0")
    complete = True
    for index, item in enumerate(items or []):
        if not isinstance(item, dict):
            gaps.append(gap("INPUT-ITEM", "{}[{}] 结构不合法".format(field_prefix, index), "缺少 key/amount", "objects"))
            complete = False
            continue
        if item.get("amount") is None:
            gaps.append(gap("INPUT-AMOUNT", "{}.{} 缺金额".format(field_prefix, item.get("key", index)),
                            "输入未给 amount", "amount（万元）"))
            complete = False
            continue
        total += amount2(item["amount"], "{}.{}".format(field_prefix, item.get("key", index)))
    return (total if complete else None)


def gap(code, what, why, need):
    return {"code": code, "what": what, "why": why, "need": need, "owner": "user"}


def tolerance_wan(rules):
    """对账容差（万元）：先取口径 YAML anchors.tolerance.amountWan，缺省 1 万元（契约 G-invest-5）。"""
    value = (((rules or {}).get("anchors") or {}).get("tolerance") or {}).get("amountWan")
    return TOLERANCE_WAN if value is None else _lib.dec(value)


def ratio_percent(part, whole):
    """占比（%）：part/whole×100，两位；缺数或分母为 0 → None（不补 0）。"""
    if part is None or whole is None:
        return None
    base = _lib.dec(whole)
    if base == 0:
        return None
    return _lib.fmt(_lib.q2(_lib.dec(part) / base * HUNDRED))


def split_even(total, years):
    """按年均匀分摊：前 n-1 年取两位商，末年吸收尾差，保证合计 = total。"""
    if total is None or years <= 0:
        return [None] * max(years, 0)
    amount = _lib.q2(_lib.dec(total))
    per_year = _lib.q2(amount / Decimal(years))
    values = [per_year] * (years - 1)
    values.append(amount - per_year * Decimal(years - 1))
    return values


def _fmt_or_none(value):
    return _lib.fmt(_lib.q2(value)) if value is not None else None


def build_tables(engineering, other_fees, other_total, contingency, interest,
                 construction, total, funding, funding_total, plan, build_years,
                 tolerance=TOLERANCE_WAN, declared=None):
    """投资估算派生表（v2/T19）：构成表 + 分年计划 + 资金来源表。

    全部由本引擎已算出的 Decimal 值派生；缺数为 None 的单元格保留 null。
    """
    composition_rows = []

    def add_row(key, name, level, amount, parent_total=None):
        row = {"key": key, "name": name, "level": level,
               "amount": _fmt_or_none(amount),
               "ratio": ratio_percent(amount, total),
               "subtotal": level == 0}
        if parent_total is not None:
            row["ratioOfParent"] = ratio_percent(amount, parent_total)
        composition_rows.append(row)

    add_row("engineering", "工程费用", 1, engineering)
    add_row("otherFees", "工程建设其他费", 1, other_total)
    for item in other_fees:
        amount = _lib.dec(item["amount"]) if item.get("amount") is not None else None
        add_row("otherFees.{}".format(item.get("key")), item.get("name") or item.get("key"),
                2, amount, parent_total=other_total)
    add_row("contingency", "预备费", 1, contingency)
    add_row("constructionInvestment", "建设投资（工程费用+工程建设其他费+预备费）", 0, construction)
    add_row("constructionInterest", "建设期利息", 1, interest)
    add_row("totalInvestment", "总投资（建设投资+建设期利息）", 0, total)

    # 分年投资计划：各构成项按建设期年数均匀分摊，末年吸收尾差
    years = build_years if build_years > 0 else 0
    engineering_by_year = split_even(engineering, years)
    other_by_year = split_even(other_total, years)
    contingency_by_year = split_even(contingency, years)
    construction_by_year = split_even(construction, years)
    interest_by_year = [None] * years
    draw_by_year = [None] * years
    for index, row in enumerate(plan or []):
        if index < years:
            if row.get("interest") is not None:
                interest_by_year[index] = _lib.dec(row["interest"])
            if row.get("draw") is not None:
                draw_by_year[index] = _lib.dec(row["draw"])
    # 输入未给分年利息/提款时，按建设期均匀分摊总额、末年吸收尾差。
    # 早先这两列直接留 null，连带 total/ratio 也全成 null，分年计划四列全是「—」。
    if years and interest is not None and any(value is None for value in interest_by_year):
        interest_by_year = split_even(interest, years)
    draw_total = funding.get("loan")
    if years and draw_total is not None and any(value is None for value in draw_by_year):
        draw_by_year = split_even(draw_total, years)
    total_by_year = []
    for index in range(years):
        parts = [construction_by_year[index], interest_by_year[index]]
        total_by_year.append(sum(parts) if all(value is not None for value in parts) else None)
    year_rows = []
    for index in range(years):
        year_rows.append({
            "year": index + 1,
            "phase": "建设期",
            "engineering": _fmt_or_none(engineering_by_year[index]),
            "otherFees": _fmt_or_none(other_by_year[index]),
            "contingency": _fmt_or_none(contingency_by_year[index]),
            "constructionInvestment": _fmt_or_none(construction_by_year[index]),
            "constructionInterest": _fmt_or_none(interest_by_year[index]),
            "total": _fmt_or_none(total_by_year[index]),
            "ratio": ratio_percent(total_by_year[index], total),
            "loanDraw": _fmt_or_none(draw_by_year[index]),
        })

    def _sum(values):
        if values and all(value is not None for value in values):
            return _lib.q2(sum(values))
        return None

    year_summary = {
        "engineering": _fmt_or_none(_sum(engineering_by_year)),
        "otherFees": _fmt_or_none(_sum(other_by_year)),
        "contingency": _fmt_or_none(_sum(contingency_by_year)),
        "constructionInvestment": _fmt_or_none(_sum(construction_by_year)),
        "constructionInterest": _fmt_or_none(_sum(interest_by_year)),
        "total": _fmt_or_none(_sum(total_by_year)),
        "ratio": ratio_percent(_sum(total_by_year), total),
        "loanDraw": _fmt_or_none(_sum(draw_by_year)),
    }

    funding_rows = []
    funding_names = [
        ("equity", "项目资本金（自有资金）"),
        ("loan", "专项融资（借款）"),
        ("subsidy", "政策性补助"),
    ]
    for key, name in funding_names:
        value = funding.get(key)
        # 上方合计口径：三项至少一项有值时，未填项按 0 计入合计。
        # 明细行必须同样写 0——否则同一张表里「合计按 0 算、明细写 —」自相矛盾。
        if value is None and funding_total is not None:
            amount = Decimal("0")
        else:
            amount = _lib.dec(value) if value is not None else None
        funding_rows.append({
            "key": key,
            "name": name,
            "amount": _fmt_or_none(amount),
            "ratio": ratio_percent(amount, funding_total),
        })
    balance = None
    if funding_total is not None and total is not None:
        balance = _lib.q2(_lib.dec(funding_total) - _lib.dec(total))
    # 申报口径（信息卡）与投资需求口径的差额：不抹平、不并入现金流，只做披露。
    # 差额本身是结论的一部分——它说明申报额度大于实际需求，业主须先解决额度匹配问题。
    reconciliation = None
    declared_funding = (declared or {}).get("funding")
    declared_loan = (declared or {}).get("loan")
    if declared_funding is not None and construction is not None and total is not None:
        declared_total = amount2(declared_funding, "finance.declaredFunding")
        reconciliation = {
            "declaredFunding": _lib.fmt(declared_total),
            "declaredLoan": (_lib.fmt(amount2(declared_loan, "finance.declaredLoan"))
                             if declared_loan is not None else None),
            "constructionInvestment": _fmt_or_none(construction),
            "totalInvestment": _fmt_or_none(total),
            "undeclaredGap": _lib.fmt(_lib.q2(declared_total - construction)),
            "undeclaredGapTotal": _lib.fmt(_lib.q2(declared_total - total)),
            "basis": "申报口径合计 − 投资需求（未提款额度，不计入项目现金流、不作为偿债来源）",
        }

    return {
        "schema": TABLES_SCHEMA,
        "schemaVersion": 1,
        "note": TABLES_NOTE,
        "unit": "万元",
        "ratioUnit": "%",
        "composition": {
            "basis": "占总投比例；level 0 为小计/合计行，level 1 为大类，level 2 为明细",
            "rows": composition_rows,
            "engineering": _fmt_or_none(engineering),
            "otherFeeTotal": _fmt_or_none(other_total),
            "contingency": _fmt_or_none(contingency),
            "constructionInvestment": _fmt_or_none(construction),
            "constructionInterest": _fmt_or_none(interest),
            "totalInvestment": _fmt_or_none(total),
        },
        "yearPlan": {
            "basis": "建设期 {} 年；建设投资各构成项按年均匀分摊、末年吸收尾差".format(years),
            "constructionYears": years,
            "rows": year_rows,
            "summary": year_summary,
            "investmentSum": _fmt_or_none(sum(total_by_year)) if years and all(v is not None for v in total_by_year) else None,
        },
        "fundSource": {
            "basis": "分母 = 资金筹措合计（资本金+借款+补助）",
            "rows": funding_rows,
            "fundingTotal": _fmt_or_none(funding_total),
            "totalInvestment": _fmt_or_none(total),
            "balanceDiff": _fmt_or_none(balance),
            "balanced": (balance is not None and abs(balance) <= tolerance),
        },
        "fundingReconciliation": reconciliation,
    }


def resolve_build_years(estimate, rules, gaps):
    """建设期年数：estimate.interest.buildYears → 口径 period.constructionYears。"""
    params = (estimate or {}).get("interest") or {}
    years = params.get("buildYears")
    if years is None:
        years = ((rules or {}).get("period") or {}).get("constructionYears")
    if years is None:
        gaps.append(gap("INPUT-BUILD-YEARS", "建设期年数缺失", "输入与口径 YAML 均未给建设期年数", "buildYears"))
        return 0
    try:
        years = int(years)
    except (TypeError, ValueError):
        raise ValueError_("建设期年数不是整数：{!r}".format(years))
    if years <= 0:
        raise ValueError_("建设期年数必须为正：{}".format(years))
    return years


def compute_interest(estimate, finance, rules, build_years, gaps):
    """建设期利息：当年借款额 × 年利率 × 计息系数（建设期半年计息，系数取口径 YAML），逐年计息、不计复利。

    参数优先级：estimate.interest > finance（loan/loanRate）。
    逐年借款额按年均匀分摊、末年吸收尾差（合计 = 借款额）；逐年利息定到两位后累加，
    保证 plan 逐年之和 = 建设期利息，yearPlan 可逐项复算回总投资。
    """
    params = (estimate or {}).get("interest") or {}
    explicit = params.get("amount")
    if explicit is not None:
        # 上游已给定建设期融资费用（如可研附表1 的专项债利息+发行费）→ 直接采用，不按利率反推
        total = amount2(explicit, "estimate.interest.amount")
        years = build_years
        draws = split_even(amount2(params.get("loanAmount", (finance or {}).get("loan")), "estimate.interest.loanAmount"), years) \
            if params.get("loanAmount", (finance or {}).get("loan")) is not None else [ZERO] * years
        per = (total / Decimal(years)).quantize(Decimal("0.01")) if years > 0 else total
        plan, accumulated = [], Decimal("0")
        for index in range(1, years + 1):
            value = per if index < years else (total - accumulated)
            accumulated += value
            plan.append({"year": index, "draw": _lib.fmt(draws[index - 1]), "interest": _lib.fmt(value)})
        return _lib.q2(total), plan
    loan_amount = params.get("loanAmount", (finance or {}).get("loan"))
    loan_rate = params.get("loanRate", (finance or {}).get("loanRate"))
    if loan_amount is None and loan_rate is None:
        # 口径修正（实测踩过）：**无借款即无建设期利息**，这是事实不是缺料。
        # 纯财政补助项目（实测某企业饮品可研：资金来源为财政资金，无借款）
        # 的建设期利息就是 0。原逻辑把它报成 INPUT-INTEREST 缺项，
        # 02 步因此 exit 3 卡住，且总投资算不出来。
        # 只有「有借款额但缺利率」才算真的缺数据。
        return Decimal("0"), []
    if loan_amount is None or loan_rate is None:
        gaps.append(gap("INPUT-INTEREST", "建设期利息参数缺失",
                        "有借款额缺利率（或反之）", "loanAmount/loanRate"))
        return None, []
    amount = amount2(loan_amount, "estimate.interest.loanAmount")
    rate = dec_positive(loan_rate, "estimate.interest.loanRate")
    years = build_years
    factor = None
    if params.get("halfYear") is False:
        factor = Decimal("1")
    else:
        factor = params.get("halfYearFactor")
        if factor is None:
            factor = (((rules or {}).get("investment") or {}).get("constructionInterest") or {}).get("halfYearFactor")
        if factor is None:
            factor = Decimal("0.5")
    factor = dec_positive(factor, "estimate.interest.halfYearFactor")
    draws = split_even(amount, years)
    plan = []
    total = Decimal("0")
    for year in range(1, years + 1):
        interest = _lib.q2(draws[year - 1] * rate * factor)
        total += interest
        plan.append({"year": year, "draw": _lib.fmt(draws[year - 1]), "interest": _lib.fmt(interest)})
    return _lib.q2(total), plan


def build_estimate(payload, rules):
    gaps = []
    estimate = payload.get("estimate") or {}
    finance = payload.get("finance") or {}
    if not isinstance(estimate, dict):
        raise InputError("estimate 必须是对象")
    # 建设期年数优先取「项目」表的建设期：财务参数.xlsx 的 建设期(年) 已填时，
    # 不能退到口径 YAML 的默认 3 年，否则 02 的分年计划与 03/05 的计算期对不上。
    project = payload.get("project") or {}
    if isinstance(estimate.get("interest"), dict) and \
            estimate["interest"].get("buildYears") is None and \
            project.get("constructionYears") is not None:
        estimate["interest"]["buildYears"] = project["constructionYears"]
    if not isinstance(estimate.get("engineering", []), list):
        raise InputError("estimate.engineering 必须是数组")
    if not isinstance(estimate.get("otherFees", []), list):
        raise InputError("estimate.otherFees 必须是数组")
    if not isinstance(estimate.get("contingency", {}), dict):
        raise InputError("estimate.contingency 必须是对象")
    if not isinstance(finance, dict):
        raise InputError("finance 必须是对象")
    tolerance = tolerance_wan(rules)

    # ---- 工程费：单项两位累加，缺项 → None
    engineering = sum_items(estimate.get("engineering"), "estimate.engineering", gaps)

    # ---- 工程建设其他费：支持 amount 或 rate×basis；缺项 → None
    other_fees = []
    other_total = Decimal("0")
    other_complete = True
    for index, item in enumerate(estimate.get("otherFees") or []):
        if not isinstance(item, dict):
            gaps.append(gap("INPUT-ITEM", "estimate.otherFees[{}] 结构不合法".format(index),
                            "缺少 key/amount", "objects"))
            other_complete = False
            other_fees.append({"key": index, "name": "", "amount": None})
            continue
        key = item.get("key", index)
        amount = item.get("amount")
        if amount is None and item.get("rate") is not None and item.get("basis") is not None:
            try:
                amount = _lib.dec(item["basis"]) * _lib.dec(item["rate"])
            except (InvalidOperation, TypeError, ValueError):
                gaps.append(gap("INPUT-OTHERFEE", "其他费 {} 费率/基数非法".format(key),
                                "rate×basis 无法解析为数值", "amount 或 rate+basis"))
                other_complete = False
                continue
        if amount is None:
            gaps.append(gap("INPUT-OTHERFEE", "其他费 {} 缺金额/费率".format(key),
                            "既无 amount 也无 rate×basis", "amount 或 rate+basis"))
            other_complete = False
            other_fees.append({"key": key, "name": item.get("name", ""), "amount": None})
            continue
        value = amount2(amount, "estimate.otherFees.{}".format(key))
        other_total += value
        other_fees.append({"key": key, "name": item.get("name", ""), "amount": _lib.fmt(value)})
    if not other_complete:
        other_total = None

    # ---- 预备费：（工程费 + 工程建设其他费）× 预备费率
    contingency_rate = (estimate.get("contingency") or {}).get("rate")
    contingency = None
    if contingency_rate is None:
        gaps.append(gap("INPUT-CONTINGENCY", "预备费率缺失", "estimate.contingency.rate 未给", "预备费率"))
    elif engineering is None or other_total is None:
        gaps.append(gap("INPUT-CONTINGENCY", "预备费缺基数", "工程费或工程建设其他费缺数", "工程费/其他费明细"))
    else:
        rate = dec_positive(contingency_rate, "estimate.contingency.rate")
        contingency = _lib.q2((engineering + other_total) * rate)

    # ---- 建设期利息
    build_years = resolve_build_years(estimate, rules, gaps)
    interest, plan = compute_interest(estimate, finance, rules, build_years, gaps)

    # ---- 逐级合计：由已定到两位的上一级推出，展示值可逐级加总复核
    construction = None
    if engineering is not None and other_total is not None and contingency is not None:
        construction = _lib.q2(engineering + other_total + contingency)
    total = _lib.q2(construction + interest) if (construction is not None and interest is not None) else None

    # ---- 资金筹措：三项齐全才给合计；缺项只报缺项
    # 口径修正（实测踩过）：**三项并非必须齐全**。
    # 纯财政补助项目（实测某企业饮品可研：补助 5300，无资本金、无借款）是合法的，
    # 「没有某一项」不等于「缺该项数据」。原逻辑要求三项都非 None，
    # 会把这类项目误判为缺项（INPUT-FUNDING），02 步因此 exit 3 卡住。
    # 改为：**至少一项有值**即可算合计，未填项按 0 计入；
    # 只有三项**全空**（真的一点资金来源都没给）才报缺项。
    funding_values = {}
    for key in ("equity", "loan", "subsidy"):
        value = finance.get(key)
        funding_values[key] = None if value is None else amount2(value, "finance.{}".format(key))
    provided = [value for value in funding_values.values() if value is not None]
    funding_total = None
    if provided:
        funding_total = _lib.q2(sum(value or Decimal("0") for value in funding_values.values()))
        if total is not None and abs(funding_total - total) > tolerance:
            gaps.append(gap("FUNDING-MISMATCH", "资金筹措与总投资不平",
                            "资金筹措合计 {} 与总投资 {} 差 {} 万元（容差 {}）".format(
                                funding_total, total, abs(funding_total - total), tolerance),
                            "调整资本金/借款/补助"))
    else:
        gaps.append(gap("INPUT-FUNDING", "资金筹措缺项",
                        "capital/loan/subsidy 三项均未提供", "至少一项资金来源金额"))

    status = "ok"
    if gaps:
        status = "needs_confirm"
    tables = build_tables(engineering, other_fees, other_total, contingency, interest,
                          construction, total, funding_values, funding_total, plan, build_years,
                          tolerance=tolerance,
                          declared={"funding": finance.get("declaredFunding"),
                                    "loan": finance.get("declaredLoan")})
    engineering_rows = []
    for item in estimate.get("engineering") or []:
        row = {"key": item.get("key"), "name": item.get("name"),
               "amount": _fmt_or_none(item.get("amount"))}
        if item.get("unit"):
            row["unit"] = item["unit"]
        if item.get("quantity") is not None:
            row["quantity"] = item["quantity"]
        engineering_rows.append(row)
    return {
        "engineering": _fmt_or_none(engineering),
        "engineeringRows": engineering_rows,
        "otherFees": other_fees,
        "otherFeeTotal": _fmt_or_none(other_total),
        "contingency": _fmt_or_none(contingency),
        "constructionInterest": _fmt_or_none(interest),
        "constructionInvestment": _fmt_or_none(construction),
        "totalInvestment": _fmt_or_none(total),
        "funding": {key: (_lib.fmt(value) if value is not None else None)
                    for key, value in funding_values.items()},
        "fundingTotal": _fmt_or_none(funding_total),
        "plan": plan,
        "fundingReconciliation": tables.get("fundingReconciliation"),
        "tables": tables,
        "status": status,
        "note": None if total is not None else NOTE_INCOMPLETE,
    }, gaps


def fingerprint(payload) -> str:
    canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def rule_ref(rules, *path):
    """从口径 YAML 取一条引用（用于 trace 的来源锚点）。"""
    node = rules
    for key in path:
        node = (node or {}).get(key) if isinstance(node, dict) else None
    return node


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 投资匡算（工程费/其他费/预备费/建设期利息 → 总投资）")
    parser.add_argument("--project-dir", default=None, help="项目目录（任意路径）")
    parser.add_argument("--input", default=None, help="财务输入 JSON（缺省 <项目>/05_过程/投资决策/财务输入.json）")
    parser.add_argument("--rules", default=None, help="规则目录（缺省=skill 自带 规则/）")
    parser.add_argument("--out", default=None, help="输出目录（缺省 <项目>/05_过程/投资决策）")
    parser.add_argument("--format", choices=("text", "json"), default="text")
    parser.add_argument("--no-timestamp", action="store_true", help="幂等：不写时间戳")
    args = parser.parse_args(argv)

    rules_dir = Path(args.rules) if args.rules else RULES_DIR
    try:
        rules, rules_path = load_rules(rules_dir)
        project = resolve_project(args)
        payload, input_path = load_input(args, project)
        estimate, gaps = build_estimate(payload, rules)
    except InputError as error:
        print("[calc_invest] {}".format(error))
        return 2
    except ValueError_ as error:
        print("[calc_invest] 数值非法：{}".format(error))
        return 1

    out_dir = Path(args.out) if args.out else project / OUTPUT_REL
    out_dir.mkdir(parents=True, exist_ok=True)
    result = _lib.envelope("cardo/invest-estimate/v1", 1, producer="cardo-invest-estimate/" + ENGINE_VERSION,
                           sources=[{"file": _lib.rel_posix(input_path, project),
                                     "locator": "estimate/finance", "sha256": _lib.sha256_file(input_path)}],
                           gaps=gaps, no_timestamp=args.no_timestamp)
    result.update(estimate)
    factor = rule_ref(rules, "investment", "constructionInterest", "halfYearFactor")
    explicit_interest = (((payload.get("estimate") or {}).get("interest") or {}).get("amount") is not None)
    result["trace"] = {
        "ruleVersion": rules.get("version"),
        "engineVersion": ENGINE_VERSION,
        "inputFingerprint": fingerprint(payload),
        "rulesFile": RULE_FILE,
        "rulesSha256": _lib.sha256_file(rules_path),
        "toleranceWan": _lib.fmt(tolerance_wan(rules)),
        "steps": [
            {"id": "STEP-1", "name": "工程费合计", "formula": "sum(engineering[].amount)",
             "value": estimate["engineering"],
             "ruleSource": rule_ref(rules, "investment", "engineeringCost", "formula")},
            {"id": "STEP-2", "name": "其他费合计", "formula": "sum(otherFees[])",
             "value": estimate["otherFeeTotal"],
             "ruleSource": rule_ref(rules, "investment", "otherFees", "formula")},
            {"id": "STEP-3", "name": "预备费", "formula": "（工程费+其他费）× 预备费率",
             "value": estimate["contingency"],
             "ruleSource": rule_ref(rules, "investment", "contingency", "formula")},
            {"id": "STEP-4", "name": "建设期利息",
             "formula": ("上游给定的建设期融资费用（不按利率反推）"
                         if explicit_interest else "逐年借款额×年利率×计息系数"),
             "value": estimate["constructionInterest"],
             "ruleSource": ("上游可研附表（estimate.interest.amount）" if explicit_interest
                            else rule_ref(rules, "investment", "constructionInterest", "sourceCell")),
             "halfYearFactor": _lib.fmt(_lib.dec(factor)) if factor is not None else None},
            {"id": "STEP-5", "name": "建设投资", "formula": "工程费+其他费+预备费",
             "value": estimate["constructionInvestment"],
             "ruleSource": rule_ref(rules, "investment", "constructionInvestment", "formula")},
            {"id": "STEP-6", "name": "总投资", "formula": "建设投资+建设期利息",
             "value": estimate["totalInvestment"],
             "ruleSource": rule_ref(rules, "investment", "totalInvestment", "formula")},
        ],
    }
    out_path = out_dir / "投资估算.json"
    _lib.write_json(out_path, result, no_timestamp=args.no_timestamp)

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("[calc_invest] 工程费={} 其他费={} 预备费={} 建设期利息={}".format(
            estimate["engineering"], estimate["otherFeeTotal"], estimate["contingency"], estimate["constructionInterest"]))
        print("[calc_invest] 建设投资={} 总投资={}（{}）".format(
            estimate["constructionInvestment"], estimate["totalInvestment"], estimate["status"]))
        if gaps:
            print("[calc_invest] gaps {} 项（不补数）：{}".format(len(gaps), ", ".join(item["code"] for item in gaps)))
        print("[calc_invest] 写出 {}".format(_lib.rel_posix(out_path, project)))

    if gaps:
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
