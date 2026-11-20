"""收入 / 成本 / 折旧摊销 / 税 / 借款还本付息薄引擎（T6）。

输入：`<项目>/05_过程/投资决策/财务输入.json`（cardo/finance-input/v1）
      + `投资估算.json`（cardo/invest-estimate/v1，T5 产物；缺则 gap）
口径：`../cardo-invest-spec/规则/财务口径.yaml`（revenue/cost/tax/profit 段）
输出：`<项目>/05_过程/投资决策/财务测算.json`（cardo/finance/v1 的底表部分；
      现金流与指标由 T7 `calc_metrics.py` 补全）

用法：
    python calc_finance.py --help
    python calc_finance.py --project-dir <项目目录>
    python calc_finance.py --input <财务输入.json> --estimate <投资估算.json> --out <目录> --no-timestamp

退出码：0 算出；1 数值非法（负值/税率非法/除零）；2 输入或口径缺失/坏；3 待确认（缺数 gaps）
金额单位：万元，Decimal 两位。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from decimal import Decimal
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))  # 自包含化：_lib.py 与脚本同目录

import _lib  # noqa: E402

SKILL_DIR = SCRIPT_DIR.parent
RULES_DIR_DEFAULT = SKILL_DIR / "规则"  # 自包含化：口径在本 skill 的 规则/ 下

ENGINE_VERSION = "calc_finance/0.2.0"
RULE_FILE = "财务口径.yaml"
INPUT_REL = Path("05_过程") / "投资决策" / "财务输入.json"
ESTIMATE_REL = Path("05_过程") / "投资决策" / "投资估算.json"
OUTPUT_REL = Path("05_过程") / "投资决策"
HUNDRED = Decimal("100")
TABLES_SCHEMA = "cardo/invest-finance-tables/v1"
TABLES_NOTE = (
    "派生块（v2/T19）：只由本引擎产出的 revenue/cost 底表按财务口径 cost 段汇总，不改原字段、不重定义口径。"
    "统计区间 = 运营期（第 constructionYears+1 年起）；金额单位万元（Decimal 两位字符串），比率单位 %（两位）。"
    "costStructure.ratio = 该明细占经营成本合计比例；revenueStructure.ratio = 该收入项占营业收入合计比例；"
    "avgAnnual = 合计 / 运营年数。底表缺失（cost 为 null）时 items 为空并标注 status=gap，不补数。"
)
DEFAULT = {
    "constructionYears": 3,
    "operationYears": 30,
    "firstYearHalfFactor": Decimal("0.5"),
    "serviceVat": Decimal("0.06"),
    "rentalVat": Decimal("0.11"),
    "utilitiesRate": Decimal("0.03"),
    "managementRate": Decimal("0.02"),
    "repairRate": Decimal("0.015"),
    "repairRateAfter10y": Decimal("0.02"),
    "otherRate": Decimal("0.015"),
    "overhaulRate": Decimal("0.04"),
    "overhaulCycle": 10,
    "buildingYears": 30,
    "buildingResidual": Decimal("0.05"),
    "incomeTaxRate": Decimal("0.25"),
    "urbanRate": Decimal("0.07"),
    "educationRate": Decimal("0.05"),
    "stampRentalRate": Decimal("0.001"),
    "landUseTaxPerSqm": Decimal("8"),
    "propertyTaxRate": Decimal("0.12"),
    "loanTerm": 10,
    "graceYears": 3,
}
ZERO = Decimal("0")


class InputError(Exception):
    """输入/口径缺失或结构错误 → exit 2。"""


class ValueError_(Exception):
    """数值非法 → exit 1。"""


def gap(code, what, why, need):
    return {"code": code, "what": what, "why": why, "need": need, "owner": "user"}


def dec(value, field, allow_negative=False):
    try:
        number = _lib.dec(value)
    except (TypeError, ValueError):
        raise ValueError_("{} 不是合法数值：{!r}".format(field, value))
    if not allow_negative and number < 0:
        raise ValueError_("{} 为负值：{}".format(field, number))
    return number


def rate(value, field, upper=Decimal("1")):
    number = dec(value, field)
    if number > upper:
        raise ValueError_("{} 超范围（>1）：{}".format(field, number))
    return number


def load_rules(rules_dir):
    path = Path(rules_dir) / RULE_FILE
    if not path.is_file():
        raise InputError("缺少口径文件：{}".format(_lib.rel_posix(path, Path.cwd())))
    rules = _lib.load_yaml(path)
    if not rules or rules.get("schema") != "cardo/invest-finance-rules/v1":
        raise InputError("口径 schema 不正确")
    return rules


def resolve_project(args):
    if args.project_dir:
        project = Path(args.project_dir)
        if not project.is_dir():
            raise InputError("项目目录不存在：{}".format(args.project_dir))
        return project
    if args.input:
        return Path.cwd()
    current = Path.cwd()
    if (current / "00_项目索引.md").is_file() or (current / ".cardo").is_dir():
        return current
    raise InputError("未传 --project-dir，且当前目录不像项目（缺 00_项目索引.md/.cardo）")


def load_json(path, label):
    if not path.is_file():
        raise InputError("缺少{}：{}".format(label, _lib.rel_posix(path, Path.cwd())))
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("{}读取失败：{}".format(label, error))


def vat_rate_for(key, rules, item):
    if item.get("vatRate") is not None:
        return rate(item["vatRate"], "revenue.{}.vatRate".format(key))
    revenue_rules = rules.get("revenue") or {}
    override = (revenue_rules.get("items") or [])
    for entry in override:
        if entry.get("key") == key and entry.get("vatRate") is not None:
            return Decimal(str(entry["vatRate"]))
    if key in ("shop-rent", "office-rent"):
        return DEFAULT["rentalVat"]
    return DEFAULT["serviceVat"]


def revenue_base(item, key, gaps):
    """单年收入基数（万元）。优先 annualAmount；否则按 key 公式（price×quantity×周期/10000）。"""
    if item.get("annualAmount") is not None:
        return dec(item["annualAmount"], "revenue.{}.annualAmount")
    price = item.get("price")
    quantity = item.get("quantity")
    if price is None or quantity is None:
        gaps.append(gap("INPUT-REVENUE", "收入项 {} 缺 price/quantity 或 annualAmount".format(key),
                        "无法计算年收入", "annualAmount 或 price+quantity"))
        return None
    price = dec(price, "revenue.{}.price".format(key))
    quantity = dec(quantity, "revenue.{}.quantity".format(key))
    if key == "parking":
        days = dec(item.get("days") or 365, "revenue.parking.days")
        return price * quantity * days / Decimal("10000")
    months = dec(item.get("months") or 12, "revenue.{}.months".format(key))
    return price * quantity * months / Decimal("10000")


def build_revenue(payload, rules, years):
    gaps = []
    operations = [index for index, info in enumerate(years) if info["phase"] == "operation"]
    items_out = []
    revenue_by_year = [ZERO] * len(years)
    output_vat_by_year = [ZERO] * len(years)
    for index, item in enumerate(payload.get("revenue") or []):
        key = item.get("key") or "item{}".format(index)
        base = revenue_base(item, key, gaps)
        vat = vat_rate_for(key, rules, item)
        growth = dec(item.get("growth") or 0, "revenue.{}.growth".format(key), allow_negative=True)
        yearly = [ZERO] * len(years)
        if base is not None:
            for op_index, year_index in enumerate(operations):
                factor = Decimal(str((1 + growth) ** op_index))
                if op_index == 0:
                    factor *= DEFAULT["firstYearHalfFactor"]
                yearly[year_index] = _lib.q2(base * factor)
        entry = {"key": key, "name": item.get("name") or key, "vatRate": _lib.fmt(vat),
                 "yearly": [_lib.fmt(value) for value in yearly],
                 "total": _lib.fmt(_lib.q2(sum(yearly)))}
        if item.get("basisText"):
            entry["basis"] = item["basisText"]
        items_out.append(entry)
        for year_index, value in enumerate(yearly):
            revenue_by_year[year_index] += value
            output_vat_by_year[year_index] += value / (1 + vat) * vat
    return items_out, [_lib.q2(v) for v in revenue_by_year], [_lib.q2(v) for v in output_vat_by_year], gaps


def build_loan(finance, rules, years, gaps):
    loan = finance.get("loan")
    if loan is None:
        gaps.append(gap("INPUT-LOAN", "借款额缺失", "finance.loan 未给", "借款额"))
        return None
    principal = dec(loan, "finance.loan")
    loan_rate = rate(finance.get("loanRate") or rules_loan_rate(rules), "finance.loanRate")
    term = int(finance.get("loanTerm") or DEFAULT["loanTerm"])
    grace = int(finance.get("graceYears") or DEFAULT["graceYears"])
    if term <= 0:
        raise ValueError_("还款年限必须为正")
    build_years = sum(1 for info in years if info["phase"] == "build")
    draws = [ZERO] * len(years)
    if build_years:
        per_year = principal / Decimal(build_years)
        for index in range(build_years):
            draws[index] = _lib.q2(per_year)
        draws[build_years - 1] += principal - sum(draws)
    schedule = []
    opening = ZERO
    repay_start = build_years + grace
    repay_years = len(years) - repay_start
    repay_years = min(term, max(repay_years, 0))
    annual_principal = principal / Decimal(repay_years) if repay_years > 0 else ZERO
    # 建设期利息口径必须与 02 匡算 / 财务口径.yaml（当年借款额×年利率×0.5）一致，
    # 否则 03 还本付息表的建设期利息（按年初余额+当年半额计）会大于 02 计入总投资的
    # 建设期利息，同一个项目出现两个「建设期利息」，投资估算与还本付息表对不上。
    half_factor = Decimal(str(
        ((rules.get("investment") or {}).get("constructionInterest") or {}).get("halfYearFactor") or 0.5))
    for index, info in enumerate(years):
        interest = (draws[index] * loan_rate * half_factor if index < build_years
                    else (opening + draws[index] / 2) * loan_rate)
        repay = ZERO
        if repay_years > 0 and repay_start <= index < repay_start + repay_years:
            repay = annual_principal
        closing = opening + draws[index] - repay
        if index == len(years) - 1:
            repay += closing
            closing = ZERO
        schedule.append({"year": info["year"], "opening": _lib.fmt(_lib.q2(opening)),
                         "draw": _lib.fmt(draws[index]), "interest": _lib.fmt(_lib.q2(interest)),
                         "principal": _lib.fmt(_lib.q2(repay)), "closing": _lib.fmt(_lib.q2(closing))})
        opening = closing
    return {"loanRate": _lib.fmt(loan_rate), "loanTerm": term, "graceYears": grace, "schedule": schedule}


def rules_loan_rate(rules):
    return "0.027"


def build_cost_and_tax(payload, estimate, rules, years, revenue_total, output_vat, loan, gaps):
    finance = payload.get("finance") or {}
    cost_input = payload.get("cost") or {}
    tax_rules = rules.get("tax") or {}
    construction = estimate.get("constructionInvestment")
    total_investment = estimate.get("totalInvestment")
    if construction is None or total_investment is None:
        gaps.append(gap("INPUT-ESTIMATE", "投资估算不完整", "无法计算折旧/大修", "补齐投资估算"))
        return None
    construction_d = dec(construction, "投资估算.constructionInvestment")
    total_d = dec(total_investment, "投资估算.totalInvestment")
    build_years = sum(1 for info in years if info["phase"] == "build")
    operations = [index for index, info in enumerate(years) if info["phase"] == "operation"]
    interest_by_year = [ZERO] * len(years)
    if loan:
        for index, row in enumerate(loan["schedule"]):
            interest_by_year[index] = dec(row["interest"], "loan.interest", allow_negative=True)

    utilities = [ZERO] * len(years)
    management = [ZERO] * len(years)
    repair = [ZERO] * len(years)
    other = [ZERO] * len(years)
    overhaul = [ZERO] * len(years)
    labor = [ZERO] * len(years)
    operating = [ZERO] * len(years)
    labor_input = cost_input.get("labor") or {}
    # `dict.get(k, default)` 在**键存在但值为 None** 时返回 None（默认值不生效），
    # dec(None) 会抛「不是合法数值」。故一律用 `or` 链兜底：
    # 缺成本数据 → 用口径默认值 → 链路不中断（缺项由 gaps 登记，不补数）。
    crew = dec(labor_input.get("crew") or 6, "cost.labor.crew")
    salary = dec(labor_input.get("salaryWan") or 7, "cost.labor.salaryWan")
    welfare = dec(labor_input.get("welfareFactor") or 1.2, "cost.labor.welfareFactor")
    labor_growth = dec(labor_input.get("growthRate") or 0.03, "cost.labor.growthRate", allow_negative=True)
    for op_index, year_index in enumerate(operations):
        revenue = dec(revenue_total[year_index], "revenue.total", allow_negative=True)
        factor = Decimal(str((1 + labor_growth) ** op_index)) if op_index else DEFAULT["firstYearHalfFactor"]
        labor[year_index] = _lib.q2(crew * salary * welfare * factor)
        utilities[year_index] = _lib.q2(revenue * DEFAULT["utilitiesRate"])
        management[year_index] = _lib.q2(revenue * DEFAULT["managementRate"])
        repair_rate = DEFAULT["repairRate"] if op_index < 10 else DEFAULT["repairRateAfter10y"]
        repair[year_index] = _lib.q2(revenue * repair_rate)
        other[year_index] = _lib.q2(revenue * DEFAULT["otherRate"])
        if op_index > 0 and op_index % DEFAULT["overhaulCycle"] == 0:
            overhaul[year_index] = _lib.q2(total_d * DEFAULT["overhaulRate"])
        operating[year_index] = utilities[year_index] + management[year_index] + repair[year_index] \
            + other[year_index] + overhaul[year_index] + labor[year_index]

    # 折旧摊销：房屋建筑物 30 年、残值 5%、首年半年。
    # 折旧年限不得超过运营期，否则期末账面余值大于可回收残值（R3 P2-18）。
    dep_years = min(int(DEFAULT["buildingYears"]), max(len(operations), 1))
    dep_units = Decimal(dep_years) - Decimal("1") + Decimal(str(DEFAULT["firstYearHalfFactor"]))
    annual_dep = (total_d * (1 - DEFAULT["buildingResidual"])) / dep_units
    depreciation = [ZERO] * len(years)
    for op_index, year_index in enumerate(operations):
        depreciation[year_index] = _lib.q2(annual_dep * (DEFAULT["firstYearHalfFactor"] if op_index == 0 else Decimal("1")))

    # 进项税：建设期按建安 9% + 其他费 6%（按年均匀分摊）；运营期按成本项税率
    input_vat = [ZERO] * len(years)
    if build_years:
        construction_per_year = construction_d / Decimal(build_years)
        other_fee_total = dec(estimate.get("otherFeeTotal") or 0, "estimate.otherFeeTotal")
        other_per_year = other_fee_total / Decimal(build_years)
        for index in range(build_years):
            input_vat[index] = _lib.q2(construction_per_year / Decimal("1.09") * Decimal("0.09")
                                       + other_per_year / Decimal("1.06") * Decimal("0.06"))
    for year_index in operations:
        input_vat[year_index] = _lib.q2(
            utilities[year_index] / Decimal("1.09") * Decimal("0.09")
            + labor[year_index] / Decimal("1.13") * Decimal("0.13")
            + (repair[year_index] + overhaul[year_index]) / Decimal("1.13") * Decimal("0.13"))

    vat_payable = [_lib.q2(max(ZERO, output_vat[index] - input_vat[index])) for index in range(len(years))]
    project = payload.get("project") or {}
    area = dec(project.get("landTaxArea") or project.get("area") or 0, "project.landTaxArea")
    rental_keys = {"mobile-food", "outdoor-dining", "shop-rent", "office-rent"}
    rental_revenue = [ZERO] * len(years)
    for item in payload.get("revenue") or []:
        if item.get("key") in rental_keys:
            base = revenue_base(item, item.get("key"), [])
            if base is None:
                continue
            growth = dec(item.get("growth") or 0, "revenue.growth", allow_negative=True)
            for op_index, year_index in enumerate(operations):
                factor = Decimal(str((1 + growth) ** op_index))
                if op_index == 0:
                    factor *= DEFAULT["firstYearHalfFactor"]
                rental_revenue[year_index] += base * factor
    surcharges = {"urban": [ZERO] * len(years), "education": [ZERO] * len(years),
                  "stamp": [ZERO] * len(years), "landUse": [ZERO] * len(years),
                  "property": [ZERO] * len(years)}
    property_enabled = bool(((tax_rules.get("propertyTax") or {}).get("enabled") or False))
    for index in range(len(years)):
        surcharges["urban"][index] = _lib.q2(vat_payable[index] * DEFAULT["urbanRate"])
        surcharges["education"][index] = _lib.q2(vat_payable[index] * DEFAULT["educationRate"])
        surcharges["stamp"][index] = _lib.q2(revenue_total[index] * DEFAULT["stampRentalRate"])
        surcharges["landUse"][index] = _lib.q2(area * DEFAULT["landUseTaxPerSqm"] / Decimal("10000"))
        if property_enabled:
            surcharges["property"][index] = _lib.q2(rental_revenue[index] * DEFAULT["propertyTaxRate"])
    surcharge_total = [sum(surcharges[key][index] for key in surcharges) for index in range(len(years))]

    # 利润与所得税
    taxable = [ZERO] * len(years)
    income_tax = [ZERO] * len(years)
    adjusted_tax = [ZERO] * len(years)
    ebit = [ZERO] * len(years)
    ebitda = [ZERO] * len(years)
    for index in range(len(years)):
        net_revenue = revenue_total[index] - vat_payable[index] - surcharge_total[index]
        total_cost = operating[index] + depreciation[index] + interest_by_year[index]
        taxable[index] = _lib.q2(net_revenue - total_cost)
        if taxable[index] > 0:
            income_tax[index] = _lib.q2(taxable[index] * DEFAULT["incomeTaxRate"])
        ebit[index] = _lib.q2(net_revenue - operating[index] - depreciation[index])
        ebitda[index] = _lib.q2(ebit[index] + depreciation[index])
        adjusted_tax[index] = _lib.q2(max(ZERO, ebit[index]) * DEFAULT["incomeTaxRate"])

    return {
        "items": {
            "utilities": [_lib.fmt(value) for value in utilities],
            "labor": [_lib.fmt(value) for value in labor],
            "management": [_lib.fmt(value) for value in management],
            "repair": [_lib.fmt(value) for value in repair],
            "other": [_lib.fmt(value) for value in other],
            "overhaul": [_lib.fmt(value) for value in overhaul],
        },
        "operatingCost": [_lib.fmt(value) for value in operating],
        "depreciation": [_lib.fmt(value) for value in depreciation],
        "interest": [_lib.fmt(value) for value in interest_by_year],
        "totalCost": [_lib.fmt(_lib.q2(operating[i] + depreciation[i] + interest_by_year[i])) for i in range(len(years))],
        "inputVat": [_lib.fmt(value) for value in input_vat],
        "vatPayable": [_lib.fmt(value) for value in vat_payable],
        "surcharges": {key: [_lib.fmt(value) for value in values] for key, values in surcharges.items()},
        "surchargeTotal": [_lib.fmt(value) for value in surcharge_total],
        "taxableIncome": [_lib.fmt(value) for value in taxable],
        "incomeTax": [_lib.fmt(value) for value in income_tax],
        "adjustedIncomeTax": [_lib.fmt(value) for value in adjusted_tax],
        "ebit": [_lib.fmt(value) for value in ebit],
        "ebitda": [_lib.fmt(value) for value in ebitda],
        "netRevenue": [_lib.fmt(_lib.q2(revenue_total[i] - vat_payable[i] - surcharge_total[i])) for i in range(len(years))],
    }


def cost_item_names(rules):
    """从口径 YAML 取经营成本明细项名称（口径唯一来源；缺失时用中文兜底名）。"""
    names = {}
    items = (((rules.get("cost") or {}).get("operatingCost") or {}).get("items")) or []
    for entry in items:
        if isinstance(entry, dict) and entry.get("key"):
            names[entry["key"]] = entry.get("name") or entry["key"]
    fallback = {"utilities": "外购燃料及动力费", "labor": "员工工资及福利费",
                "management": "管理费用", "repair": "日常修理费",
                "other": "其他费用", "overhaul": "大修费"}
    for key, name in fallback.items():
        names.setdefault(key, name)
    return names


def ratio_percent(part, whole):
    if part is None or whole is None:
        return None
    base = _lib.dec(whole)
    if base == 0:
        return None
    return _lib.fmt(_lib.q2(_lib.dec(part) / base * HUNDRED))


def build_tables(payload, rules, years, revenue_items, revenue_total, output_vat, cost):
    """财务测算派生表（v2/T19）：成本结构表 + 收入结构表 + 总成本费用构成。"""
    operation_index = [index for index, info in enumerate(years) if info["phase"] == "operation"]
    op_years = len(operation_index)

    def sum_operation(series):
        values = series or []
        return _lib.q2(sum((_lib.dec(values[index]) for index in operation_index if index < len(values)), ZERO))

    def avg_annual(total):
        return _lib.q2(total / Decimal(op_years)) if op_years else None

    period = {"constructionYears": len(years) - op_years, "operationYears": op_years,
              "basis": "运营期合计（第 {}–{} 年）".format(len(years) - op_years + 1, len(years)) if years else "—"}

    names = cost_item_names(rules)
    cost_rows = []
    cost_total = None
    if cost and cost.get("items"):
        cost_items = cost["items"]
        order = [key for key in names if key in cost_items] + [key for key in cost_items if key not in names]
        cost_total = sum_operation(cost.get("operatingCost"))
        for key in order:
            total = sum_operation(cost_items.get(key))
            cost_rows.append({
                "key": key,
                "name": names.get(key, key),
                "total": _lib.fmt(total),
                "avgAnnual": _lib.fmt(avg_annual(total)),
                "ratio": ratio_percent(total, cost_total),
            })
    ratio_sum = None
    if cost_rows and all(row["ratio"] is not None for row in cost_rows):
        ratio_sum = _lib.fmt(_lib.q2(sum(_lib.dec(row["ratio"]) for row in cost_rows)))

    cost_structure = {
        "status": "ok" if cost_rows else "gap",
        "unit": "万元",
        "ratioUnit": "%",
        "total": _lib.fmt(cost_total) if cost_total is not None else None,
        # 年均值取分项之和，不取总额÷年数：表格里分项与合计必须对得上，
        # 否则「222.41（总额/29）」与「222.42（分项相加）」会差一分（R3 P2-5）。
        "avgAnnual": _lib.fmt(_lib.q2(sum(_lib.dec(row["avgAnnual"]) for row in cost_rows)))
        if cost_rows else (_lib.fmt(avg_annual(cost_total)) if cost_total is not None else None),
        "ratioSum": ratio_sum,
        "rows": cost_rows,
    }

    depreciation_total = sum_operation((cost or {}).get("depreciation"))
    interest_total = sum_operation((cost or {}).get("interest"))
    full_rows = []
    full_total = None
    if cost_total is not None:
        full_total = _lib.q2(cost_total + depreciation_total + interest_total)
        for key, name, value in (("operatingCost", "经营成本", cost_total),
                                 ("depreciation", "折旧摊销", depreciation_total),
                                 ("interest", "利息支出", interest_total)):
            full_rows.append({"key": key, "name": name, "total": _lib.fmt(value),
                              "avgAnnual": _lib.fmt(avg_annual(value)),
                              "ratio": ratio_percent(value, full_total)})
    full_cost = {
        "status": "ok" if full_rows else "gap",
        "unit": "万元",
        "ratioUnit": "%",
        "total": _lib.fmt(full_total) if full_total is not None else None,
        "avgAnnual": _lib.fmt(avg_annual(full_total)) if full_total is not None else None,
        "rows": full_rows,
    }

    revenue_total_op = sum_operation([_lib.dec(value) for value in revenue_total])
    revenue_rows = []
    for item in revenue_items:
        total = _lib.q2(sum((_lib.dec(item["yearly"][index]) for index in operation_index
                             if index < len(item["yearly"])), ZERO))
        revenue_rows.append({
            "key": item.get("key"),
            "name": item.get("name"),
            "total": _lib.fmt(total),
            "avgAnnual": _lib.fmt(avg_annual(total)),
            "ratio": ratio_percent(total, revenue_total_op),
            "vatRate": item.get("vatRate"),
        })
    revenue_structure = {
        "status": "ok" if revenue_rows else "gap",
        "unit": "万元",
        "ratioUnit": "%",
        "total": _lib.fmt(revenue_total_op),
        "avgAnnual": _lib.fmt(avg_annual(revenue_total_op)),
        "outputVatTotal": _lib.fmt(sum_operation([_lib.dec(value) for value in output_vat])),
        "rows": revenue_rows,
    }

    return {
        "schema": TABLES_SCHEMA,
        "schemaVersion": 1,
        "note": TABLES_NOTE,
        "unit": "万元",
        "ratioUnit": "%",
        "period": period,
        "costStructure": cost_structure,
        "fullCostStructure": full_cost,
        "revenueStructure": revenue_structure,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 收入/成本/折旧摊销/税引擎（T6）")
    parser.add_argument("--project-dir", default=None, help="项目目录（任意路径）")
    parser.add_argument("--input", default=None, help="财务输入 JSON")
    parser.add_argument("--estimate", default=None, help="投资估算 JSON（缺省 <项目>/05_过程/投资决策/投资估算.json）")
    parser.add_argument("--rules", default=None, help="规则目录（缺省=skill 自带 规则/）")
    parser.add_argument("--out", default=None, help="输出目录（缺省 <项目>/05_过程/投资决策）")
    parser.add_argument("--format", choices=("text", "json"), default="text")
    parser.add_argument("--no-timestamp", action="store_true", help="幂等：不写时间戳")
    args = parser.parse_args(argv)

    rules_dir = Path(args.rules) if args.rules else RULES_DIR_DEFAULT
    try:
        rules = load_rules(rules_dir)
        project = resolve_project(args)
        input_path = Path(args.input) if args.input else project / INPUT_REL
        estimate_path = Path(args.estimate) if args.estimate else project / ESTIMATE_REL
        payload = load_json(input_path, "财务输入")
        estimate = load_json(estimate_path, "投资估算")
        # 防御：建设期可能缺失（可研未写工期）→ 用口径默认值，不让 int(None) 崩
        build_years = int(((payload.get("project") or {}).get("constructionYears")
                           or DEFAULT["constructionYears"]))
        operation_years = int(((payload.get("project") or {}).get("operationYears")
                               or DEFAULT["operationYears"]))
        if build_years < 0 or operation_years <= 0:
            raise ValueError_("计算期非法：建设期 {} / 运营期 {}".format(build_years, operation_years))
        years = [{"year": index + 1, "phase": "build" if index < build_years else "operation"}
                 for index in range(build_years + operation_years)]
        revenue_items, revenue_total, output_vat, gaps_rev = build_revenue(payload, rules, years)
        loan = build_loan(payload.get("finance") or {}, rules, years, gaps_rev)
        cost = build_cost_and_tax(payload, estimate, rules, years, revenue_total, output_vat, loan, gaps_rev)
    except InputError as error:
        print("[calc_finance] {}".format(error))
        return 2
    except ValueError_ as error:
        print("[calc_finance] 数值非法：{}".format(error))
        return 1

    gaps = gaps_rev
    out_dir = Path(args.out) if args.out else project / OUTPUT_REL
    out_dir.mkdir(parents=True, exist_ok=True)
    result = _lib.envelope("cardo/finance/v1", 1, producer="cardo-invest-finance/" + ENGINE_VERSION,
                           sources=[{"file": _lib.rel_posix(input_path, project), "sha256": _lib.sha256_file(input_path)},
                                    {"file": _lib.rel_posix(estimate_path, project), "sha256": _lib.sha256_file(estimate_path)}],
                           gaps=gaps, no_timestamp=args.no_timestamp)
    result["period"] = {"constructionYears": build_years, "operationYears": operation_years,
                        "totalYears": len(years), "years": years,
                        "discountRate": _lib.fmt(_lib.dec((payload.get("project") or {}).get("discountRate") or 0.035))}
    result["revenue"] = {"items": revenue_items,
                         "total": [_lib.fmt(value) for value in revenue_total],
                         "outputVat": [_lib.fmt(value) for value in output_vat]}
    result["cost"] = cost if cost else None
    result["loan"] = loan
    result["tables"] = build_tables(payload, rules, years, revenue_items, revenue_total, output_vat, cost)
    result["status"] = "needs_confirm" if gaps else "ok"
    result["trace"] = {"ruleVersion": rules.get("version"), "engineVersion": ENGINE_VERSION,
                       "inputFingerprint": hashlib.sha256((json.dumps(payload, ensure_ascii=False, sort_keys=True)
                                                           + json.dumps(estimate, ensure_ascii=False, sort_keys=True)).encode("utf-8")).hexdigest()}
    out_path = out_dir / "财务测算.json"
    _lib.write_json(out_path, result, no_timestamp=args.no_timestamp)

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("[calc_finance] 收入合计={} 运营成本合计={} 所得税合计={}".format(
            _lib.fmt(sum(revenue_total)), _lib.fmt(sum(dec(v, "cost") for v in (cost or {}).get("operatingCost") or [0])),
            _lib.fmt(sum(dec(v, "tax") for v in (cost or {}).get("incomeTax") or [0]))))
        print("[calc_finance] 状态={} gaps={}".format(result["status"], len(gaps)))
        print("[calc_finance] 写出 {}".format(_lib.rel_posix(out_path, project)))
    return 3 if gaps else 0


if __name__ == "__main__":
    sys.exit(main())