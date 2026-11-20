"""现金流与指标引擎（T7）：由财务测算底表拼装全部投资/资本金现金流，计算 10 项指标与计算轨迹。

输入：`<项目>/05_过程/投资决策/{财务测算.json,投资估算.json,财务输入.json}`
口径：`../cardo-invest-spec/规则/财务口径.yaml`（cashflow / indicators / anchors 段；折现率 3.5%）
输出：回写 `财务测算.json`（cashflows + indicators）并写 `计算轨迹.json`（cardo/finance-trace/v1）

用法：
    python calc_metrics.py --help
    python calc_metrics.py --project-dir <项目目录> --no-timestamp

退出码：0 算出；1 校验失败（现金流长度不符/除零）；2 输入或口径缺失/坏；3 口径待确认（指标为 gap）
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

ENGINE_VERSION = "calc_metrics/0.2.0"
RULE_FILE = "财务口径.yaml"
DEC = Path("05_过程") / "投资决策"
FINANCE_REL = DEC / "财务测算.json"
ESTIMATE_REL = DEC / "投资估算.json"
INPUT_REL = DEC / "财务输入.json"
TRACE_REL = DEC / "计算轨迹.json"
ZERO = Decimal("0")
IRR_TOLERANCE = Decimal("1e-8")
INDICATOR_KEYS = ["firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity",
                  "ptEquity", "roe", "roi", "icr", "dscr"]
ONE = Decimal("1")
HUNDRED = Decimal("100")
TABLES_SCHEMA = "cardo/invest-cashflow-tables/v1"
FINANCE_TABLES_SCHEMA = "cardo/invest-finance-tables/v1"
TABLES_NOTE = (
    "派生块（v2/T19）：只增字段、不改旧字段、不重定义口径。"
    "costStructure/fullCostStructure/revenueStructure 由 calc_finance.py 按财务口径 cost 段汇总（运营期合计）；"
    "cashflows/debtSchedule 由 calc_metrics.py 用已算出的现金流底表与借款计划复算。"
    "手工构造的底表可能缺 costStructure 子块，此时章节按缺项处理，不补数。"
    "金额单位万元（Decimal 两位字符串）；比率单位 %（两位）。"
)
CASHFLOW_NOTE = (
    "税前 = 所得税前净现金流量（口径 cashflow.*.ebt）；税后 = 税前 - 所得税"
    "（全部投资用调整所得税，资本金用所得税）。折现系数 = 1/(1+折现率)^t（t 为年序，从 1 起；四位小数）；"
    "净现值累计 = 折现净现金流逐年累加（末行 = 税后 FNPV）。"
)
SENSITIVITY_SCHEMA = "cardo/invest-sensitivity/v1"
SENSITIVITY_NOTE = (
    "派生块（v2/T19）：单因素敏感性分析，基准 = 全部投资所得税后净现金流量（口径 indicators.firrAll 同一现金流）。"
    "要素法：仅调整该因素对应的现金流分量（投资→建设投资；收入→营业收入；经营成本→经营成本），"
    "调整所得税、回收固定资产余值等其余分量按基准情形冻结，用于比较相对影响，不替代重新测算；不改口径。"
    "变化率单位 %（deltaPercent 为百分数整数）；FIRR 为小数（与 indicators 一致，×100 即 %）；FNPV 万元两位；回收期年两位。"
)
SENSITIVITY_FACTORS = [("investment", "投资"), ("revenue", "收入"), ("cost", "经营成本")]
SENSITIVITY_DELTAS = ("-10", "0", "10")
# 临界点的可实现性判定：
#   * 只有 ±10% 是实测区间，区间外是线性外推，不能当成可执行结论；
#   * 每个因素还有自己的方向（收入提高 / 投资压减 / 成本下降才使 FNPV 上升），
#     方向反了说明符号或口径错了，必须单独标出来；
#   * 经营成本的物理上限是 −100%（成本不可能为负）——超过它的"临界点"不是结论，
#     只能用来回答"把成本压到零够不够"。
TEST_RANGE = Decimal("10")
EXPECTED_DIRECTION = {"investment": "down", "revenue": "up", "cost": "down"}
COST_FLOOR_DELTA = Decimal("-1")


class InputError(Exception):
    """输入/口径缺失 → exit 2。"""


class ValueError_(Exception):
    """校验失败 → exit 1。"""


def gap(code, what, why, need):
    return {"code": code, "what": what, "why": why, "need": need, "owner": "user"}


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
    current = Path.cwd()
    if (current / "00_项目索引.md").is_file() or (current / ".cardo").is_dir():
        return current
    raise InputError("未传 --project-dir，且当前目录不像项目")


def load_json(path, label):
    if not path.is_file():
        raise InputError("缺少{}：{}".format(label, _lib.rel_posix(path, Path.cwd())))
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("{}读取失败：{}".format(label, error))


def dec(value):
    return _lib.dec(value)


def irr(series):
    """二分法求 IRR（年序列，已含第 1..N 年现金流）。返回 Decimal 或 None。"""
    values = [float(value) for value in series]

    def npv(rate):
        total = 0.0
        for index, value in enumerate(values, start=1):
            total += value / ((1 + rate) ** index)
        return total

    previous_rate, previous_value = -0.9, npv(-0.9)
    roots = []
    for step in range(1, 600):
        rate = -0.9 + step * 0.005
        value = npv(rate)
        if abs(value) < 1e-9:
            roots.append(rate)
        elif previous_value * value < 0:
            low, high = previous_rate, rate
            f_low = previous_value
            for _ in range(120):
                mid = (low + high) / 2
                mid_value = npv(mid)
                if abs(mid_value) < 1e-10:
                    break
                if f_low * mid_value < 0:
                    high = mid
                else:
                    low = mid
                    f_low = mid_value
            roots.append((low + high) / 2)
        previous_rate, previous_value = rate, value
    if not roots:
        return None
    # Excel IRR 默认 guess=10%：多根时取最靠近 0.1 的根（避免选到负根）
    best = min(roots, key=lambda item: abs(item - 0.1))
    return Decimal(str(best))


def fnpv(series, rate):
    total = ZERO
    for index, value in enumerate(series, start=1):
        total += value / (Decimal("1") + rate) ** index
    return _lib.q2(total)


def payback(cumulative):
    """静态投资回收期：累计首次转正年 - 1 + |上年累计|/当年净现金流。"""
    for index, value in enumerate(cumulative):
        if value is not None and value >= 0 and index > 0:
            previous = cumulative[index - 1]
            current = value - (previous or ZERO)
            if current == 0:
                # 当年净现金流为 0（如资本金为零时的建设期）：累计没有因当年流入转正，
                # 真正转正要等到有净流入的年份，不能把「0」当成已回收（会少算一年）。
                continue
            # 模板口径：回收期 = 转正年序号 - |上年累计|/当年净现金流（附表6!G20 =W4+V16/W15）
            return _lib.q2(Decimal(index + 1) - (abs(previous or ZERO) / current))
    return None


def build_cashflows(finance, estimate, payload, rules):
    period = finance.get("period") or {}
    years = period.get("years") or []
    if not years:
        raise ValueError_("财务测算.json 缺 period.years")
    build_years = int(period.get("constructionYears") or 0)
    funding = payload.get("finance") or {}
    # 财务评价投资基数：默认全口径总投资；可研另有口径时用 finance.evaluationInvestment（附表8）
    total = dec(funding.get("evaluationInvestment") or estimate.get("totalInvestment"))
    construction = dec(funding.get("evaluationInvestment") or estimate.get("constructionInvestment"))
    equity = dec(funding.get("equity") or 0)
    length = len(years)
    overrides = finance.get("componentFlows") or {}

    def series(name, default):
        values = overrides.get(name)
        if values is not None:
            if len(values) != length:
                raise ValueError_("componentFlows.{} 长度不符（{} != {}）".format(name, len(values), length))
            return [dec(value) for value in values]
        return default

    residual = dec(overrides.get("residualValue") or total * Decimal("0.05"))
    revenue = [dec(value) for value in finance["revenue"]["total"]]
    cost = finance.get("cost") or {}
    operating = series("operating", [dec(value) for value in cost.get("operatingCost") or []])
    vat = series("vatSurcharge", [dec(a) + dec(b) for a, b in zip(cost.get("vatPayable") or [], cost.get("surchargeTotal") or [])])
    adjusted_tax = series("adjustedTax", [dec(value) for value in cost.get("adjustedIncomeTax") or []])
    income_tax = series("incomeTax", [dec(value) for value in cost.get("incomeTax") or []])
    interest = series("interest", [dec(value) for value in cost.get("interest") or []])
    loan = finance.get("loan") or {}
    # 本金序列来自借款还本付息表；**无借款时该表为空（长度 0）**，
    # 直接 principal[index] 会 IndexError（实测纯财政投资项目崩在这里）。
    # 长度不足一律补零（无借款 → 每年还本 0，本就是事实）。
    _principal_raw = [dec(row["principal"]) for row in (loan.get("schedule") or [])]
    if len(_principal_raw) < length:
        _principal_raw = _principal_raw + [ZERO] * (length - len(_principal_raw))
    principal = series("principal", _principal_raw)
    equity_draw_override = overrides.get("equityDraw")
    if equity_draw_override is not None:
        if len(equity_draw_override) != length:
            raise ValueError_("componentFlows.equityDraw 长度不符")
        equity_draw_override = [dec(value) for value in equity_draw_override]
    construction_series = overrides.get("construction")
    if construction_series is not None:
        if len(construction_series) != length:
            raise ValueError_("componentFlows.construction 长度不符")
        construction_by_year = [dec(value) for value in construction_series]
    else:
        construction_by_year = [ZERO] * length
        if build_years:
            per_year = _lib.q2(construction / Decimal(build_years))
            for index in range(build_years):
                construction_by_year[index] = per_year
            # 末年吸收尾差：与 02 分年投资计划同一规则。否则两年各 11,886.05、
            # 建设期流出合计 23,772.10，与总投资 23,772.09 差 0.01。
            construction_by_year[build_years - 1] = _lib.q2(
                construction - per_year * Decimal(build_years - 1))

    def assemble(equity_mode):
        inflow, outflow, net = [ZERO] * length, [ZERO] * length, [ZERO] * length
        for index in range(length):
            inflow[index] = revenue[index] + (residual if index == length - 1 else ZERO)
            if equity_mode:
                if index < build_years:
                    # 建设期只出现资本金投入；建设期增值税及附加已含在建设投资内，
                    # 再计一次会让资本金口径凭空多出 31.72 万元（与总投资对不上）。
                    year_out = (equity_draw_override[index] if equity_draw_override is not None
                                else (_lib.q2(equity / Decimal(build_years)) if build_years else equity))
                else:
                    year_out = operating[index] + vat[index] + principal[index] + interest[index]
                tax = income_tax[index]
            else:
                # 全部投资口径：建设期流出 = 分年建设投资，不含建设期增值税及附加（已含在投资内）
                if index < build_years:
                    year_out = construction_by_year[index]
                else:
                    year_out = operating[index] + vat[index]
                tax = adjusted_tax[index]
            outflow[index] = _lib.q2(year_out)
            net[index] = inflow[index] - outflow[index]
        cumulative = []
        running = ZERO
        for value in net:
            running += value
            cumulative.append(_lib.q2(running))
        return inflow, outflow, net, cumulative

    all_in = assemble(False)
    equity_flow = assemble(True)

    def rows(values, discounted=False, rate=None):
        result = []
        for index, item in enumerate(values):
            row = {"year": years[index]["year"], "value": _lib.fmt(_lib.q2(item))}
            if discounted:
                row["discounted"] = _lib.fmt(_lib.q2(item / (Decimal("1") + rate) ** (index + 1)))
            result.append(row)
        return result

    discount_rate = dec((payload.get("project") or {}).get("discountRate") or 0.035)
    all_scalars = {"inflow": all_in[0], "outflow": all_in[1], "net": all_in[2], "cumulative": all_in[3]}
    eq_scalars = {"inflow": equity_flow[0], "outflow": equity_flow[1], "net": equity_flow[2], "cumulative": equity_flow[3]}
    cashflows = {
        "allInvestment": {key: rows(values) for key, values in all_scalars.items()},
        "equity": {key: rows(values) for key, values in eq_scalars.items()},
        "discountRate": _lib.fmt(discount_rate),
        "residualValue": _lib.fmt(residual),
    }
    cashflows["allInvestment"]["netAfterTax"] = rows([all_in[2][i] - adjusted_tax[i] for i in range(length)],
                                                     discounted=True, rate=discount_rate)
    cashflows["equity"]["netAfterTax"] = rows([eq_scalars["net"][i] - income_tax[i] for i in range(length)],
                                              discounted=True, rate=discount_rate)

    def cumulative_of(series):
        result = []
        running = ZERO
        for value in series:
            running += value
            result.append(_lib.q2(running))
        return result

    after_all = [all_in[2][i] - adjusted_tax[i] for i in range(length)]
    after_equity = [eq_scalars["net"][i] - income_tax[i] for i in range(length)]
    return cashflows, {"all": all_in, "equity": equity_flow,
                       "afterAll": after_all, "afterEquity": after_equity,
                       "afterAllCum": cumulative_of(after_all), "afterEquityCum": cumulative_of(after_equity),
                       "incomeTax": income_tax, "interest": interest, "principal": principal,
                       "operating": operating, "depreciation": [dec(v) for v in cost.get("depreciation") or []],
                       "ebit": [dec(v) for v in cost.get("ebit") or []], "ebitda": [dec(v) for v in cost.get("ebitda") or []],
                       "discountRate": discount_rate,
                       # 派生表/敏感性用分量（不写盘，仅供本脚本内复算）
                       "revenue": revenue, "vatSurcharge": vat, "construction": construction_by_year,
                       "adjustedTax": adjusted_tax, "residual": residual, "buildYears": build_years,
                       "years": [info.get("year") for info in years]}, length


def compute_indicators(finance, estimate, payload, rules, flows):
    indicators = {}
    gaps = []
    after_all, after_equity = flows["afterAll"], flows["afterEquity"]
    discount = flows["discountRate"]
    firr_all = irr(after_all)
    firr_eq = irr(after_equity)
    indicators["firrAll"] = _lib.fmt(_lib.q4(firr_all)) if firr_all is not None else None
    indicators["firrEquity"] = _lib.fmt(_lib.q4(firr_eq)) if firr_eq is not None else None
    indicators["fnpvAll"] = _lib.fmt(fnpv(after_all, discount))
    indicators["fnpvEquity"] = _lib.fmt(fnpv(after_equity, discount))
    indicators["ptAll"] = _lib.fmt(payback(flows["afterAllCum"])) if payback(flows["afterAllCum"]) is not None else None
    # 资本金 ≤0 时不存在「资本金回收期」：早先 payback() 对 0,0,377.01… 的序列
    # 机械返回 3 年，资本金为零却对外讲「3 年回本」，与 FNPV 为负并列更讲不通。
    equity = dec((payload.get("finance") or {}).get("equity") or 0)
    pt_equity = payback(flows["afterEquityCum"]) if equity > 0 else None
    indicators["ptEquity"] = _lib.fmt(pt_equity) if pt_equity is not None else None
    total = dec(estimate.get("totalInvestment"))
    cost = finance.get("cost") or {}
    taxable = [dec(value) for value in cost.get("taxableIncome") or []]
    income_tax = flows["incomeTax"]
    op_years = int(finance["period"].get("operationYears") or len(taxable) or 1)
    net_profit = [taxable[i] - income_tax[i] for i in range(min(len(taxable), len(income_tax)))]
    operation_net = net_profit[-op_years:] if op_years <= len(net_profit) else net_profit
    avg_net_profit = sum(operation_net) / Decimal(len(operation_net)) if operation_net else ZERO
    ebit = flows["ebit"]
    operation_ebit = ebit[-op_years:] if op_years <= len(ebit) else ebit
    avg_ebit = sum(operation_ebit) / Decimal(len(operation_ebit)) if operation_ebit else ZERO
    indicators["roe"] = _lib.fmt(_lib.q4(avg_net_profit / equity)) if equity else None
    indicators["roi"] = _lib.fmt(_lib.q4(avg_ebit / total)) if total else None
    interest = flows["interest"]
    principal = flows["principal"]
    build_years = int((finance.get("period") or {}).get("constructionYears") or 0)
    op_ebit = sum(ebit[build_years:]) if len(ebit) > build_years else sum(ebit)
    op_interest = sum(interest[build_years:]) if len(interest) > build_years else sum(interest)
    total_interest = sum(interest)
    # 汇总口径（运营期合计）；逐年值另见 coverage 数组
    indicators["icr"] = _lib.fmt(_lib.q4(op_ebit / op_interest)) if op_interest else None
    debt_service = sum(principal) + total_interest
    indicators["dscr"] = _lib.fmt(_lib.q4((sum(flows["ebitda"]) - sum(income_tax)) / debt_service)) if debt_service else None
    coverage = {"basis": "aggregate-operation-period",
                "note": "ICR=运营期ΣEBIT/Σ利息；DSCR=(ΣEBITDA-Σ所得税)/Σ(还本+利息)；逐年值见 icrByYear/dscrByYear",
                "icrByYear": [], "dscrByYear": []}
    for index in range(len(ebit)):
        year_interest = interest[index] if index < len(interest) else ZERO
        year_principal = principal[index] if index < len(principal) else ZERO
        year_tax = income_tax[index] if index < len(income_tax) else ZERO
        year_ebitda = flows["ebitda"][index] if index < len(flows["ebitda"]) else ZERO
        coverage["icrByYear"].append(_lib.fmt(_lib.q4(ebit[index] / year_interest)) if year_interest else None)
        coverage["dscrByYear"].append(_lib.fmt(_lib.q4((year_ebitda - year_tax) / (year_principal + year_interest)))
                                       if (year_principal + year_interest) else None)
    # 大修年缺口与累计结余：还本期首个「EBITDA − 所得税 < 应付本息」的年份（大修年）。
    # 只写「需靠结余」不够，必须回答结余够不够——这是全表最脆弱的一年。
    ebitda = flows["ebitda"]
    overhaul = ((cost.get("items") or {}).get("overhaul") or [])
    repair = None
    for index in range(build_years, len(ebitda)):
        year_principal = principal[index] if index < len(principal) else ZERO
        if year_principal <= 0:
            continue
        year_interest = interest[index] if index < len(interest) else ZERO
        year_tax = income_tax[index] if index < len(income_tax) else ZERO
        year_ebitda = ebitda[index] if index < len(ebitda) else ZERO
        service = year_principal + year_interest
        if year_ebitda - year_tax >= service:
            continue
        reserve = ZERO
        for year_index in range(build_years, index):
            prior_service = ((principal[year_index] if year_index < len(principal) else ZERO)
                             + (interest[year_index] if year_index < len(interest) else ZERO))
            prior_tax = income_tax[year_index] if year_index < len(income_tax) else ZERO
            prior_ebitda = ebitda[year_index] if year_index < len(ebitda) else ZERO
            reserve += prior_ebitda - prior_tax - prior_service
        shortfall = service + year_tax - year_ebitda
        repair = {
            "year": index + 1,
            "overhaul": _lib.fmt(_lib.q2(dec(overhaul[index])))
            if index < len(overhaul) and overhaul[index] is not None else None,
            "ebitda": _lib.fmt(_lib.q2(year_ebitda)),
            "incomeTax": _lib.fmt(_lib.q2(year_tax)),
            "debtService": _lib.fmt(_lib.q2(service)),
            "shortfall": _lib.fmt(_lib.q2(shortfall)),
            "reserveBefore": _lib.fmt(_lib.q2(reserve)),
            "reserveAfter": _lib.fmt(_lib.q2(reserve - shortfall)),
            "covered": bool(reserve >= shortfall),
            "dscr": coverage["dscrByYear"][index] if index < len(coverage["dscrByYear"]) else None,
            "fromYear": build_years + 1,
            "toYear": index,
            "basis": "还本期首个 EBITDA 不足以覆盖所得税与本息的年份；结余=此前各年（EBITDA−所得税−本息）累计",
        }
        break
    coverage["repairYear"] = repair
    # 申报上限全额发行反证：判断「申报额度与实际需求是否匹配」，不进入基准现金流。
    loan_info = finance.get("loan") or {}
    declared_loan = dec((payload.get("finance") or {}).get("declaredLoan") or 0)
    loan_rate = dec(loan_info.get("loanRate") or 0)
    loan_term = int(dec(loan_info.get("loanTerm") or 0))
    normal_values = [ebitda[index] for index in range(build_years, len(ebitda))
                     if not (index < len(overhaul) and dec(overhaul[index] or 0) > 0)]
    normal = sum(normal_values) / Decimal(len(normal_values)) if normal_values else None
    declared = None
    if declared_loan > 0 and loan_rate > 0 and loan_term > 0 and normal is not None:
        annual_service = declared_loan / Decimal(loan_term) + declared_loan * loan_rate
        declared = {
            "declaredLoan": _lib.fmt(_lib.q2(declared_loan)),
            "loanRate": _lib.fmt(loan_rate),
            "loanTerm": loan_term,
            "annualDebtService": _lib.fmt(_lib.q2(annual_service)),
            "ebitdaNormal": _lib.fmt(_lib.q2(normal)),
            "dscr": _lib.fmt(_lib.q4(normal / annual_service)) if annual_service else None,
            "basis": "若专项债按申报上限全额发行：等额还本+当年利息；EBITDA 取运营期均值、剔除大修年",
        }
    coverage["declaredScenario"] = declared
    for key in INDICATOR_KEYS:
        if indicators.get(key) is None and key not in ("ptAll", "ptEquity"):
            gaps.append(gap("INDICATOR-GAP", "指标 {} 未能计算（现金流/除零）".format(key), "缺数据", "补充输入"))
        elif indicators.get(key) is None:
            indicators[key] = None  # 计算期内未回收：显式 null，不补数、不算 gap
    return indicators, gaps, coverage


def cumulative(values):
    """逐年累计（两位）。"""
    result = []
    running = ZERO
    for value in values:
        running += value
        result.append(_lib.q2(running))
    return result


def discounted(values, rate):
    """折现：values[t-1]/(1+rate)^t；返回未取整的 Decimal 序列（累计后再取整，保证末行 = FNPV）。"""
    return [value / (ONE + rate) ** (index + 1) for index, value in enumerate(values)]


def fmt_or_none(value, quantize=_lib.q2):
    return _lib.fmt(quantize(value)) if value is not None else None


def build_cashflow_tables(flows, length, build_years, discount_rate):
    """现金流表（税前/税后 + 折现系数 + 净现值累计），全部投资与资本金各一张。"""
    years = flows.get("years") or [index + 1 for index in range(length)]

    def build_table(name, title, basis, inflow, outflow, net_before, cum_before, tax, net_after, cum_after):
        disc_before = discounted(net_before, discount_rate)
        disc_after = discounted(net_after, discount_rate)
        cum_disc_before = cumulative(disc_before)
        cum_disc_after = cumulative(disc_after)
        rows = []
        for index in range(length):
            factor = _lib.q4(ONE / (ONE + discount_rate) ** (index + 1))
            rows.append({
                "year": years[index],
                "phase": "建设期" if index < build_years else "运营期",
                "inflow": fmt_or_none(inflow[index]),
                "outflow": fmt_or_none(outflow[index]),
                "netBeforeTax": fmt_or_none(net_before[index]),
                "cumulativeBeforeTax": fmt_or_none(cum_before[index]),
                "discountFactor": _lib.fmt(factor),
                "discountedBeforeTax": fmt_or_none(disc_before[index]),
                "cumulativeDiscountedBeforeTax": fmt_or_none(cum_disc_before[index]),                "incomeTax": fmt_or_none(tax[index]),
                "netAfterTax": fmt_or_none(net_after[index]),
                "cumulativeAfterTax": fmt_or_none(cum_after[index]),
                "discountedAfterTax": fmt_or_none(disc_after[index]),
                "cumulativeDiscountedAfterTax": fmt_or_none(cum_disc_after[index]),
            })
        before_firr = irr(net_before)
        after_firr = irr(net_after)
        summary = {
            "firrBeforeTax": _lib.fmt(_lib.q4(before_firr)) if before_firr is not None else None,
            "firrAfterTax": _lib.fmt(_lib.q4(after_firr)) if after_firr is not None else None,
            "fnpvBeforeTax": _lib.fmt(fnpv(net_before, discount_rate)),
            "fnpvAfterTax": _lib.fmt(fnpv(net_after, discount_rate)),
            "ptBeforeTax": fmt_or_none(payback(cum_before)),
            "ptAfterTax": fmt_or_none(payback(cum_after)),
        }
        return {"key": name, "title": title, "basis": basis, "rows": rows, "summary": summary}

    all_in = flows["all"]
    equity = flows["equity"]
    return {
        "schema": TABLES_SCHEMA,
        "schemaVersion": 1,
        "note": CASHFLOW_NOTE,
        "unit": "万元",
        "discountRate": _lib.fmt(discount_rate),
        "totalYears": length,
        "constructionYears": build_years,
        "allInvestment": build_table(
            "allInvestment", "项目投资现金流量表（全部投资）",
            "现金流入 = 营业收入 + 回收固定资产余值；现金流出 = 建设投资 + 经营成本 + 增值税及附加（税后再扣调整所得税）",
            all_in[0], all_in[1], all_in[2], all_in[3], flows["adjustedTax"],
            flows["afterAll"], flows["afterAllCum"]),
        "equity": build_table(
            "equity", "项目资本金现金流量表",
            "现金流入 = 营业收入 + 回收固定资产余值；现金流出 = 项目资本金 + 经营成本 + 增值税及附加 + 借款本金偿还 + 借款利息支付（税后再扣所得税）",
            equity[0], equity[1], equity[2], equity[3], flows["incomeTax"],
            flows["afterEquity"], flows["afterEquityCum"]),
    }


def build_debt_schedule(finance, build_years):
    """偿债表：年初余额 / 当年提款 / 当年还本 / 当年利息 / 年末余额（来自借款还本付息计划）。"""
    loan = finance.get("loan") or {}
    schedule = loan.get("schedule") or []
    rows = []
    total_principal = ZERO
    total_interest = ZERO
    for row in schedule:
        opening = dec(row.get("opening") or 0)
        draw = dec(row.get("draw") or 0)
        principal = dec(row.get("principal") or 0)
        interest = dec(row.get("interest") or 0)
        closing = dec(row.get("closing") or 0)
        total_principal += principal
        total_interest += interest
        rows.append({
            "year": row.get("year"),
            "phase": "建设期" if (row.get("year") or 0) <= build_years else "运营期",
            "opening": fmt_or_none(opening),
            "draw": fmt_or_none(draw),
            "principal": fmt_or_none(principal),
            "interest": fmt_or_none(interest),
            "debtService": fmt_or_none(principal + interest),
            "closing": fmt_or_none(closing),
        })
    return {
        "schema": TABLES_SCHEMA,
        "schemaVersion": 1,
        "note": "派生块（v2/T19）：逐行来自 财务测算.json#loan.schedule（口径 附表8 借款还本付息计划），"
                "不重算利息、不改还本方式；年初余额 + 当年提款 - 当年还本 = 年末余额。金额单位万元。",
        "unit": "万元",
        "loanRate": loan.get("loanRate"),
        "loanTerm": loan.get("loanTerm"),
        "graceYears": loan.get("graceYears"),
        "rows": rows,
        "summary": {
            "totalDraw": fmt_or_none(sum((dec(row.get("draw") or 0) for row in schedule), ZERO)),
            "totalPrincipal": fmt_or_none(total_principal),
            "totalInterest": fmt_or_none(total_interest),
            "totalDebtService": fmt_or_none(total_principal + total_interest),
            "closingEnd": fmt_or_none(dec(schedule[-1].get("closing") or 0)) if schedule else None,
        },
    }


def factor_series(flows, length, key, delta):
    """按因素生成要素法调整后的全部投资税后净现金流量（其余分量冻结）。

    逐年构成必须与 build_cashflows.assemble(False) 完全一致：建设期流出只有
    分年建设投资（建设期增值税及附加已含在投资内，不再单列），运营期流出
    为经营成本 + 增值税及附加。此前建设期多扣一次 vat，导致基准序列 FNPV
    比 indicators.fnpvAll 少 30.13 万元，敏感性基准与指标表对不上。
    """
    revenue = flows["revenue"]
    construction = flows["construction"]
    operating = flows["operating"]
    vat = flows["vatSurcharge"]
    adjusted_tax = flows["adjustedTax"]
    residual = flows["residual"]
    build_years = int(flows.get("buildYears") or 0)
    series = []
    for index in range(length):
        inflow = revenue[index]
        if index == length - 1:
            inflow += residual
        construction_year = construction[index]
        operating_year = operating[index]
        if key == "investment":
            construction_year = construction_year * (ONE + delta)
        elif key == "revenue":
            inflow = revenue[index] * (ONE + delta)
            if index == length - 1:
                inflow += residual
        elif key == "cost":
            operating_year = operating_year * (ONE + delta)
        if index < build_years:
            series.append(inflow - construction_year - adjusted_tax[index])
        else:
            series.append(inflow - construction_year - operating_year - vat[index] - adjusted_tax[index])
    return series


def build_sensitivity(flows, length, discount_rate, base_indicators):
    """敏感性分析：投资 / 收入 / 经营成本 三因素，-10% / 0 / +10% 对 FIRR、FNPV、静态回收期。

    单一锚点：基准一律取 indicators（firrAll / fnpvAll / ptAll），基准序列直接复用
    flows["afterAll"]，delta 一律按「本行值 − 基准值」现算，不再另存一套序列值。
    """
    base_series = [dec(value) for value in flows["afterAll"]]
    base_fnpv = dec(base_indicators.get("fnpvAll"))
    base_firr = dec(base_indicators.get("firrAll"))
    factors = []
    ranking = []
    for key, name in SENSITIVITY_FACTORS:
        rows = []
        for delta_text in SENSITIVITY_DELTAS:
            delta = Decimal(delta_text) / HUNDRED
            series = base_series if delta == ZERO else factor_series(flows, length, key, delta)
            firr_value = irr(series)
            npv_value = fnpv(series, discount_rate)
            pt_value = payback(cumulative(series))
            rows.append({
                "deltaPercent": delta_text,
                "label": "基准" if delta == ZERO else ("{}%".format(delta_text) if Decimal(delta_text) < 0
                                                        else "+{}%".format(delta_text)),
                "firr": _lib.fmt(_lib.q4(firr_value)) if firr_value is not None else None,
                "fnpv": _lib.fmt(npv_value),
                "pt": fmt_or_none(pt_value),
                "fnpvDelta": _lib.fmt(_lib.q2(npv_value - base_fnpv)),
                "firrDelta": _lib.fmt(_lib.q4(firr_value - base_firr)) if (firr_value is not None and base_firr is not None) else None,
            })
        negative = rows[0]
        positive = rows[-1]
        zero = None
        direction = None
        consistent = None
        if negative["fnpv"] is not None and positive["fnpv"] is not None:
            npv_low = dec(negative["fnpv"])
            npv_high = dec(positive["fnpv"])
            if npv_high != npv_low:
                # 线性外推求 FNPV=0 的变化率：区间内是插值，区间外是外推。
                # 外推出来的幅度本来就该写出来（"收入要提高 95.6%" 是结论），
                # 早先只在 ±10% 内出现变号才给值，正文就只能写"临界点未形成"，
                # 等于把"要涨一倍才持平"这个最重要的判断藏掉。
                zero = _lib.fmt(_lib.q2(Decimal("-10") + (ZERO - npv_low)
                                        / (npv_high - npv_low) * Decimal("20")))
                direction = "up" if npv_high > npv_low else "down"
                consistent = direction == EXPECTED_DIRECTION[key]
        inside = zero is not None and abs(dec(zero)) <= TEST_RANGE
        # 临界点能不能当结论用：方向必须对，且必须落在实测的 ±10% 区间内。
        # 区间外是外推值，只能作量级证据（写进正文时改成否定式结论）。
        feasible = bool(inside and consistent)
        at_limit = None
        if key == "cost":
            # 成本压减的物理上限：压到零（−100%）。用它回答"最多能压多少、够不够"。
            floor_fnpv = fnpv(factor_series(flows, length, key, COST_FLOOR_DELTA), discount_rate)
            at_limit = {"deltaPercent": "-100", "fnpv": _lib.fmt(_lib.q2(floor_fnpv))}
        factors.append({
            "key": key,
            "name": name,
            "rows": rows,
            "fnpvZeroDeltaPercent": zero,
            "fnpvZeroDirection": direction,
            "fnpvZeroExpectedDirection": EXPECTED_DIRECTION[key],
            "fnpvZeroDirectionOk": consistent,
            "fnpvZeroInTestRange": inside,
            "fnpvZeroFeasible": feasible,
            "fnpvZeroAtLimit": at_limit,
            "spans": {
                "firrSpan": _lib.fmt(_lib.q4(abs(dec(positive["firr"]) - dec(negative["firr"]))))
                if (positive["firr"] is not None and negative["firr"] is not None) else None,
                "fnpvSpan": _lib.fmt(_lib.q2(abs(dec(positive["fnpv"]) - dec(negative["fnpv"])))),
                "ptSpan": _lib.fmt(_lib.q2(abs(dec(positive["pt"]) - dec(negative["pt"]))))
                if (positive["pt"] is not None and negative["pt"] is not None) else None,
            },
        })
        ranking.append((key, abs(dec(positive["fnpv"]) - dec(negative["fnpv"]))))
    ranking.sort(key=lambda item: item[1], reverse=True)
    return {
        "schema": SENSITIVITY_SCHEMA,
        "schemaVersion": 1,
        "note": SENSITIVITY_NOTE,
        "basis": "全部投资所得税后净现金流量（与 indicators.firrAll / fnpvAll / ptAll 同口径）",
        "method": "单因素要素法",
        "discountRate": _lib.fmt(discount_rate),
        "variations": [int(text) for text in SENSITIVITY_DELTAS],
        "base": {
            "firr": base_indicators.get("firrAll"),
            "fnpv": base_indicators.get("fnpvAll"),
            "pt": base_indicators.get("ptAll"),
        },
        "factors": factors,
        "ranking": [key for key, _ in ranking],
        "rankingBasis": "按 ±10% 区间 FNPV 变动幅度降序（幅度越大越敏感）",
        "zeroPointBasis": ("±10% 区间内为线性插值，可作为临界点使用；区间外为线性外推，"
                           "只作量级证据，不构成可实现结论（fnpvZeroFeasible=false）。"
                           "经营成本另给 −100%（成本压至零）时的 FNPV，用于判断物理上限是否够用。"),
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 现金流与指标引擎（T7）")
    parser.add_argument("--project-dir", default=None)
    parser.add_argument("--rules", default=None)
    parser.add_argument("--format", choices=("text", "json"), default="text")
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args(argv)

    rules_dir = Path(args.rules) if args.rules else RULES_DIR_DEFAULT
    try:
        rules = load_rules(rules_dir)
        project = resolve_project(args)
        finance = load_json(project / FINANCE_REL, "财务测算")
        estimate = load_json(project / ESTIMATE_REL, "投资估算")
        payload = load_json(project / INPUT_REL, "财务输入")
        cashflows, flows, length = build_cashflows(finance, estimate, payload, rules)
        indicators, gaps, coverage = compute_indicators(finance, estimate, payload, rules, flows)
    except InputError as error:
        print("[calc_metrics] {}".format(error))
        return 2
    except ValueError_ as error:
        print("[calc_metrics] 校验失败：{}".format(error))
        return 1

    finance["cashflows"] = cashflows
    finance["indicators"] = {key: (_lib.dec(value) if value is not None else None) for key, value in indicators.items()}
    finance["indicators"] = indicators
    finance["coverage"] = coverage
    build_years = int((finance.get("period") or {}).get("constructionYears") or 0)
    existing_tables = finance.get("tables") if isinstance(finance.get("tables"), dict) else {}
    tables = {
        "schema": FINANCE_TABLES_SCHEMA,
        "schemaVersion": 1,
        "note": TABLES_NOTE,
        "unit": "万元",
        "ratioUnit": "%",
    }
    for key in ("period", "costStructure", "fullCostStructure", "revenueStructure"):
        if key in existing_tables:
            tables[key] = existing_tables[key]
    tables["cashflows"] = build_cashflow_tables(flows, length, build_years, flows["discountRate"])
    tables["debtSchedule"] = build_debt_schedule(finance, build_years)
    finance["tables"] = tables
    finance["sensitivity"] = build_sensitivity(flows, length, flows["discountRate"], indicators)
    finance["status"] = "needs_confirm" if gaps else "ok"
    existing = finance.get("gaps") or []
    finance["gaps"] = existing + [item for item in gaps if item not in existing]
    trace = {
        "schema": "cardo/finance-trace/v1",
        "schemaVersion": 1,
        "generatedAt": _lib.iso_now(args.no_timestamp),
        "producer": "cardo-invest-finance/" + ENGINE_VERSION,
        "ruleVersion": rules.get("version"),
        "engineVersion": ENGINE_VERSION,
        "inputFingerprint": hashlib.sha256(json.dumps(
            {"finance": finance.get("revenue"), "estimate": estimate, "payload": payload},
            ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest(),
        "anchors": (rules.get("anchors") or {}).get("items"),
        "steps": [
            {"id": "CF-1", "name": "全部投资税后净现金流", "formula": "（收入+余值）-（建设投资+经营成本+增值税及附加+调整所得税）", "years": length},
            {"id": "CF-2", "name": "资本金税后净现金流", "formula": "（收入+余值）-（资本金+经营成本+增值税及附加+还本+利息+所得税）", "years": length},
            {"id": "IND-1", "name": "FIRR", "formula": "二分法求 NPV=0（1e-8）", "values": {key: indicators[key] for key in ("firrAll", "firrEquity")}},
            {"id": "IND-2", "name": "FNPV", "formula": "Σ 净现金流/(1+折现率)^t", "discountRate": _lib.fmt(flows["discountRate"]),
             "values": {key: indicators[key] for key in ("fnpvAll", "fnpvEquity")}},
            {"id": "IND-3", "name": "静态回收期", "formula": "累计首次转正年-1+|上年累计|/当年净现金流",
             "values": {key: indicators[key] for key in ("ptAll", "ptEquity")}},
            {"id": "IND-4", "name": "ROE/ROI/ICR/DSCR", "formula": "见财务口径 indicators",
             "values": {key: indicators[key] for key in ("roe", "roi", "icr", "dscr")}},
            {"id": "SENS-1", "name": "敏感性分析（派生表）",
             "formula": "单因素要素法：投资/收入/经营成本 各 -10%/0/+10% → 重算 FIRR/FNPV/静态回收期；其余分量冻结",
             "basis": "全部投资所得税后净现金流量", "factors": [key for key, _ in SENSITIVITY_FACTORS],
             "variations": [int(text) for text in SENSITIVITY_DELTAS], "output": "财务测算.json#sensitivity"},
        ],
    }
    finance["trace"] = trace
    out_dir = project / DEC
    out_dir.mkdir(parents=True, exist_ok=True)
    _lib.write_json(out_dir / "财务测算.json", finance, no_timestamp=args.no_timestamp)
    _lib.write_json(out_dir / "计算轨迹.json", trace, no_timestamp=args.no_timestamp)

    if args.format == "json":
        print(json.dumps({"indicators": indicators, "gaps": gaps}, ensure_ascii=False, indent=2))
    else:
        print("[calc_metrics] FIRR 全部={} 资本金={}；FNPV 全部={} 资本金={}；PT 全部={} 资本金={}".format(
            indicators["firrAll"], indicators["firrEquity"], indicators["fnpvAll"],
            indicators["fnpvEquity"], indicators["ptAll"], indicators["ptEquity"]))
        print("[calc_metrics] ROE={} ROI={} ICR={} DSCR={}".format(
            indicators["roe"], indicators["roi"], indicators["icr"], indicators["dscr"]))
        print("[calc_metrics] 派生表：现金流表（税前/税后）/ 偿债表 / 敏感性（{} 因素 × {}）".format(
            len(SENSITIVITY_FACTORS), "/".join(text + "%" for text in SENSITIVITY_DELTAS)))
        print("[calc_metrics] 写出 {} / 计算轨迹.json（gaps={}）".format(
            _lib.rel_posix(out_dir / "财务测算.json", project), len(gaps)))
    return 3 if gaps else 0


if __name__ == "__main__":
    sys.exit(main())