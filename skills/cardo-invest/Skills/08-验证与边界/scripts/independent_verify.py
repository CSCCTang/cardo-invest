"""08 · 独立验证与边界（不复用引擎函数、不读模板结果单元格）。

读（相对工作区根）
  过程文件/00-总控/上游机读件/权重表.json（AHP 只读对账）
  过程文件/01-上游取料/财务输入.json
  过程文件/02-匡算/投资估算.json
  过程文件/03-财务评价/财务测算.json
  过程文件/04-三方案比选/方案比选.json
  过程文件/05-章节成文/分章/*.md
  成果文件/项目投资决策建议书.docx、投资估算.xlsx、财务评价.xlsx、方案比选.xlsx
  记忆与规则/验证基线-CardoInvest.json（可删）
写
  过程文件/08-验证与边界/独立验证结果.json
  过程文件/08-验证与边界/产物指纹.json
  记忆与规则/验证基线-CardoInvest.json（--baseline write 或首次 auto 且不存在时）

口径（写死在 规则/验证口径.yaml 同值，这里不读取规则，避免与引擎共用解析）
  锚点容差：0.5% / 1 万元 / 0.5 年；成品容差：0.04pp / 0.01 万元 / 0.01 年
  IRR：自实现二分（根取最接近 0.1）；NPV：Σ CFt/(1+r)^t（t 从 1 起）；PT：累计转正年 − 1 + |上年累计|/当年净现金流

用法：python independent_verify.py --workspace . [--baseline auto|write|check|none]
退出码：0 全过；1 有失败项；2 输入缺失。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from decimal import Decimal, getcontext
from pathlib import Path

getcontext().prec = 28

# 外来专名扫描器与 07 回读自检共用一份实现（skills/07-回读自检/scripts/check_invest.py）
SKILLS_ROOT = Path(__file__).resolve().parents[2]
if str(SKILLS_ROOT / "07-回读自检" / "scripts") not in sys.path:
    sys.path.insert(0, str(SKILLS_ROOT / "07-回读自检" / "scripts"))

import _lib  # noqa: E402
try:
    from check_invest import scan_foreign_names  # noqa: E402
except Exception:  # pragma: no cover - 07 缺失时显式失败，不静默放行
    scan_foreign_names = None


def load_foreign_whitelist(workspace):
    """本项目合法专名的来源：上游摘要全文 + 财务输入（都来自本项目材料）。

    编制依据库的机关名/文件全称是规范引用，不是外来专名：连同运行期核验留痕一起
    纳入合法词表，与 07 的 DOC-FOREIGN 口径保持一致。
    """
    parts = []
    for candidate in (workspace / "过程文件" / "01-上游取料" / "上游摘要.json",
                      workspace / "过程文件" / "01-上游取料" / "财务输入.json",
                      # 三份社区台账汇总（示例片区B/示例社区B与各院落名）属本项目提资，
                      # 不是外来专名；与 07 的合法词表保持同一口径。
                      workspace / "过程文件" / "00-总控" / "引擎项目" / "05_过程" / "投资决策" / "现状台账.json",
                      workspace / "过程文件" / "05-章节成文" / "依据核验.json",
                      SKILLS_ROOT.parent / "assets" / "依据库" / "依据库.json"):
        if candidate.is_file():
            try:
                parts.append(candidate.read_text(encoding="utf-8", errors="replace"))
            except OSError:
                continue
    return "\n".join(parts)


SCHEMA_RESULT = "cardo-invest/independent-verification/v1"
SCHEMA_BASELINE = "cardo-invest/verification-baseline/v1"
PACKAGE = "CardoInvest"
TOL_FIRR_PP = Decimal("0.04")     # 成品容差：百分点
TOL_AMOUNT = Decimal("0.01")      # 成品容差：万元
TOL_YEARS = Decimal("0.01")       # 成品容差：年
TOL_ANCHOR_AMOUNT = Decimal("1")
TOL_ANCHOR_FIRR = Decimal("0.005")
TOL_ANCHOR_YEARS = Decimal("0.5")
CHAPTER_TITLES = [
    "1 项目概述",
    "2 项目背景与建设必要性",
    "3 现状调研分析",
    "4 市场分析与功能定位",
    "5 项目建设范围、改造内容与建设方案",
    "6 组织架构与运营模式",
    "7 投资估算与资金筹措方案",
    "8 财务评价",
    "9 基于AHP层次分析法的投资决策指标体系量化分析",
    "10 附件清单",
]
INPUT_ALLOWED = {".docx", ".doc", ".pdf", ".xlsx", ".xls", ".pptx", ".json", ".zip",
                 # 真实用户会一股脑把整个「输入」文件夹拷进来：现场照片、公示图、CAD、压缩包
                 # 都是合法输入材料，不能因为扩展名不是 Office 就判边界违规。
                 ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff",
                 ".dwg", ".dxf", ".rar", ".7z", ".txt", ".md", ".csv", ".geojson", ".gpkg"}
OUTPUT_ALLOWED = {".docx", ".xlsx", ".pdf"}
SKILL_ALLOWED = {".py", ".md", ".yaml", ".yml"}
# 盘符：字母 + 冒号 + 斜杠/反斜杠。前面不能紧跟字母/数字/+/./-，
# 否则 `https:/`、`ftp:/`、`file:/` 这类 URL scheme 会被误判成盘符（实测踩过）。
DRIVE_PATTERN = re.compile(r"(?<![A-Za-z0-9+.\-])[A-Za-z]:[\\/]")
PLACEHOLDER_PATTERNS = ("{{", "}}", "待补", "TODO", "XXX", "（数据来源：")


# ---------------------------------------------------------------- 工具

def read_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return None


def write_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def dec(value):
    if value is None:
        return None
    return Decimal(str(value))


def fnum(value):
    """展示值 → Decimal；"—"/None → None；带单位（「3.1 年」）取数值部分。

    表位表会把单位写进结果格（「3.1 年」），早先整串 Decimal 解析失败 → None →
    被当成「docx 缺值」误判。这里按第一个数值取数，与 07 的回读口径一致。
    """
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("%", "")
    if text in ("", "—", "-"):
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return Decimal(match.group(0))
    except Exception:
        return None


def rel(path, base):
    try:
        return Path(path).resolve().relative_to(Path(base).resolve()).as_posix()
    except ValueError:
        return Path(path).name


class Checker:
    def __init__(self):
        self.checks = []

    def add(self, code, name, ok, detail="", warn=False):
        # warn=True 表示「通过但必须披露」：ok 仍为真（不拦流水线），
        # 单独进 summary.warned / result.warnings，验证报告里不许删行。
        self.checks.append({"id": code, "name": name, "ok": bool(ok),
                            "warn": bool(warn) and bool(ok), "detail": str(detail)[:400]})

    def near(self, code, name, actual, expected, tol, unit="", detail=""):
        actual_dec, expected_dec = dec(actual), dec(expected)
        if actual_dec is None or expected_dec is None:
            self.add(code, name, False, "缺值：actual={} expected={}".format(actual, expected))
            return
        diff = abs(actual_dec - expected_dec)
        self.add(code, name, diff <= tol,
                 "实际 {} / 期望 {} / 差 {} {}（容差 {}） {}".format(
                     actual_dec, expected_dec, diff, unit, tol, detail))


# ---------------------------------------------------------------- 独立算法

def irr(values):
    """自实现二分法：在 [-0.9, 1.0] 上扫根，多根取最接近 10% 的根。"""
    numbers = [float(value) for value in values]

    def npv(rate):
        return sum(value / ((1 + rate) ** index) for index, value in enumerate(numbers, start=1))

    roots = []
    low_rate, low_value = -0.9, npv(-0.9)
    for step in range(1, 761):
        rate = -0.9 + step * 0.0025
        value = npv(rate)
        if abs(value) < 1e-9:
            roots.append(rate)
        elif low_value * value < 0:
            low, high, f_low = low_rate, rate, low_value
            for _ in range(160):
                middle = (low + high) / 2
                f_middle = npv(middle)
                if abs(f_middle) < 1e-11:
                    break
                if f_low * f_middle < 0:
                    high = middle
                else:
                    low, f_low = middle, f_middle
            roots.append((low + high) / 2)
        low_rate, low_value = rate, value
    if not roots:
        return None
    return Decimal(str(min(roots, key=lambda item: abs(item - 0.1))))


def npv(values, rate):
    total = Decimal("0")
    for index, value in enumerate(values, start=1):
        total += dec(value) / (Decimal("1") + rate) ** index
    return total.quantize(Decimal("0.01"))


def payback(values):
    cumulative = Decimal("0")
    for index, value in enumerate(values, start=1):
        previous = cumulative
        cumulative += dec(value)
        if cumulative >= 0:
            current = dec(value)
            if current == 0:
                # 与 03 引擎同口径：当年净现金流为 0 不算回收，等真正有净流入的年份。
                continue
            return (Decimal(index) - abs(previous) / current).quantize(Decimal("0.01"))
    return None


# ---------------------------------------------------------------- 检查：匡算

def check_estimate(checker, payload, estimate):
    engineering_items = (payload.get("estimate") or {}).get("engineering") or []
    engineering = sum(dec(item.get("amount") or 0) for item in engineering_items)
    checker.near("EST-ENGINEERING", "工程费合计独立复算", engineering, estimate.get("engineering"),
                 TOL_AMOUNT, "万元", "sum(engineering[].amount)")

    other = Decimal("0")
    other_complete = True
    for item in (payload.get("estimate") or {}).get("otherFees") or []:
        if item.get("amount") is not None:
            other += dec(item.get("amount"))
        elif item.get("basis") is not None and item.get("rate") is not None:
            other += dec(item.get("basis")) * dec(item.get("rate"))
        else:
            other_complete = False
    checker.add("EST-OTHER-COMPLETE", "其他费明细齐全（无缺数项）", other_complete)
    checker.near("EST-OTHER", "工程建设其他费独立复算", other, estimate.get("otherFeeTotal"),
                 TOL_AMOUNT, "万元", "amount 或 取费基数×费率")

    rate = ((payload.get("estimate") or {}).get("contingency") or {}).get("rate")
    contingency = None
    if rate is not None and other_complete:
        contingency = ((engineering + other) * dec(rate)).quantize(Decimal("0.01"))
    checker.near("EST-CONTINGENCY", "预备费独立复算", contingency, estimate.get("contingency"),
                 TOL_AMOUNT, "万元", "（工程费+其他费）×预备费率")

    interest_info = (payload.get("estimate") or {}).get("interest") or {}
    loan = dec(interest_info.get("loanAmount")) or Decimal("0")
    annual_rate = dec(interest_info.get("loanRate")) or Decimal("0")
    default_years = ((payload.get("estimate") or {}).get("interest") or {}).get("buildYears") or 3
    years = int(default_years)
    half = bool(interest_info.get("halfYear", True))
    interest = Decimal("0")
    explicit_interest = interest_info.get("amount")
    if explicit_interest is not None:
        # 上游可研已给定建设期融资费用（附表1）→ 直接对账，不按利率反推
        interest = dec(explicit_interest)
        note = "上游给定的建设期融资费用（estimate.interest.amount）"
    else:
        if years > 0:
            draw = loan / years
            for _ in range(years):
                # 引擎口径（财务口径.yaml）：建设期利息不计复利，只对当年提款按半年计息
                interest += (draw * annual_rate * (Decimal("0.5") if half else Decimal("1"))).quantize(Decimal("0.01"))
        note = "Σ(当年提款×年利率×{})，不计复利".format("0.5" if half else "1")
    checker.near("EST-INTEREST", "建设期利息独立复算", interest, estimate.get("constructionInterest"),
                 TOL_AMOUNT, "万元", note)

    total = engineering + (other if other_complete else Decimal("0")) + (contingency or Decimal("0")) + interest
    if not other_complete:
        checker.add("EST-TOTAL", "总投资独立复算", False, "其他费有缺数项，合计不可信（不补数）")
    else:
        checker.near("EST-TOTAL", "总投资独立复算", total, estimate.get("totalInvestment"),
                     TOL_AMOUNT, "万元", "工程费+其他费+预备费+建设期利息")

    funding = estimate.get("funding") or {}
    funding_total = sum(dec(funding.get(key) or 0) for key in ("equity", "loan", "subsidy"))
    actual_total = dec(estimate.get("totalInvestment"))
    if actual_total is None:
        checker.add("EST-FUNDING", "资金筹措与总投资平衡", False, "总投资缺")
    else:
        checker.add("EST-FUNDING", "资金筹措与总投资平衡（容差 1 万元）",
                    abs(funding_total - actual_total) <= TOL_ANCHOR_AMOUNT,
                    "筹措 {} / 总投资 {} / 差 {}".format(funding_total, actual_total, funding_total - actual_total))


# ---------------------------------------------------------------- 检查：财务

def check_finance(checker, payload, estimate, finance):
    indicators = finance.get("indicators") or {}
    period = finance.get("period") or {}
    cash = ((finance.get("tables") or {}).get("cashflows") or {})
    discount = dec(period.get("discountRate")) or Decimal("0.035")
    equity = dec((payload.get("finance") or {}).get("equity")) or Decimal("0")

    for key, label in (("allInvestment", "全部投资"), ("equity", "资本金")):
        block = cash.get(key) or {}
        rows = block.get("rows") or []
        if not rows:
            checker.add("FIN-{}".format(key), "{}现金流序列存在".format(label), False, "缺 rows")
            continue
        series = [dec(row.get("netAfterTax") or 0) for row in rows]
        my_irr = irr(series)
        my_npv = npv(series, discount)
        my_pt = payback(series)
        prefix = "all" if key == "allInvestment" else "equity"
        actual_irr = dec(indicators.get("firr" + ("All" if prefix == "all" else "Equity")))
        actual_pt = dec(indicators.get("pt" + ("All" if prefix == "all" else "Equity")))
        if my_irr is None and actual_irr is None:
            checker.add("FIN-FIRR-{}".format(prefix), "{}FIRR 独立复算".format(label), True,
                        "两边均无解（现金流不能回收），与引擎一致")
        elif my_irr is None or actual_irr is None:
            checker.add("FIN-FIRR-{}".format(prefix), "{}FIRR 独立复算".format(label), False,
                        "一边有解一边无解：实际 {} / 独立 {}".format(actual_irr, my_irr))
        else:
            checker.add("FIN-FIRR-{}".format(prefix), "{}FIRR 独立复算".format(label),
                        abs(actual_irr - my_irr) * 100 <= TOL_FIRR_PP,
                        "实际 {} / 独立 {} / 差 {}pp（容差 0.04pp）".format(
                            actual_irr, my_irr.quantize(Decimal("0.0001")),
                            (abs(actual_irr - my_irr) * 100).quantize(Decimal("0.0001"))))
        checker.near("FIN-FNPV-{}".format(prefix), "{}FNPV 独立复算".format(label),
                     indicators.get("fnpv" + ("All" if prefix == "all" else "Equity")),
                     my_npv, TOL_AMOUNT, "万元", "折现率 {}".format(discount))
        if prefix == "equity" and equity <= 0:
            # 资本金为 0 时没有出资可回收：序列 0,0,377.01… 会被 payback 机械算成 3 年。
            checker.add("FIN-PT-equity", "资本金静态回收期（资本金为 0，不适用）",
                        indicators.get("ptEquity") is None,
                        "项目资本金为 0，资本金回收期无经济含义；引擎 ptEquity={}（应为 null）".format(
                            indicators.get("ptEquity")))
        elif my_pt is None and actual_pt is None:
            checker.add("FIN-PT-{}".format(prefix), "{}静态回收期独立复算".format(label), True,
                        "两边均不可回收，与引擎一致")
        elif my_pt is None or actual_pt is None:
            checker.add("FIN-PT-{}".format(prefix), "{}静态回收期独立复算".format(label), False,
                        "一边有值一边无值：实际 {} / 独立 {}".format(actual_pt, my_pt))
        else:
            checker.near("FIN-PT-{}".format(prefix), "{}静态回收期独立复算".format(label),
                         actual_pt, my_pt, TOL_YEARS, "年", "累计首次转正插值")

    # FNPV 为负说明项目在计算期内没有回本，任何静态/动态回收期都不该是有限正数。
    fnpv_all = dec(indicators.get("fnpvAll"))
    if fnpv_all is not None and fnpv_all < 0:
        checker.add("FIN-PT-NPV", "FNPV<0 时不得输出有限的静态/动态投资回收期",
                    indicators.get("ptAll") is None,
                    "FNPV {} 万元 < 0，ptAll={}（应为 null/未转正）".format(
                        fnpv_all, indicators.get("ptAll")))
    cost = finance.get("cost") or {}
    taxable = [dec(value) for value in cost.get("taxableIncome") or []]
    income_tax = [dec(value) for value in cost.get("incomeTax") or []]
    operation_years = int(period.get("operationYears") or 0)
    if taxable and income_tax and operation_years > 0:
        net = [taxable[index] - income_tax[index] for index in range(min(len(taxable), len(income_tax)))]
        operation_net = net[-operation_years:]
        my_roe = (sum(operation_net) / Decimal(len(operation_net)) / equity).quantize(Decimal("0.0001")) if equity else None
        if equity:
            checker.near("FIN-ROE", "ROE 可复算（年均净利/资本金）", indicators.get("roe"), my_roe,
                         Decimal("0.0002"), "", "运营期年均净利润 ÷ 资本金")
        else:
            # 纯补助项目资本金为 0，ROE 的分母为 0：指标不适用，不是缺值。
            checker.add("FIN-ROE", "ROE 可复算（年均净利/资本金）",
                        indicators.get("roe") is None,
                        "项目资本金为 0，ROE 不适用（引擎与独立复算均为 None）")
    else:
        checker.add("FIN-ROE", "ROE 可复算（年均净利/资本金）", False, "缺 taxableIncome/incomeTax 序列")

    ebit = [dec(value) for value in cost.get("ebit") or []]
    total = dec(estimate.get("totalInvestment"))
    if ebit and operation_years > 0 and total:
        my_roi = (sum(ebit[-operation_years:]) / Decimal(len(ebit[-operation_years:])) / total).quantize(Decimal("0.0001"))
        checker.near("FIN-ROI", "ROI 可复算（年均息税前利润/总投资）", indicators.get("roi"), my_roi,
                     Decimal("0.0002"), "", "运营期年均 EBIT ÷ 总投资")
    else:
        checker.add("FIN-ROI", "ROI 可复算（年均息税前利润/总投资）", False, "缺 EBIT 序列")

    ebitda = [dec(value) for value in cost.get("ebitda") or []]
    debt_rows = ((finance.get("tables") or {}).get("debtSchedule") or {}).get("rows") or []
    build_years = int(period.get("constructionYears") or 0)
    interest_series = [dec(row.get("interest") or 0) for row in debt_rows]
    principal_series = [dec(row.get("principal") or 0) for row in debt_rows]
    if debt_rows and interest_series:
        op_ebit = sum(ebit[build_years:]) if len(ebit) > build_years else sum(ebit)
        op_interest = sum(interest_series[build_years:]) if len(interest_series) > build_years else sum(interest_series)
        debt_service = sum(principal_series) + sum(interest_series)
        if op_interest:
            my_icr = (op_ebit / op_interest).quantize(Decimal("0.0001"))
            checker.near("FIN-ICR", "ICR 可复算（运营期ΣEBIT/Σ利息）", indicators.get("icr"), my_icr,
                         Decimal("0.01"), "", "运营期汇总口径")
        else:
            # 无借款 → 利息为 0，ICR 分母为 0：不适用，不是缺值。
            checker.add("FIN-ICR", "ICR 可复算（运营期ΣEBIT/Σ利息）",
                        indicators.get("icr") is None,
                        "项目无借款（运营期利息合计 0），ICR 不适用")
        if debt_service:
            my_dscr = ((sum(ebitda) - sum(income_tax)) / debt_service).quantize(Decimal("0.0001"))
            checker.near("FIN-DSCR", "DSCR 可复算（(ΣEBITDA−Σ所得税)/Σ还本付息）", indicators.get("dscr"), my_dscr,
                         Decimal("0.01"), "", "全期汇总口径")
        else:
            checker.add("FIN-DSCR", "DSCR 可复算（(ΣEBITDA−Σ所得税)/Σ还本付息）",
                        indicators.get("dscr") is None,
                        "项目无还本付息（还本与付息合计 0），DSCR 不适用")

    tax_rate_rules = check_discount_caliber(checker, finance, discount)


def check_sensitivity(checker, finance):
    """敏感性单锚点闸门：基准必须等于指标表基准，每行 delta = 本行值 − 基准。"""
    indicators = finance.get("indicators") or {}
    sensitivity = finance.get("sensitivity") or {}
    base = sensitivity.get("base") or {}
    base_fnpv = dec(base.get("fnpv"))
    indicator_fnpv = dec(indicators.get("fnpvAll"))
    # 双锚点残留（旧字段 *FromSeries）本身就是「两个值并存」，一并判失败。
    dual = sorted(key for key in base if str(key).lower().endswith("fromseries"))
    if base_fnpv is None or indicator_fnpv is None:
        checker.add("SENS-BASE", "敏感性基准与指标表同口径（差 ≤0.01 万元）", False,
                    "缺值：sensitivity.base.fnpv={} / indicators.fnpvAll={}".format(
                        base.get("fnpv"), indicators.get("fnpvAll")))
    else:
        diff = abs(base_fnpv - indicator_fnpv)
        checker.add("SENS-BASE", "敏感性基准与指标表同口径（差 ≤0.01 万元）",
                    diff <= TOL_AMOUNT and not dual,
                    "敏感性基准 {} / 指标表 {} / 差 {} 万元；双锚点残留 {}".format(
                        base_fnpv, indicator_fnpv, diff, "、".join(dual) if dual else "无"))
    problems = []
    if base_fnpv is None:
        problems.append("缺基准 FNPV")
    else:
        for factor in sensitivity.get("factors") or []:
            for row in factor.get("rows") or []:
                value, delta = dec(row.get("fnpv")), dec(row.get("fnpvDelta"))
                if value is None or delta is None:
                    problems.append("{} {}：缺 fnpv 或 fnpvDelta".format(
                        factor.get("key"), row.get("label")))
                    continue
                expect = value - base_fnpv
                if abs(delta - expect) > TOL_AMOUNT:
                    problems.append("{} {}：delta {} ≠ 值 {} − 基准 {} = {}".format(
                        factor.get("key"), row.get("label"), delta, value, base_fnpv, expect))
    if not sensitivity.get("factors"):
        problems.append("缺 sensitivity.factors")
    checker.add("SENS-DELTA", "敏感性每行 delta = 本行 FNPV − 基准 FNPV（差 ≤0.01 万元）",
                not problems, "；".join(problems[:4]) if problems else "全部自洽")


BASIS_UNIT_RE = re.compile(
    r"([\d.]+)\s*(万)?\s*(kWh|千瓦时|平方米|㎡|m2|个|车位|灯箱|台|套)\s*[×xX*]\s*"
    r"([\d.]+)(?:\s*[—–~-]\s*([\d.]+))?\s*元\s*/\s*(kWh|千瓦时|平方米|㎡|m2|个)")
BASIS_MONTH_RE = re.compile(r"[×xX*]\s*12\s*个月")


def basis_annual(basis_text):
    """按输入表「测算依据」的「数量×单价」复算年收入（万元）；不可闭合返回 None。

    只认三种可闭合乘式：元/kWh、元/平方米·月、元/个·月|年。
    「元/小时×天数×使用率」这类需要中间假设的口径不猜，返回 None 交给占比判据兜底。
    """
    if not basis_text:
        return None
    total = Decimal("0")
    found = False
    for segment in re.split(r"[＋+]", str(basis_text)):
        match = BASIS_UNIT_RE.search(segment)
        if not match:
            continue
        quantity = Decimal(match.group(1))
        ten_thousand = bool(match.group(2))
        low = Decimal(match.group(4))
        price = (low + Decimal(match.group(5))) / 2 if match.group(5) else low
        value = quantity * price
        if not ten_thousand:
            value = value / Decimal("10000")
        if BASIS_MONTH_RE.search(segment):
            value *= Decimal("12")
        total += value
        found = True
    return total if found else None


REV_DISCLOSE_DOCX = Path("成果文件") / "成品版" / "项目投资决策建议书" / "项目投资决策建议书.docx"
REV_DISCLOSE_XLSX = Path("成果文件") / "内部件" / "待确认清单" / "待确认清单.xlsx"


def _docx_plain(path):
    try:
        import docx
        document = docx.Document(str(path))
    except Exception:
        return ""
    parts = [item.text for item in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def _xlsx_plain(path):
    try:
        import openpyxl
        workbook = openpyxl.load_workbook(str(path), data_only=True)
    except Exception:
        return ""
    parts = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(values_only=True):
            parts.extend(str(value) for value in row if value not in (None, ""))
    return "\n".join(parts)


def _has_number(text, value):
    """文本里是否有与 value 相等的数值（千分位写法也认）。"""
    target = dec(value)
    if target is None:
        return False
    for raw in re.findall(r"\d[\d,]*(?:\.\d+)?", text or ""):
        try:
            if Decimal(raw.replace(",", "")) == target:
                return True
        except Exception:
            continue
    return False


def revenue_disclosure(workspace, item, share):
    """单位疑误项的两处披露核对（成品脚注 + 内部件《待确认清单》）。

    两处同时齐备才允许 REV-MAG 带警告通过：缺一处即按失败处理，
    避免「只在内部件记一笔、正文照旧静默改数」。
    """
    name = str(item.get("name") or item.get("key") or "")
    key = str(item.get("key") or "")
    suspect = item.get("unitSuspect") or {}
    lacking = []
    body = _docx_plain(Path(workspace) / REV_DISCLOSE_DOCX)
    if not body:
        lacking.append("成品 docx 读不到")
    else:
        if not name or name not in body:
            lacking.append("正文缺收入项名「{}」".format(name))
        if "按元计列" not in body:
            lacking.append("正文缺「按元计列」口径")
        if "核实" not in body:
            lacking.append("正文缺核实说明")
        if not _has_number(body, item.get("annualAmount")):
            lacking.append("正文缺按元计列后的值")
    pending = _xlsx_plain(Path(workspace) / REV_DISCLOSE_XLSX)
    if not pending:
        lacking.append("内部件《待确认清单》读不到")
    else:
        if key and key not in pending:
            lacking.append("待确认清单缺项目编号 {}".format(key))
        if not _has_number(pending, suspect.get("originalAmount")):
            lacking.append("待确认清单缺原值")
        if "元填成万元" not in pending:
            lacking.append("待确认清单缺单位疑误判定")
        if not _has_number(pending, item.get("annualAmount")):
            lacking.append("待确认清单缺按元计列后的值")
        if share is not None and "{:.2f}%".format(share * 100) not in pending:
            lacking.append("待确认清单缺占比 {}".format("{:.2f}%".format(share * 100)))
        if "核实" not in pending:
            lacking.append("待确认清单缺核实要求")
    return (not lacking), "；".join(lacking)


def check_revenue_magnitude(checker, workspace, payload, finance):
    """REV-MAG 收入量级闸门：量价复算差 10 倍、或占比异常即拦。

    8.13 万元/年的充电桩收入连过五轮：01 只把输入表的年均收入搬进 JSON，03 见
    annualAmount 就不再走量价公式，07 只核对「正文数值能否在来源里找到同一个值」——
    错值自身自洽，全链没有一处复算量级。这里补两道独立判据：
      1) 测算依据可复算时，申报年收入 / 复算年收入 必须落在 [0.1, 10]；
      2) 无法复算时，该收入项占运营期营业收入 <2% 或 >90% 视为量级/结构异常。

    判据不变；唯一例外是「提资单位疑误」项（01 已在财务输入标注 unitSuspect）：
    此时按「必须披露的警告」放行，但成品脚注与内部件《待确认清单》两处必须同时写明
    原值、单位疑误判定、按元计列值与占比、待核实要求，缺一处仍判失败。
    """
    revenue = finance.get("revenue") or {}
    build = int((finance.get("period") or {}).get("constructionYears") or 0)
    totals = [dec(value) or Decimal("0") for value in revenue.get("total") or []]
    operation_total = sum(totals[build:], Decimal("0"))
    if operation_total <= 0:
        checker.add("REV-MAG", "收入量级闸门（量价复算差 10 倍 / 占比 <2% 或 >90%）",
                    False, "运营期营业收入合计为 0，无法判断量级")
        return
    declared_by_key = {str(item.get("key")): item for item in (payload.get("revenue") or [])}
    problems = []
    warnings = []
    for entry in revenue.get("items") or []:
        key = str(entry.get("key"))
        name = entry.get("name") or key
        yearly = [dec(value) or Decimal("0") for value in entry.get("yearly") or []]
        item_total = sum(yearly[build:], Decimal("0"))
        share = item_total / operation_total
        source = declared_by_key.get(key) or {}
        declared = dec(source.get("annualAmount"))
        expected = basis_annual(source.get("basisText"))
        if expected is not None and expected > 0 and declared is not None:
            ratio = declared / expected
            if ratio > Decimal("10") or ratio < Decimal("0.1"):
                times = ratio if ratio >= 1 else Decimal("1") / ratio
                problems.append("{} {}：申报 {:.2f} 万元/年，量价复算 {:.2f} 万元/年，相差 {:.1f} 倍".format(
                    key, name, declared, expected, times))
        elif share < Decimal("0.02") or share > Decimal("0.90"):
            message = "{} {}：无可用量价依据，占营业收入 {:.2f}%（须在 2%—90% 之间）".format(
                key, name, share * 100)
            if source.get("unitSuspect"):
                disclosed, lacking = revenue_disclosure(workspace, source, share)
                if disclosed:
                    warnings.append("{}；提资单位疑误（元填成万元），已在成品脚注与"
                                    "内部件《待确认清单》披露原值、按元计列值与占比".format(message))
                else:
                    problems.append("{}；单位疑误未按要求披露：{}".format(message, lacking))
            else:
                problems.append(message)
    if problems:
        checker.add("REV-MAG", "收入量级闸门（量价复算差 10 倍 / 占比 <2% 或 >90%）",
                    False, "；".join(problems[:3]))
    elif warnings:
        checker.add("REV-MAG", "收入量级闸门（量价复算差 10 倍 / 占比 <2% 或 >90%）",
                    True, "；".join(warnings), warn=True)
    else:
        checker.add("REV-MAG", "收入量级闸门（量价复算差 10 倍 / 占比 <2% 或 >90%）",
                    True, "各项量级与量价依据自洽")


def check_reference_consistency(checker, workspace, finance):
    """引擎自算值 vs 可研对照值：口径一致性检查。

    本报告指标一律取引擎自算值；可研对照值只在成品中单列为对照表（表 8-8）。
    此处校验两件事：
      1) 引擎值与可研值差异超限时，必须已在成品中显式说明并登记缺项（否则读者无从判断口径）；
      2) 成品中不得残留「同一指标两套数」——即正文/表格出现的指标必须是引擎值。
    """
    payload = read_json(workspace / "过程文件/01-上游取料/财务输入.json") or {}
    reference = (payload.get("finance") or {}).get("referenceIndicators") or {}
    reference = {key: value for key, value in reference.items() if value is not None}
    indicators = finance.get("indicators") or {}
    if not reference:
        checker.add("REF-CONSISTENT", "本报告指标口径统一（无可研对照值，全部取引擎自算值）", True,
                    "引擎值直接用于正文与表格")
        return

    diffs = []
    for key, label, unit, limit in (
            ("firrAll", "全部投资FIRR", "%", Decimal("0.005")),
            ("fnpvAll", "全部投资FNPV", "万元", Decimal("0.10")),
            ("ptAll", "全部投资回收期", "年", Decimal("0.10")),
            ("firrEquity", "资本金FIRR", "%", Decimal("0.005")),
            ("fnpvEquity", "资本金FNPV", "万元", Decimal("0.10")),
            ("ptEquity", "资本金回收期", "年", Decimal("0.10"))):
        mine, theirs = dec(indicators.get(key)), dec(reference.get(key))
        if mine is None or theirs is None:
            continue
        base = abs(theirs) if abs(theirs) > Decimal("0.000001") else Decimal("1")
        rel = abs(mine - theirs) / base
        if rel > limit:
            diffs.append("{}：引擎 {} vs 可研 {}".format(label, mine, theirs))
    if diffs:
        # 差异必须已在成品里显式说明（对照表 + 口径提示），否则算失败
        docx_path = workspace / "成果文件" / "项目投资决策建议书.docx"
        disclosed = False
        if docx_path.is_file():
            try:
                import docx
                document = docx.Document(str(docx_path))
                text = "\n".join([item.text for item in document.paragraphs] +
                                 [" ".join(cell.text for cell in row.cells)
                                  for table in document.tables for row in table.rows])
                disclosed = ("可研" in text and "对照" in text)
            except Exception:
                disclosed = False
        checker.add("REF-CONSISTENT", "引擎值与可研对照值差异已在成品中显式说明并登记缺项",
                    disclosed,
                    "差异：{}；成品已披露对照口径：{}".format("；".join(diffs), disclosed))
    else:
        checker.add("REF-CONSISTENT", "引擎值与可研对照值一致（口径无冲突）", True,
                    "共比对 {} 项指标".format(len(reference)))


def check_discount_caliber(checker, finance, discount):
    rows = (((finance.get("tables") or {}).get("cashflows") or {}).get("allInvestment") or {}).get("rows") or []
    ok = True
    for index, row in enumerate(rows[:3], start=1):
        expected = (Decimal("1") / (Decimal("1") + discount) ** index).quantize(Decimal("0.0001"))
        actual = fnum(row.get("discountFactor"))
        if actual is None or abs(actual - expected) > Decimal("0.0001"):
            ok = False
    checker.add("FIN-DISCOUNT-FACTOR", "折现系数 = 1/(1+r)^t（前 3 年抽检）", ok,
                "r={}".format(discount))
    return ok


# ---------------------------------------------------------------- 检查：方案比选

def check_schemes(checker, schemes):
    table = schemes.get("compareTable") or {}
    rows = table.get("rows") or []
    check_keys = {"investment", "annualRevenue", "annualCost", "firr", "fnpv", "pt", "score"}
    columns = {item.get("key") for item in table.get("columns") or []}
    checker.add("SCH-3", "三方案同口径（3 套且编号唯一）",
                len(rows) == 3 and len({row.get("id") for row in rows}) == 3,
                "方案：" + "、".join(str(row.get("id")) for row in rows))
    checker.add("SCH-DIMS", "7 维度齐全（投资/收入/成本/IRR/NPV/回收期/得分）",
                check_keys <= columns, "缺：" + "、".join(sorted(check_keys - columns)))
    ranks = sorted(row.get("scoreRank") for row in rows if row.get("scoreRank") is not None)
    checker.add("SCH-RANK", "得分排名连续 1..3", ranks == [1, 2, 3], str(ranks))
    winner = (table.get("recommendation") or {}).get("winner")
    top = next((row.get("id") for row in rows if row.get("scoreRank") == 1), None)
    checker.add("SCH-WINNER", "推荐结论与排名第 1 一致", bool(winner) and winner == top,
                "推荐 {} / 排名第 1 {}".format(winner, top))
    firrs = [fnum(row.get("firr")) for row in rows if fnum(row.get("firr")) is not None]
    checker.add("SCH-FIRR-RANGE", "方案 FIRR 在 0–60% 合理区间（防口径错）",
                all(Decimal("-0.05") <= item <= Decimal("0.6") for item in firrs), str(firrs))
    checker.add("SCH-CALIBER", "注明方案指标为参数缩放口径",
                "参数缩放" in json.dumps(table, ensure_ascii=False), "compareTable.caliberNote/note")


# ---------------------------------------------------------------- 检查：AHP 不重算

def read_ahp_office_weights(workspace):
    """独立读用户拷入的《AHP权重表.xlsx》「导出值」表；读不到返回 None。"""
    input_root = workspace / "输入文件"
    candidates = sorted(input_root.rglob("*AHP权重表*.xlsx")) if input_root.is_dir() else []
    if not candidates:
        candidates = sorted(input_root.rglob("*权重表*.xlsx")) if input_root.is_dir() else []
    if not candidates:
        return None
    path = candidates[0]
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None
    try:
        sheet = workbook["导出值"] if "导出值" in workbook.sheetnames else workbook.worksheets[-1]
        values = {}
        for row in sheet.iter_rows(values_only=True):
            if not row or len(row) < 2 or row[0] is None:
                continue
            key = str(row[0]).strip()
            try:
                values[key] = Decimal(str(row[1]))
            except Exception:
                continue
    finally:
        workbook.close()
    criteria = {key.split(".", 1)[1]: value for key, value in values.items() if key.startswith("准则权重.")}
    indicators = {key.split(".", 1)[1]: value for key, value in values.items() if key.startswith("指标综合.")}
    return {"criteria": criteria, "indicators": indicators} if criteria and indicators else None


def check_ahp(checker, workspace, payload):
    weights = (payload.get("ahpWeights") or {}).get("criteria") or []
    total = sum(dec(item.get("weight") or 0) for item in weights)
    checker.add("AHP-SUM", "AHP 准则权重合计≈1（只读不重算）", abs(total - 1) <= Decimal("0.001"),
                "合计 {} / {} 项".format(total, len(weights)))

    if (payload.get("ahpWeights") or {}).get("temporary"):
        checker.add("AHP-NO-RECALC", "AHP 未重算（缺上游时用显式默认权重）", True,
                    "缺上游场景：default 权重已标 temporary，不涉及 AHP 对账")
        return

    office = read_ahp_office_weights(workspace)
    indicators = (payload.get("ahpWeights") or {}).get("indicatorWeights") or []
    if office:
        expected_criteria = list(office["criteria"].values())
        actual_criteria = [dec(item.get("weight")) for item in weights]
        # 引擎侧权重保留 4 位小数，上游 xlsx 为全精度：按 0.001 容差比对，
        # 不能要求逐位相等（否则 0.1887 与 0.188690742525 会被判「重算」）。
        criteria_ok = (len(expected_criteria) == len(actual_criteria) and
                       all(abs(actual - expected) <= Decimal("0.001")
                           for actual, expected in zip(actual_criteria, expected_criteria)))
        expected_indicators = office["indicators"]
        actual_indicators = {item.get("indicatorId"): dec(item.get("compositeWeight"))
                             for item in indicators if item.get("indicatorId")}
        # 成品版权重保留 6 位小数、溯源版是全精度：同样按 0.001 容差比对，
        # 不能因 0.063447 与 0.063446641857 差 3.6e-7 就判「重算」。
        mismatched = [key for key, value in expected_indicators.items()
                      if key not in actual_indicators
                      or abs(actual_indicators[key] - value) > Decimal("0.001")]
        checker.add("AHP-NO-RECALC", "权重与用户拷入的 AHP xlsx 逐项一致（未重算 AHP）",
                    criteria_ok and not mismatched,
                    "准则 {} 项{}；指标 {} 项，不一致 {}".format(
                        len(actual_criteria), "一致" if criteria_ok else "不一致（{}≠{}）".format(
                            actual_criteria, expected_criteria),
                        len(expected_indicators), mismatched[:4]))
        return

    staging = read_json(workspace / "过程文件/00-总控/引擎项目/05_过程/AHP/权重表.json") or {}
    expected = {item.get("key"): dec(item.get("weight")) for item in staging.get("criteria") or []}
    if not expected:
        checker.add("AHP-NO-RECALC", "权重与上游来源逐项一致（未重算 AHP）", False,
                    "既无 AHP xlsx 也无引擎侧权重表，无法对账")
        return
    mismatched = [item.get("key") for item in weights
                  if item.get("key") in expected and dec(item.get("weight")) != expected[item.get("key")]]
    checker.add("AHP-NO-RECALC", "权重与上游来源逐项一致（未重算 AHP）", not mismatched,
                "{}；{} 项".format("不一致 " + str(mismatched[:4]) if mismatched else "全部一致", len(expected)))


def check_ahp_in_report(checker, workspace):
    """AHP 成果是否真的进了第 9 章：16 项指标名 + 3 个方案合成得分都要在分章稿里出现。"""
    staging_scores = workspace / "过程文件/00-总控/引擎项目/05_过程/AHP/评分表.json"
    if not staging_scores.is_file():
        checker.add("AHP-IN-REPORT", "第 9 章纳入 AHP 成果", True,
                    "缺上游场景：跳过（评分段按缺项占位）")
        return
    path = workspace / "过程文件/05-章节成文/分章/CH-09.md"
    if not path.is_file():
        checker.add("AHP-IN-REPORT", "第 9 章纳入 AHP 成果", False, "缺 CH-09.md")
        return
    text = path.read_text(encoding="utf-8")
    upstream_json = workspace / ((read_json(workspace / "工作区参数.json") or {}).get("paths", {})
                                 .get("upstreamJson") or "过程文件/00-总控/上游机读件")
    weights = (read_json(workspace / "过程文件/00-总控/引擎项目/05_过程/AHP/权重表.json")
               or read_json(upstream_json / "权重表.json") or {})
    scores = read_json(staging_scores) or read_json(upstream_json / "评分表.json") or {}
    indicators = [item.get("indicator") for item in weights.get("indicatorWeights") or []]
    missing_indicators = [name for name in indicators if name and name not in text]
    composites = [str(item.get("composite")) for item in scores.get("schemes") or []
                  if item.get("composite") is not None]
    # 正文按两位小数呈现（7.26/6.09/5.11），不能要求出现 17 位原始小数：
    # 只要两位小数形态在正文里能找到即视为纳入。
    missing_scores = [value for value in composites
                      if ("{:.2f}".format(float(dec(value))) not in text) and (value not in text)]
    checker.add("AHP-IN-REPORT", "第 9 章纳入 AHP 指标与三方案得分",
                not missing_indicators and not missing_scores,
                "{} 项指标全部入章：{}；方案得分 {}：缺 {}".format(
                    len(indicators), not missing_indicators, composites, missing_scores))


# ---------------------------------------------------------------- 检查：docx 回读

def check_docx(checker, workspace, finance, estimate):
    try:
        import docx
    except ImportError:
        checker.add("DOC-LIB", "python-docx 可用", False, "pip install python-docx")
        return None
    path = workspace / "成果文件/成品版/项目投资决策建议书/项目投资决策建议书.docx"
    if not path.is_file():           # 兼容旧的扁平产出
        path = workspace / "成果文件/项目投资决策建议书.docx"
    if not path.is_file():
        checker.add("DOC-EXISTS", "成品 docx 存在", False,
                    "成果文件/成品版/项目投资决策建议书/项目投资决策建议书.docx")
        return None
    checker.add("DOC-EXISTS", "成品 docx 存在", True, rel(path, workspace))
    document = docx.Document(str(path))
    headings = [paragraph.text.strip() for paragraph in document.paragraphs
                if paragraph.style.name == "Heading 1" and paragraph.text.strip()]
    def _norm_title(text):
        cleaned = re.sub(r"[\s（）()、，,\-—]", "", text or "")
        return re.sub(r"^(?:\d+|[一二三四五六七八九十]+)(?=[^\d一二三四五六七八九十])", "", cleaned)

    chapters_ok = (len(headings) == len(CHAPTER_TITLES)
                   and all(_norm_title(a) == _norm_title(b) for a, b in zip(headings, CHAPTER_TITLES)))
    checker.add("DOC-CHAPTERS", "成品一级标题 = 10 章且顺序与契约一致", chapters_ok,
                "实测 {} 章：{}".format(len(headings), headings[:3]))
    slot_all = next((table for table in document.tables
                     if len(table.rows[0].cells) >= 4
                     and table.rows[0].cells[3].text.strip() == "计算结果（全部投资）"), None)
    slot_equity = next((table for table in document.tables
                        if len(table.rows[0].cells) >= 4
                        and table.rows[0].cells[3].text.strip() == "计算结果（资本金）"), None)
    checker.add("DOC-TABLES", "成品表位 1–6 齐全（含 20+ 张分析表）",
                len(document.tables) >= 20 and slot_all is not None and slot_equity is not None,
                "实测 {} 张；表位1 {}；表位2 {}".format(
                    len(document.tables), "有" if slot_all is not None else "缺",
                    "有" if slot_equity is not None else "缺"))
    checker.add("DOC-FIGURES", "成品必备图 ≥8 张", len(document.inline_shapes) >= 8,
                "实测 {} 张".format(len(document.inline_shapes)))

    texts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    joined = "\n".join(texts)
    residue = [pattern for pattern in PLACEHOLDER_PATTERNS if pattern in joined]
    checker.add("DOC-PLACEHOLDER", "占位符/提示行 0 残留", not residue, "命中：{}".format(residue))
    checker.add("DOC-PATH", "成品无盘符/绝对路径", not DRIVE_PATTERN.search(joined),
                "；".join(DRIVE_PATTERN.findall(joined)[:3]))
    config = read_json(workspace / "工作区参数.json") or {}
    expected_name = str(config.get("projectName") or "").strip()
    if expected_name:
        found = expected_name in joined
        ok = found
        detail = "期望项目名「{}」{}".format(expected_name, "已出现" if found else "未出现")
    else:
        ok = True
        detail = "未配置 projectName：跳过项目名检查"
    checker.add("DOC-NAME", "成品项目名与工作区参数一致", ok, detail)

    # 第十二类挑刺「跨包表述串用」＋单件独立可交付：建议书可能单独发业主，
    # 正文不得介绍《现状分析报告》的交付构成，也不得出现内部件名称与流水线口径。
    body_texts = [paragraph.text for paragraph in document.paragraphs]
    body_texts += [cell.text for table in document.tables
                   for row in table.rows for cell in row.cells]
    cross_terms = ("内部件", "数据缺口清单", "输入材料补齐确认单", "标注版", "待确认清单")
    cross_hits = ["{}｜{}".format(term, text.strip()[:40])
                  for text in body_texts if text.strip()
                  for term in cross_terms if term in text]
    checker.add("DOC-CROSSPKG", "正文无内部件名称/他件交付构成（单件独立交付）",
                not cross_hits,
                "命中 {} 处：{}".format(len(cross_hits), "；".join(cross_hits[:5]))
                if cross_hits else "0 命中")
    upstream_hits = [text.strip()[:60] for text in body_texts if "现状分析报告" in text]
    checker.add("DOC-UPSTREAM-MENTION", "正文《现状分析报告》仅资料依据处一次性交代",
                len(upstream_hits) <= 1,
                "命中 {} 处：{}".format(len(upstream_hits), "；".join(upstream_hits[:3]))
                if upstream_hits else "0 命中")
    push_hits = [text.strip()[:60] for text in body_texts
                 if re.search(r"(详见|参阅|参见)\s*[《「][^》」]{1,30}[》」]", text)]
    checker.add("DOC-PUSHREF", "无「详见/参阅/参见＋他件名称」推诿句式", not push_hits,
                "命中 {} 处：{}".format(len(push_hits), "；".join(push_hits[:3]))
                if push_hits else "0 命中")

    own_corpus = load_foreign_whitelist(workspace)
    if scan_foreign_names is None:
        checker.add("DOC-FOREIGN", "成品不含外来专名", False, "扫描器加载失败（07-回读自检 缺失）")
    else:
        foreign = scan_foreign_names(joined, own_corpus, [expected_name])
        checker.add("DOC-FOREIGN", "成品不含外来专名（地名/机构名须能在本项目材料中找到）",
                    not foreign, "疑似外来专名：" + "、".join(foreign[:8]) if foreign else "0 命中")

    # C7 形式参考零串入：拿 01 登记的参考件专名/数值指纹扫成品全文（docx + 两张成品表）
    fingerprint_path = workspace / "过程文件" / "01-上游取料" / "形式参考指纹.json"
    if fingerprint_path.is_file():
        fingerprint = read_json(fingerprint_path) or {}
        nouns = [str(item) for item in (fingerprint.get("properNouns") or [])]
        numbers = [str(item) for item in (fingerprint.get("numbers") or [])]
        blob = joined
        try:
            import openpyxl
            for relative in ("成果文件/成品版/匡算表/匡算表.xlsx",
                             "成果文件/成品版/财务评价表/财务评价表.xlsx"):
                book_path = workspace / relative
                if not book_path.is_file():
                    continue
                book = openpyxl.load_workbook(book_path, data_only=True)
                for sheet in book.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        for value in row:
                            if value is not None:
                                blob += "\n" + str(value)
                book.close()
        except Exception:  # noqa: BLE001 - xlsx 读不了由 07 的成品表回读兜底
            pass
        # 形式参考件常与本项目同片区，「××街道」这类上级地名两边都会用到；
        # 本项目材料（上游摘要/财务输入/台账）里已有的专名属于合法口径，不计串入。
        hit_nouns = [item for item in nouns if item in blob and item not in own_corpus]
        hit_numbers = [item for item in numbers if item in blob]
        checker.add("REF-LEAK", "形式参考零串入（专名/数值指纹全文扫描）",
                    not (hit_nouns or hit_numbers),
                    "专名命中 {} 个（{}）；数值命中 {} 个（{}）".format(
                        len(hit_nouns), "、".join(hit_nouns[:5]) or "—",
                        len(hit_numbers), "、".join(hit_numbers[:5]) or "—")
                    if (hit_nouns or hit_numbers) else "0 命中（专名 {} 个、数值 {} 个）".format(
                        len(nouns), len(numbers)))
    else:
        checker.add("REF-LEAK", "形式参考零串入（专名/数值指纹全文扫描）", True,
                    "未投放形式参考件，跳过（01 无 形式参考指纹.json）")

    if slot_all is not None and slot_equity is not None:
        engine_indicators = finance.get("indicators") or {}
        # 资本金为 0 时，资本金口径指标没有经济含义，成品写「不适用」是对的；
        # 但必须核到资金结构里资本金确为 0，不能拿「不适用」掩盖真实数值。
        _inputs = read_json(workspace / "过程文件/01-上游取料/财务输入.json") or {}
        _fin = (_inputs.get("finance") or {})
        _funding = _fin if "equity" in _fin else (_inputs.get("funding") or {})
        try:
            zero_equity = ("equity" in _funding
                           and float(_funding.get("equity") or 0) <= 0)
        except (TypeError, ValueError):
            zero_equity = False
        # 口径纪律：表位 1/2 与正文、敏感性分析同为引擎自算值，对账基准取引擎值。
        # 可研对照值不参与（否则等于拿可研值对可研值，篡改引擎值也照样通过）。
        indicators = dict(engine_indicators)
        mapping = [
            ("DOC-NUM-ALL-FIRR", "表位 1 全部投资 FIRR", slot_all, 1, indicators.get("firrAll"), True, "pp"),
            ("DOC-NUM-ALL-FNPV", "表位 1 全部投资 FNPV", slot_all, 2, indicators.get("fnpvAll"), False, "万元"),
            ("DOC-NUM-ALL-PT", "表位 1 全部投资回收期", slot_all, 3, indicators.get("ptAll"), False, "年"),
            ("DOC-NUM-EQ-FIRR", "表位 2 资本金 FIRR", slot_equity, 1, indicators.get("firrEquity"), True, "pp"),
            ("DOC-NUM-EQ-FNPV", "表位 2 资本金 FNPV", slot_equity, 2, indicators.get("fnpvEquity"), False, "万元"),
            ("DOC-NUM-EQ-PT", "表位 2 资本金回收期", slot_equity, 3, indicators.get("ptEquity"), False, "年"),
        ]
        for code, name, table, row_index, expected, is_percent, unit in mapping:
            if table is None:
                checker.add(code, name, False, "未按表头定位到表位表")
                continue
            if row_index >= len(table.rows):
                checker.add(code, name, False, "表行缺失")
                continue
            cells = table.rows[row_index].cells
            cell_text = cells[3].text.strip() if len(cells) > 3 else ""
            shown = fnum(cell_text) if len(cells) > 3 else None
            if expected is None or (zero_equity and code.startswith("DOC-NUM-EQ-")):
                # 算不出的指标允许写定性结论（不适用／未转正），但不得夹带数值冒充结果
                qualitative = any(word in cell_text for word in ("不适用", "未转正", "不可回收"))
                checker.add(code, name, cell_text in ("—", "-", "") or qualitative,
                            ("资本金为 0，表为 {}".format(cell_text)
                             if zero_equity and expected is not None
                             else "JSON 无值（算不出），表为 {}".format(cell_text)))
                continue
            expected_dec = dec(expected)
            if shown is None or expected_dec is None:
                checker.add(code, name, False, "缺值：docx={} json={}".format(shown, expected))
                continue
            # 比率换算与 07 用同一口径（按单位/标记判断，不靠数值大小猜）——
            # 早先这里是「无条件 ×100」，07 是「≤1 才 ×100」，两者互斥，
            # FIRR=150% 时必有一个误判。
            target = _lib.ratio_to_percent(expected_dec, "%" if is_percent else "", is_percent)
            target = target.quantize(Decimal("0.01"))
            tolerance = TOL_FIRR_PP if is_percent else (TOL_YEARS if unit == "年" else TOL_AMOUNT)
            checker.add(code, name, abs(shown - target) <= tolerance,
                        "docx {} / json {} / 容差 {}{}".format(shown, target, tolerance, unit))
    return document


# ---------------------------------------------------------------- 检查：分章与来源

def check_chapters(checker, workspace):
    directory = workspace / "过程文件/05-章节成文/分章"
    files = sorted(directory.glob("CH-*.md")) if directory.is_dir() else []
    expected = ["CH-{:02d}.md".format(index) for index in range(1, 11)]
    checker.add("CH-FILES", "分章稿 10 份且章序齐全", [item.name for item in files] == expected,
                "实测 {} 份".format(len(files)))
    source_map = read_json(directory / "来源映射.json") or {}
    entries = source_map.get("chapters") or []
    checker.add("CH-SOURCE", "每章都有来源行（来源映射 ≥10 条）", len(entries) >= 10, "{} 条".format(len(entries)))
    without = [item.get("chapter") for item in entries if not (item.get("sources") or "").strip()]
    checker.add("CH-SOURCE-EMPTY", "来源行为空 = 0", not without, str(without))
    missing = 0
    for item in files:
        missing += item.read_text(encoding="utf-8").count("（资料缺失")
    (workspace / "过程文件/08-验证与边界").mkdir(parents=True, exist_ok=True)
    checker.add("CH-GAP-MARK", "缺失处显式标注（占位计数登记）", True, "显式缺项 {} 处".format(missing))
    return missing


# ---------------------------------------------------------------- 检查：基线

def baseline_from(estimate, finance, schemes):
    indicators = finance.get("indicators") or {}
    table = schemes.get("compareTable") or {}
    return {
        "schema": SCHEMA_BASELINE,
        "note": "本工作区基线（02/03/04 产物）。删掉本文件即只做结构自洽验证；改口径后重跑会自动更新前须删文件。",
        "totalInvestment": estimate.get("totalInvestment"),
        "funding": estimate.get("funding"),
        "indicators": {key: indicators.get(key) for key in (
            "firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity", "ptEquity", "roe", "roi", "icr", "dscr")},
        "schemes": [{"id": row.get("id"), "score": row.get("score"), "scoreRank": row.get("scoreRank")}
                    for row in table.get("rows") or []],
        "winner": (table.get("recommendation") or {}).get("winner"),
        "docx": {"chapters": 10, "tables": "按表头定位表位 1–6（≥20 张）", "figures": "≥6 张"},
    }


def check_baseline(checker, workspace, estimate, finance, schemes, mode):
    # 基线按包命名空间隔离：四包可能同机共存，文件名带包名 + 内容带 package，
    # 避免 A 包的基线被 B 包当成自己的（实测：Status 基线被 Invest 读走，本步 8 项失败）。
    path = workspace / "记忆与规则/验证基线-CardoInvest.json"
    current = baseline_from(estimate, finance, schemes)
    current["package"] = PACKAGE
    if mode == "none":
        checker.add("BASE-SKIP", "基线对账", True, "mode=none（跳过）")
        return {"mode": "none", "status": "skipped"}
    if mode == "write" or (mode == "auto" and not path.is_file()):
        path.parent.mkdir(parents=True, exist_ok=True)
        write_json(path, current)
        checker.add("BASE-WRITE", "基线已写入（首次/重写）", True, rel(path, workspace))
        return {"mode": "write", "status": "written", "file": rel(path, workspace)}
    if not path.is_file():
        checker.add("BASE-MISSING", "基线文件存在", False,
                    "记忆与规则/验证基线-CardoInvest.json 缺失（--baseline write 生成）")
        return {"mode": mode, "status": "missing"}

    saved = read_json(path) or {}
    if saved.get("package") != PACKAGE:
        checker.add("BASE-SKIP", "基线对账", True,
                    "非本包基线（package={}），按首跑处理".format(saved.get("package")))
        return {"mode": mode, "status": "skipped", "reason": "foreign-baseline"}
    checker.near("BASE-TOTAL", "基线对账：总投资", estimate.get("totalInvestment"),
                 saved.get("totalInvestment"), TOL_ANCHOR_AMOUNT, "万元", "锚点容差 1 万元")
    for key, tol, unit in (("firrAll", TOL_ANCHOR_FIRR, ""), ("fnpvAll", TOL_ANCHOR_AMOUNT, "万元"),
                           ("ptAll", TOL_ANCHOR_YEARS, "年"), ("firrEquity", TOL_ANCHOR_FIRR, ""),
                           ("fnpvEquity", TOL_ANCHOR_AMOUNT, "万元"), ("ptEquity", TOL_ANCHOR_YEARS, "年"),
                           ("roe", TOL_ANCHOR_FIRR, ""), ("roi", TOL_ANCHOR_FIRR, "")):
        actual = (finance.get("indicators") or {}).get(key)
        saved_value = (saved.get("indicators") or {}).get(key)
        if actual is None and saved_value is None:
            checker.add("BASE-" + key.upper(), "基线对账：{}".format(key), True, "两边均无值（算不出）")
        elif actual is None or saved_value is None:
            checker.add("BASE-" + key.upper(), "基线对账：{}".format(key), False,
                        "一边有值一边无值：本次 {} / 基线 {}".format(actual, saved_value))
        else:
            checker.near("BASE-" + key.upper(), "基线对账：{}".format(key), actual,
                         saved_value, tol, unit, "锚点容差")
    saved_schemes = {item.get("id"): item for item in saved.get("schemes") or []}
    current_schemes = {item.get("id"): item for item in current["schemes"]}
    scheme_ok = saved_schemes.keys() == current_schemes.keys() and all(
        abs(dec(saved_schemes[key].get("score")) - dec(current_schemes[key].get("score"))) <= Decimal("0.01")
        for key in saved_schemes)
    checker.add("BASE-SCHEMES", "基线对账：三方案得分（容差 0.01）", scheme_ok,
                "保存 {} / 本次 {}".format(saved_schemes, current_schemes))
    checker.add("BASE-WINNER", "基线对账：推荐方案一致", saved.get("winner") == current.get("winner"),
                "保存 {} / 本次 {}".format(saved.get("winner"), current.get("winner")))
    return {"mode": mode, "status": "checked", "file": rel(path, workspace)}


# ---------------------------------------------------------------- 检查：边界

def scan_extensions(root, allowed, json_subdir=None):
    violations, checked = [], 0
    if not root.is_dir():
        return violations, checked
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        checked += 1
        suffix = path.suffix.lower()
        if suffix in allowed:
            continue
        if json_subdir and suffix == ".json":
            try:
                path.resolve().relative_to((root / json_subdir).resolve())
                continue
            except ValueError:
                pass
        violations.append(rel(path, root.parent))
    return violations, checked


def check_boundary(checker, workspace):
    violations, checked_in = scan_extensions(workspace / "输入文件", INPUT_ALLOWED)
    checker.add("BND-INPUT", "输入文件只放用户资料（docx/xlsx/pdf/pptx/json/zip）", not violations,
                "违规：{}".format(violations[:8]))
    violations_out, checked_out = scan_extensions(workspace / "成果文件", OUTPUT_ALLOWED)
    checker.add("BND-OUTPUT", "成果文件只放 Office（docx/xlsx/pdf）", not violations_out,
                "违规：{}".format(violations_out[:8]))

    skill_violations, checked_skill = [], 0
    for path in (workspace / "Skills").rglob("*") if (workspace / "Skills").is_dir() else []:
        if not path.is_file() or "__pycache__" in path.parts:
            continue
        checked_skill += 1
        suffix = path.suffix.lower()
        if suffix in SKILL_ALLOWED:
            continue
        if suffix == ".docx" and path.parent.name == "模板":
            continue
        if suffix == ".pyc":
            continue
        skill_violations.append(rel(path, workspace))
    checker.add("BND-SKILLS", "Skills 只放脚本/文档/规则（模板资产限 模板/ 目录）", not skill_violations,
                "违规：{}".format(skill_violations[:8]))

    # 过程文本产物不得出现盘符绝对路径
    # 豁免「运行日志.json」：它按设计记录各步命令与 stdout（stdout 含本机绝对路径），
    # 是流水线自己的执行痕迹，不是交付内容。豁免按绝对路径判定，
    # 避免用户自建的同名文件被误豁免。
    absolute = []
    log_abs = (workspace / "过程文件/00-总控/运行日志.json").resolve()
    proc = workspace / "过程文件"
    skills = workspace / "Skills"
    own_dirs = {p.name for p in skills.iterdir() if p.is_dir()} if skills.is_dir() else set()
    own_dirs.add("00-总控")
    roots = [proc / name for name in sorted(own_dirs) if (proc / name).is_dir()]
    for root in roots:
        for path in root.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in (".json", ".md", ".yaml", ".txt"):
                continue
            if path.resolve() == log_abs:
                continue
            try:
                content = path.read_text(encoding="utf-8", errors="ignore")
            except OSError:
                continue
            if DRIVE_PATTERN.search(content):
                absolute.append(rel(path, workspace))
    # 只扫本包自有步骤目录：四包可能共用一个工作区，别包的过程文件不是本包的交付物
    # （实测：同工作区下本检查会把 CardoStatus 的过程文件扫进来）。
    checker.add("BND-OWN", "只扫本包自有过程目录（不扫别包）", bool(roots),
                "扫描 {} 个目录：{}".format(len(roots), "、".join(r.name for r in roots)))
    checker.add("BND-PATH", "过程产物无盘符/绝对路径（运行日志.json 除外：执行痕迹非交付内容）",
                not absolute, "违规：{}".format(absolute[:8]))

    staging = workspace / "过程文件/00-总控/引擎项目"
    ledger_ok = True
    detail = "引擎项目在工作区内"
    try:
        staging.resolve().relative_to(workspace.resolve())
    except ValueError:
        ledger_ok = False
        detail = "引擎项目越界：{}".format(staging)
    checker.add("BND-STAGING", "引擎中间目录不越出工作区", ledger_ok, detail)
    return {"violations": violations + violations_out + skill_violations + absolute,
            "checked": {"input": checked_in, "output": checked_out, "skills": checked_skill}}


# ---------------------------------------------------------------- 指纹

def fingerprints(workspace):
    targets = [
        "成果文件/成品版/项目投资决策建议书/项目投资决策建议书.docx",
        "成果文件/成品版/项目投资决策建议书/项目投资决策建议书.pdf",
        "成果文件/内部件/投资估算/投资估算.xlsx", "成果文件/内部件/财务评价/财务评价.xlsx",
        "成果文件/内部件/方案比选/方案比选.xlsx",
        "成果文件/内部件/待确认清单/待确认清单.xlsx",
        "过程文件/01-上游取料/财务输入.json", "过程文件/02-匡算/投资估算.json",
        "过程文件/03-财务评价/财务测算.json", "过程文件/03-财务评价/计算轨迹.json",
        "过程文件/04-三方案比选/方案比选.json", "过程文件/05-章节成文/分章/CH-08.md",
        "过程文件/05-章节成文/分章/来源映射.json",
    ]
    payload = {}
    for name in targets:
        path = workspace / name
        if path.is_file():
            payload[name] = sha256_file(path)
    return payload


# ---------------------------------------------------------------- 主流程

def check_year_plan(checker, estimate):
    """分年投资计划复算：各年合计 = 建设投资 + 建设期利息，四列不得留空。

    早先 02 把建设期利息/借款提款留 null，连带 total/ratio 全空，
    分年计划四列在 xlsx 与 docx 里都是「—」，可合计明明算得出来。
    """
    def wan(value):
        try:
            return Decimal(str(value))
        except Exception:
            return None

    tables = estimate.get("tables") or {}
    rows = (tables.get("yearPlan") or {}).get("rows") or []
    problems = []
    if not rows:
        problems.append("yearPlan 无行")
    else:
        empty = [row.get("year") for row in rows
                 if any(wan(row.get(key)) is None for key in ("constructionInterest", "total", "ratio", "loanDraw"))]
        if empty:
            problems.append("第 {} 年四列有空值".format(empty[:5]))
        sums = {}
        for key in ("constructionInvestment", "constructionInterest", "total", "loanDraw"):
            values = [wan(row.get(key)) for row in rows]
            sums[key] = sum((value for value in values if value is not None), Decimal("0"))
        expect_total = (wan(estimate.get("constructionInvestment")) or Decimal("0")) + \
            (wan(estimate.get("constructionInterest")) or Decimal("0"))
        if abs(sums["total"] - expect_total) > Decimal("0.02"):
            problems.append("分年合计 {:.2f} ≠ 建设投资+利息 {:.2f}".format(sums["total"], expect_total))
        expect_build = wan(estimate.get("constructionInvestment"))
        if expect_build is not None and abs(sums["constructionInvestment"] - expect_build) > Decimal("0.02"):
            problems.append("建设投资分年合计 {:.2f} ≠ {:.2f}".format(sums["constructionInvestment"], expect_build))
        expect_loan = wan((estimate.get("funding") or {}).get("loan"))
        if expect_loan is not None and abs(sums["loanDraw"] - expect_loan) > Decimal("0.05"):
            problems.append("借款提款合计 {:.2f} ≠ 专项债 {:.2f}".format(sums["loanDraw"], expect_loan))
        ratio_sum = sum((wan(row.get("ratio")) or Decimal("0") for row in rows), Decimal("0"))
        if abs(ratio_sum - Decimal("100")) > Decimal("0.1"):
            problems.append("占比合计 {:.2f}% ≠ 100%".format(ratio_sum))
    checker.add("YEAR-PLAN", "分年投资计划四列齐全且可复算（合计=建设投资+利息）", not problems,
                "；".join(problems[:3]))


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 独立验证与边界扫描")
    parser.add_argument("--workspace", default=".")
    parser.add_argument("--baseline", choices=("auto", "write", "check", "none"), default="auto")
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve()

    payload = read_json(workspace / "过程文件/01-上游取料/财务输入.json")
    estimate = read_json(workspace / "过程文件/02-匡算/投资估算.json")
    finance = read_json(workspace / "过程文件/03-财务评价/财务测算.json")
    schemes = read_json(workspace / "过程文件/04-三方案比选/方案比选.json")
    missing = [name for name, data in (("财务输入", payload), ("投资估算", estimate),
                                       ("财务测算", finance), ("方案比选", schemes)) if not data]
    if missing:
        print("[independent_verify] 缺过程产物：{}（先跑 01→04）".format(missing))
        return 2

    checker = Checker()
    check_estimate(checker, payload, estimate)
    check_year_plan(checker, estimate)
    check_finance(checker, payload, estimate, finance)
    check_revenue_magnitude(checker, workspace, payload, finance)
    check_sensitivity(checker, finance)
    check_reference_consistency(checker, workspace, finance)
    check_schemes(checker, schemes)
    check_ahp(checker, workspace, payload)
    check_ahp_in_report(checker, workspace)
    check_docx(checker, workspace, finance, estimate)
    check_chapters(checker, workspace)
    baseline_info = check_baseline(checker, workspace, estimate, finance, schemes, args.baseline)
    boundary = check_boundary(checker, workspace)

    failed = [item for item in checker.checks if not item["ok"]]
    warned = [item for item in checker.checks if item.get("warn")]
    summary = {"total": len(checker.checks), "passed": len(checker.checks) - len(failed),
               "failed": len(failed), "warned": len(warned)}
    result = {
        "schema": SCHEMA_RESULT,
        "generatedAt": None,
        "workspace": ".",
        "summary": summary,
        "checks": checker.checks,
        "warnings": warned,
        "baseline": baseline_info,
        "boundary": boundary,
        "caliber": {
            "anchorTolerance": {"ratio": "0.5%", "amountWan": 1, "years": 0.5},
            "deliverableTolerance": {"pp": 0.04, "amountWan": 0.01, "years": 0.01},
            "antiCheat": "指标从成分现金流序列重算，不读模板结果单元格，不调用引擎函数",
        },
    }
    out_dir = workspace / "过程文件/08-验证与边界"
    write_json(out_dir / "独立验证结果.json", result)

    prints = fingerprints(workspace)
    write_json(out_dir / "产物指纹.json", {"schema": "cardo-invest/artifact-fingerprints/v1",
                                          "generatedAt": None, "sha256": prints})

    print("[independent_verify] {}/{} 通过，失败 {} 项；警告 {} 项；边界 {} 条".format(
        summary["passed"], summary["total"], summary["failed"],
        summary["warned"], len(boundary["violations"])))
    for item in failed:
        print("  ✗ {} {}：{}".format(item["id"], item["name"], item["detail"]))
    for item in warned:
        print("  ! {} {}：{}".format(item["id"], item["name"], item["detail"]))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
