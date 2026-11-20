"""三方案快速生成与比选（T12）：拆除重建 / 微更新 / 功能置换。

输入：<项目>/05_过程/投资决策/{财务输入.json,投资估算.json,财务测算.json,财务参数.json?}
口径：skill/规则/方案参数表.yaml（默认参数/默认权重）+ skill/规则/财务口径.yaml（期数/折现率兜底）
输出：<out>/方案比选.json（cardo/scheme-compare/v1）+ <out>/方案对比表.md

用法：python gen_schemes.py --project-dir <项目目录> [--params <参数.json>]
                              [--out <输出目录>] [--rules <规则目录>] [--no-timestamp]
退出码：0 生成 3 套；1 参数校验失败；2 输入缺失（无基础假设）；3 方案参数缺口待确认（产物仍落盘）。

自包含化：不依赖插件目录树；_lib.py 与 calc_metrics.py 与本脚本同目录，规则在本 skill 的 规则/ 下。
"""
from __future__ import annotations

import argparse
import json
import sys
from decimal import Decimal
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))  # 自包含化：_lib.py / calc_metrics.py 与本脚本同目录

import _lib  # noqa: E402
import calc_metrics  # noqa: E402

ENGINE_VERSION = "gen_schemes/0.4.0"
DEC = Path("05_过程") / "投资决策"
RULES_DIR_DEFAULT = SKILL_DIR / "规则"  # 自包含化：口径在本 skill 的 规则/ 下
PARAMS_TABLE_FILE = "方案参数表.yaml"
FINANCE_RULES_FILE = "财务口径.yaml"
AHP_WEIGHTS_REL = Path("05_过程") / "AHP" / "权重表.json"
TABLE_SCHEMA = "cardo/scheme-compare-table/v1"
SCORE_CRITERIA = ("投资经济", "风险可控")
RISK_BONUS = {"微更新": 1.0, "拆除重建": 0.6, "功能置换": 0.6}
SCHEME_GAP_CODE = "SCHEME-GAP"
AHP_GAP_CODE = "AHP-WEIGHTS"
CALIBER_NOTE = (
    "本产物的方案投资 / 收入 / 指标均为**参数缩放口径**"
    "（基准总投资 × 方案系数 / 基准年运营收入 × 方案系数，由简化现金流派生），"
    "与 7.1 投资估算（逐项工程量口径）**不同源**；"
    "引用时必须注明口径差异、不得互相替代（见 cardo-invest-chapter-05 SKILL.md §4 第 6 条）。"
)
TABLE_NOTE = (
    "派生块（v2/T19）：方案对比表数据，全部由本引擎已算出的方案参数与指标派生，不改原字段、不重定义口径。"
    "对比维度 = 投资 / 年运营收入 / 年运营成本 / FIRR / FNPV / 静态回收期 / 综合得分；"
    "年值为运营期年均口径；annualCost = 年运营收入 × 基准经营成本率（与简化现金流同一口径）；"
    "该比率另存 operatingCostRatioExact（六位），annualCost 用它复算，与 annualRevenue 相乘四舍五入到两位后一致；"
    "operatingCostRatio 为两位展示值，仅供人读，不可用于反算金额；"
    "指标由参数缩放后的简化现金流派生（G-invest-4 未确认，非逐项重算工程量）；金额万元、比率 %、期数年。"
    "recommendReason 把推荐理由拆成可比字段（排名/得分差/指标差/投资与成本相对水平/风险说明），供章节直接取用。"
    "比率字段（firr）为小数，与 财务测算.json#indicators 一致（×100 即 %）；firrPercent 为已换算的百分数。"
    + CALIBER_NOTE
)


class InputError(Exception):
    """输入缺失 / 坏（exit 2）。"""


class ParamError(Exception):
    """参数校验失败（exit 1）。"""


def read_json(path, label, required=True):
    if not path.is_file():
        if required:
            raise InputError("缺少{}：{}".format(label, _lib.rel_posix(path, Path.cwd())))
        return None
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("{}读取失败：{}".format(label, error))


def load_rules(rules_dir):
    """读规则目录：方案参数表（默认参数/默认权重）+ 财务口径（期数/折现率兜底）。缺任一 → exit 2。"""
    directory = Path(rules_dir)
    table_path = directory / PARAMS_TABLE_FILE
    finance_path = directory / FINANCE_RULES_FILE
    if not table_path.is_file():
        raise InputError("缺少方案参数表：{}".format(_lib.rel_posix(table_path, Path.cwd())))
    if not finance_path.is_file():
        raise InputError("缺少财务口径：{}".format(_lib.rel_posix(finance_path, Path.cwd())))
    try:
        table = _lib.load_yaml(table_path)
        finance = _lib.load_yaml(finance_path)
    except Exception as error:  # YAML 解析错误（_lib.YamlError 等）
        raise InputError("规则文件解析失败：{}".format(error))
    if not isinstance(table, dict) or not isinstance(table.get("modes"), list) or not table["modes"]:
        raise InputError("方案参数表结构非法：{}".format(_lib.rel_posix(table_path, Path.cwd())))
    if not isinstance(finance, dict) or not isinstance(finance.get("period"), dict):
        raise InputError("财务口径结构非法：{}".format(_lib.rel_posix(finance_path, Path.cwd())))
    return table, finance


def validate_params(params, table):
    """校验用户参数覆盖（任一不满足 → exit 1）。

    - params.modes（可选）：声明本次要比选的三种模式；必须恰好 3 套、互不重复、且都属固定三模式。
    - params.schemes.<S1|S2|S3>：该方案的系数 / 关键参数覆盖（可只覆盖部分）。
    """
    if params is None:
        return {}
    if not isinstance(params, dict):
        raise ParamError("参数文件根节点必须是对象")
    canonical = [str(item.get("name") or "") for item in table["modes"]]
    declared = params.get("modes")
    if declared is not None:
        if not isinstance(declared, list):
            raise ParamError("参数 modes 必须是数组")
        if len(declared) != 3:
            raise ParamError("方案模式必须恰好 3 套（{}），收到 {} 套：{}".format(
                "、".join(canonical), len(declared), declared))
        names = [str(item) for item in declared]
        if len(set(names)) != 3:
            raise ParamError("方案模式重复：{}".format(names))
        unknown = [name for name in names if name not in canonical]
        if unknown:
            raise ParamError("方案模式不在固定三模式内（{}）：{}".format("、".join(canonical), unknown))
    overrides = params.get("schemes")
    if overrides is None:
        overrides = {}
    if not isinstance(overrides, dict):
        raise ParamError("参数 schemes 必须是对象")
    valid_ids = [str(item.get("id")) for item in table["modes"]]
    unknown_ids = [str(key) for key in overrides if str(key) not in valid_ids]
    if unknown_ids:
        raise ParamError("参数 schemes 含未知方案编号（{}）：{}".format("、".join(valid_ids), unknown_ids))
    for key, value in overrides.items():
        if not isinstance(value, dict):
            raise ParamError("参数 schemes.{} 必须是对象".format(key))
        name = value.get("name")
        if name is not None and str(name) not in canonical:
            raise ParamError("参数 schemes.{}.name 必须是固定三模式之一（{}）：{}".format(
                key, "、".join(canonical), name))
    # 生效模式名集合必须恰好 3 套且互不重复（重复模式 → exit 1）
    effective = []
    for mode in table["modes"]:
        override = overrides.get(str(mode.get("id"))) or {}
        effective.append(str(override.get("name") or mode.get("name")))
    if len(set(effective)) != 3:
        raise ParamError("方案模式重复：{}".format(effective))
    return overrides


def resolve_ahp_weights(payload, project):
    """解析 AHP 权重。

    优先级：财务输入.json#ahpWeights → 05_过程/AHP/权重表.json（ahp-weights/v1）→ 默认权重。
    权重须含评分公式实际使用的两个准则（投资经济 / 风险可控），否则视为不可用 → 回落默认并标注临时。
    返回 (weights, source, temporary, detail)。权重不做归一（保持上游口径），合计记入 detail 供追溯。
    """
    candidates = []
    ahp = (payload or {}).get("ahpWeights") or {}
    if isinstance(ahp, dict) and ahp.get("criteria"):
        candidates.append(("财务输入.json#ahpWeights", ahp, bool(ahp.get("temporary"))))
    ahp_file = project / AHP_WEIGHTS_REL
    if ahp_file.is_file():
        try:
            data = _lib.read_json(ahp_file)
        except (json.JSONDecodeError, OSError):
            data = None
        if isinstance(data, dict) and data.get("criteria"):
            candidates.append((_lib.rel_posix(ahp_file, Path.cwd()), data, bool(data.get("temporary"))))

    for label, data, temporary in candidates:
        weights = {}
        for item in data.get("criteria") or []:
            key = item.get("key") or item.get("name")
            if not key:
                continue
            try:
                weights[str(key)] = _lib.dec(item.get("weight"))
            except (ValueError, TypeError):
                weights[str(key)] = Decimal("0")
        if [key for key in SCORE_CRITERIA if key not in weights]:
            continue
        total = sum(weights.values())
        if total <= 0:
            continue
        source = "ahp-temporary" if temporary else "ahp"
        return weights, source, bool(temporary), {"file": label, "criteria": list(weights.keys()),
                                                  "weightSum": _lib.fmt(total)}
    return {}, "default", True, {"file": None, "criteria": [], "weightSum": None}


def build_cashflow(investment, revenue, operating_ratio, years, build_years, discount_rate):
    """简化现金流：建设期按年摊投资，运营期按（收入 − 经营成本）计入。指标由本现金流派生。"""
    series = []
    for index in range(years):
        if index < build_years:
            series.append(-investment / build_years)
        else:
            series.append(revenue * (1 - operating_ratio))
    # **无运营收入时财务指标不可算**（政府投资项目：现金流全为投资支出）。
    # 此时返回 None 而非负值 —— 负 IRR 是"算得出来但很难看"，
    # None 才是"材料里没有这个数"，两者不能混。
    if revenue <= 0:
        return series, None, None, None
    rate = calc_metrics.irr(series)
    npv = calc_metrics.fnpv([Decimal(str(value)) for value in series], Decimal(str(discount_rate)))
    cumulative, running = [], 0.0
    for value in series:
        running += value
        cumulative.append(running)
    pt = calc_metrics.payback([Decimal(str(value)) for value in cumulative])
    return series, rate, npv, pt


def signed_amount(value):
    text = _lib.fmt(_lib.q2(value))
    return "+" + text if value > 0 else text


def q6(value):
    """比率六位（口径复算用）：annualCost = 年运营收入 × 基准经营成本率。
    q2 的比率不足以反算金额（偏差可达数个百分点），故另存高精度值供复算。"""
    return _lib.dec(value).quantize(Decimal("0.000001"), rounding="ROUND_HALF_UP")


def signed_pct_points(value):
    """比率差（小数→百分点，两位，带正负号）。"""
    points = _lib.q2(value * 100)
    text = _lib.fmt(points)
    return "+" + text if points > 0 else text


def build_recommend_reason(item, winner, base, weights_source):
    """单方案推荐理由拆字段：排名/得分差/指标差/投资与成本相对水平/风险说明。"""
    if item.get("status") != "ok" or winner is None:
        return {
            "scoreRank": item.get("scoreRank"),
            "recommended": False,
            "gapReason": item.get("gapReason"),
            "weightsSource": weights_source,
            "summary": "方案参数缺口（{}），未参与比选与排序，未借用其它方案/项目数字。".format(
                item.get("gapReason") or "缺关键参数"),
        }
    investment = float(item["investment"])
    revenue = float(item["annualRevenue"])
    cost = float(item["annualCost"])
    firr = float(item["indicators"]["firr"] or 0)
    fnpv = float(item["indicators"]["fnpv"] or 0)
    winner_investment = float(winner["investment"])
    winner_revenue = float(winner["annualRevenue"])
    winner_cost = float(winner["annualCost"])
    winner_firr = float(winner["indicators"]["firr"] or 0)
    winner_fnpv = float(winner["indicators"]["fnpv"] or 0)
    base_investment = float(base.get("investment") or 0)
    base_firr = float(base.get("firr") or 0)
    base_fnpv = float(base.get("fnpv") or 0)
    result = {
        "scoreRank": item.get("scoreRank"),
        "recommended": item.get("recommended", False),
        "scoreGapVsWinner": signed_amount(float(item["score"]) - float(winner["score"])),
        "firrGapVsWinnerPp": signed_pct_points(firr - winner_firr),
        "fnpvGapVsWinner": signed_amount(fnpv - winner_fnpv),
        "investmentGapVsWinner": signed_amount(investment - winner_investment),
        "costGapVsWinner": signed_amount(cost - winner_cost),
        "revenueGapVsWinner": signed_amount(revenue - winner_revenue),
        "investmentRatioVsBase": _lib.fmt(_lib.q2(investment / base_investment)) if base_investment else None,
        "firrGapVsBasePp": signed_pct_points(firr - base_firr),
        "fnpvGapVsBase": signed_amount(fnpv - base_fnpv),
        "riskNote": item.get("note"),
        "weightsSource": weights_source,
    }
    if item.get("recommended"):
        # 权重口径写成读者话术：正文里出现「ahp」这类引擎字段值，专家会问这是什么。
        weight_text = ("层次分析法（AHP）专家打分" if weights_source == "ahp"
                       else "临时权重（待 AHP 结果替换）")
        result["summary"] = ("综合得分 {} 排名第 {}（权重依据：{}）：投资 {} 万元（基准的 {} 倍），"
                             "年运营收入 {} 万元、年运营成本 {} 万元，FIRR {}、FNPV {} 万元、静态回收期 {} 年。").format(
            item["score"], item["scoreRank"], weight_text, item["investment"],
            result["investmentRatioVsBase"], item["annualRevenue"], item["annualCost"],
            item["indicators"]["firr"], item["indicators"]["fnpv"], item["indicators"]["pt"])
    else:
        result["summary"] = ("综合得分 {} 排名第 {}（较推荐方案得分差 {}，FIRR 差 {} 个百分点）：{}").format(
            item["score"], item["scoreRank"], result["scoreGapVsWinner"],
            result["firrGapVsWinnerPp"], item.get("note") or "")
    return result


def build_compare_table(results, winner, weights_source, base, operating_ratio):
    """方案对比表数据（v2/T19）：列定义 + 数据行 + 推荐结论（理由拆字段）。"""
    columns = [
        {"key": "id", "name": "方案编号", "unit": ""},
        {"key": "name", "name": "方案名称", "unit": ""},
        {"key": "investment", "name": "总投资", "unit": "万元"},
        {"key": "annualRevenue", "name": "年运营收入", "unit": "万元"},
        {"key": "annualCost", "name": "年运营成本", "unit": "万元"},
        {"key": "firr", "name": "财务内部收益率（FIRR）", "unit": "%", "format": "ratio-decimal（×100 即 %）"},
        {"key": "firrPercent", "name": "财务内部收益率（百分数）", "unit": "%"},
        {"key": "fnpv", "name": "财务净现值（FNPV）", "unit": "万元"},
        {"key": "pt", "name": "静态投资回收期", "unit": "年"},
        {"key": "score", "name": "综合得分", "unit": ""},
        {"key": "scoreRank", "name": "得分排名", "unit": ""},
        {"key": "recommended", "name": "推荐结论", "unit": ""},
    ]
    rows = []
    for item in results:
        firr = item["indicators"]["firr"]
        rows.append({
            "id": item["id"],
            "name": item["name"],
            "investment": item["investment"],
            "annualRevenue": item["annualRevenue"],
            "annualCost": item["annualCost"],
            "firr": firr,
            "firrPercent": (_lib.fmt(_lib.q2(_lib.dec(firr) * 100)) if firr is not None else None),
            "fnpv": item["indicators"]["fnpv"],
            "pt": item["indicators"]["pt"],
            "score": item["score"],
            "scoreRank": item.get("scoreRank"),
            "recommended": item.get("recommended", False),
            "recommendReason": item.get("recommendReason"),
        })
    ranked = sorted([item for item in results if item.get("status") == "ok"],
                    key=lambda entry: entry.get("scoreRank") or 99)
    runner_up = None
    if winner is not None:
        runner_up = next((entry for entry in ranked if entry["id"] != winner["id"]), None)
    return {
        "schema": TABLE_SCHEMA,
        "schemaVersion": 1,
        "note": TABLE_NOTE,
        "caliberNote": CALIBER_NOTE,
        "unit": "万元",
        "ratioUnit": "%",
        "columns": columns,
        "rows": rows,
        "winner": winner["id"] if winner else None,
        "winnerName": winner["name"] if winner else None,
        "weightsSource": weights_source,
        "operatingCostRatio": _lib.fmt(_lib.q2(operating_ratio)),
        "operatingCostRatioExact": _lib.fmt(q6(operating_ratio)),
        "base": base,
        "runnerUp": ({"id": runner_up["id"], "name": runner_up["name"],
                      "score": runner_up["score"],
                      "scoreGap": signed_amount(float(runner_up["score"]) - float(winner["score"]))}
                     if runner_up else None),
        "recommendation": ({
            "winner": winner["id"],
            "winnerName": winner["name"],
            "score": winner["score"],
            "scoreRank": winner.get("scoreRank"),
            "why": winner.get("recommendReason"),
        } if winner else None),
    }


def render_markdown(results, winner, weights_source, gaps):
    lines = ["# 三方案成本收益对比（T12）", "",
             "| 方案 | 总投资（万元） | 年运营收入（万元） | 年运营成本（万元） | FIRR | FNPV（万元） | 回收期（年） | 综合得分 | 排名 |",
             "|---|---|---|---|---|---|---|---|---|"]
    for item in results:
        if item.get("status") != "ok":
            lines.append("| {} {} | 缺关键参数 | — | — | — | — | — | — | — |".format(
                item["id"], item["name"]))
            continue
        lines.append("| {} {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
            item["id"], item["name"], item["investment"], item["annualRevenue"], item["annualCost"],
            item["indicators"]["firr"], item["indicators"]["fnpv"], item["indicators"]["pt"],
            item["score"], item["scoreRank"]))
    lines += ["", "> 口径提示：{}".format(CALIBER_NOTE)]
    if winner is not None:
        lines += ["", "推荐方案：**{} {}**（综合得分最高）".format(winner["id"], winner["name"]),
                  "", "> 推荐理由（拆字段见 方案比选.json#compareTable.rows[].recommendReason）：{}".format(
                      (winner.get("recommendReason") or {}).get("summary", ""))]
    else:
        lines += ["", "推荐方案：**无**（全部方案存在参数缺口，待确认）"]
    lines += ["", "> 权重来源：{}（{}）".format(
        weights_source, "临时权重，待 AHP 插件输出替换" if weights_source != "ahp" else "AHP 插件")]
    if gaps:
        lines += ["", "> 待确认事项：{}".format("；".join(
            "{}（{}）".format(item.get("what"), item.get("need")) for item in gaps))]
    return "\n".join(lines) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 三方案生成与比选（T12）")
    parser.add_argument("--project-dir", required=True)
    parser.add_argument("--params", default=None)
    parser.add_argument("--out", default=None, help="输出目录（缺省 = <项目>/05_过程/投资决策）")
    parser.add_argument("--rules", default=None, help="规则目录（缺省 = skill/规则）")
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args(argv)
    project = Path(args.project_dir)
    if not project.is_dir():
        print("[gen_schemes] 项目目录不存在：{}".format(args.project_dir))
        return 2
    dec = project / DEC
    out_dir = Path(args.out) if args.out else dec
    rules_dir = Path(args.rules) if args.rules else RULES_DIR_DEFAULT

    try:
        table, finance_rules = load_rules(rules_dir)
        payload = read_json(dec / "财务输入.json", "财务输入")
        estimate = read_json(dec / "投资估算.json", "投资估算")
        finance = read_json(dec / "财务测算.json", "财务测算")
        if args.params:
            params = read_json(Path(args.params), "参数")
        else:
            params = read_json(dec / "财务参数.json", "参数", required=False)
        overrides = validate_params(params, table)
    except ParamError as error:
        print("[gen_schemes] 参数校验失败：{}".format(error))
        return 1
    except InputError as error:
        print("[gen_schemes] {}".format(error))
        return 2

    if not finance.get("indicators"):
        print("[gen_schemes] 财务测算.json 缺 indicators（先跑 T7 calc_metrics.py）")
        return 2

    base_investment = float(estimate.get("totalInvestment") or 0)
    if base_investment <= 0:
        print("[gen_schemes] 总投资缺失，无法生成方案")
        return 2

    periods = finance.get("period") or {}
    rule_period = finance_rules.get("period") or {}
    total_years = int(periods.get("totalYears") or rule_period.get("totalYears") or 33)
    build_years = int(periods.get("constructionYears") or rule_period.get("constructionYears") or 3)
    if total_years <= build_years:
        print("[gen_schemes] 运营期年数必须为正：totalYears={} constructionYears={}".format(
            total_years, build_years))
        return 1

    # 年值口径：运营期年均（绝不把全部期数合计当年值）
    revenue_series = [float(value) for value in ((finance.get("revenue") or {}).get("total") or [])]
    cost_series = [float(value) for value in ((finance.get("cost") or {}).get("operatingCost") or [])]
    revenue_op = revenue_series[build_years:] or revenue_series or [0]
    cost_op = cost_series[build_years:] or cost_series or [0]
    revenue_base = sum(revenue_op) / len(revenue_op)
    operating_cost = sum(cost_op) / len(cost_op)
    if revenue_base <= 0:
        # 口径修正（实测踩过）：**无运营收入 ≠ 无法比选**。
        # 政府投资项目（实测某企业饮品可研：纯财政资金、无经营性收入）
        # 的三方案仍可按**投资额与实施周期**比选，只是 FIRR/FNPV/回收期算不出。
        # 原逻辑 exit 2 把整条链卡在 04，产出不了建议书。
        # 改为：继续生成方案，财务类指标如实留空并登记缺项（不编造）。
        revenue_base = 0.0
        no_revenue = True
        print("[gen_schemes] 无运营收入（政府投资项目）：按投资额与实施周期比选，"
              "财务指标（FIRR/FNPV/回收期）将标缺项")
    else:
        no_revenue = False
    operating_ratio = (operating_cost / revenue_base) if revenue_base else 0.3
    discount_rate = float((payload.get("project") or {}).get("discountRate")
                          or rule_period.get("discountRate") or 0.035)

    weights, weights_source, ahp_temporary, ahp_detail = resolve_ahp_weights(payload, project)
    if not weights:
        weights = {str(key): _lib.dec(value) for key, value in (table.get("defaultWeights") or {}).items()}

    gaps = []
    results = []
    for mode in table["modes"]:
        scheme_id = str(mode.get("id"))
        name = str(mode.get("name"))
        override = overrides.get(scheme_id) or {}
        factors = dict(mode.get("factors") or {})
        param_kind = {}
        for key in ("costFactor", "revenueFactor", "subsidyFactor", "periodYears"):
            param_kind[key] = "user" if key in override else "default"
            factors[key] = override.get(key, factors.get(key))
        try:
            cost_factor = _lib.dec(factors["costFactor"])
            revenue_factor = _lib.dec(factors["revenueFactor"])
            subsidy_factor = _lib.dec(factors["subsidyFactor"])
            period_years = int(_lib.dec(factors["periodYears"]))
        except (ValueError, TypeError, KeyError) as error:
            print("[gen_schemes] 方案参数非法（{}）：{}".format(scheme_id, error))
            return 1
        if cost_factor <= 0 or revenue_factor <= 0:
            print("[gen_schemes] 方案系数必须为正（{}）".format(scheme_id))
            return 1

        # 关键参数完整性：本模式声明了 keyParams 就必须给全（缺一 → 该方案 gap，不借用其它方案/项目数字）
        key_params = override.get("keyParams")
        required = [str(item) for item in (mode.get("requiredKeyParams") or [])]
        gap_reason = None
        key_params_out = None
        if key_params is not None:
            if not isinstance(key_params, dict):
                print("[gen_schemes] 方案参数非法（{}）：keyParams 必须是对象".format(scheme_id))
                return 1
            missing = [key for key in required if key_params.get(key) is None]
            key_params_out = {key: key_params.get(key) for key in required}
            if missing:
                gap_reason = "缺关键参数：{}".format("、".join(missing))

        params_out = {"costFactor": _lib.fmt(cost_factor), "revenueFactor": _lib.fmt(revenue_factor),
                      "subsidyFactor": _lib.fmt(subsidy_factor), "periodYears": period_years}
        if gap_reason:
            results.append({
                "id": scheme_id, "name": name, "note": mode.get("note"),
                "params": params_out, "paramKind": param_kind, "keyParams": key_params_out,
                "investment": None, "annualRevenue": None, "annualCost": None, "annualNetIncome": None,
                "indicators": {"firr": None, "fnpv": None, "pt": None},
                "status": "gap", "gapReason": gap_reason,
                "score": None, "scoreRank": None, "recommended": False,
            })
            gaps.append({"code": SCHEME_GAP_CODE, "what": gap_reason,
                         "why": "{}（{}）关键参数不全，未参与比选与排序，未借用其它方案/项目数字".format(name, scheme_id),
                         "need": "补齐该方案关键参数后重跑", "owner": "用户/项目"})
            continue

        investment = base_investment * float(cost_factor)
        revenue = revenue_base * float(revenue_factor)
        annual_cost = revenue * operating_ratio
        series, rate, npv, pt = build_cashflow(investment, revenue, operating_ratio,
                                                total_years, build_years, discount_rate)
        missing_metrics = no_revenue and rate is None
        results.append({
            "id": scheme_id, "name": name, "note": mode.get("note"),
            "params": params_out, "paramKind": param_kind, "keyParams": key_params_out,
            # 缩放口径：方案投资 = 基准投资 × costFactor。正文凡引用方案投资或方案财务指标，
            # 必须同时写明缩放系数，否则 13,074.65 这类数会被误读为独立估算的投资额。
            "scaleFactor": _lib.fmt(cost_factor),
            "scaleBase": _lib.fmt(_lib.q2(base_investment)),
            "investment": _lib.fmt(_lib.q2(investment)),
            "annualRevenue": _lib.fmt(_lib.q2(revenue)),
            "annualCost": _lib.fmt(_lib.q2(annual_cost)),
            "annualNetIncome": _lib.fmt(_lib.q2(revenue - annual_cost)),
            "indicators": {"firr": _lib.fmt(_lib.q4(rate)) if rate is not None else None,
                           "fnpv": _lib.fmt(npv) if npv is not None else None,
                           "pt": _lib.fmt(pt) if pt is not None else None},
            # 无收入 → 财务指标不可算，标 gap（但仍参与投资额与周期比选）
            "status": "gap" if missing_metrics else "ok",
            "gapReason": ("项目无经营性收入，财务指标（FIRR/FNPV/回收期）不可算；"
                          "本方案按总投资与实施周期参与比选") if missing_metrics else None,
        })

    # 综合评分（默认权重；AHP 缺失显式标注临时）——公式与上游一致，不做口径改动
    # 评分口径：
    #   * 有运营收入 → 按 FIRR 归一化（投资经济）+ 风险加分（原口径，不变）
    #   * 无运营收入 → 财务指标不可算，改按**总投资**归一（投资越低得分越高）
    #     + 风险加分；此时全部方案同口径，不掺入编造的财务数。
    scored = [item for item in results if item.get("status") in ("ok", "gap")]
    use_investment_score = no_revenue
    if use_investment_score:
        investments = [float(item["investment"] or 0) for item in scored]
        low, high = (min(investments), max(investments)) if investments else (0, 0)
        for item in scored:
            value = float(item["investment"] or 0)
            normalized = (high - value) / (high - low) if high > low else 1.0
            score = (float(weights.get("投资经济", 0.4)) * normalized
                     + float(weights.get("风险可控", 0.3)) * RISK_BONUS.get(item["name"], 0.6))
            item["score"] = _lib.fmt(_lib.q4(score))
    else:
        firrs = [float(item["indicators"]["firr"] or 0) for item in scored]
        low, high = (min(firrs), max(firrs)) if firrs else (0, 0)
        for item, firr in zip(scored, firrs):
            normalized = (firr - low) / (high - low) if high > low else 1.0
            score = (float(weights.get("投资经济", 0.4)) * normalized
                     + float(weights.get("风险可控", 0.3)) * RISK_BONUS.get(item["name"], 0.6))
            item["score"] = _lib.fmt(_lib.q4(score))
    if no_revenue:
        gaps.append({"code": SCHEME_GAP_CODE,
                     "what": "项目无经营性收入，三方案财务指标不可算",
                     "why": "无运营收入则 FIRR/FNPV/静态回收期在数学上无解；"
                            "已改按总投资与实施周期比选，未借用任何外部数字",
                     "need": "若项目有经营性收入，补充收入口径后重跑", "owner": "用户/项目"})
    winner = max(scored, key=lambda item: float(item["score"])) if scored else None
    ranked = sorted(scored, key=lambda item: (-float(item["score"]), item["id"]))
    for index, item in enumerate(ranked, start=1):
        item["scoreRank"] = index
        item["recommended"] = winner is not None and item["id"] == winner["id"]

    base = {
        "investment": _lib.fmt(_lib.q2(base_investment)),
        "annualRevenue": _lib.fmt(_lib.q2(revenue_base)),
        "annualCost": _lib.fmt(_lib.q2(operating_cost)),
        "firr": (finance.get("indicators") or {}).get("firrAll"),
        "fnpv": (finance.get("indicators") or {}).get("fnpvAll"),
        "pt": (finance.get("indicators") or {}).get("ptAll"),
        "operatingCostRatio": _lib.fmt(_lib.q2(operating_ratio)),
        "operatingCostRatioExact": _lib.fmt(q6(operating_ratio)),
        "basis": "基准 = 财务测算.json 的现行方案（总投资/年运营收入/年运营成本/指标）",
        "caliberNote": CALIBER_NOTE,
    }
    for item in results:
        item["recommendReason"] = build_recommend_reason(item, winner, base, weights_source)
    compare_table = build_compare_table(results, winner, weights_source, base, operating_ratio)
    compare = {"dimensions": ["总投资", "年运营收入", "FIRR", "FNPV", "静态回收期", "综合得分"],
               "matrix": [[item["id"], item["name"], item["investment"], item["annualRevenue"],
                           item["indicators"]["firr"], item["indicators"]["fnpv"],
                           item["indicators"]["pt"], item["score"]] for item in results]}
    if weights_source != "ahp":
        gaps.append({"code": AHP_GAP_CODE, "what": "AHP 权重缺失或不可用",
                     "why": "使用默认权重评分（temporary=true，待 AHP 插件输出替换）",
                     "need": "AHP 插件权重表", "owner": "ahp"})
    output = {
        "schema": "cardo/scheme-compare/v1",
        "schemaVersion": 1,
        "generatedAt": _lib.iso_now(args.no_timestamp),
        "producer": "cardo-invest-scheme/" + ENGINE_VERSION,
        "weightsSource": weights_source,
        "temporary": weights_source != "ahp",
        "ahpSource": ahp_detail,
        "caliberNote": CALIBER_NOTE,
        "schemes": results,
        "compare": compare,
        "compareTable": compare_table,
        "winner": winner["id"] if winner else None,
        "reportMapping": {"chapter": 9, "section": "9.3", "table": "方案对比表",
                          "compareTable": "compareTable"},
        "gaps": gaps,
    }
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "方案比选.json").write_text(
        json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (out_dir / "方案对比表.md").write_text(
        render_markdown(results, winner, weights_source, gaps), encoding="utf-8")
    print("[gen_schemes] 3 套方案 → 方案比选.json / 方案对比表.md（推荐 {}，权重 {}，对比表列 {}，缺口 {}）".format(
        winner["id"] if winner else "无", weights_source, len(compare_table["columns"]), len(gaps)))
    # 退出码 3 = 方案参数缺口待确认（AHP 临时权重不算缺口，仍 exit 0，与上游 T13 验收一致）
    return 3 if any(item.get("code") == SCHEME_GAP_CODE for item in gaps) else 0


if __name__ == "__main__":
    sys.exit(main())
