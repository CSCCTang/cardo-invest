"""上游取料与手动槽位（T9）：现状报告 + 现状事实库 + AHP 权重/评分 + 低空指标 + 财务参数 + 槽位
→ 财务输入底表（cardo/finance-input/v1）+ 缺项表。

纪律（与 SKILL.md §5 一一对应）：
  只收不猜——上游有则取用并记 sources[]，没有则 gaps[] + 显式占位（null/空表），链不中断；
  金额经 Decimal 归一（不做 float 算术、不改数值），口径与精度以 规则/财务口径.yaml 为准。

输入（全部只读，路径相对项目根）：
  04_成果报告/*现状分析报告*.md|.docx        现状分析报告（无机器契约，按内容抽取）
  05_过程/现状事实库/*.json                  cardo/facts/v1（占位名；只判存在，不解析）
  05_过程/AHP/权重表.json                     cardo/ahp-weights/v1（缺失→显式默认权重）
  05_过程/AHP/评分表.json                     cardo/ahp-scores/v1（缺失→第五/九章占位 + gap）
  05_过程/低空解析/现状指标.json               cardo/lowair-indicators/v1（只判存在，不解析）
  05_过程/投资决策/财务参数.json               cardo/invest-params/v1（缺失→占位）
  05_过程/投资决策/槽位.json                   cardo/invest-slots/v1（可选；--slots 优先）
输出（默认写 05_过程/投资决策/）：
  财务输入.json                               cardo/finance-input/v1
  缺项表.md                                   code / 缺什么 / 为什么缺 / 需要谁补 / 补到哪

用法：python collect_inputs.py --project-dir <项目目录> [--slots <槽位.json>] [--out <输出目录>] [--no-timestamp]
退出码：0 生成（可有 gaps）；1 输入坏/路径非法；2 项目目录无法识别。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

import _lib  # noqa: E402

ENGINE_VERSION = "collect_inputs/0.2.0"
RULES_REL = Path("规则") / "财务口径.yaml"
INVEST_REL = Path("05_过程") / "投资决策"
OUTPUT_JSON = "财务输入.json"
OUTPUT_GAPS = "缺项表.md"

# 下游（chapter-01..05 / estimate / finance / scheme）实际读取的项目字段
PROJECT_KEYS = ("name", "direction", "area", "constructionYears", "operationYears", "discountRate")
# estimate 中由输入给值的项 → 契约形状（依据 cardo-invest-estimate/calc_invest.py 实际读取：
# engineering[]/otherFees[] 为数组，contingency 是 {rate} 对象、interest 是 {loanAmount,loanRate,buildYears} 对象）
ESTIMATE_SHAPES = {
    "engineering": (list, []),
    "otherFees": (list, []),
    "otherFeeTotal": (None, None),
    "contingency": (dict, {"rate": None}),
    "interest": (dict, {}),
}
# estimate 中由下游按财务口径计算、本 skill 只占位不计算的字段
ESTIMATE_DERIVED_KEYS = ("constructionInvestment", "totalInvestment")
DERIVED_REASON = "由 cardo-invest-estimate 按财务口径计算（本 skill 不做任何计算，只占位）"
SHAPE_LABEL = {list: "数组", dict: "对象"}

DEFAULT_WEIGHTS = {
    "criteria": [
        {"key": "现状情况", "name": "现状适配性", "weight": 0.1},
        {"key": "投资经济", "name": "投资经济性", "weight": 0.3},
        {"key": "规划合规", "name": "规划合规性", "weight": 0.2},
        {"key": "风险可控", "name": "风险可控性", "weight": 0.4},
    ],
    "temporary": True,
    "note": "临时权重，待 AHP 插件输出替换",
}
# 指标层默认子权重（模板表 6 结构：每准则 4 项，子权重 0.5/0.25/0.15/0.1）
DEFAULT_INDICATORS = {
    "现状情况": [("房屋设施老化程度", 0.5), ("安全风险等级", 0.25), ("现状建筑密度", 0.15), ("现状容积率", 0.1)],
    "投资经济": [("改造总成本", 0.5), ("预期内部收益率", 0.25), ("投资回收期", 0.15), ("资产增值潜力", 0.1)],
    "规划合规": [("城市规划符合度", 0.4), ("政策合规", 0.3), ("基础设施配套率", 0.2), ("生态环保标准", 0.1)],
    "风险可控": [("政策风险", 0.5), ("成本超支风险", 0.25), ("社会稳定风险", 0.15), ("工程延误风险", 0.1)],
}

# code → (owner, 补到哪)
GAP_META = {    "UPSTREAM-STATUS": ("status", "04_成果报告/<项目>现状分析报告.md"),
    "UPSTREAM-FACTS": ("status", "05_过程/现状事实库/*.json"),
    "UPSTREAM-AHP": ("ahp", "05_过程/AHP/权重表.json"),
    "UPSTREAM-AHP-INVALID": ("ahp", "05_过程/AHP/权重表.json"),
    "UPSTREAM-AHP-SCORES": ("ahp", "05_过程/AHP/评分表.json"),
    "UPSTREAM-OSGB": ("osgb", "05_过程/低空解析/现状指标.json"),
    "INPUT-PARAMS": ("invest", "05_过程/投资决策/财务参数.json"),
    "INPUT-PARAMS-SCHEMA": ("invest", "05_过程/投资决策/财务参数.json"),
    "INPUT-SLOTS": ("user", "05_过程/投资决策/槽位.json"),
    "INPUT-PROJECT-KEY": ("user", "05_过程/投资决策/财务参数.json#project"),
    "INPUT-ESTIMATE-KEY": ("user", "05_过程/投资决策/财务参数.json#estimate"),
    "INPUT-REVENUE-KEY": ("user", "05_过程/投资决策/财务参数.json#revenue[]"),
    "INPUT-BAD-NUMBER": ("user", "05_过程/投资决策/财务参数.json"),
    "DEP-PYTHON-DOCX": ("user", "运行环境：pip install python-docx"),
    "RULE-FILE": ("invest", "skill/规则/财务口径.yaml"),
}

# 缺项表/ gaps[] 的稳定展示顺序（上游 → 输入 → 环境/口径）
GAP_ORDER = [
    "UPSTREAM-STATUS", "UPSTREAM-FACTS", "UPSTREAM-AHP", "UPSTREAM-AHP-INVALID", "UPSTREAM-AHP-SCORES",
    "UPSTREAM-OSGB", "INPUT-PARAMS", "INPUT-PARAMS-SCHEMA", "INPUT-PROJECT-KEY", "INPUT-ESTIMATE-KEY",
    "INPUT-REVENUE-KEY", "INPUT-BAD-NUMBER", "INPUT-SLOTS", "DEP-PYTHON-DOCX", "RULE-FILE",
]

SKIP_SOURCE_KEYS = {
    "key", "name", "unit", "note", "schema", "schemaVersion", "producer", "generatedAt",
    "gaps", "sources", "source", "formula", "basis", "sourceRow", "sourceCell", "vatRateNote",
}
# ahpWeights 只透出契约内字段，避免上游附加键（如夹具注记）污染产物
AHP_KEEP_KEYS = {
    "schema", "schemaVersion", "generatedAt", "producer", "method", "methodNote",
    "criteria", "criteriaMatrix", "criteriaWeights", "indicatorWeights", "consistency", "provenance",
}
ENVELOPE_KEYS = {"schema", "schemaVersion", "generatedAt", "producer", "source", "sources", "gaps"}

NUM_RE = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$")
AMOUNT_HINT = ("amount", "price", "total", "cost", "fee", "investment", "contingency", "interest",
               "loan", "equity", "value", "capital", "salary", "profit", "tax", "revenue", "income")
RATIO_HINT = ("rate", "ratio", "growth", "factor", "percent", "share")
COUNT_HINT = ("area", "quantity", "days", "months", "years", "crew", "term", "buildyears", "graceyears")
NUMERIC_HINT = AMOUNT_HINT + RATIO_HINT + COUNT_HINT


class InputError(Exception):
    """输入坏/路径非法 → exit 1。"""


class ProjectError(Exception):
    """项目目录无法识别 → exit 2。"""


# ------------------------------------------------------------------ 基础

def make_gap(code, what, why, need, owner=None, where=None):
    meta_owner, meta_where = GAP_META.get(code, ("user", "—"))
    return {
        "code": code,
        "what": what,
        "why": why,
        "need": need,
        "owner": owner or meta_owner,
        "where": where or meta_where,
    }


def read_json_file(path, label):
    """读 JSON；文件不存在返回 None；解析失败 → InputError（exit 1）。"""
    if not path.is_file():
        return None
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError, UnicodeDecodeError) as error:
        raise InputError("{} 读取失败（{}）：{}".format(label, _lib.rel_posix(path, path.parent), error))


def field_kind(key):
    lowered = str(key).lower()
    if any(hint in lowered for hint in RATIO_HINT):
        return "ratio"
    if any(hint in lowered for hint in AMOUNT_HINT):
        return "amount"
    return "other"


def dec_or_none(value):
    """数值 → Decimal；非数值返回 (None, 原文)。"""
    if isinstance(value, bool) or value is None:
        return None, value
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)), value
        except InvalidOperation:
            return None, value
    if isinstance(value, str) and NUM_RE.match(value.strip()):
        try:
            return Decimal(value.strip()), value
        except InvalidOperation:
            return None, value
    return None, value


def normalize_tree(node, prefix, file_rel, kind, ctx):
    """递归规整：数值经 Decimal 归一后原值输出；同时登记 sources[] 与占位。"""
    if isinstance(node, dict):
        out = {}
        for key, value in node.items():
            child = "{}.{}".format(prefix, key) if prefix else str(key)
            out[key] = normalize_tree(value, child, file_rel, kind, ctx)
        return out
    if isinstance(node, list):
        return [normalize_tree(value, "{}[{}]".format(prefix, index), file_rel, kind, ctx)
                for index, value in enumerate(node)]

    if node is None or node == "":
        ctx["placeholders"].append({
            "field": prefix or "—",
            "gap": ctx.get("active_gap"),
            "reason": "上游缺项占位（不补数）" if ctx.get("active_gap") else "输入未给值（不补数）",
        })
        return node

    number, _raw = dec_or_none(node)
    key = prefix.rsplit(".", 1)[-1].split("[")[0]
    if number is not None and key not in SKIP_SOURCE_KEYS:
        # 口径核对：金额小数位超出 财务口径 rounding.amountDecimals 时提示，不改数
        decimals = -number.as_tuple().exponent if number.as_tuple().exponent < 0 else 0
        limit = ctx["limits"]["amount"]
        if limit is not None and field_kind(key) == "amount" and decimals > int(limit):
            ctx["ruleChecks"].append({
                "code": "RULE-DECIMALS", "field": prefix, "severity": "warn",
                "detail": "小数位 {} 超出财务口径 {} 位（保留原值，未改数）".format(decimals, limit),
            })
        ctx["sources"].append({
            "field": prefix,
            "source": {"file": file_rel, "locator": "$." + prefix, "quote": _lib.fmt(number)},
            "kind": kind,
        })
        # Decimal → JSON 数值：整数值落 int，其余落 float（≤15 位有效数字可无损往返）
        if number == number.to_integral_value():
            return int(number)
        return float(number)
    if isinstance(node, str) and key not in SKIP_SOURCE_KEYS:
        # 本地修正：slots.* 是文本槽位（功能定位/运营模式等），不参与数值型字段扫描；
        # 否则 operationModel 里的 "ratio" 会被 NUMERIC_HINT 误判成比率字段。
        if not prefix.startswith("slots.") and any(hint in str(key).lower() for hint in NUMERIC_HINT):
            # 数值型字段给了非数值 → 报输入坏，不把坏值透传给下游
            ctx["gaps"].append(make_gap("INPUT-BAD-NUMBER", "{} 不是数值".format(prefix),
                                        "该字段应为数值，实际为文本 {!r}".format(node[:30]),
                                        "改为数值（单位见规则/财务口径.yaml units）"))
        else:
            ctx["sources"].append({
                "field": prefix,
                "source": {"file": file_rel, "locator": "$." + prefix, "quote": node[:60]},
                "kind": kind,
            })
    return node


# ------------------------------------------------------------------ AHP

def valid_ahp(payload):
    """结构校验：criteria 非空、每项 weight 可解析、权重和 > 0。"""
    if not isinstance(payload, dict):
        return False
    criteria = payload.get("criteria")
    if not isinstance(criteria, list) or not criteria:
        return False
    total = Decimal("0")
    for item in criteria:
        if not isinstance(item, dict):
            return False
        number, _ = dec_or_none(item.get("weight"))
        if number is None or number < 0:
            return False
        total += number
    return total > 0


def enrich_ahp(ahp, kind):
    """把权重表补成表 3–6 可渲染结构；缺矩阵时派生 rank-1 矩阵并标注来源。"""
    criteria = [dict(item) for item in (ahp.get("criteria") or [])]
    weights = []
    for item in criteria:
        number, _ = dec_or_none(item.get("weight"))
        weights.append(float(number or 0))
    total = sum(weights) or 1.0
    weights = [value / total for value in weights]
    full = ahp.get("criteriaMatrix")
    derived = full is None
    if full is None:
        full = [[round(weights[row] / weights[col], 4) if weights[col] else 0 for col in range(len(weights))]
                for row in range(len(weights))]
    geomean_all = 1.0
    for value in weights:
        geomean_all *= value
    geomean_all = geomean_all ** (1.0 / len(weights)) if geomean_all > 0 else 1.0
    criteria_weights = ahp.get("criteriaWeights")
    if criteria_weights:
        # 上游权重表只给权重、不给行几何平均时，按判别矩阵补算该展示列；
        # 仅作表格展示中间量，不参与权重与排序，计算口径写入 provenance。
        criteria_weights = [dict(item) for item in criteria_weights]
        for index, item in enumerate(criteria_weights):
            if item.get("geometricMean") is not None or index >= len(full):
                continue
            row = [dec_or_none(value)[0] for value in full[index][:len(weights)]]
            if row and all(number is not None and number > 0 for number in row):
                product = 1.0
                for number in row:
                    product *= float(number)
                item["geometricMean"] = round(product ** (1.0 / len(row)), 5)
    else:
        criteria_weights = [{"key": item.get("key"), "name": item.get("name"), "weight": round(weights[index], 4),
                             "geometricMean": round((weights[index] / geomean_all) if geomean_all else 0, 5)}
                            for index, item in enumerate(criteria)]
    indicator_weights = ahp.get("indicatorWeights")
    if not indicator_weights:
        indicator_weights = []
        for index, item in enumerate(criteria):
            key = item.get("key") or item.get("name")
            for name, sub in DEFAULT_INDICATORS.get(key, []):
                indicator_weights.append({"criterion": key, "criterionWeight": round(weights[index], 4),
                                          "indicator": name, "subWeight": sub,
                                          "compositeWeight": round(weights[index] * sub, 4)})
    consistency = ahp.get("consistency")
    if consistency is None:
        consistency = {"cr": None, "method": "derived-rank1" if derived else "provided",
                       "note": "权重为派生/默认值，未做 AHP 一致性检验" if derived else "以 AHP 插件输出为准"}
    enriched = {key: value for key, value in ahp.items() if key in AHP_KEEP_KEYS}
    enriched.update({
        "kind": kind,                                  # default | upstream
        "criteria": [{"key": item.get("key"), "name": item.get("name"), "weight": round(weights[index], 4)}
                     for index, item in enumerate(criteria)],
        "criteriaMatrix": full,
        "criteriaWeights": criteria_weights,
        "indicatorWeights": indicator_weights,
        "consistency": consistency,
        "derived": derived,
        "temporary": kind == "default",
        "provenance": ahp.get("provenance") or ("默认权重（AHP 缺失）" if kind == "default" else
                                                ("准则权重来自 AHP 权重表；指标层矩阵/子权重为派生" if derived
                                                 else "AHP 权重表")),
        "note": "临时权重，待 AHP 插件输出替换" if kind == "default" else ahp.get("note"),
    })
    if enriched["note"] is None:
        enriched.pop("note")
    return enriched


# ------------------------------------------------------------------ 现状报告

def read_report(project, ctx):
    base = project / "04_成果报告"
    candidates = []
    if base.is_dir():
        for pattern in ("*现状分析报告*.md", "*现状分析报告*.docx"):
            candidates = sorted(base.glob(pattern))
            if candidates:
                break
    if not candidates:
        ctx["gaps"].append(make_gap("UPSTREAM-STATUS", "现状分析报告缺失",
                                    "04_成果报告 下无 *现状分析报告*.md/.docx",
                                    "现状分析报告源文件（.md 或 .docx）"))
        return None
    path = candidates[0]
    relative = _lib.rel_posix(path, project)
    text, note = None, None
    if path.suffix.lower() == ".md":
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError as error:
            raise InputError("现状分析报告读取失败：{}".format(error))
    else:
        try:
            import docx  # noqa: F401
        except ImportError:
            ctx["gaps"].append(make_gap("DEP-PYTHON-DOCX", "缺 python-docx，.docx 现状报告未解析",
                                        "运行环境无 docx 模块（依赖未安装）",
                                        "pip install python-docx，或提供同内容 .md 版"))
            note = "docx-not-parsed"
        else:
            document = docx.Document(str(path))
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    ctx["sources"].append({
        "field": "upstream.statusReport",
        "source": {"file": relative, "locator": "全文（抽取 基本信息/改造内容/投资相关）",
                   "quote": (text or "").strip().replace("\n", " ")[:60] or "—"},
        "kind": "upstream",
    })
    return {"file": relative, "sha256": _lib.sha256_file(path), "text": text, "note": note}


# ------------------------------------------------------------------ 主流程

def extract_slots(payload):
    if not isinstance(payload, dict):
        return {}
    inner = payload.get("slots")
    if isinstance(inner, dict) and inner:
        return inner
    return {key: value for key, value in payload.items() if key not in ENVELOPE_KEYS}


def recognize_project(project):
    marks = [project / "00_项目索引.md", project / ".cardo", project / "04_成果报告", project / "05_过程"]
    return any(mark.exists() for mark in marks)


def build(args):
    project = Path(args.project_dir)
    if not project.exists():
        raise ProjectError("项目目录不存在：{}".format(args.project_dir))
    if not project.is_dir():
        raise ProjectError("不是目录：{}".format(args.project_dir))
    if not recognize_project(project):
        raise ProjectError("项目目录无法识别（缺 00_项目索引.md / .cardo / 04_成果报告 / 05_过程）：{}"
                           .format(args.project_dir))

    # 路径合法性先判（exit 1），避免白跑取料
    if args.slots and not Path(args.slots).is_file():
        raise InputError("--slots 指向的文件不存在：{}".format(args.slots))

    # 输入固定取自项目内 05_过程/投资决策；--out 只决定写盘位置（不得反向影响取料）
    in_dir = project / INVEST_REL
    out_dir = Path(args.out) if args.out else in_dir
    ctx = {"gaps": [], "sources": [], "placeholders": [], "ruleChecks": [],
           "limits": {"amount": None, "ratio": None}, "active_gap": None}

    # ---- 口径文件（唯一口径来源；不读则无法核对精度）
    rules, rules_rel, rules_sha = None, None, None
    rules_path = SKILL_ROOT / RULES_REL
    if rules_path.is_file():
        try:
            rules = _lib.load_yaml(rules_path)
            rules_rel = RULES_REL.as_posix()
            rules_sha = _lib.sha256_file(rules_path)
        except (OSError, _lib.YamlError) as error:
            raise InputError("财务口径.yaml 解析失败：{}".format(error))
        rounding = (rules or {}).get("rounding") or {}
        ctx["limits"]["amount"] = rounding.get("amountDecimals")
        ctx["limits"]["ratio"] = rounding.get("ratioDecimals")
        ctx["sources"].append({
            "field": "__rules__",
            "source": {"file": rules_rel, "locator": "rounding/revenue.items/tax", "quote": "字段口径与精度"},
            "kind": "template",
        })
    else:
        ctx["gaps"].append(make_gap("RULE-FILE", "财务口径.yaml 缺失",
                                    "skill/规则/财务口径.yaml 不存在，无法核对字段口径",
                                    "从 cardo-invest-spec/规则 同步 财务口径.yaml"))

    # ---- 财务参数（人工给定的项目输入）
    params = read_json_file(in_dir / "财务参数.json", "财务参数")
    params_rel = _lib.rel_posix(in_dir / "财务参数.json", project)
    if params is None:
        ctx["active_gap"] = "INPUT-PARAMS"
        ctx["gaps"].append(make_gap("INPUT-PARAMS", "财务参数缺失",
                                    "缺 {}".format(params_rel),
                                    "project/estimate/revenue/cost/tax/finance 参数（cardo/invest-params/v1）"))
    elif not isinstance(params, dict):
        raise InputError("财务参数必须是 JSON 对象：{}".format(params_rel))
    elif params.get("schema") not in (None, "cardo/invest-params/v1"):
        ctx["gaps"].append(make_gap("INPUT-PARAMS-SCHEMA", "财务参数 schema 不匹配",
                                    "{} 的 schema={}（期望 cardo/invest-params/v1）".format(
                                        params_rel, params.get("schema")),
                                    "按 cardo/invest-params/v1 重出财务参数；本次不采用该文件"))
        params = None
        ctx["active_gap"] = "INPUT-PARAMS"

    blocks = {}
    source_kind = "user"
    for name in ("project", "estimate", "revenue", "cost", "tax", "finance", "profit"):
        value = (params or {}).get(name)
        if params is None:
            blocks[name] = None
        else:
            blocks[name] = value
        if value is not None:
            blocks[name] = normalize_tree(value, name, params_rel, source_kind, ctx)

    # 缺省块 → 显式占位（null/空表），不补数
    project_block = blocks.get("project")
    if project_block is None:
        project_block = {key: None for key in PROJECT_KEYS}
        ctx["gaps"].append(make_gap("INPUT-PROJECT-KEY", "project 关键字段全缺",
                                    "财务参数未提供 project 块（或缺文件）",
                                    "project.{} 六项".format("/".join(PROJECT_KEYS))))
        for key in PROJECT_KEYS:
            ctx["placeholders"].append({"field": "project.{}".format(key), "gap": "INPUT-PROJECT-KEY"})
    else:
        missing = [key for key in PROJECT_KEYS if project_block.get(key) in (None, "")]
        if missing:
            ctx["gaps"].append(make_gap("INPUT-PROJECT-KEY", "project 关键字段缺失：{}".format(", ".join(missing)),
                                        "下游 chapter/estimate 需读 project.{}".format("/".join(missing)),
                                        "补齐 project.{}".format("/".join(missing))))
            for key in missing:
                ctx["placeholders"].append({"field": "project.{}".format(key), "gap": "INPUT-PROJECT-KEY"})

    estimate_block = blocks.get("estimate")
    if estimate_block is None:
        estimate_block = {key: (list() if default is not None and isinstance(default, list)
                                else (dict(default) if isinstance(default, dict) else None))
                          for key, (_kind, default) in ESTIMATE_SHAPES.items()}
        for key, (_kind, default) in ESTIMATE_SHAPES.items():
            ctx["placeholders"].append({"field": "estimate.{}".format(key), "gap": "INPUT-PARAMS",
                                        "reason": "财务参数缺失（不补数）"})
    else:
        # 形状硬校验：形状不符=输入坏（exit 1），不静默转换、不丢数
        problems = []
        for key, (kind, _default) in ESTIMATE_SHAPES.items():
            value = estimate_block.get(key)
            if kind is None or value is None:
                continue
            if not isinstance(value, kind):
                problems.append("estimate.{} 必须是{}，实际为 {}".format(
                    key, SHAPE_LABEL[kind], type(value).__name__))
        for key in ("engineering", "otherFees"):
            for index, item in enumerate(estimate_block.get(key) or []):
                if not isinstance(item, dict):
                    problems.append("estimate.{}[{}] 必须是对象".format(key, index))
        if problems:
            raise InputError("财务参数形状与下游契约不符（{}）：{}；请按 cardo/invest-params/v1 修正".format(
                params_rel, "；".join(problems)))
        missing = [key for key in ESTIMATE_SHAPES if key not in estimate_block]
        if missing:
            ctx["gaps"].append(make_gap("INPUT-ESTIMATE-KEY", "estimate 缺关键项：{}".format(", ".join(missing)),
                                        "财务参数未给 estimate.{}".format("/".join(missing)),
                                        "补齐 estimate.{}".format("/".join(missing))))
            for key in missing:
                kind, default = ESTIMATE_SHAPES[key]
                estimate_block[key] = (dict(default) if isinstance(default, dict)
                                       else (list() if isinstance(default, list) else None))
                ctx["placeholders"].append({"field": "estimate.{}".format(key), "gap": "INPUT-ESTIMATE-KEY",
                                            "reason": "上游缺项占位（不补数）"})
        # 对象内必需键缺失 → 记缺项（下游按 rate/借款额出 gaps，此处提前暴露）
        if isinstance(estimate_block.get("contingency"), dict) and estimate_block["contingency"].get("rate") is None:
            estimate_block["contingency"] = {"rate": None}
            ctx["gaps"].append(make_gap("INPUT-ESTIMATE-KEY", "预备费率缺失",
                                        "estimate.contingency.rate 未给，预备费无法按口径计算",
                                        "预备费率（estimate.contingency.rate）"))
            ctx["placeholders"].append({"field": "estimate.contingency.rate", "gap": "INPUT-ESTIMATE-KEY",
                                        "reason": "上游缺项占位（不补数）"})
    for key in ESTIMATE_DERIVED_KEYS:
        estimate_block.setdefault(key, None)
        ctx["placeholders"].append({"field": "estimate.{}".format(key), "gap": None, "reason": DERIVED_REASON})

    revenue_block = blocks.get("revenue")
    if revenue_block is None:
        revenue_block = []
        if params is not None:
            ctx["placeholders"].append({"field": "revenue", "gap": "INPUT-PARAMS"})
    elif isinstance(revenue_block, list):
        for index, item in enumerate(revenue_block):
            if not isinstance(item, dict):
                ctx["gaps"].append(make_gap("INPUT-REVENUE-KEY", "revenue[{}] 结构不合法".format(index),
                                            "收入项不是对象", "revenue[{}] 需 {{key,name,price,quantity,...}}".format(index)))
                continue
            has_amount = item.get("annualAmount") not in (None, "")
            has_base = item.get("price") not in (None, "") and item.get("quantity") not in (None, "")
            if not (has_amount or has_base):
                label = item.get("key") or "revenue[{}]".format(index)
                ctx["gaps"].append(make_gap("INPUT-REVENUE-KEY", "收入项 {} 缺价量".format(label),
                                            "既无 annualAmount，也缺 price/quantity",
                                            "annualAmount（万元/年）或 price+quantity"))
                ctx["placeholders"].append({"field": "revenue[{}].price/quantity".format(index),
                                            "gap": "INPUT-REVENUE-KEY"})
            elif item.get("vatRate") is not None and rules:
                expected = None
                for entry in ((rules.get("revenue") or {}).get("items") or []):
                    if isinstance(entry, dict) and entry.get("key") == item.get("key"):
                        expected = entry.get("vatRate")
                        break
                given, _ = dec_or_none(item.get("vatRate"))
                want, _ = dec_or_none(expected)
                if given is not None and want is not None and given != want:
                    ctx["ruleChecks"].append({
                        "code": "RULE-VAT", "field": "revenue[{}].vatRate".format(index), "severity": "warn",
                        "detail": "输入 {} 与财务口径 {} 不一致（保留输入值，不静默改写）".format(
                            _lib.fmt(given), _lib.fmt(want)),
                    })
    ctx["active_gap"] = None

    # ---- AHP 权重
    ahp_raw = read_json_file(project / "05_过程" / "AHP" / "权重表.json", "AHP 权重表")
    ahp_kind = "default"
    if ahp_raw is not None and valid_ahp(ahp_raw):
        ahp_kind = "upstream"
        ahp_file = _lib.rel_posix(project / "05_过程" / "AHP" / "权重表.json", project)
        for index, item in enumerate((ahp_raw.get("criteria") or [])):
            number, _ = dec_or_none(item.get("weight"))
            ctx["sources"].append({
                "field": "ahpWeights.criteria[{}].weight".format(item.get("key") or index),
                "source": {"file": ahp_file, "locator": "$.criteria[{}].weight".format(index),
                           "quote": _lib.fmt(number) if number is not None else "—"},
                "kind": "upstream",
            })
        for index, item in enumerate((ahp_raw.get("indicatorWeights") or [])):
            if isinstance(item, dict) and item.get("subWeight") is not None:
                number, _ = dec_or_none(item.get("subWeight"))
                ctx["sources"].append({
                    "field": "ahpWeights.indicatorWeights[{}].subWeight".format(item.get("indicator") or index),
                    "source": {"file": ahp_file, "locator": "$.indicatorWeights[{}].subWeight".format(index),
                               "quote": _lib.fmt(number) if number is not None else "—"},
                    "kind": "upstream",
                })
    elif ahp_raw is not None:
        ctx["gaps"].append(make_gap("UPSTREAM-AHP-INVALID", "AHP 权重表结构不合规",
                                    "05_过程/AHP/权重表.json 缺 criteria 或权重无法解析",
                                    "按 cardo/ahp-weights/v1 重出权重表；本次用显式默认权重"))
    else:
        ctx["gaps"].append(make_gap("UPSTREAM-AHP", "AHP 权重缺失",
                                    "05_过程/AHP/权重表.json 不存在",
                                    "AHP 插件权重表（cardo/ahp-weights/v1）"))
    ahp_source = ahp_raw if ahp_kind == "upstream" else DEFAULT_WEIGHTS
    if ahp_kind == "default":
        ctx["sources"].append({
            "field": "ahpWeights.criteria",
            "source": {"file": "skill/scripts/collect_inputs.py", "locator": "DEFAULT_WEIGHTS",
                       "quote": "临时权重，待 AHP 插件输出替换"},
            "kind": "default",
        })
    ahp = enrich_ahp(ahp_source, ahp_kind)
    ahp["sourceFile"] = (_lib.rel_posix(project / "05_过程" / "AHP" / "权重表.json", project)
                         if ahp_kind == "upstream" else None)

    # ---- AHP 评分（缺失 → 第五/九章占位 + gap）
    scores_path = project / "05_过程" / "AHP" / "评分表.json"
    scores = read_json_file(scores_path, "AHP 评分表")
    ahp_scores = None
    if scores is None:
        ctx["gaps"].append(make_gap("UPSTREAM-AHP-SCORES", "AHP 评分表缺失",
                                    "05_过程/AHP/评分表.json 不存在，第五章/第九章评分段占位",
                                    "AHP 评分表（cardo/ahp-scores/v1）"))
        ctx["placeholders"].append({"field": "ahpScores", "gap": "UPSTREAM-AHP-SCORES"})
    else:
        ahp_scores = {"file": _lib.rel_posix(scores_path, project), "sha256": _lib.sha256_file(scores_path)}
        ctx["sources"].append({
            "field": "ahpScores",
            "source": {"file": ahp_scores["file"], "locator": "全文", "quote": "AHP 评分"},
            "kind": "upstream",
        })

    # ---- 低空指标（只判存在，不解析内容）
    lowair_path = project / "05_过程" / "低空解析" / "现状指标.json"
    lowair = None
    if lowair_path.is_file():
        lowair = {"status": "present", "file": _lib.rel_posix(lowair_path, project)}
        ctx["sources"].append({
            "field": "upstream.lowairIndicators",
            "source": {"file": lowair["file"], "locator": "—（本 skill 不解析，透传下游）", "quote": "存在"},
            "kind": "upstream",
        })
    else:
        ctx["gaps"].append(make_gap("UPSTREAM-OSGB", "低空现状指标缺失",
                                    "05_过程/低空解析/现状指标.json 不存在",
                                    "低空解析插件指标（cardo/lowair-indicators/v1）"))
        ctx["placeholders"].append({"field": "upstream.lowairIndicators", "gap": "UPSTREAM-OSGB"})

    # ---- 现状事实库（只判存在，不解析内容）
    facts_dir = project / "05_过程" / "现状事实库"
    facts_files = sorted(_lib.rel_posix(path, project) for path in facts_dir.glob("*.json")) if facts_dir.is_dir() else []
    if not facts_files:
        ctx["gaps"].append(make_gap("UPSTREAM-FACTS", "现状事实库缺失",
                                    "05_过程/现状事实库 下无 *.json",
                                    "现状事实库 JSON（cardo/facts/v1，契约占位名）"))
        ctx["placeholders"].append({"field": "upstream.facts", "gap": "UPSTREAM-FACTS"})

    # ---- 现状报告
    report = read_report(project, ctx)

    # ---- 槽位（--slots 优先，项目内 槽位.json 兜底）
    slots_payload, slots_file, slots_origin = None, None, None
    if args.slots:
        slots_path = Path(args.slots)
        if not slots_path.is_file():
            raise InputError("--slots 指向的文件不存在：{}".format(args.slots))
        slots_payload = read_json_file(slots_path, "槽位（--slots）")
        slots_origin = "cli"
        given = args.slots.replace("\\", "/")
        relative = _lib.rel_posix(slots_path, project)
        if relative != slots_path.name:
            slots_file = relative                     # 项目内 → 相对项目根 posix
        elif re.match(r"^([A-Za-z]:|/)", given):
            slots_file = slots_path.name               # 绝对路径 → 只留文件名（不写盘符）
        else:
            slots_file = given                         # 调用方给的相对路径 → 原样记录
    else:
        candidate = in_dir / "槽位.json"
        if candidate.is_file():
            slots_payload = read_json_file(candidate, "槽位")
            slots_file = _lib.rel_posix(candidate, project)
            slots_origin = "project"
    slots_block = extract_slots(slots_payload) if slots_payload is not None else {}
    if slots_block:
        slots_block = normalize_tree(slots_block, "slots", slots_file, "user", ctx)
    else:
        slots_block = {}
        ctx["gaps"].append(make_gap("INPUT-SLOTS", "手动槽位缺失",
                                    "未传 --slots，且 {} 不存在或为空".format(_lib.rel_posix(in_dir / "槽位.json", project)),
                                    "功能定位/业态/时序/组织/资金结构等槽位（cardo/invest-slots/v1）"))
        ctx["placeholders"].append({"field": "slots", "gap": "INPUT-SLOTS"})

    # ---- 汇总
    for item in ctx["placeholders"]:            # 统一补齐 reason（缺项占位 / 下游计算占位）
        item.setdefault("reason", "上游缺项占位（不补数）" if item.get("gap") else DERIVED_REASON)
    ctx["gaps"].sort(key=lambda item: (GAP_ORDER.index(item["code"]) if item["code"] in GAP_ORDER
                                       else len(GAP_ORDER), item["code"]))
    payload = {
        "schema": "cardo/finance-input/v1",
        "schemaVersion": 1,
        "generatedAt": _lib.iso_now(args.no_timestamp),
        "producer": "cardo-invest-intake/" + ENGINE_VERSION.split("/")[-1],
        "project": project_block,
        "estimate": estimate_block,
        "revenue": revenue_block,
        "cost": blocks.get("cost") or {},
        "tax": blocks.get("tax") or {},
        "finance": blocks.get("finance") or {},
        "profit": blocks.get("profit") or {},
        "slots": slots_block,
        "slotsMeta": {"file": slots_file, "origin": slots_origin},
        "ahpWeights": ahp,
        "ahpScores": ahp_scores,
        "upstream": {
            "statusReport": ({"file": report["file"], "sha256": report["sha256"],
                              "textAvailable": report["text"] is not None,
                              "note": report["note"]} if report else None),
            "facts": {"files": facts_files} if facts_files else None,
            "lowairIndicators": lowair,
            "excerpt": (report["text"][:4000] if report and report["text"] else None),
        },
        "rules": ({"schema": (rules or {}).get("schema"), "version": (rules or {}).get("version"),
                   "file": rules_rel, "scope": "skill", "sha256": rules_sha} if rules else None),
        "sources": ctx["sources"],
        "placeholders": ctx["placeholders"],
        "ruleChecks": ctx["ruleChecks"],
        "gaps": ctx["gaps"],
    }
    return payload, out_dir


def render_gaps(payload):
    gaps = payload["gaps"]
    lines = ["# 缺项表（cardo-invest-intake）", "",
             "- 项目：`{}`".format(payload["project"].get("name") or "（未提供，占位）"),
             "- 产物：`cardo/finance-input/v1`（schemaVersion {}）".format(payload["schemaVersion"]),
             "- 缺口：{} 项；占位字段：{} 个；口径提示：{} 条".format(
                 len(gaps), len(payload["placeholders"]), len(payload["ruleChecks"])),
             "", "## 1. 缺项（对应 `财务输入.json` 的 `gaps[]`）", "",
             "| code | 缺什么 | 为什么缺 | 需要谁补 | 补到哪 |", "|---|---|---|---|---|"]
    if gaps:
        for item in gaps:
            lines.append("| `{}` | {} | {} | {} | `{}` |".format(
                item["code"], item["what"], item["why"], item["owner"], item["where"]))
    else:
        lines.append("| — | 无缺项 | — | — | — |")

    lines += ["", "## 2. 占位字段（值为 null/空表，未补数）", "",
              "| 字段 | 关联 gap | 占位原因 |", "|---|---|---|"]
    if payload["placeholders"]:
        for item in payload["placeholders"]:
            lines.append("| `{}` | {} | {} |".format(
                item["field"],
                "`{}`".format(item["gap"]) if item.get("gap") else "—",
                item.get("reason") or "—"))
    else:
        lines.append("| — | — | — |")

    lines += ["", "## 3. 口径核对提示（不阻断，不改数）", "",
              "| code | 字段 | 说明 |", "|---|---|---|"]
    if payload["ruleChecks"]:
        for item in payload["ruleChecks"]:
            lines.append("| `{}` | `{}` | {} |".format(item["code"], item["field"], item["detail"]))
    else:
        lines.append("| — | — | 与 `规则/财务口径.yaml` 无冲突 |")

    lines += ["", "## 4. 纪律说明", "",
              "1. **有缺项不是失败**：本表与 `gaps[]` 一一对应，命令仍以退出码 0 结束，链不中断。",
              "2. **不补数**：缺失字段一律置 `null`/空表（见 §2），不填行业经验值、不引用相似项目数字。",
              "3. **不改上游**：修复建议只写在本表，不回写现状报告/AHP/低空产物。",
              "4. **来源可追**：已取到的每个值在 `财务输入.json` 的 `sources[]` 有 `{field, source:{file, locator, quote}, kind}`。",
              "5. 补齐后重跑：`python scripts/collect_inputs.py --project-dir <项目目录>`。",
              ""]
    return "\n".join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 上游取料与手动槽位（T9）")
    parser.add_argument("--project-dir", required=True, help="项目根目录")
    parser.add_argument("--slots", default=None, help="手动槽位 JSON（优先于项目内 05_过程/投资决策/槽位.json）")
    parser.add_argument("--out", default=None, help="输出目录（默认 <项目>/05_过程/投资决策）")
    parser.add_argument("--no-timestamp", action="store_true", help="去除时间戳（两次跑 sha256 一致）")
    args = parser.parse_args(argv)

    try:
        payload, out_dir = build(args)
    except ProjectError as error:
        print("[collect_inputs] {}".format(error))
        return 2
    except InputError as error:
        print("[collect_inputs] {}".format(error))
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    _lib.write_json(out_dir / OUTPUT_JSON, payload, no_timestamp=args.no_timestamp)
    (out_dir / OUTPUT_GAPS).write_text(render_gaps(payload), encoding="utf-8")
    print("[collect_inputs] 写出 {}/{}（gaps={}, placeholders={}, sources={}）".format(
        _lib.rel_posix(out_dir, Path(args.project_dir)) if Path(args.project_dir).is_dir() else out_dir,
        OUTPUT_JSON, len(payload["gaps"]), len(payload["placeholders"]), len(payload["sources"])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
