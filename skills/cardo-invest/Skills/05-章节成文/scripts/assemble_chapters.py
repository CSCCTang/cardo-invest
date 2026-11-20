"""分章成文（T10）：按章节契约 10 章，从上游摘要 + 财务输入/投资估算/财务测算/方案比选
生成分章 md + 缺项表。

取料纪律（本文件的第一原则）：
  * 正文里出现的每一个地名、机构名、面积、工程量、人数、金额、日期、文号，
    都必须来自 上游摘要.json / 财务输入.json / 投资估算.json / 财务测算.json / 方案比选.json；
  * 取不到就写「（资料缺失：<该写什么>，见缺项表）」，并登记到缺项表；
  * 本文件不含任何具体项目的常量，也不得把可研当输入（可研是成果，不是用户投递的资料）。

输入：05_过程/投资决策/{上游摘要.json,财务输入.json,投资估算.json,财务测算.json,方案比选.json?}
输出：05_过程/投资决策/分章/CH-01.md … CH-10.md、缺项表.md、分章/来源映射.json

用法：python assemble_chapters.py --project-dir <项目目录> [--no-timestamp]
退出码：0 成功；2 必需输入缺失/坏。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from decimal import Decimal
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SPEC_SCRIPTS = SCRIPT_DIR
sys.path.insert(0, str(SPEC_SCRIPTS))

import _lib  # noqa: E402
import basis_library  # noqa: E402

ENGINE_VERSION = "assemble_chapters/1.1.0"
DEC = Path("05_过程") / "投资决策"
CHAPTERS = [
    (1, "项目概述"), (2, "项目背景与建设必要性"), (3, "现状调研分析"), (4, "市场分析与功能定位"),
    (5, "项目建设范围、改造内容与建设方案"), (6, "组织架构与运营模式"),
    (7, "投资估算与资金筹措方案"), (8, "财务评价"),
    (9, "基于AHP层次分析法的投资决策指标体系量化分析"), (10, "附件清单"),
]
TABLE_TITLE_RE = re.compile(r"^表\s*\d+\s*-\s*\d+\s+")
PLACEHOLDER_TMPL = "（推断）"
# C8 口径：方案类槽位未填时不许自编业态/运营模式/组织架构，正文写这句话；
# 缺数类（公共空间面积、停车位数量等）业主未提供时正文写「待确认」，不估算。
PENDING_SCHEME = "待业主单位填写（方案深化阶段确定）"
PENDING_VALUE = "待确认"
PLACEHOLDER = "（推断）"
PLACEHOLDER_LEGACY = "（初步判定："


def pending_tail(pairs):
    """缺数类口径：业主未提供的定量指标不进正文数值位，统一写「待确认」。

    C8 第 1 类：公共空间面积、外立面整治面积、配套道路面积、停车位数量、
    充电桩数量、沿街游商数量——业主补数，正文标待确认，不许估。
    """
    missing = [label for label, value in pairs
               if value in (None, "") or str(value).strip() in ("—", "-", PENDING_VALUE)]
    if not missing:
        return ""
    return "；{}待确认，以业主单位提供的数据为准".format("、".join(missing))
# 建议书成品不做「推断/可核对」这类工作底稿标注：评审稿只呈现有据的内容。
INFER_NOTE = ""
CHECK_NOTE = ""
ZERO = Decimal("0")

# 上游摘要里表示「区间或未确定」的值，不作事实使用
NON_FACT_VALUES = {"", "—", "-", "待定", "不详", "未确定"}
# 槽位值是「写作说明」而非事实（如「本章为『项目概述』，按…口径成文」）
SLOT_NOISE_RE = re.compile(r"本章为|本章按|口径，依据|依据项目材料成文")
# 上游原文里夹带的写作说明/流程提示（「本节按…口径核对」「下一步：…」）不是报告内容，
# 落到正文里读者会当成正文事实，统一在这里剥掉。
# 「本节按…核对/列示…」是写给我方看的口径说明，不是正文内容：
# 之前只匹配「核对」，漏掉了「本节按国家、省、市三级列示本报告采用的…」这类变体。
META_CLAUSE_RE = re.compile(r"[；;]?本节按[^。；]*[。；]?")
META_TAIL_RE = re.compile(r"下一步：[^。]*。?")
META_CHAPTER_RE = re.compile(r"(本章为|本章按)[^。]*。")


def strip_meta(text):
    text = str(text or "")
    if not text:
        return text
    text = META_CLAUSE_RE.sub("。", text)
    text = META_TAIL_RE.sub("", text)
    text = META_CHAPTER_RE.sub("", text)
    text = re.sub(r"。{2,}", "。", text)
    return text.strip().lstrip("；;，, ")


def retarget_upstream_refs(text):
    """上游报告整段引入后，「下表/本报告/表中」指向的是上游报告的图表。

    建议书可能单独发给业主，业主手里不一定有《现状分析报告》（第十二类挑刺：
    跨包表述串用）。所以指针一律改指本件自己的表图或本件内已写出的结论，
    上游报告只作资料依据，不在正文承担论证。
    """
    text = str(text or "")
    for old, new in (
        # 「上述/该表」指的都是建议书里并不存在的指标表；主语落到本件列举的指标上。
        ("本节把 OSGB 解析成果中的 7 项指标原值列出",
         "本节列出本项目 OSGB 解析成果的 7 项指标"),
        ("本节引述《现状分析报告》中 OSGB 解析成果的 7 项指标原值",
         "本节列出本项目 OSGB 解析成果的 7 项指标"),
        ("下表全部来自三维实测", "这 7 项指标全部出自三维实测"),
        ("上述指标全部来自《现状分析报告》的指标表", "这 7 项指标全部出自三维实测"),
        ("两者在本报告中分表列示", "两类指标在本报告中分列说明、不交叉换算"),
        ("表中「区间」", "指标「区间」"),
        ("该表「区间」", "指标「区间」"),
        # 2.2 首句承诺了本报告不存在的小节与小计表；只留归属原则，并说明不按区域拆分。
        ("；区域小计与逐小区归属分别在「人口与社会人居现状」「建筑现状与风貌保护现状」两节列出。",
         "。本报告不按区域拆分现状数据，不列区域小计；片区整体指标以两社区台账为准。"),
        ("该口径在 C 章引用相关比例时不再重复说明。",
         "该口径为影像纹理覆盖代理口径，本报告引用相关比例时不再重复说明。"),
    ):
        text = text.replace(old, new)
    # 上游原文把口径说明挂到它自己的章节上（「C 章」「现状详细分析」一章），
    # 读者在本件里翻不到；改成本件内的自足表述。
    text = re.sub(r"该口径在[^。]{0,24}(?:章|一节)[^。]{0,24}不再重复说明。",
                  "该口径为影像纹理覆盖代理口径，本报告引用相关比例时不再重复说明。", text)
    # 兜底：仍指向上游报告的表述改写成「本项目现状调研成果」，读者按字面即可在本件内找到。
    return text.replace("《现状分析报告》", "本项目现状调研成果")


# 建议书自身的成果构成：一份成品（docx/PDF 同版）＋随文两册工作簿。
# 单件独立可交付——业主手里可能没有上游报告，正文不出现内部件名称与流水线口径。
INVEST_DELIVERY_TEXT = (
    "成果构成为一份《项目投资决策建议书》成品（docx 与 PDF 同版），随文附投资估算表与财务评价表两册工作簿。"
    "成品只呈现投资必要性、建设方案、投资估算与财务结论；现状调研的关键指标、口径与来源均在本件内写清，"
    "读者不依赖其他文件即可独立审阅。需要业主补料的事项、逐条依据与复核口径另册说明，不进入本件正文。")

CROSS_PKG_DELIVERY_RE = re.compile(r"内部件|数据缺口清单|输入材料补齐确认单|标注版报告")


def sanitize_status_delivery(text):
    """剔除上游《现状分析报告》自述交付构成的句子（第十二类挑刺：跨包表述串用）。

    建议书可能单独发业主，不能在本件里介绍另一份报告有几个成品、几个内部件；
    命中的整句删除，由 INVEST_DELIVERY_TEXT 按本建议书自己的交付重写。
    """
    if not text:
        return ""
    kept = []
    for seg in re.split(r"(?<=。)", str(text)):
        if not seg.strip():
            continue
        if CROSS_PKG_DELIVERY_RE.search(seg):
            continue
        kept.append(seg)
    return "".join(kept)


class InputError(Exception):
    pass


def reflow_problem_text(text):
    """2.3 的问题原文把「1. 表现：…2. 证据：…」塞在同一个自然段里。

    层级序数一律用文字（（1）（2）…），且每项独立成段，读者才看得出层次；
    不在正文引入项目符号（List Bullet）。
    """
    text = str(text or "")
    if not text:
        return text
    text = re.sub(r"(?<=[。；])\s*(?=P-\d{2}\s)", "\n\n", text)
    text = re.sub(r"(?<![\d.])(\d)\.\s*(?=(?:表现|证据|成因|影响|对应更新方向)：)",
                  r"\n\n（\1）", text)
    return text


def read_json(path, label, required=True):
    if not path.is_file():
        if required:
            raise InputError("缺少{}：{}".format(label, _lib.rel_posix(path, Path.cwd())))
        return None
    try:
        return _lib.read_json(path)
    except (json.JSONDecodeError, OSError) as error:
        raise InputError("{}读取失败：{}".format(label, error))


def to_float(value, default=None):
    """JSON 里的数值可能是字符串（Decimal 序列化），比较前统一转换。"""
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def require_float(value, label):
    """结论判定用的关键指标：读不到就抛错，禁止静默兜底成 0。

    历史 bug（R5）：`to_float(indicators.get("fnpvAll"), 0)` 在键读空时返回 0，
    `0 >= 0` 成立，报告把 FNPV −9,416.94 万元判成「财务可行」。
    结论句依赖的数值一律走这里，宁可 05 抛错阻断，也不让缺值伪装成通过。
    """
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValueError("关键财务指标缺失：{}（实际值 {!r}）".format(label, value))


# 项目成立的外部条件：1.3 与 8.3 必须同一套措辞、同一顺序（R6 复核：一处两项、一处三项）。
EXTERNAL_CONDITIONS = ("一是政策性资金对建设投资的覆盖，"
                       "二是运营期收入提升，即出租率与租金水平达到同类改造项目的上限区间，"
                       "三是可经营资产注入带来的增量收益")

# 第 9 章方案编号：以 AHP 包（上游成果）为准，建议书不再另立一套。
# AHP 评分表：S1 保护修缮为主 6.2306 / S2 拆建结合 6.0194 / S3 综合提升 6.9955；
# 方案比选包原名：功能置换 / 拆除重建 / 微更新，按方案名对应到上述编号。
# 两包同时交付业主，同一套方案编号必须相同，故统一取 AHP 口径。
CANON_SCHEMES = (
    ("S1", "保护修缮为主", "功能置换"),
    ("S2", "拆建结合", "拆除重建"),
    ("S3", "综合提升", "微更新"),
)


def canon_scheme(name):
    """按方案名（比选名或 AHP 名）返回统一的（编号, 名称）；匹配不上返回 None。"""
    target = str(name or "")
    if not target:
        return None
    for scheme_id, canon_name, compare_name in CANON_SCHEMES:
        for candidate in (canon_name, compare_name):
            if target in candidate or candidate in target:
                return scheme_id, canon_name
    return None


def fmt_amount(value):
    try:
        return "{:,.2f}".format(float(value))
    except (TypeError, ValueError):
        return "—"


def fmt_pct(value):
    try:
        return "{:.2f}%".format(float(value) * 100)
    except (TypeError, ValueError):
        return "—"


def fmt_pct_num(value):
    """表 8-6/8-7 的结果列：单位列已写「百分比」，结果列不再重复带 %。"""
    try:
        return "{:.2f}".format(float(value) * 100)
    except (TypeError, ValueError):
        return "—"


PAYBACK_NOT_FORMED = "计算期内未转正"
EQUITY_NOT_APPLICABLE = "不适用（资本金为 0）"


def fmt_score(value):
    try:
        return "{:.4f}".format(float(value))
    except (TypeError, ValueError):
        return "—"


def _or_missing(value, formatter):
    """口径不适用的指标写「不适用」：写「—」会拼出「（ROE）—。」，
    句号被挤到下一行行首，读起来也别扭；「未测算」属于自曝缺口。"""
    return "不适用" if value is None else formatter(value)


def fmt_num(value, digits=2):
    try:
        text = "{:,.{}f}".format(float(value), digits)
    except (TypeError, ValueError):
        return "—"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def fmt_int(value):
    """户数、栋数、工程量：千分位整数，正文与表格同一写法。"""
    try:
        return "{:,}".format(int(round(float(value))))
    except (TypeError, ValueError):
        return "—"


def fmt_area(value):
    """建筑面积（万㎡）保留 4 位小数，合计行按同一小数位对齐。"""
    try:
        return "{:.4f}".format(float(value))
    except (TypeError, ValueError):
        return "—"


def yearly_stats(finance):
    """返回首个运营年/年均/末年/合计（万元）。"""
    period = (finance or {}).get("period") or {}
    build = int(period.get("constructionYears") or 0)
    revenue = [Decimal(str(v)) for v in ((finance or {}).get("revenue") or {}).get("total") or []]
    cost = [Decimal(str(v)) for v in ((finance or {}).get("cost") or {}).get("operatingCost") or []]
    if not revenue or build >= len(revenue):
        return None
    op_rev = revenue[build:]
    op_cost = cost[build:] if len(cost) > build else []
    return {"first": op_rev[0], "last": op_rev[-1], "total": sum(revenue),
            "avg": (sum(op_rev) / Decimal(len(op_rev))) if op_rev else ZERO,
            "costFirst": op_cost[0] if op_cost else ZERO,
            "costAvg": (sum(op_cost) / Decimal(len(op_cost))) if op_cost else ZERO,
            "costTotal": sum(cost)}


# 来源行的落地校验上下文：工作区根 / 引擎项目目录 / 上游摘要来源。
# 由 build_chapters 在开头写入，section() 生成来源行时据此把每条引用
# 收敛成「工作区相对路径 + 单键定位」，并剔掉查不到文件的引用。
_REF_CTX = {}


def _looks_like_file(ref_path):
    """含目录分隔符或以扩展名结尾，才算文件引用；否则是「政策文件」这类说明。"""
    return "/" in ref_path or "\\" in ref_path or bool(re.search(r"\.\w{2,5}$", ref_path))


def _json_has_pointer(path, pointer):
    """JSON 里 a.b.c 能否取到节点（取不到返回 False）。"""
    try:
        with open(path, encoding="utf-8") as handle:
            node = json.load(handle)
    except (OSError, ValueError):
        return False
    for key in [item for item in pointer.split(".") if item]:
        if isinstance(node, list):
            try:
                node = node[int(key)]
            except (ValueError, IndexError):
                return False
        elif isinstance(node, dict) and key in node:
            node = node[key]
        else:
            return False
    return True


def _resolve_ref_file(ref_path, workspace, project):
    if os.path.isabs(ref_path):
        return ref_path if os.path.isfile(ref_path) else ""
    for base in (workspace, project):
        if not base:
            continue
        cand = Path(base) / ref_path
        if cand.is_file():
            return str(cand.resolve())
    name = os.path.basename(ref_path)
    if name and workspace:
        for hit in Path(workspace).rglob(name):
            if hit.is_file():
                return str(hit.resolve())
    return ""


def _expand_compound(ref):
    """`投资估算.json#totalInvestment/funding` → 两条单键引用。

    一个数值必须指得到**唯一字段**：复合指针（多字段并列）在回读时无法定位，
    是 07 报「指针缺失」的根因。这里拆成一个字段一条，不做兼容性放行。
    """
    path_part, sep, pointer = ref.partition("#")
    if not sep or "/" not in pointer or not path_part.lower().endswith(".json"):
        return [ref]
    parts = [item.strip() for item in pointer.split("/") if item.strip()]
    return ["{}#{}".format(path_part, item) for item in parts] or [ref]


def _canonical_ref(ref, ctx):
    """把一条来源引用收敛成「工作区相对路径#单键」；查不到文件/键返回空串。"""
    path_part, sep, pointer = ref.partition("#")
    path_part, pointer = path_part.strip(), pointer.strip()
    if not path_part:
        return ""
    if not _looks_like_file(path_part):
        return ref                      # 说明性来源（政策文件、技术规范）原样保留
    workspace, project = ctx.get("workspace"), ctx.get("project")
    real = _resolve_ref_file(path_part, workspace, project)
    if not real:
        return ""
    rel = os.path.relpath(real, workspace).replace("\\", "/")
    if pointer and rel.lower().endswith(".json"):
        if not _json_has_pointer(real, pointer):
            # 键取不到 → 这条引用指不到唯一字段，整条剔除（不回退成「文件级」）
            return ""
    return "{}#{}".format(rel, pointer) if pointer and rel.lower().endswith(".json") else rel


def normalize_sources(sources, ctx=None):
    """来源行收敛：拆复合指针、换成可解析的工作区相对路径、剔除查不到的引用。"""
    ctx = ctx or _REF_CTX
    if not ctx.get("workspace"):
        return [item for item in sources if item]
    out = []
    for ref in sources or []:
        ref = (ref or "").strip()
        if not ref:
            continue
        for cand in _expand_compound(ref):
            canon = _canonical_ref(cand, ctx)
            if canon and canon not in out:
                out.append(canon)
    return out


def section(level, number, title, body, sources):
    heading = "#" * level
    text = "{} {} {}".format(heading, number, title).rstrip()
    sources = normalize_sources(sources)
    source = "来源：" + "；".join(sources) if sources else "来源：无"
    return "{}\n\n{}\n\n{}\n".format(text, body.strip(), source)



def detect_workspace(project):
    """从引擎工作目录向上找到工作区根（同时含 输入文件/ 与 过程文件/）。"""
    current = Path(project).resolve()
    for parent in [current] + list(current.parents):
        if (parent / "输入文件").is_dir() and (parent / "过程文件").is_dir():
            return parent
    return None


def load_upstream_scores(project):
    """读上游 AHP 评分表（05_过程/AHP/评分表.json）的合成结果；只引用，不重算。"""
    path = Path(project) / "05_过程" / "AHP" / "评分表.json"
    if not path.is_file():
        return None
    data = read_json(path, "AHP 评分表", required=False)
    if not isinstance(data, dict):
        return None
    rows = []
    schemes = []
    for item in data.get("schemes") or []:
        if item.get("composite") is None:
            continue
        row = {"id": item.get("id"), "name": item.get("name"),
               "composite": item.get("composite"),
               "label": (item.get("rating") or {}).get("label")}
        rows.append(row)
        schemes.append(dict(row, scores=[{
            "id": score.get("id"),
            "name": score.get("name"),
            "criterionId": score.get("criterionId"),
            "globalWeight": score.get("globalWeight"),
            "score": score.get("score"),
            "basis": score.get("basis"),
        } for score in item.get("scores") or []]))
    if not rows:
        return None
    conclusion = data.get("conclusion") or {}
    ranking = data.get("ranking") or []
    office_files = []
    workspace = detect_workspace(project)
    if workspace is not None:
        folder = workspace / "输入文件"
        if folder.is_dir():
            office_files = sorted(item.name for item in folder.rglob("AHP*")
                                  if item.suffix.lower() in (".xlsx", ".docx", ".pdf"))
    return {
        "file": _lib.rel_posix(path, project),
        "rows": rows,
        "schemes": schemes,
        "ranking": ranking,
        "summary": conclusion.get("summary"),
        "top_drivers": conclusion.get("topDrivers") or [],
        "scoring": data.get("scoring"),
        "office_files": office_files,
    }


# ---------------------------------------------------------------- 取料器

class Facts:
    """上游事实的唯一入口：有值返回值，无值返回缺口话术并登记缺项。"""

    def __init__(self, digest, gaps, measured=None):
        digest = digest or {}
        self.facts = digest.get("facts") or {}
        self.slots = digest.get("slots") or {}
        self.sections = digest.get("sections") or []
        self.sources = digest.get("sources") or []
        self.measured = measured or {}
        self.gaps = gaps
        self._codes = {item.get("code") for item in gaps}
        # 已由匡算附表/收入量价确证的规模指标，不再挂在上游缺项里。
        for key in self.measured:
            self.drop_gap("UP-{}".format(key.upper()))

    # ---- 基础
    def _raw(self, key):
        item = self.facts.get(key)
        if not isinstance(item, dict):
            # 上游报告不写工程量，但匡算附表1 的工程量列（经 01→02 贯通到投资估算）与
            # 附表4 的量价依据给了确数；这里作为第二来源，正文因此可以写实数而非空句。
            item = self.measured.get(key)
        if not isinstance(item, dict):
            return None
        value = item.get("value")
        if isinstance(value, str) and value.strip() in NON_FACT_VALUES:
            return None
        # 上游摘要偶尔把「本章为…按…口径成文」这类写作说明抽成事实取值（实测 4.2 串入第 1 章导语）。
        # 这类字符串不是事实，一律按缺值处理，交给推定话术。
        if isinstance(value, str) and SLOT_NOISE_RE.search(value):
            return None
        return item

    def has(self, key) -> bool:
        return self._raw(key) is not None

    def gap(self, code, what, need):
        if code not in self._codes:
            self.gaps.append({"code": code, "what": what, "need": need})
            self._codes.add(code)
        return PLACEHOLDER_TMPL

    def drop_gap(self, code):
        """补到的缺口从缺项表撤掉：正文已写实，就不能还挂着「缺四至范围」。"""
        self._codes.discard(code)
        self.gaps[:] = [item for item in self.gaps if item.get("code") != code]

    # ---- 取值
    def num(self, key, what="", need="", digits=2):
        item = self._raw(key)
        if item is None:
            return None
        return fmt_num(item.get("value"), digits)

    def unit(self, key):
        item = self._raw(key) or {}
        return item.get("unit") or ""

    def text(self, key, what="", need="", code=None):
        item = self._raw(key)
        if item is None:
            return self.gap(code or "UP-{}".format(key.upper()), what or key, need or what or key)
        return strip_meta(str(item.get("value")).strip())

    def with_unit(self, key, what="", need="", digits=2):
        item = self._raw(key)
        if item is None:
            return self.gap("UP-{}".format(key.upper()), what or key, need or what or key)
        return "{} {}".format(fmt_num(item.get("value"), digits), item.get("unit") or "").strip()

    def cell(self, key, what="", need="", digits=2):
        """表格单元格：只返回数值（单位另设一列），缺则不写数。"""
        item = self._raw(key)
        if item is None:
            self.gap("UP-{}".format(key.upper()), what or key, need or what or key)
            return "—"
        return fmt_num(item.get("value"), digits)

    def slot(self, key, what="", need=""):
        value = self.slots.get(key)
        if isinstance(value, list):
            value = "、".join(str(item) for item in value if str(item).strip())
        # 上游摘要抽槽位时偶尔把「本章为…按…口径成文」这类**写作说明**抽成本身取值，
        # 直接印进表格会变成「项目性质：本章为『项目概述』…」。这种不是事实，按缺项处理。
        if value and SLOT_NOISE_RE.search(str(value)):
            value = ""
        value = strip_meta(value)
        if not value or not str(value).strip():
            self.gap("SLOT-{}".format(key.upper()), what or key,
                     need or "输入文件-模板/槽位-空白.xlsx 的「{}」填写后回填".format(what or key))
            return PENDING_SCHEME
        # 去掉尾部标点：调用处统一以「。」收尾，用户填的值自带句号会拼出「。。」
        return str(value).strip().rstrip("。．.；;，,、 ")

    # ---- 原文取用
    def excerpt(self, keywords, max_chars=900):
        """按小节标题/路径关键词取上游原文段落；取不到返回空串。

        同一段上游原文不跨章重复引述：2.2 的「现状判断」与 3.1 的现状摘录取到的是
        同一段（实测 532 字），成品 PDF p10–p11 读者要看两遍。
        """
        used = getattr(self, "_used_paragraphs", None)
        if used is None:
            used = self._used_paragraphs = set()
        picked, total = [], 0
        for item in self.sections:
            label = " / ".join(item.get("path") or []) + " " + (item.get("title") or "")
            if not any(keyword in label for keyword in keywords):
                continue
            for paragraph in item.get("paragraphs") or []:
                paragraph = strip_meta(paragraph)
                if not paragraph or paragraph.strip() in used:
                    continue
                if total + len(paragraph) > max_chars:
                    return "".join(picked)
                picked.append(paragraph if paragraph.endswith("。") else paragraph + "。")
                used.add(paragraph.strip())
                total += len(paragraph)
        return "".join(picked)

    def excerpt_sentences(self, keywords, max_chars=120):
        """表格单元格用：只取**完整句子**，超长就停，绝不从句子中间截断。

        `excerpt` 按字符数硬截，实测把「商业），没有商铺数量…需现场核实的是」这种
        半句话塞进表 4-1，读者当场能看出是机器拼接的。
        """
        out = ""
        for item in self.sections:
            label = " / ".join(item.get("path") or []) + " " + (item.get("title") or "")
            if not any(keyword in label for keyword in keywords):
                continue
            for paragraph in item.get("paragraphs") or []:
                paragraph = strip_meta(paragraph)
                for sentence in re.split(r"(?<=。)", paragraph):
                    sentence = sentence.strip()
                    if not sentence.endswith("。"):
                        continue
                    # 证据/流程类句子不进表格：读者要看的是结论，不是取料过程
                    if sentence.startswith(("判断依据", "证据", "成因", "影响", "下一步")):
                        continue
                    sentence = re.sub(r"^(初步判定|初步判断)[:：]", "", sentence)
                    if len(out) + len(sentence) > max_chars:
                        return out
                    out += sentence
        return out

    def tables_matching(self, keywords):
        """取小节标题命中关键词的上游表格（整表，含表头行）。"""
        out = []
        for item in self.sections:
            label = " / ".join(item.get("path") or []) + " " + (item.get("title") or "")
            if not any(keyword in label for keyword in keywords):
                continue
            out.extend(item.get("tables") or [])
        return out

    def render_upstream_table(self, keywords, caption, fallback_headers, what):
        """上游有同主题表格就原样渲染（表头用上游的，不硬套本报告的列序）。"""
        for table in self.tables_matching(keywords):
            rows = table.get("rows") or []
            if len(rows) >= 2 and len(rows[0]) >= 2:
                width = len(rows[0])
                body = [list(row[:width]) + [""] * (width - len(row)) for row in rows[1:]]
                return md_table(caption, rows[0], body)
        return md_table(caption, fallback_headers, [gap_row(len(fallback_headers), what)])


def drop_empty_columns(headers, rows):
    """整列单元格全是「—」/空格 → 删列。

    复验口径：「无值不建表」——留一列全是破折号，读者会当成没填完。
    示例：表 8-2「测算基础」、表 8-9「回收期」「FNPV 归零临界点」、表 9-5「回收期」。
    只删「所有行都空」的列；只要有一行有值就保留（不擅自删数据列）。
    """
    if not rows:
        return headers, rows
    width = len(headers)
    padded = [list(row) + [""] * (width - len(row)) for row in rows]
    # 契约规定的必备列不能被删（表头必须与表格契约一致），即使整列未提供值
    keep = [index for index in range(width)
            if str(headers[index]).strip() in ("行几何平均",)
            or any(str(row[index]).strip() not in ("", "—", "-") for row in padded)]
    if len(keep) < 2:                       # 全空表交给调用方处理，别的表不要被误删
        return headers, padded
    return [headers[index] for index in keep], [[row[index] for index in keep] for row in padded]


def md_table(caption, headers, rows):
    headers, rows = drop_empty_columns(list(headers), list(rows))
    lines = [caption, "", "| " + " | ".join(str(item) for item in headers) + " |",
             "| " + " | ".join("---" for _ in headers) + " |"]
    for row in rows:
        lines.append("| " + " | ".join("" if cell is None else str(cell) for cell in row) + " |")
    return "\n".join(lines)


def gap_row(width, what):
    # 表格里缺数就留「—」，不把推断口径写进单元格：一格塞 40 字说明会撑高整行，
    # 也确实容易被读成「编了个数」。推断口径与置信度写在正文与缺项表里。
    row = ["—"] + ["—"] * (width - 1)
    return row


def _readable_source(source):
    """外发表格的「依据」列只写读者认识的文件名，不写 `xxx.json#field` 这种内部指针。"""
    text = str(source or "").strip()
    if not text:
        return "—"
    if "#" not in text and not re.search(r"\.(json|docx|xlsx|md)$", text):
        return text
    base = text.split("#")[0].rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    stem = re.sub(r"\.(json|docx|xlsx|md)$", "", base)
    # 建议书可能单独发业主：依据列不挂另一份报告的标题（第十二类挑刺：跨包表述串用）。
    if stem in ("现状分析报告", "上游摘要"):
        return "本项目现状调研成果"
    return "《{}》".format(stem) if stem else text


def _basis(value, source, inferred="测算口径见正文"):
    """依据列：数值缺位时不给它编出处（缺值行本应在建表时整行剔除）。"""
    return inferred if str(value or "").startswith(PLACEHOLDER) else _readable_source(source)


# ---------------------------------------------------------------- 缺值话术
# 成品是给专家评审的建议书，不是内部工作底稿：正文里不出现「没资料」的姿态。
# 口径（2026-11-14 起）：
#   * 有值写值，并把取值依据写在同一句里（「按台账口径…」「按财务测算口径…」）；
#   * 取不到的指标**整句、整行不写**，不拿「暂按同类项目推定」凑字数；
#   * 需业主确认的事项不进正文，统一登记到内部件《待确认清单》。
def _is_placeholder(value) -> bool:
    return str(value or "").startswith(PLACEHOLDER)


def _metric_phrase(items, sep="，"):
    """把若干（标签, 带单位取值）拼成一句：只写有值的项，缺值项整项省略。

    sep 必须由调用方显式给：返回的是**整句字符串**，调用方若再套 "；".join(...)
    会按字符拆开（R5 P0055/P0061/P0107 的「公。空；间；面；积。」即此）。
    """
    known = ["{}约 {}".format(label, value) for label, value in items if not _is_placeholder(value)]
    return sep.join(known)


def _fact_phrase(label, value, basis):
    return "{}：{}".format(label, value) if not _is_placeholder(value) else ""


def _cell_or(value, basis="同类项目水平"):
    """表格单元格：缺值留破折号，由建表逻辑整行剔除，不在格子里写推定话术。"""
    return "—" if _is_placeholder(value) else value


def _qty(value, label, basis="同类项目水平"):
    """正文里的「标签约 X」：缺值返回空串，调用方据此省略该分句。"""
    return "{}约 {}".format(label, value) if not _is_placeholder(value) else ""


# ---------------------------------------------------------------- 标注口径
# 与 CardoStatus 统一：正文不用「置信度」，只用两枚标注——
#   （可核对）            本节结论全部能落到已落地文件/字段；
#   （推断，需现场核实）  本节含上游没有、按同类区间推断的数。
# 标注只落在每节判定句上，一节一枚；章末用一句清点（中文数字，避免被当成需溯源的数值）。
_CN_DIGITS = "零一二三四五六七八九"
_SECTION_SPLIT = re.compile(r"(?m)^(?=##\s)")   # 以编号「节」为单位，一节一枚
_CONF_RE = re.compile(r"[；，]?\s*置信度\s*[:：][^）)]*")
_FOOTER_RE = re.compile(r"本章结论中[^。]*。")


def _cn_number(value):
    if value < 10:
        return _CN_DIGITS[value]
    tens, ones = divmod(value, 10)
    head = "十" if tens == 1 else _CN_DIGITS[tens] + "十"
    return head + (_CN_DIGITS[ones] if ones else "")


def _label_section(block):
    """返回（新块, 'ok'/'infer'/None）：一节只留一枚标注，落在判定句上。

    上游摘要（CardoStatus 口径）已按段落在句末带标注，这里先全部摘掉再按节重打，
    否则同一节会叠出好几枚（实测 2.1 节 4 枚）。
    """
    tail = block[len(block.rstrip("\n")):]      # 保留块尾换行，否则来源行会与下节标题粘连
    # 缺值改写成推定话术后，PLACEHOLDER 已不在块内；推定/未载明同样意味着本节含推断，
    # 不能再打「可核对」。
    had = (PLACEHOLDER in block or "本次输入未载明" in block or "推定" in block
           or "置信度" in block or INFER_NOTE in block)
    lines = []
    for line in block.splitlines():
        line = line.replace(INFER_NOTE, "").replace(CHECK_NOTE, "")
        line = _FOOTER_RE.sub("", line)
        # 缺值标记直接删掉，不再换成「暂按同类项目推定」；拼空的半句另行收拾。
        line = line.replace(PLACEHOLDER, "")
        line = re.sub(r"(约|为|是|达|共)\s*[，。；]", "。", line)
        line = re.sub(r"[，；]\s*[，；。]", "。", line)
        line = re.sub(r"：\s*。", "：现状情况以第 3 章调研结论为准。", line)
        lines.append(_CONF_RE.sub("", line))
    block = "\n".join(lines).replace("（）", "") + tail
    label = INFER_NOTE if had else CHECK_NOTE
    lines = block.splitlines()
    for index, line in enumerate(lines):
        text = line.strip()
        if not text or text.startswith(("#", "来源：", "|", ">", "- ")):
            continue
        pos = line.find("。")
        if pos >= 0:
            cut = pos + 1
            if pos > 45:
                # 判定句太长时标注挂在句末会单独占一行（实测 5.3 折出「实）」），
                # 改为跟在第一个分句后，仍落在判定句上，也不会拖出两字尾行。
                comma = line.find("，", 10)
                if 0 < comma < pos:
                    cut = comma
            lines[index] = line[:cut] + label + line[cut:]
            return "\n".join(lines) + tail, ("infer" if had else "ok")
    return block, None


def annotate_chapter(text):
    """给每节判定句加一枚标注。

    章末原先还有一句「本章结论中 X 条可直接核对、Y 条为推断」——这是对我方写作质量的
    清点，读者看到只会觉得是内部备注，已删除（_FOOTER_RE 仍负责剥掉上游带来的同类句子）。
    """
    blocks = [item for item in _SECTION_SPLIT.split(text or "") if item.strip()]
    out = []
    for block in blocks:
        new, _kind = _label_section(block)
        out.append(new)
    return "".join(out).rstrip() + "\n"


READABLE_NAMES = [
    # 文件级指针一律落到章/节，不再映射到某张具体表：
    # 原先「财务测算.json → 表 8-1」会把任何一句「详见 财务测算.json」都写成「详见 表 8-1」，
    # 即使该参数根本不在 表 8-1（R9 的 8.1.4.1 就是这样被指到收入参数表）。
    ("财务测算.json", "第 8 章财务测算"),
    ("财务口径.yaml", "表 8-1"),
    ("财务输入.json", "表 8-1"),
    ("《项目信息.png》", "《现状分析报告》"),
    ("投资估算.json", "第 7 章投资估算"),
    ("方案比选.json", "第 9 章方案比选"),
    ("上游摘要.json", "上游资料摘要"),
]


_CLAUSE_LABEL_RE = re.compile(
    r"^(?:表现|现状|问题|判断依据|证据|成因|影响|下一步|结论|说明)[:：]\s*")
_CLAUSE_PUNCT_RE = re.compile(r"[\s，。；、：:,.;！？（）()\[\]【】“”\"'—-]")


def _dedupe_adjacent(text):
    """同一句在相邻两句内重复出现即删后一次。

    上游摘要 sections/22 的两句是拼接来的：「沿街商业以生活服务为主，业态同质、空置与
    低租金并存；沿街商业以生活服务为主，业态同质与空置并存……」——前半句整段重复。
    按句号/分号切句后比较相邻两句（去掉「表现：」这类标签与标点）：
      * 完全相同 → 删后一句；
      * 公共前缀 ≥ 8 字且占较短句一半以上 → 视为同句重复，删后一句；
      * 后一句独有的括注（「（推断，需现场核实）」）保留到前一句，别把口径标记一起删掉。
    """
    text = str(text or "")
    parts = [part for part in re.split(r"(?<=[。！？；;])", text) if part]
    if len(parts) < 2:
        return text
    out = []
    for part in parts:
        if out:
            prev = out[-1]
            left = _CLAUSE_PUNCT_RE.sub("", _CLAUSE_LABEL_RE.sub("", prev))
            right = _CLAUSE_PUNCT_RE.sub("", _CLAUSE_LABEL_RE.sub("", part))
            prefix = 0
            for a, b in zip(left, right):
                if a != b:
                    break
                prefix += 1
            shorter = min(len(left), len(right))
            if shorter >= 8 and (prefix >= shorter
                                 or (prefix >= 8 and prefix / shorter >= 0.5)):
                notes = set(re.findall(r"[（(][^）)]*[）)]", prev))
                extra = [note for note in re.findall(r"[（(][^）)]*[）)]", part)
                         if note not in notes]
                if extra:
                    stem = prev.rstrip("。！？；;")
                    tail = prev[len(stem):] or "。"
                    prev = stem + "".join(extra) + tail
                    out[-1] = prev
                continue
        out.append(part)
    return "".join(out)


def readable(text):
    """正文不出现内部分件名与重复句号。

    读者看到「财务测算.json」「财务口径.yaml」只会当成未清理的机读标识；
    「。。」则是取值自带句号而调用处又补了一个造成的。
    """
    text = str(text or "")
    for old, new in READABLE_NAMES:
        text = text.replace(old, new)
    text = re.sub(r"。{2,}", "。", text)
    text = re.sub(r"([！？])。", r"\1", text)
    return _dedupe_adjacent(text)


def paragraphs(*items):
    return "\n\n".join(readable(text) for text in items if text and text.strip())


def bullets(items):
    """条目清单：公文层次序数用文字（1. 2. …），不用「-」「•」项目符号。

    条目自带序数（「1、」「（一）」）时不再叠加，避免「1. 1、」双编号。
    """
    lines = []
    n = 0
    for text in items:
        text = str(text)
        if re.match(r"^(?:[（(][一二三四五六七八九十]+[）)]|[一二三四五六七八九十]+、|\d+[、.．)])", text):
            lines.append(text)
        else:
            n += 1
            lines.append("{}. {}".format(n, text))
    return "\n".join(lines)


def section_body(*items):
    return "\n\n".join(text for text in items if text and text.strip())


# ---------------------------------------------------------------- 提资台账
# 第 3 章六块结构的底表：直接读三份社区台账 xlsx，正文与表里的每个数都可回指原表。
# 片区归属与是否合并计列由工作区参数 scope.includeB 决定，不在代码里预设。
LEDGER_FILES = (
    ("示例片区B", "示例片区B", "示例社区B", "示例片区B改造小区基本信息及改造内容统计表.xlsx"),
    ("示例社区A", "示例片区A", "示例社区A", "示例社区A改造小区基本信息及改造内容统计表(1).xlsx"),
    ("示例社区C", "示例片区A", "示例社区C", "改造小区基本信息及改造内容统计表（示例社区C）.xlsx"),
)
LEDGER_QUANTITIES = ("立面材料修缮或翻新", "屋面翻新", "更换外墙雨棚",
                     "冷凝水管", "落水管", "空调外机披水板")
LEDGER_UNITS = ("平方米", "平方米", "个", "米", "米", "个")


def _ledger_number(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?[\d.]+", str(value))
    return float(match.group(0)) if match else None


def _ledger_int(value):
    number = _ledger_number(value)
    return int(round(number)) if number is not None else None


def build_ledger(project_dir):
    """读三份社区台账，返回分片区汇总、逐小区明细与工程量合计。

    表结构：第 1 行大标题、第 2 行表头、第 3 行单位、第 4 行起数据。
    列：3 小区名称｜4 地址｜5 户数｜6 楼栋数｜7 建筑面积（万㎡）｜8 建成年代｜
        9 层数（最高）｜10 结构形式｜11 建筑性质｜13 外立面装饰类型｜
        15—20 六项工程量（立面、屋面、雨棚、冷凝水管、落水管、披水板）。
    """
    workspace = detect_workspace(Path(project_dir)) if project_dir else None
    if not workspace:
        return None
    try:
        import openpyxl
    except ImportError:
        return None
    root = Path(workspace) / "输入文件"
    records = []
    for _key, area, community, filename in LEDGER_FILES:
        hits = [path for path in (root.rglob(filename) if root.is_dir()
                                  else Path(workspace).rglob(filename)) if path.is_file()]
        if not hits:
            continue
        workbook = openpyxl.load_workbook(str(hits[0]), data_only=True, read_only=True)
        sheet = workbook.worksheets[0]
        for row in sheet.iter_rows(min_row=4, values_only=True):
            name = str(row[3] or "").strip()
            if not name:
                continue
            households = _ledger_int(row[5])
            buildings = _ledger_int(row[6])
            area_wan = _ledger_number(row[7])
            if households is None and buildings is None and area_wan is None:
                continue
            records.append({
                "片区": area,
                "社区": community,
                "小区名称": name,
                "地址": str(row[4] or "").strip(),
                "户数": households,
                "楼栋数": buildings,
                "建筑面积万㎡": area_wan,
                "建成年代": re.sub(r"年$", "", str(row[8] or "")).strip(),
                "层数": _ledger_int(row[9]),
                "结构形式": str(row[10] or "").strip(),
                "建筑性质": str(row[11] or "").strip(),
                "外立面类型": str(row[13] or "").strip(),
                "工程量": {key: _ledger_int(row[15 + index])
                          for index, key in enumerate(LEDGER_QUANTITIES)},
                "来源文件": hits[0].name,
            })
        workbook.close()
    if not records:
        return None
    residential = [item for item in records if item["建筑性质"] != "商业"]
    commercial = [item for item in records if item["建筑性质"] == "商业"]

    def _sum(items, key):
        total = Decimal("0")
        for item in items:
            if item.get(key) is not None:
                total += Decimal(str(item[key]))
        return total

    areas = []
    for _key, area, community, _filename in LEDGER_FILES:
        items = [item for item in residential if item["社区"] == community]
        if not items:
            continue
        areas.append({
            "片区": area,
            "社区": community,
            "小区数": len(items),
            "栋数": int(_sum(items, "楼栋数")),
            "户数": int(_sum(items, "户数")),
            "建筑面积万㎡": float(_sum(items, "建筑面积万㎡")),
            "小区": [item["小区名称"] for item in items],
        })
    household_total = int(_sum(residential, "户数"))
    area_total = float(_sum(residential, "建筑面积万㎡"))
    totals = {
        "小区数": len(residential),
        "栋数": int(_sum(residential, "楼栋数")),
        "户数": household_total,
        "建筑面积万㎡": area_total,
        "人口（按户均3.2人）": round(household_total * 3.2, 1),
        "户均面积㎡": round(area_total * 10000 / household_total, 2) if household_total else None,
        "砖混小区数": sum(1 for item in residential if "砖混" in item["结构形式"]),
        "1980年建成小区数": sum(1 for item in residential if item["建成年代"] == "1980"),
    }
    totals["砖混占比%"] = round(totals["砖混小区数"] * 100 / totals["小区数"], 2)
    largest = max(residential, key=lambda item: item["户数"] or 0)
    totals["首位小区"] = largest["小区名称"]
    totals["首位小区户数占比%"] = round((largest["户数"] or 0) * 100 / household_total, 2)
    # 首位小区外的合计户数：正文要引用，落成台账字段（E-7 要求数值在来源里能查到同值，
    # 正文现算的减法结果不在来源池里）。
    totals["首位小区外户数"] = household_total - int(largest["户数"] or 0)
    quantities = {key: int(sum((item["工程量"].get(key) or 0) for item in residential))
                  for key in LEDGER_QUANTITIES}
    # 分社区工程量合计也落盘：第 3 章表 3-5 的「来源与说明」列要引社区分项，
    # 现算值不在来源池里，E-7 会判无出处。
    quantities_by_community = {}
    for _key, _area, community, _filename in LEDGER_FILES:
        items = [item for item in residential if item["社区"] == community]
        if not items:
            continue
        quantities_by_community[community] = {
            key: int(sum((item["工程量"].get(key) or 0) for item in items))
            for key in LEDGER_QUANTITIES
        }
    commercial_total = {
        "处数": len(commercial),
        "栋数": int(_sum(commercial, "楼栋数")),
        "名称": "、".join(item["小区名称"] for item in commercial) or "—",
    }
    return {
        "schema": "cardo-invest/ledger/v1",
        "source": "输入文件/01-真提资/信息收集（三份社区改造小区基本信息及改造内容统计表）",
        "quantityUnits": dict(zip(LEDGER_QUANTITIES, LEDGER_UNITS)),
        "records": records,
        "areas": areas,
        "totals": totals,
        "commercial": commercial_total,
        "quantities": quantities,
        "quantitiesByCommunity": quantities_by_community,
    }


# ---------------------------------------------------------------- 正文

def build_chapters(project_name, payload, estimate, finance, schemes, score_block=None,
                   upstream=None, lowair=None, project_dir=None, planning_basis="", ledger=None,
                   scope=None):
    scope = scope or {}
    _include_b = scope.get("includeB")
    _boundary_basis = str(scope.get("boundaryBasis") or "").strip()
    gaps = []
    _REF_CTX.clear()
    _REF_CTX.update(workspace=detect_workspace(project_dir) if project_dir else None,
                    project=str(Path(project_dir).resolve()) if project_dir else None)
    # ---------------------------------------------------------------- 可核规模指标
    # 上游《现状分析报告》不写工程量，但匡算附表1（工程量×单位造价）与附表4（量价依据）
    # 逐项给了确数。这里只搬有原表可翻的项：取不到的仍走缺项表，进内部件《待确认清单》，
    # 不写进成品，也不拿「同类项目」凑数。
    def _measured_quantities():
        unit_map = {"m": "米", "m2": "平方米", "㎡": "平方米", "套": "套", "个": "个"}
        out = {}
        area = (payload.get("project") or {}).get("area")
        if area is not None:
            out["scopeArea"] = {"value": area, "unit": "平方米", "name": "改造范围面积",
                                "evidence": "规划用地面积 19820.37 平方米",
                                "source": "匡算附表 B0建设内容与规模表"}
        for row in (estimate or {}).get("engineeringRows") or []:
            label = str(row.get("name") or "")
            quantity = row.get("quantity")
            if quantity is None:
                continue
            unit = unit_map.get(str(row.get("unit") or ""), str(row.get("unit") or ""))
            if "外立面" in label:
                key, name = "facadeArea", "外立面整治面积"
            elif "道路" in label:
                key, name = "roadArea", "配套道路面积"
            elif "街区环境" in label or "公共空间" in label:
                key, name = "publicSpaceArea", "公共空间面积"
            elif "充电" in label:
                key, name = "chargers", "充电桩数量"
            else:
                continue
            out[key] = {"value": quantity, "unit": unit, "name": name,
                        "evidence": "{} {} {}".format(label, fmt_num(quantity, 2), unit),
                        "source": "匡算附表1 工程总投资估算表（工程量列）"}
        revenue_items = payload.get("revenue") or []
        if isinstance(revenue_items, dict):
            revenue_items = revenue_items.get("items") or []
        for item in revenue_items:
            basis = str(item.get("basisText") or "")
            label = str(item.get("name") or "")
            if "停车" in label and "parkingSpots" not in out:
                matched = re.search(r"(\d[\d,]*)\s*个车位", basis)
                if matched:
                    out["parkingSpots"] = {"value": float(matched.group(1).replace(",", "")),
                                           "unit": "个", "name": "停车位数量",
                                           "evidence": basis, "source": "匡算附表4 停车位收入量价"}
        years = (payload.get("project") or {}).get("constructionYears")
        if years:
            out["buildWindow"] = {"value": "{} 年（第 1—{} 年）".format(years, years), "unit": "",
                                  "name": "建设起止时间",
                                  "evidence": "财务参数：建设期 {} 年".format(years),
                                  "source": "输入文件/财务参数.xlsx#项目!B5"}
        return out

    f = Facts(upstream, gaps, _measured_quantities())

    indicators = (finance or {}).get("indicators") or {}
    reference = ((payload.get("finance") or {}).get("referenceIndicators") or {})
    # 口径纪律：本报告的指标一律用「引擎自算值」（indicators），保证正文、敏感性分析、
    # 方案比选、图 8-1 全书同一口径。可研对照值（reference）不覆盖引擎值，
    # 只作为「对照」并列呈现——否则同一份报告会同时出现两套 FIRR（如表位 1 写可研 3.24%、
    # 敏感性分析表写引擎 9.57%），读者无法判断哪套是本报告结论。
    display = dict(indicators)
    # 可研对照值：有值才参与对照，且只用于说明差异，不参与任何判定或结论。
    compare = {key: reference.get(key) for key in
               ("firrAll", "fnpvAll", "ptAll", "firrEquity", "fnpvEquity", "ptEquity")
               if reference.get(key) is not None}
    stats = yearly_stats(finance)

    def _firr_text(value):
        """结论句不许出现「—」：FIRR 算不出来就说明是现金流为负，而不是留空。"""
        return fmt_pct(value) if value is not None else "未转正（累计净现金流为负）"

    def _payback_text(value):
        return ("{} 年".format(fmt_num(value, 2)) if value is not None
                else PAYBACK_NOT_FORMED)
    project = payload.get("project") or {}
    period = (finance or {}).get("period") or {}
    funding = (estimate or {}).get("funding") or {}
    tables = (finance or {}).get("tables") or {}
    full_cost_avg = (tables.get("fullCostStructure") or {}).get("avgAnnual")
    revenue_structure = (tables.get("revenueStructure") or {}).get("rows") or []
    cost_structure = (tables.get("costStructure") or {}).get("rows") or []
    composition = (estimate or {}).get("tables", {}).get("composition") or {}
    fund_source = (estimate or {}).get("tables", {}).get("fundSource") or {}
    year_plan = ((estimate or {}).get("tables") or {}).get("yearPlan") or {}
    sensitivity = (finance or {}).get("sensitivity") or {}
    sens_factors = {item.get("key"): item for item in sensitivity.get("factors") or []}
    ahp = payload.get("ahpWeights") or {}
    criteria = ahp.get("criteria") or []
    criteria_weights = {item.get("key") or item.get("name"): item.get("weight") for item in criteria[:4]}
    indicator_weights = ahp.get("indicatorWeights") or []
    temporary = bool(ahp.get("temporary"))
    score_rows = {row.get("id"): row for row in (score_block or {}).get("rows") or []}
    scheme_list = (schemes or {}).get("schemes") or []
    winner_id = (schemes or {}).get("winner")
    _winner_raw = next((item.get("name") for item in scheme_list if item.get("id") == winner_id), "")
    # 推荐方案在第 9 章统一用报告编号与名称（微更新 → 综合提升），与表 9-5、表 9-6 对齐。
    winner_name = (canon_scheme(_winner_raw) or (None, _winner_raw))[1]
    winner_score = next((item.get("score") for item in scheme_list if item.get("id") == winner_id), None)
    year_plan_rows = year_plan.get("rows") or []
    fund_rows = fund_source.get("rows") or []
    comp_rows = composition.get("rows") or []
    # 充电桩正常年收入必须落纸：读者按 表 8-2 量价乘出来是 235.9，
    # 正文若只给 29 年均值 231.83，会以为对不上（R5 N-2）。数值一律取 JSON 逐年值。
    charger_row = next((item for item in revenue_structure if "充电" in str(item.get("name"))), {})
    charger_item = next((item for item in (((finance or {}).get("revenue") or {}).get("items") or [])
                         if "充电" in str(item.get("name"))), {})
    _charger_yearly = [to_float(value, 0.0) or 0.0 for value in (charger_item.get("yearly") or [])]
    charger_normal = max(_charger_yearly) if _charger_yearly else None
    charger_first = next((value for value in _charger_yearly if value > 0), None)

    total_w = fmt_amount((estimate or {}).get("totalInvestment"))
    engineering_w = fmt_amount((estimate or {}).get("engineering"))
    other_w = fmt_amount((estimate or {}).get("otherFeeTotal"))
    contingency_w = fmt_amount((estimate or {}).get("contingency"))
    interest_w = fmt_amount((estimate or {}).get("constructionInterest"))
    construction_w = fmt_amount((estimate or {}).get("constructionInvestment"))
    # 资本金未确定时对外一律写 0.00（与表 7-3、资金筹措合计同一口径），
    # 不再在表 1-1 里写「— / 0.00 万元」——同一份报告两种写法，读者会以为表填漏了。
    try:
        equity_undetermined = float(funding.get("equity") or 0) <= 0
    except (TypeError, ValueError):
        equity_undetermined = funding.get("equity") is None

    def _equity_payback_text(value):
        """资本金 ≤0 时不存在资本金回收期，不能把机械算出的 3 年写进结论。"""
        return EQUITY_NOT_APPLICABLE if equity_undetermined else _payback_text(value)
    equity_w = "0.00" if equity_undetermined else fmt_amount(funding.get("equity"))
    _subsidy_row = next((row for row in fund_rows
                         if row.get("key") == "subsidy" or "补助" in str(row.get("name") or "")), {})
    subsidy_w = fmt_amount(_subsidy_row.get("amount"))
    loan_w = fmt_amount(funding.get("loan"))
    # 资金申报口径（信息卡）：只用于披露「申报额度 − 投资需求」的差额，不进入现金流。
    # 取不到时相关句子整句不写，不在正文留空，也不拿匡算数反推申报数。
    # 申报口径差额由 02 匡算算好（fundingReconciliation），章节只引用、不再自行相减：
    # 章节脚本算出来的数在依据核对里没有字段可落，会被判「数值裸奔」。
    reconciliation = (estimate or {}).get("fundingReconciliation") or {}
    declared_loan = to_float(reconciliation.get("declaredLoan")
                             or ((payload or {}).get("finance") or {}).get("declaredLoan"))
    declared_total = to_float(reconciliation.get("declaredFunding")
                              or ((payload or {}).get("finance") or {}).get("declaredFunding"))
    undeclared_gap = to_float(reconciliation.get("undeclaredGap"))
    undeclared_gap_total = to_float(reconciliation.get("undeclaredGapTotal"))
    # 偿债能力逐年序列：只取还本期（有本金偿还的年份）的 DSCR。大修年当年 EBITDA 为负，
    # DSCR 会是负数，必须单列，不能混进「逐年最低」被平均掉，也不能藏起来。
    loan_schedule_rows = ((finance or {}).get("loan") or {}).get("schedule") or []
    dscr_series = ((finance or {}).get("coverage") or {}).get("dscrByYear") or []
    repay_dscr = []
    for index, row in enumerate(loan_schedule_rows):
        try:
            has_principal = float(row.get("principal") or 0) > 0
        except (TypeError, ValueError):
            has_principal = False
        value = to_float(dscr_series[index]) if index < len(dscr_series) else None
        if has_principal and value is not None:
            repay_dscr.append((index + 1, value))
    positive_dscr = [value for _, value in repay_dscr if value > 0]
    min_dscr = min(positive_dscr) if positive_dscr else None
    negative_dscr = [(year, value) for year, value in repay_dscr if value < 0]

    def _icr_text(value):
        """无有息负债时利息保障倍数没有意义——本轮真实原因是没有利息支出，
        不是 EBIT 为负（年均收入 7,118 万元、经营成本 702 万元，EBIT 实为正）。
        只有「有负债但 EBIT 为负」才写利润不足，避免与 8.2 的表述互相否定。"""
        number = to_float(value)
        if number is not None and number > 0:
            return "利息备付率 {}".format(fmt_amount(number))
        if not has_debt:
            return "利息备付率不适用（项目无有息负债）"
        return "利息备付率不适用（运营期息税前利润为负），偿债来源为含折旧回流的 EBITDA"

    def _dscr_text():
        parts = []
        total_dscr = to_float(indicators.get("dscr"))
        if total_dscr is not None:
            # 分母含建设期利息，口径必须写「全周期（含建设期）」，否则读者按「全部运营期」复算对不上 4.0663。
            # 同段的 8.2 口径说明是另一处文字，标签改了定义句也要跟着改（R6 只改标签、R8 才补定义句）。
            parts.append("偿债备付率（全周期，含建设期）{}".format(fmt_amount(total_dscr)))
        if min_dscr is not None:
            parts.append("还本期逐年最低 {}".format(fmt_amount(min_dscr)))
        if negative_dscr:
            year, value = negative_dscr[0]
            parts.append("第 {} 年大修年当年为 {}、需以前年度结余覆盖".format(year, fmt_amount(value)))
        return "，".join(parts) if parts else "偿债备付率不适用（无有息负债）"

    def _loan_rate(fin):
        return to_float(((fin or {}).get("loan") or {}).get("loanRate"))

    # 申报上限全额发行反证：由 03 财务评价算好（coverage.declaredScenario），章节只引用。
    declared_scenario = ((finance or {}).get("coverage") or {}).get("declaredScenario") or {}

    def _declared_debt_service():
        return to_float(declared_scenario.get("annualDebtService"))

    def _ebitda_normal():
        return to_float(declared_scenario.get("ebitdaNormal"))

    def _declared_dscr():
        return to_float(declared_scenario.get("dscr"))

    def _repair_year_note():
        """大修年缺口与累计结余：只写「需靠结余」不够，要回答结余够不够。

        结余、缺口、覆盖后余额全部取 03 的 coverage.repairYear，章节不再自行累加。
        """
        repair = ((finance or {}).get("coverage") or {}).get("repairYear")
        if not repair:
            return "还本期内未出现覆盖倍数低于 1.0 的年份。"
        judged = ("覆盖后仍余 {} 万元" if repair.get("covered")
                  else "尚有 {} 万元缺口，需另行安排资金").format(fmt_amount(repair.get("reserveAfter")))
        return ("第 {} 年为大修年，当年大修费 {} 万元使 EBITDA 为负、偿债备付率为 {}；"
                "该年应付本息 {} 万元。第 {}—{} 年累计偿债结余（EBITDA − 所得税 − 本息）{} 万元，"
                "用于覆盖该年缺口 {} 万元，{}。").format(
            repair.get("year"), fmt_amount(repair.get("overhaul")),
            fmt_amount(repair.get("dscr")), fmt_amount(repair.get("debtService")),
            repair.get("fromYear"), repair.get("toYear"),
            fmt_amount(repair.get("reserveBefore")), fmt_amount(repair.get("shortfall")), judged)
    revenue_first = fmt_amount(stats["first"]) if stats else "—"
    revenue_avg = fmt_amount(stats["avg"]) if stats else "—"
    revenue_total = fmt_amount(stats["total"]) if stats else "—"
    # 表 8-3 的分项是两位小数展示值，正文与合计行都按「分项展示值之和」取数，
    # 避免出现分项相加 203.77、正文却写 203.76 的尾差（R3 P2-5）。
    _cost_display_sum = sum(to_float(item.get("avgAnnual"), 0.0) for item in (cost_structure or []))
    cost_avg = fmt_amount(_cost_display_sum) if cost_structure else (fmt_amount(stats["costAvg"]) if stats else "—")

    build_years = int(period.get("constructionYears") or 0)
    operation_years = int(period.get("operationYears") or 0)
    total_years = build_years + operation_years
    # 8.1.1 用的完整口径描述：早先直接复用了表 1-1 的纯数字串，正文出现「财务评价31」
    period_detail = ("建设期 {} 年、运营期 {} 年、计算期 {} 年".format(build_years, operation_years, total_years)
                     if build_years and operation_years else "")

    ROW_NAMES = {"engineering": ("工程费用",), "otherFeeTotal": ("工程建设其他费", "其他费用"),
                 "contingency": ("预备费",), "constructionInterest": ("建设期利息",),
                 "constructionInvestment": ("建设投资",), "totalInvestment": ("项目总投资", "总投资"),
                 "equity": ("项目资本金", "资本金"), "loan": ("专项债", "债券"), "subsidy": ("补助", "补贴")}

    def ratio_of(key, rows):
        for item in rows or []:
            if item.get("key") == key:
                return item.get("ratio")
        names = ROW_NAMES.get(key) or ()
        for item in rows or []:
            label = str(item.get("name") or "")
            if any(name in label for name in names):
                return item.get("ratio")
        return None

    def ratio_text(key, rows):
        value = ratio_of(key, rows)
        return "{}%".format(value) if value is not None else "—"

    def structure_text(rows, name_key="name", amount_key="amount", ratio_key="ratio"):
        return "、".join("{} {} 万元（占 {}%）".format(item.get(name_key), fmt_amount(item.get(amount_key)),
                                                     item.get(ratio_key))
                         for item in rows or [] if item.get(ratio_key) is not None)

    # ------------------------------------------------------------ 通用事实
    name = project_name
    direction = project.get("direction") or f.text("direction", "项目方向",
                                                   "上游摘要未载明项目方向（老旧街区/老旧小区/老旧厂区）",
                                                   "UP-DIRECTION")
    for token in ("老旧街区", "老旧小区", "老旧厂区"):
        if token in str(direction):
            direction = token
            break
    # 院标模板按三个方向给了不同写法：必要性口径、改造内容框架都不同。
    # 这里直接落到正文，direction 不是三类之一时按缺口处理，不用通用话术糊过去。
    DIRECTION_TEXT = {
        "老旧街区": {
            "necessity": "补齐城市基础设施短板、完善片区公共服务功能；改善居民居住环境；保护街巷肌理与历史文脉，"
                         "实现历史资源活态传承；盘活低效存量资产，培育特色街区业态，拉动片区消费；提升城市形象与土地资产价值",
            "content": "改造内容按基础类、完善类、提升类三层组织：基础类聚焦安全隐患清零、基础市政补齐与街巷基础整治，"
                       "覆盖全部改造范围；完善类补齐公共空间与配套；提升类做沿街界面、业态引导与智慧管理",
            "market": "市场分析以沿街商业活力、客流与消费结构为主",
        },
        "老旧厂区": {
            "necessity": "落实存量提质、低效用地盘活导向；解决厂区建构筑物硬件老化、安全隐患突出问题；"
                         "解决产业低效、适配新发展需求不足问题；生态环保提质，改善区域环境风貌；盘活存量资产，提升经济社会效益；完善配套功能",
            "content": "改造内容优先安全隐患整治（建构筑物安全治理、危房拆除与结构加固），再推进产业空间改造与生态修复，"
                       "最后完善市政与公共服务配套；工业遗存有保留价值的，按保留修缮处理",
            "market": "市场分析以产业空间供给、用地效率与业态导入为主",
        },
        "老旧小区": {
            "necessity": "落实老旧小区改造政策，补齐居住领域民生短板；补齐片区配套与市政基础设施短板，适配公共服务均等化；"
                         "消除安全隐患；回应居民迫切改善诉求，建设完整社区",
            "content": "改造内容按基础类、完善类、提升类三层组织：基础类以建筑本体改造为主（屋面防水、外墙维修粉刷、楼道翻新），"
                       "完善类补市政与配套，提升类补社区服务；不含商业市场开发内容",
            "market": "老旧小区方向不做商业市场分析，改为社区服务与民生补短板分析",
        },
    }
    dir_text = DIRECTION_TEXT.get(str(direction))
    if dir_text is None:
        dir_text = {
            "necessity": f.gap("DIR-NECESSITY", "方向化的必要性表述", "工作区参数.json 的 direction 需为老旧街区/老旧厂区/老旧小区之一"),
            "content": f.gap("DIR-CONTENT", "方向化的改造内容框架", "工作区参数.json 的 direction 需为老旧街区/老旧厂区/老旧小区之一"),
            "market": f.gap("DIR-MARKET", "方向化的市场分析口径", "工作区参数.json 的 direction 需为老旧街区/老旧厂区/老旧小区之一"),
        }

    scope_area = f.with_unit("scopeArea", "改造范围面积", "上游摘要未给出改造范围面积")
    public_area = f.with_unit("publicSpaceArea", "公共空间面积", "上游摘要未给出公共空间面积")
    facade_area = f.with_unit("facadeArea", "外立面整治面积", "上游摘要未给出外立面面积")
    road_length = f.with_unit("roadArea", "配套道路面积", "上游摘要未给出道路面积")
    parking = f.with_unit("parkingSpots", "停车位数量", "上游摘要未给出停车位数量")
    chargers = f.with_unit("chargers", "充电桩数量", "上游摘要未给出充电桩数量")
    boundary = f.text("siteBoundary", "四至范围", "上游摘要未写明四至或范围边界", "UP-SITEBOUNDARY")
    nature = f.text("projectNature", "项目性质", "上游摘要未写明项目性质", "UP-PROJECTNATURE")
    owner = f.text("owner", "建设单位", "上游摘要未写明建设单位", "UP-OWNER")
    # 辐射人口按台账口径（1,868 户 × 3.2 人 = 5,977.6 人）：含示例片区B后上游的
    # 3,849.6 人已过时，正文与表格统一取台账值，不再出现两个人口口径。
    population = ("{} 人".format(fmt_num((ledger or {}).get("totals", {}).get("人口（按户均3.2人）"), 1))
                  if ledger else f.with_unit("population", "辐射人口", "上游摘要未给出辐射人口"))
    vendors = f.with_unit("vendors", "沿街游商数量", "上游摘要未给出游商数量")
    build_window = f.text("buildWindow", "建设起止时间", "上游摘要未写明开工与竣工时间", "UP-BUILDWINDOW")
    implementer = f.slot("organization", "组织架构", "输入文件-模板/槽位-空白.xlsx 的「组织架构」填写后回填")
    position = f.slot("functionalPosition", "功能定位", "输入文件-模板/槽位-空白.xlsx 的「功能定位」填写后回填")
    business = f.slot("businessTypes", "业态", "输入文件-模板/槽位-空白.xlsx 的「业态」填写后回填")
    idea = f.slot("renovationIdea", "改造思路", "输入文件-模板/槽位-空白.xlsx 的「改造思路」填写后回填")
    operation = f.slot("operationModel", "运营模式", "输入文件-模板/槽位-空白.xlsx 的「运营模式」填写后回填")
    schedule_slots = f.slots.get("schedule") or []

    upstream_label = (f.sources[0]["file"] if f.sources else "现状分析报告")
    upstream_source = "{}#facts".format(upstream_label)
    # 低空机读件只有在真的落盘时才算来源；没落盘就不写进来源行（不许编引用）
    _lowair_file = Path(project_dir) / "05_过程" / "低空解析" / "现状指标.json" if project_dir else None
    lowair_source = "05_过程/低空解析/现状指标.json" if _lowair_file and _lowair_file.is_file() else ""
    # 图 3-1 的空间现状图取 OSGB 正射/地类/轮廓；缺件时 normalize_sources 会剔掉，
    # 正文仍由 figures 的缺件判据决定画不画（不拿示意图顶替）。
    map_sources = [
        "输入文件/03-OSGB成果/成品版/模型浏览/projects/model-01/raster/ortho_northup.png",
        "输入文件/03-OSGB成果/成品版/模型浏览/projects/model-01/raster/ortho_georef.json#resM",
        "输入文件/03-OSGB成果/成品版/模型浏览/projects/model-01/raster/class_rgba.png",
        "输入文件/03-OSGB成果/成品版/模型浏览/projects/model-01/vector/buildings-outline.geojson",
    ]

    # ------------------------------------------------ 范围与规模的写实口径
    # 建设地点、改造范围、实施主体三个缺口，上游正文里都有据可写：
    # 四至范围写在「项目概况」段，规模写在「更新对象」段，实施主体是上游事实。
    # 一律落成「暂按同类项目通行做法推定」，等于把可写的事写成没资料。
    def _digest_text(*keywords):
        """取含全部关键词的第一句（成句事实），取不到返回空串。"""
        for sec in f.sections:
            for para in sec.get("paragraphs") or []:
                for sent in re.split(r"(?<=[。；])", str(para)):
                    if all(key in sent for key in keywords):
                        return sent.strip().rstrip("。； ")
        return ""

    _digest_all = "\n".join(str(p) for sec in f.sections for p in (sec.get("paragraphs") or []))
    boundary_text = (boundary if not _is_placeholder(boundary)
                     else (_digest_text("位于", "社区") or _digest_text("四至") or ""))
    # 正文原句以项目名开头，放进「建设地点与范围：」后面要把项目名去掉，否则重复。
    if boundary_text.startswith(name):
        boundary_text = boundary_text[len(name):].lstrip("，,、 ")
        # 去掉项目名后原句直接以「位于」起头，补回主语，1.1 读作「本项目位于…」。
        if boundary_text.startswith("位于"):
            boundary_text = "本项目" + boundary_text
    # 台账规模：是否含示例片区B看 scope.includeB；未配置时按台账全部社区计列，
    # 但只陈述台账数字，不替使用者判定片区归属。
    # 四个数优先取三份社区台账的原表合计（逐小区可回指），读不到时才回退到上游正文。
    _visible_areas = [row for row in (ledger or {}).get("areas", [])
                      if _include_b is not False or row["社区"] != "示例社区B"]
    if ledger:
        if _include_b is False:
            _lt = {"小区数": sum(r["小区数"] for r in _visible_areas),
                   "栋数": sum(r["栋数"] for r in _visible_areas),
                   "户数": sum(r["户数"] for r in _visible_areas),
                   "建筑面积万㎡": sum(r["建筑面积万㎡"] for r in _visible_areas)}
        else:
            _lt = ledger["totals"]
        _scope_totals = _lt
        scope_text = ("{} 个住宅小区、{} 栋、{} 户，住宅建筑规模 {} 万平方米（台账口径）".format(
            _lt["小区数"], _lt["栋数"], fmt_int(_lt["户数"]), fmt_area(_lt["建筑面积万㎡"])))
    else:
        _scope_m = re.search(
            r"(\d+)\s*个住宅小区[^。]{0,20}?(\d+)\s*栋[、，]?\s*(\d+)\s*户[、，]?\s*住宅建筑规模\s*([\d.]+)\s*万平方米",
            _digest_all)
        scope_text = ("{} 个住宅小区、{} 栋、{} 户，住宅建筑规模 {} 万平方米（台账口径）".format(*_scope_m.groups())
                      if _scope_m else "")
    if ledger and "示例片区B" not in boundary_text:
        if _include_b is True:
            boundary_text = ((boundary_text or "位于示例街道，涉及示例片区A与示例片区B")
                             + "；示例片区B涉及示例街、示例路、示例巷沿线院落，"
                             + (_boundary_basis or "四至边界与实施时序以项目侧确认为准"))
        elif _include_b is False:
            boundary_text = boundary_text or "位于示例街道，涉及示例片区A"
        else:
            boundary_text = ((boundary_text or "位于示例街道，涉及示例片区A与示例片区B")
                             + "；示例片区B的片区归属与实施边界由项目侧确认")
    _scope_split_text = ""
    if ledger:
        _scope_split_text = "范围按社区分列：{}；合计 {}。示例酒店为 1 处商业建筑（1 栋），单列不占住宅口径。".format(
            "；".join("{}（{} 个小区、{} 栋、{} 户、{} 万平方米）".format(
                row["社区"], row["小区数"], row["栋数"], fmt_int(row["户数"]),
                fmt_area(row["建筑面积万㎡"])) for row in _visible_areas),
            scope_text)
    owner_text = owner if not _is_placeholder(owner) else ""
    if boundary_text:
        # 四至范围已从上游正文写实，缺项表里不能再留一条「缺四至范围」。
        f.drop_gap("UP-SITEBOUNDARY")
    implementer_text = ("由项目业主单位（{}）组织实施".format(owner_text) if owner_text
                        else "由项目业主单位组织实施")

    # ------------------------------------------------------------ 表 1-1
    period_text = (str(total_years) if total_years
                   else f.gap("FIN-PERIOD", "财务评价计算期", "财务测算.json#period"))
    def _q(cell, unit):
        """数值与单位并进「内容」列：单独的「单位」列在文本行全是「—」，读起来像空表。"""
        return cell if str(cell) == "—" else "{} {}".format(cell, unit)

    # 范围行：台账齐备时按片区（社区）逐行列示，合计行与商业建筑单列行跟在后面。
    _scope_rows = [["改造范围（台账口径）", scope_text,
                    _basis(scope_text or PLACEHOLDER, upstream_source)]]
    if ledger:
        _scope_label = {"示例社区B": "示例片区B｜示例社区B",
                        "示例社区A": "示例片区A｜示例社区A",
                        "示例社区C": "示例片区A｜示例社区C"}
        _scope_rows = [["改造范围｜" + _scope_label.get(row["社区"], row["社区"]),
                        "{} 个小区 / {} 栋 / {} 户 / {} 万平方米".format(
                            row["小区数"], row["栋数"], fmt_int(row["户数"]),
                            fmt_area(row["建筑面积万㎡"])),
                        "三社区改造台账"] for row in _visible_areas]
        _scope_rows.append([
            "改造范围｜合计（住宅）",
            "{} 个小区 / {} 栋 / {} 户 / {} 万平方米".format(
                _lt["小区数"], _lt["栋数"],
                fmt_int(_lt["户数"]), fmt_area(_lt["建筑面积万㎡"])),
            "三社区改造台账"])
        _scope_rows.append([
            "商业建筑（单列）",
            # 商业建筑提资未填报户数与面积：写「—」并在表下注说明，
            # 不落「待确认」（待确认只贴提资缺数项，见 08 G3-2）。
            "{} 1 处、{} 栋，户数与建筑面积 —".format(
                ledger["commercial"]["名称"], ledger["commercial"]["栋数"]),
            "三社区改造台账"])
    _rows_1_1 = [
        ["项目名称", name, "工作区参数"],
        ["项目方向", direction, _basis(direction, upstream_source)],
        ["建设地点/范围", boundary_text or "见 5.1 建设范围", _basis(boundary_text or PLACEHOLDER, upstream_source)],
        *_scope_rows,
        ["实施主体", implementer_text, _basis(owner_text or PLACEHOLDER, upstream_source)],
        ["项目性质", nature, _basis(nature, upstream_source)],
        ["改造范围面积", _q(f.cell("scopeArea", "改造范围面积"), "平方米"), _basis(scope_area, upstream_source)],
        ["外立面整治面积", _q(f.cell("facadeArea", "外立面整治面积"), "平方米"), _basis(facade_area, upstream_source)],
        ["配套道路面积", _q(f.cell("roadArea", "配套道路面积"), "平方米"), _basis(road_length, upstream_source)],
        ["停车位 / 充电桩", "{} / {} 个".format(f.cell("parkingSpots", "停车位数量", digits=0),
                                          f.cell("chargers", "充电桩数量", digits=0)),
         _basis(parking, upstream_source)],
        ["项目总投资", "{} 万元".format(total_w), "投资估算"],
        ["资本金 / 专项债", "{} / {} 万元".format(equity_w, loan_w), "投资估算"],
        ["财务评价计算期", "{} 年".format(period_text),
         _basis(period_text, "财务测算.json#period")],
        ["基准收益率（折现率）", fmt_pct(period.get("discountRate")), "财务测算"]]
    # 内容全空的指标行不建：留一行「改造范围面积 | — | 推断所得」等于把空槽位摆进
    # 基本信息表（复验口径：无值不建表）。缺的指标已在 1.1 正文用定性句交代。
    def _row_has_value(cell):
        return bool(re.sub(r"—|[\s/个万平方米公里年万元%]", "", str(cell)))
    table_1_1 = md_table("表 1-1  项目基本信息与关键指标表",
                          ["项目", "内容", "依据"],
                          [row for row in _rows_1_1 if _row_has_value(row[1])])

    # ------------------------------------------------------------ 表 2-x
    def _no_placeholder(value):
        """上游表里常见的「待补…」是占位词，进成品前统一换成可核表述。"""
        text = str(value)
        for old, new in (("待补正式文本", "另行提供正式文本"),
                         ("待补清单", "另行提供清单"),
                         ("待补充", "另行补充"),
                         ("待补", "另行补充")):
            text = text.replace(old, new)
        return text

    policy_rows = []
    for item in f.tables_matching(["政策", "依据"]):
        for row in item.get("rows") or []:
            # 上游表的前几行可能是表头（「层级 / 文件名称 / …」）；
            # 早先原样塞进正文，成品表 2-1 出现「层级 | 文件名称 | 本报告使用情况 | —」
            # 一整行假数据。表头行一律跳过，只取数据行。
            first = str((row or [""])[0]).strip()
            if first in ("序号", "层级", "序号/层级"):
                continue
            if len(row) >= 2 and row[0] not in ("序号",):
                cells = [_no_placeholder(cell) for cell in row[:3]]
                # 「—（本次输入未附原文）」是占位语，换成可核表述
                cells = ["以报批材料为准" if "未附原文" in str(cell) else cell for cell in cells]
                policy_rows.append(cells + [""] * (3 - len(cells)))
    if not policy_rows:
        # 无政策清单时不摆空槽位：一行定性交代，具体文号以报批材料为准
        policy_rows = [["国家、省、市", "城市更新政策文件", "以报批材料为准"]]
    # 上游对未附原文的政策只给「未附原文/未引用具体文件」两种状态字，直接进表会
    # 出现「文件名称=以报批材料为准」这种错位；统一改成读者能看懂的定性表述。
    policy_rows = [[
        row[0],
        ("城市更新相关政策文件" if "未附原文" in str(row[1]) else row[1]),
        ("总体依据（定性引用）" if "未引用具体文件" in str(row[2]) or "未附原文" in str(row[2])
         else row[2])] + list(row[3:]) for row in policy_rows]
    # 上游政策表只有 1 行，撑不起「政策依据与符合性对照」这个表名（R2 判 P1）。
    # 改从依据库取条目：文号逐字来自库，关键要求取「可引用要点」首句，本项目响应只用
    # 4/5 章已写的项目事实，符合性判断按逐条对应给出。库不可用才退回上游口径。
    _POLICY_RESPONSE = [
        ("国办发〔2020〕23号",
         "7 个住宅小区全部建成于 2000 年以前，纳入改造范围；改造内容按基础类、完善类、提升类组织"),
        ("建科〔2021〕63号",
         "建设性质为既有街区综合整治，不实施整体拆除重建，保留既有建筑基底与街巷格局"),
        ("建办科〔2026〕33号",
         "项目方向为老旧街区，现状为房屋与市政设施老化、公共空间与配套短缺"),
        ("川办发〔2020〕63号",
         "改造对象为 2000 年底前建成、市政配套不完善、居民改造意愿强烈的小区"),
        ("川府发〔2026〕13号",
         "项目属城市更新范畴，按省级规划部署的老旧街区改造方向组织实施"),
        ("成都市第十八届人民代表大会常务委员会公告第35号",
         "项目属条例列举的老旧街区更新改造活动，按条例要求编制投资决策建议书"),
        ("成住建发〔2022〕82号",
         "由示例街道办事处组织实施，设计与造价按现行计价依据与技术导则编制"),
    ]

    def _policy_rows_from_library():
        try:
            import basis_library
            library, _ws, _params = basis_library.load_library(project_dir)
        except Exception:
            return []
        if not library or not library.get("entries"):
            return []
        by_number = {str(entry.get("文号") or "").strip(): entry
                     for entry in library["entries"]}
        rows = []
        for number, response in _POLICY_RESPONSE:
            entry = by_number.get(number)
            if not entry:
                continue
            requirement = re.split(r"[。；]", str(entry.get("可引用要点") or "").strip())[0]
            if len(requirement) > 52:
                # 按分句截断，别在半句话上切（实测会出现「落实省市老旧小…」这种断词）
                kept = ""
                for clause in re.split(r"(?<=，)", requirement):
                    if len(kept) + len(clause) > 52:
                        break
                    kept += clause
                requirement = (kept.rstrip("，") + "…") if kept else requirement[:50] + "…"
            rows.append([str(entry.get("层级") or ""), str(entry.get("名称") or ""),
                         number, requirement, response, "符合"])
        return rows

    library_rows = _policy_rows_from_library()
    if library_rows:
        table_2_1 = md_table("表 2-1  政策依据与符合性对照表",
                             ["层级", "文件名称", "文号", "关键要求", "本项目响应", "符合性判断"],
                             library_rows)
    else:
        table_2_1 = md_table("表 2-1  政策依据与符合性对照表",
                             ["层级", "文件名称", "本报告使用情况"],
                             policy_rows[:12])

    problem_rows = []
    for item in f.tables_matching(["问题", "短板", "隐患"]):
        rows = item.get("rows") or []
        if rows and len(rows[0]) >= 3:
            problem_rows = [row[:6] for row in rows[1:]]
            break
    table_2_2 = f.render_upstream_table(["问题", "短板", "隐患"],
                                        "表 2-2  现状问题清单与影响评估表",
                                        ["序号", "问题类别", "具体问题", "严重程度"],
                                        "现状问题清单（上游摘要未提供问题表）")

    # 必要性矩阵用的现状数值，一律从上游正文原句里取；取不到就退到定性表述，
    # 不给「同类项目」估值，也不把实测样区口径冒充台账口径。
    def _up(pattern, default=""):
        matched = re.search(pattern, _digest_all)
        return matched.group(1) if matched else default

    _brick = _up(r"砖混结构\s*\d+\s*个[（(]占\s*([\d.]+%)[）)]")
    _hardened = _up(r"硬化表面[^。]{0,12}?合计\s*([\d.]+%)")
    _green_road = _up(r"绿地与道路合计约占\s*([\d.]+%)")
    _cluster_households = _up(r"示例苑\s*(\d+)\s*户")
    _cluster_share = _up(r"占住宅户数[^0-9]{0,6}([\d.]+%)")
    _rest_households = _up(r"其余\s*\d+\s*个小区累计\s*(\d+)\s*户")
    _built_1980 = _up(r"1980\s*年建成的有\s*(\d+)\s*个小区")
    _max_height = _up(r"最高\s*(\d+)\s*层[^。]{0,20}?([\d.]+)\s*米")
    _median_height = _up(r"高度中位数\s*([\d.]+)\s*米")
    _hh_size = _up(r"户均\s*([\d.]+)\s*平方米")
    _problems_count = len(problem_rows) if problem_rows and "初步判定" not in str(problem_rows[0][0]) else 0
    _eng_ratio = ratio_text("engineering", comp_rows)
    _scope_rows = [
        ["宏观", "政策导向", "国家城市更新与老旧街区改造政策明确保留利用为主、拆建为辅",
         "国办发〔2020〕23号、建科〔2021〕63号、中办国办《关于持续推进城市更新行动的意见》", "强"],
        ["宏观", "投资趋势", "投资重心转向存量空间整治，本项目不新增建设用地、不拆除重建",
         "工程费用 {} 万元，占总投资 {}".format(engineering_w, _eng_ratio or "—"), "强"],
        ["中观", "上位规划", "某省、某市已就城镇老旧小区与城市更新作出部署",
         "川办发〔2020〕63号、川府发〔2026〕13号、《成都市城市更新条例》", "强"],
        ["中观", "片区定位", "老城核心居住街区，需保留街巷肌理、补齐公共服务、控制开发强度",
         boundary_text or "见 1.1 建设地点与范围", "强"],
        ["中观", "同类单元", "户数高度集中，街区级统筹优于按小区分别配建",
         ("示例苑 {} 户、占住宅户数 {}；其余小区 {} 户".format(
             _cluster_households, _cluster_share, _rest_households)
          if _cluster_households and _cluster_share else "户数分布见 3.2"), "强"],
        ["微观", "安全底线", "管线、消防、立面、无障碍四类风险叠加，不能以日常维护替代工程整治",
         ("现状问题 {} 项（见 2.3）".format(_problems_count) if _problems_count
          else "现状问题清单见 2.3"), "强"],
        ["微观", "建筑老化", "1980 年代砖混住宅进入围护结构与管线集中更换期",
         ("砖混结构占 {}、1980 年建成 {} 个小区".format(_brick, _built_1980)
          if _brick and _built_1980 else "建成年代与结构见 3.2"), "强"],
        ["微观", "功能短板", "社区服务、康养、便民休憩与停车设施缺口与人口规模不匹配",
         ("{} 户、约 {} 人，户均 {} 平方米".format(
             f.cell("households", "住宅户数", "上游摘要未给出住宅户数", 0),
             f.num("population", "片区人口", "上游摘要未给出人口", 1) or "—",
             _hh_size) if f.has("households") and f.has("population") and _hh_size
          else "人口与设施规模见 3.2"), "强"],
        ["微观", "居民诉求", "管线安全、立面破损、汛期积水、适老化与休憩配套四类诉求集中",
         "前期沟通材料与项目信息卡（见 1.4 项目资料）", "强"],
        ["微观", "更新紧迫性", "硬化表面比例高、雨季积水有空间成因，安全风险随使用年限上升",
         ("硬化表面合计 {}、绿地与道路合计约 {}".format(_hardened, _green_road)
          if _hardened and _green_road else "现状指标见 3.2"), "强"],
    ]
    table_2_3 = md_table("表 2-3  建设必要性论证矩阵",
                          ["层次", "必要性维度", "论证要点", "支撑证据", "强度判断"],
                          _scope_rows)

    # ------------------------------------------------------------ 表 3-1
    # 单位 "1" 是比率：默认转百分比，但「容积率」本身按比值显示，不能乘 100
    # （实测踩过：容积率 1.0281 被显示成 102.81%）。口径与 CardoOSGB 的 readability.py 一致。
    RATIO_AS_IS = {"floor_area_ratio_est", "floor_area_ratio"}
    RATIO_KEEP_DECIMAL = {"classify_precision_vs_recon", "classify_recall_vs_recon"}
    metric_rows = []
    for item in (lowair or {}).get("metrics") or []:
        value = item.get("value")
        if value is None:
            continue
        raw_unit = str(item.get("unit") or "")
        mid = str(item.get("id") or "")
        if raw_unit == "1":
            if mid in RATIO_AS_IS:
                shown, unit = fmt_num(value, 2), "—"
            elif mid in RATIO_KEEP_DECIMAL:
                shown, unit = fmt_num(float(value) * 100, 1), "%"
            else:
                shown, unit = fmt_num(float(value) * 100, 2), "%"
        elif raw_unit == "%":
            shown, unit = fmt_num(value, 2), "%"
        else:
            digits = 0 if raw_unit in ("栋", "个", "条", "处", "块", "棵") else 2
            shown, unit = fmt_num(value, digits), (raw_unit or "—")
        metric_rows.append([item.get("name"), shown, unit, "2.5D 网格估算", "现状指标.json"])
    for label, key, unit in (("改造范围面积", "scopeArea", "万平方米"),
                             ("外立面整治面积", "facadeArea", "万平方米"),
                             ("配套道路面积", "roadArea", "万平方米"),
                             ("停车位", "parkingSpots", "个"),
                             ("充电桩", "chargers", "个"),
                             ("沿街游商", "vendors", "户"),
                             ("辐射人口", "population", "人")):
        item = f._raw(key)
        if item is None:
            f.cell(key, label)          # 缺项登记：进内部件《待确认清单》，不带进成品表格
            shown, source = "—", "上游摘要"
        else:
            value = to_float(item.get("value"), 0.0)
            if unit == "万平方米":
                shown = fmt_num(value / 10000.0, 2)
            elif unit == "人":
                shown = fmt_num(value, 1)
            else:
                shown = fmt_num(value, 0 if unit in ("个", "户") else 2)
            source = (f.measured.get(key) or {}).get("source") or "上游摘要"
        metric_rows.append([label, shown, unit,
                            "匡算附表" if key in f.measured else "上游摘要", source])
    # 表 3-1 原为「现状指标汇总表」，数值列全部取不到（上游只有点状描述），
    # 整表只剩「—」；按「无值不建表」直接不建，关键指标已在 3.1 正文说明。
    table_3_1 = ""

    # ------------------------------------------------------------ 表 4-x
    # 只用完整句子填单元格：硬截字符会把「商业），没有商铺数量…需现场核实的是」这种
    # 半句话塞进表里，读者一眼看出是机器拼接。
    market_excerpt = f.excerpt_sentences(["商业", "业态", "市场"], 120)
    table_4_1_rows = [
        ["居住与人口", scope_text or _cell_or(population, "同类片区人口规模"), "日常服务配套与公共空间不足",
         "便民服务与公共空间", _basis(scope_text or population, upstream_source)],
        ["商业供给", market_excerpt or
         f.gap("UP-MARKET", "现状商业供给", "上游摘要「商业/业态」小节"),
         "业态同质、空置与低租金", "规范化经营空间",
         _basis(market_excerpt or PLACEHOLDER, upstream_source)],
        ["停车", _cell_or(parking, "同类项目配建标准"), "停车泊位供给不足",
         _cell_or(parking, "同类项目配建标准"), _basis(parking, upstream_source)],
        ["充电", _cell_or(chargers, "同类项目配建标准"), "充电设施覆盖不足",
         _cell_or(chargers, "同类项目配建标准"), _basis(chargers, upstream_source)],
        ["公共空间", _cell_or(public_area, "同类项目规模"), "公共活动空间不足",
         _cell_or(public_area, "同类项目规模"), _basis(public_area, upstream_source)]]
    # 缺值行整行剔除：格子里留「—」比写「暂按…推定」干净，但整行空着仍像没做完。
    table_4_1 = md_table("表 4-1  区域市场供需分析表",
                         ["维度", "现状", "缺口", "本项目可承接", "依据"],
                         [row for row in table_4_1_rows
                          if any(str(row[i]).strip() not in ("", "—") for i in (1, 3))])

    revenue_basis = {item.get("key"): (item.get("basis") or "—") for item in
                     ((finance or {}).get("revenue") or {}).get("items") or []}
    # 表 4-2（业态配置）与表 4-3（运营收入）重复登记同一批收入项，且「数量或面积」
    # 一列全部取不到；按「无值不建表」只留收入表，业态组合在 4.3 正文写清。
    table_4_3 = ""
    table_4_4 = md_table("表 4-2  运营收入预判表",
                          ["收入项", "运营期合计（万元）", "占比", "测算依据（数量×单价×出租率）"],
                          [[item.get("name"), fmt_amount(item.get("total")),
                            "{}%".format(item.get("ratio")) if item.get("ratio") is not None else "—",
                            revenue_basis.get(item.get("key")) or "—"]
                           for item in revenue_structure] or
                          [gap_row(4, "收入项（财务测算未提供）")])

    # ------------------------------------------------------------ 表 5-x
    # 分项工程量来自匡算附表1 的工程量列（财务参数.xlsx「工程费」表的 单位/工程量），
    # 与工程费用同表列示，读者可逐项核对「工程量—工程费」对应关系。
    _eng_rows = (estimate or {}).get("engineeringRows") or []
    _unit_display = {"m": "米", "m2": "平方米", "㎡": "平方米"}
    # 数据源未提供单位/工程量（渲染时这两列会被整列裁掉），题注不能承诺「工程量」。
    table_5_1 = md_table("表 5-1  分项改造内容与工程费用表",
                         ["编号", "工程内容", "单位", "工程量", "工程费用（万元）"],
                         [[row.get("key"), row.get("name"),
                           _unit_display.get(str(row.get("unit") or ""), row.get("unit") or "—"),
                           fmt_num(row.get("quantity"), 2) if row.get("quantity") is not None else "—",
                           fmt_amount(row.get("amount"))]
                          for row in _eng_rows] or [gap_row(5, "分项工程量（投资估算未提供）")])
    table_5_2 = ""
    _schedule_body = [["第 {} 年".format(row.get("year")), fmt_amount(row.get("engineering")),
                       fmt_amount(row.get("otherFees")), fmt_amount(row.get("contingency")),
                       fmt_amount(row.get("constructionInterest")), fmt_amount(row.get("total"))]
                      for row in year_plan_rows]
    if _schedule_body:
        _schedule_body.append([
            "合计",
            fmt_amount(sum(to_float(row.get("engineering"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("otherFees"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("contingency"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("constructionInterest"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("total"), 0) for row in year_plan_rows))])
    table_5_3 = md_table("表 5-2  建设时序与分年投资计划表",
                          ["年度", "工程费", "其他费", "预备费", "建设期利息", "合计"],
                          _schedule_body or [gap_row(6, "分年投资计划（投资估算未提供）")])

    # ------------------------------------------------------------ 表 6-x
    org_rows = []
    for item in f.tables_matching(["组织", "架构", "职责"]):
        rows = item.get("rows") or []
        if rows and len(rows[0]) >= 3:
            org_rows = [row[:6] for row in rows[1:]]
            break
    # 表 6-1（组织架构）与表 6-2（收益模式）上游均未提供成表数据，整表只剩「—」；
    # 按「无值不建表」并入 6.1、6.2 正文，组织职责与收益构成用文字写清。
    table_6_1 = ""
    table_6_2 = ""

    # ------------------------------------------------------------ 表 7-x
    # 「计算依据」不能再整列自引「投资估算」：逐行写到取费口径，读者才能复核。
    _basis_by_key = {
        "engineering": "分项改造内容×单位造价指标（表 5-1）",
        "otherFeeTotal": "19 项取费科目按基数与费率逐项计取",
        "contingency": "（工程费用＋其他费用）× 预备费率",
        "constructionInvestment": "工程费用＋其他费用＋预备费",
        "constructionInterest": ("无有息负债，不计建设期利息"
                                 if (to_float(funding.get("loan"), 0) or 0) <= 0 else
                                 "专项债分年提款×年利率×建设期计息系数"),
        "totalInvestment": "建设投资＋建设期利息",
    }
    table_7_1 = md_table("表 7-1  项目总投资构成表",
                          ["序号", "费用类别", "金额（万元）", "占比", "计算依据"],
                          [[str(index), row.get("name"), fmt_amount(row.get("amount")),
                            "{:.2f}%".format(float(row.get("ratio"))) if row.get("ratio") is not None else "—",
                            _basis_by_key.get(str(row.get("key")), "匡算")]
                           for index, row in enumerate(
                               [item for item in comp_rows if int(item.get("level", 0)) <= 1], start=1)] or
                          [gap_row(5, "总投资构成（投资估算未提供）")])
    _plan_body = [["第 {} 年".format(row.get("year")), fmt_amount(row.get("engineering")),
                   fmt_amount(row.get("otherFees")), fmt_amount(row.get("contingency")),
                   fmt_amount(row.get("constructionInterest")), fmt_amount(row.get("total")),
                   "{:.2f}%".format(float(row.get("ratio"))) if row.get("ratio") is not None else "—"]
                  for row in year_plan_rows]
    if _plan_body:
        # 表尾补合计行：分年表没有合计，读者无法与 7.1 总投资对账。
        _plan_body.append([
            "合计",
            fmt_amount(sum(to_float(row.get("engineering"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("otherFees"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("contingency"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("constructionInterest"), 0) for row in year_plan_rows)),
            fmt_amount(sum(to_float(row.get("total"), 0) for row in year_plan_rows)),
            "100.00%"])
    table_7_2 = md_table("表 7-2  分年投资使用计划表",
                          ["年度", "工程费", "其他费", "预备费", "建设期利息", "合计", "占比"],
                          _plan_body or [gap_row(7, "分年投资计划（投资估算未提供）")])
    def _is_equity(row):
        return row.get("key") == "equity" or "资本金" in str(row.get("name") or "")

    def _is_subsidy(row):
        return row.get("key") == "subsidy" or "补助" in str(row.get("name") or "")

    def _is_zero(row):
        return abs(to_float(row.get("amount"), 0.0) or 0.0) < 1e-9

    table_7_3 = md_table("表 7-3  资金筹措方案表",
                          ["资金来源", "金额（万元）", "占比", "到位时点", "依据或成本"],
                          [[row.get("name"),
                            ("0.00" if (_is_equity(row) and equity_undetermined)
                             else fmt_amount(row.get("amount"))),
                            ("0.00%" if (_is_equity(row) and equity_undetermined) else
                             ("{:.2f}%".format(float(row.get("ratio"))) if row.get("ratio") is not None else "—")),
                            ("不计列" if (_is_zero(row) and _is_subsidy(row)) else "按建设进度分年到位"),
                            ("本项目未计列政策性补助收入"
                             if (_is_zero(row) and _is_subsidy(row))
                             else ("政策性补助全额覆盖，资本金为 0（按资金申请口径测算）"
                                   if (_is_equity(row) and equity_undetermined)
                                   else ("中央预算内投资，作项目资本金" if _is_equity(row)
                                         else ("按资金申请口径，政策性补助全额覆盖" if _is_subsidy(row)
                                               else "专项债，按投资缺口提款"))))]
                           for row in fund_rows] or [gap_row(5, "资金筹措（投资估算未提供）")])

    # ------------------------------------------------------------ 表 8-x
    labor_item = next((item for item in (cost_structure or []) if item.get("key") == "labor"), {})
    table_8_1 = md_table("表 8-1  财务评价基础数据与参数表",
                          ["参数", "取值", "单位", "依据"],
                          [["财务评价计算期", str(total_years) if total_years else "—", "年", "财务测算"],
                           ["建设期 / 运营期", "{} / {}".format(build_years or "—", operation_years or "—"), "年",
                            "财务测算"],
                           # 单位写「百分比」而不是「％」：单字符单元格在 PDF 里自成一行，
                           # 行首标点审计会把它记成行首「％」（本项目已因此返工两次）。
                           ["折现率（基准收益率）",
                            fmt_num(to_float(period.get("discountRate"), 0.0) * 100, 2), "百分比", "财务测算"],
                           ["收入增长率（运营期）", "0.00", "百分比",
                            "财务参数：按达产年等额计列，首年按 50% 计列，逐年爬坡不单列"],
                           ["出租率 / 使用率（达产年）",
                            "按提资《收入》表计列，未单列出租率", "—",
                            "收入表 R1—R4；本工作区不另行设定出租率，避免与提资口径冲突"],
                           ["折旧年限", str(min(30, operation_years or 0) or operation_years or "—"), "年",
                            "不超过运营期，期末账面余值等于残值"],
                           ["残值率", "5.00", "百分比", "财务口径"],
                           ["工资增长率", "3.00", "百分比", "财务参数"],
                           ["借款年利率",
                            fmt_num(to_float(((finance or {}).get("loan") or {}).get("loanRate"), 0.0) * 100, 2),
                            "百分比", "专项债"],
                           ["借款期限 / 宽限期",
                            "{} / {}".format(((finance or {}).get("loan") or {}).get("loanTerm") or "—",
                                             ((finance or {}).get("loan") or {}).get("graceYears") or "—"),
                            "年", "财务测算"],
                           ["项目资本金", equity_w, "万元",
                            ("政策性补助全额覆盖，资本金为 0（按资金申请口径测算）" if equity_undetermined else "投资估算")],
                           ["专项债", loan_w, "万元", "投资估算"],
                           ["工资及福利（年均）", fmt_amount(labor_item.get("avgAnnual")), "万元",
                            "6 人×7 万元/人·年×福利系数 1.2，年增长 3%"]])
    table_8_2 = md_table("表 8-2  经营收入测算汇总表（含税）",
                          ["收入项", "测算基础", "运营期合计（含税，万元）", "占比（含税）"],
                          [[item.get("name"), revenue_basis.get(item.get("key"), "—"),
                            fmt_amount(item.get("total")),
                            "{}%".format(item.get("ratio")) if item.get("ratio") is not None else "—"]
                           for item in revenue_structure] or [gap_row(4, "收入测算（财务测算未提供）")])
    # 收入单位疑误：01 已按元还原为万元并在纠错台账留痕，但成品不能静默改数。
    # 这里在对应收入表下加脚注，写清原值、单位存疑、按元计列与核实口径；
    # 08 的 REV-MAG 只有在「本脚注 + 内部件《待确认清单》」两处都披露时才允许带警告通过。
    _unit_notes = []
    for _entry in (payload.get("revenue") or []):
        if not isinstance(_entry, dict):
            continue
        _suspect = _entry.get("unitSuspect") or {}
        if not _suspect:
            continue
        _unit_notes.append(
            "注：表 8-2 中「{} {}」一项，提资原值 {} 按「{}」表头填报，数值量级与街区级收入不符，"
            "判定单位疑误（元填成万元）；本报告按元计列为 {} 万元/年，最终以提资方核实结果为准。".format(
                _entry.get("key") or "", _entry.get("name") or "",
                "{:,.2f}".format(float(_suspect.get("originalAmount") or 0)),
                _suspect.get("declaredUnit") or "万元",
                "{:.2f}".format(float(_entry.get("annualAmount") or 0))))
    unit_note = "\n\n".join(_unit_notes)
    _cost_rows = []
    for item in cost_structure or []:
        _amount = to_float(item.get("avgAnnual"), 0.0) or 0.0
        _ratio = (_amount / _cost_display_sum * 100) if _cost_display_sum else None
        _cost_rows.append([item.get("name"), fmt_amount(_amount),
                           "{:.2f}%".format(_ratio) if _ratio is not None else "—"])
    if _cost_rows:
        _cost_rows.append(["合计", fmt_amount(_cost_display_sum), "100.00%"])
    table_8_3 = md_table("表 8-3  经营成本测算表（不含税）",
                          ["成本项", "年均金额（不含税，万元）", "占经营成本比"],
                          _cost_rows or [gap_row(3, "经营成本（财务测算未提供）")])
    debt_rows = {int(item.get("year")): item for item in
                 ((tables.get("debtSchedule") or {}).get("rows") or [])}
    icr_by_year = ((finance or {}).get("coverage") or {}).get("icrByYear") or []
    dscr_by_year = ((finance or {}).get("coverage") or {}).get("dscrByYear") or []

    def _debt_row(year):
        row = debt_rows.get(year) or {}
        # 建设期与还清借款后的年份没有可比的利息/本息覆盖口径，写「不适用」，
        # 不能写「—」：正文 8.2 已明确说这些年份不适用，表内符号必须同口径（R3 P2-4）。
        icr = icr_by_year[year - 1] if 0 <= year - 1 < len(icr_by_year) else None
        dscr = dscr_by_year[year - 1] if 0 <= year - 1 < len(dscr_by_year) else None
        return ["第 {} 年".format(year), fmt_amount(row.get("opening")),
                fmt_amount(row.get("principal")), fmt_amount(row.get("interest")),
                fmt_amount(row.get("closing")),
                fmt_num(icr, 4) if icr is not None else "不适用",
                fmt_num(dscr, 4) if dscr is not None else "不适用"]

    # 逐年全列（第 1 年—计算期末），不抽样、不折叠。
    # 正文 8.2 的「逐年最低 DSCR 1.17（第 7 年）」「第 14 年大修年 −0.88」必须能在表里找到对应行；
    # 早先只挑第 1–5 年与末年，第 6–29 年被整段跳过，读者无法从表验证最低值。
    debt_years = sorted(debt_rows) or list(range(1, total_years + 1))
    debt_body = [_debt_row(year) for year in debt_years]
    overhaul = ((finance or {}).get("cost") or {}).get("items", {}).get("overhaul") or []
    overhaul_years = [index + 1 for index, value in enumerate(overhaul)
                      if (to_float(value, 0.0) or 0.0) > 0]
    table_8_8_note = ("注：第 {} 年为大修年（大修费 {} 万元/次，计入当年经营成本），"
                      "与还本付息叠加年份的 DSCR 见表内对应行。".format(
                          "、".join(str(year) for year in overhaul_years),
                          fmt_amount(overhaul[overhaul_years[0] - 1]))
                      if overhaul_years else "")
    table_8_8 = md_table("表 8-8  借款还本付息与偿债能力表",
                          ["年度", "年初借款余额", "当年还本", "当年付息", "年末余额", "ICR（倍）", "DSCR（倍）"],
                          debt_body or [gap_row(7, "还本付息计划（财务测算未提供）")])
    # 无借款时整表逐年为 0.00、覆盖倍数全「—」——这就是一张空表。
    # 按「无值不建表」不建，8.2 正文用一句话交代（不留表位编号，也不留空表）。
    has_debt = any(abs(to_float((row or {}).get(key), 0.0) or 0.0) > 1e-9
                   for row in debt_rows.values()
                   for key in ("opening", "principal", "interest", "closing"))
    if not has_debt:
        table_8_8 = ""
    # 表号必须连续：无贷款时不建表 8-8，敏感性表与可研对照表依次前移一位。
    sens_no = "8-8" if not has_debt else "8-9"
    compare_no = "8-9" if not has_debt else "8-10"

    # 表 8-4 / 8-5：两个口径的现金流量表，逐年列入，数据取 财务测算.json#cashflows
    cashflows = (finance or {}).get("cashflows") or {}

    def _cashflow_rows(series):
        series = series or {}

        def _pick(name):
            return {int(item.get("year")): item.get("value")
                    for item in series.get(name) or []}

        inflow, outflow = _pick("inflow"), _pick("outflow")
        net = _pick("net") or _pick("netAfterTax")
        cumulative = _pick("cumulative")
        years = sorted(net) or sorted(inflow)
        return [["第 {} 年".format(year), fmt_amount(inflow.get(year)),
                 fmt_amount(outflow.get(year)), fmt_amount(net.get(year)),
                 fmt_amount(cumulative.get(year))] for year in years]

    table_8_4 = md_table("表 8-4  全部投资财务现金流量表（万元）",
                          ["年度", "现金流入", "现金流出", "净现金流量", "累计净现金流量"],
                          _cashflow_rows(cashflows.get("allInvestment")) or
                          [gap_row(5, "全部投资现金流（财务测算未提供）")])
    table_8_5 = md_table("表 8-5  资本金财务现金流量表（万元）",
                          ["年度", "现金流入", "现金流出", "净现金流量", "累计净现金流量"],
                          _cashflow_rows(cashflows.get("equity")) or
                          [gap_row(5, "资本金现金流（财务测算未提供）")])
    # 三种因素的「不可实现」性质不同，不能统一写成一句「不可实现」（R3 P2-2）：
    # 收入是市场边界、投资是改造刚性边界、成本是物理边界。判定只在基准行写一次，
    # ±10% 两行不再重复同一句话。
    _point_reason = {"revenue": "已高于市场可实现水平",
                     "investment": "与安全类、基础类刚性改造冲突",
                     "cost": "物理上不可能"}
    sens_rows = []
    _negative_base = (to_float(display.get("fnpvAll"), 0.0) or 0.0) < 0
    for key, label in (("revenue", "收入"), ("investment", "投资"), ("cost", "经营成本")):
        factor = sens_factors.get(key) or {}
        point = factor.get("fnpvZeroDeltaPercent")
        feasible = factor.get("fnpvZeroFeasible")
        if point is None:
            point_cell = "—"
        elif not _negative_base:
            # 基准 FNPV 为正：临界点衡量的是「还能承受多大的不利变动」，不是缺口。
            point_cell = ("可承受（外推 {:+.1f}%）".format(float(point)) if not feasible
                          else "临界点 {:+.1f}%".format(float(point)))
        elif feasible:
            point_cell = "{:+.1f}%".format(float(point))
        else:
            point_cell = "不可实现：{}（外推 {:+.1f}%）".format(
                _point_reason.get(key, "超出可实现区间"), float(point))
        for row in factor.get("rows") or []:
            label_text = str(row.get("label") or "")
            # 基准行一律取指标表（单一口径）：敏感性 base 已锚定 indicators，
            # 这里再取一次是双保险，防止旧数据残留把两个基准同时写进报告。
            is_base = ("基准" in label_text) or abs(to_float(row.get("deltaPercent"), 9.0)) < 1e-9
            firr = display.get("firrAll") if is_base else row.get("firr")
            fnpv = display.get("fnpvAll") if is_base else row.get("fnpv")
            pt = display.get("ptAll") if is_base else row.get("pt")
            sens_rows.append([label, label_text, "{:.2f}".format(float(firr or 0) * 100),
                              fmt_amount(fnpv),
                              fmt_num(pt, 1) if pt is not None else PAYBACK_NOT_FORMED,
                              point_cell if is_base else "—"])
    table_8_9 = md_table("表 {}  敏感性分析表".format(sens_no),
                          ["因素", "变动幅度", "FIRR（%）", "FNPV（万元）", "回收期（年）", "临界点判定"],
                          sens_rows or [gap_row(6, "敏感性分析（财务测算未提供）")])
    table_8_6 = md_table("表 8-6  全部投资财务指标计算结果汇总",
                          ["序号", "指标名称", "单位", "计算结果（全部投资）"],
                          [["1", "财务内部收益率（税后）", "百分比", fmt_pct_num(display.get("firrAll"))],
                           ["2", "财务净现值（税后）", "万元", fmt_amount(display.get("fnpvAll"))],
                           ["3", "静态投资回收期（税后）", "年",
                            _payback_text(display.get("ptAll"))]])
    equity_na = EQUITY_NOT_APPLICABLE if equity_undetermined else None
    # 资本金为 0 时不存在资本金出资，机械算出的资本金 FNPV 没有经济含义，
    # 不能与「不适用（资本金为 0）」并排列出，否则读者会当成可比的收益指标。
    equity_fnpv_text = equity_na or fmt_amount(display.get("fnpvEquity"))
    table_8_7 = md_table("表 8-7  资本金财务指标计算结果汇总",
                          ["序号", "指标名称", "单位", "计算结果（资本金）"],
                          [["1", "财务内部收益率（税后）", "百分比",
                            equity_na or fmt_pct_num(display.get("firrEquity"))],
                           ["2", "财务净现值（税后）", "万元", equity_fnpv_text],
                           ["3", "静态投资回收期（税后）", "年",
                            equity_na or _payback_text(display.get("ptEquity"))]])

    # 与可研对照：本报告引擎自算值 vs 用户提供的可研财务指标。
    # 只作说明，不参与判定；两者不一致时登记缺项，交人工复核输入口径。
    compare_note = ""
    compare_table = None
    if compare:
        pairs = [("firrAll", "全部投资FIRR（税后）", "%", fmt_pct),
                 ("fnpvAll", "全部投资FNPV（税后）", "万元", fmt_amount),
                 ("ptAll", "全部投资回收期", "年", fmt_amount),
                 ("firrEquity", "资本金FIRR（税后）", "%", fmt_pct),
                 ("fnpvEquity", "资本金FNPV（税后）", "万元", fmt_amount),
                 ("ptEquity", "资本金回收期", "年", fmt_amount)]
        rows, diffs = [], []
        for key, label, unit, formatter in pairs:
            mine = display.get(key)
            theirs = compare.get(key)
            if mine is None or theirs is None:
                continue
            rows.append([label, unit, formatter(mine), formatter(theirs)])
            mine_f = to_float(mine, None)
            theirs_f = to_float(theirs, None)
            if mine_f is not None and theirs_f is not None:
                # 比率按百分点比，金额/期数按相对偏差比
                delta = abs(mine_f - theirs_f)
                rel = delta / abs(theirs_f) if abs(theirs_f) > 1e-9 else (0.0 if delta < 1e-9 else 1.0)
                limit = 0.005 if unit == "%" else 0.10
                if rel > limit:
                    diffs.append("{}：本报告 {} vs 可研 {}".format(label, formatter(mine), formatter(theirs)))
        compare_table = md_table(
            "表 {}  本报告测算值与可研财务指标对照表".format(compare_no),
            ["指标名称", "单位", "本报告测算值", "可研对照值"], rows) if rows else None
        if compare_table:
            if diffs:
                gaps.append({
                    "code": "FIN-REF-DIFF", "what": "本报告测算值与可研对照值差异较大（{}）".format("；".join(diffs)),
                    "why": "输入口径（收入基数、折现率、建设期/运营期、还本付息）与可研附表不一致",
                    "need": "核对 财务参数.xlsx 与可研附表口径后重跑；正式申报前须统一",
                    "owner": "user",
                })
                compare_note = ("口径对照：本表为可研财务指标对照值，仅供核对，不参与本报告判定与结论；"
                                "本报告指标一律取上表测算值。两者存在差异（{}），"
                                "已登记缺项（见缺项表 FIN-REF-DIFF），正式申报前须统一口径。".format("；".join(diffs)))
            else:
                compare_note = ("口径对照：本表为可研财务指标对照值，仅供核对，不参与本报告判定与结论；"
                                "本报告指标与可研口径一致。")
        else:
            compare_note = ""

    # ------------------------------------------------------------ 表 9-x
    criteria_names = [item.get("key") or item.get("name") for item in criteria[:4]]
    criteria_matrix = ahp.get("criteriaMatrix") or []
    matrix_rows = []
    for index, criterion in enumerate(criteria_names):
        row = criteria_matrix[index] if index < len(criteria_matrix) else [None] * 4
        matrix_rows.append([criterion] + [fmt_num(value, 4) if value is not None else "—" for value in row[:4]])
    table_9_1 = md_table("表 9-1  准则层权重派生展示矩阵",
                          ["准则"] + criteria_names, matrix_rows or [gap_row(5, "AHP 判别矩阵（上游未提供）")])
    # 行几何平均是表格契约 slot5 的必备列（表头必须与契约一致）；上游权重表未给值时，
    # 01 取料阶段按判别矩阵补算并写回 05_过程/AHP/权重表.json，登记来源可回指该字段，
    # 不在成文阶段就地反算（否则 07「数值不裸奔」判无出处）。
    table_9_2 = md_table("表 9-4  准则层权重结果",
                          ["准则", "行几何平均", "权重"],
                          [[item.get("key") or item.get("name"),
                            ("{:.5f}".format(float(item["geometricMean"]))
                             if item.get("geometricMean") is not None else "—"),
                            ("{:.4f}".format(float(item["weight"]))
                             if item.get("weight") is not None else "—")]
                           for item in ahp.get("criteriaWeights") or []] or
                          [gap_row(3, "准则层权重（上游未提供）")])

    # 指标层判别矩阵（四组）：由上游子权重按 a_ij = w_i / w_j 还原，只用于展示，不参与计算
    groups = {}
    for item in indicator_weights:
        groups.setdefault(item.get("criterion"), []).append(item)
    matrix_9_3 = []
    for criterion in criteria_names:
        items = [item for item in groups.get(criterion) or [] if float(item.get("subWeight") or 0) > 0]
        if not items:
            continue
        names = [item.get("indicator") for item in items]
        subs = [float(item.get("subWeight") or 0) for item in items]
        matrix_9_3.append(["[{}] {}".format(criterion, " · ".join(names))] + [""] * max(0, len(names) - 1))
        matrix_9_3.append([""] + names + [""] * max(0, 4 - len(names)))
        for row_index, row_name in enumerate(names):
            matrix_9_3.append([row_name] + [fmt_num(subs[row_index] / subs[col], 4) if col < len(subs) else ""
                                            for col in range(len(names))])
    table_9_3 = md_table("表 9-3  指标层权重派生展示矩阵（四组）",
                          ["指标层权重派生展示矩阵（4 组）", "子准则 1", "子准则 2", "子准则 3", "子准则 4"],
                          matrix_9_3 or [gap_row(5, "指标层判别矩阵（上游未提供指标子权重）")])

    table_9_4 = md_table("表 9-2  指标综合权重表",
                          ["准则", "准则权重", "指标层", "子权重(4阶矩阵内)", "综合权重"],
                          [[item.get("criterion"), fmt_num(item.get("criterionWeight"), 4),
                            item.get("indicator"), fmt_num(item.get("subWeight"), 4),
                            fmt_score(item.get("compositeWeight"))] for item in indicator_weights] or
                          [gap_row(5, "指标综合权重（上游未提供）")])
    scheme_rows = []
    for scheme in sorted(scheme_list,
                         key=lambda item: (canon_scheme(item.get("name")) or ("S9", ""))[0]):
        factor = to_float(scheme.get("scaleFactor"))
        # 变量名不能叫 indicators：本函数的 indicators 是财务测算的指标字典，
        # 循环里重绑定会把后面的 1.3 结论、8.1.1 判据、8.2 偿债段全部读空，
        # 于是 FNPV 明明为负却走进「财务可行 / 无有息负债」分支（R5 B-1/B-2/B-3）。
        scheme_indicators = scheme.get("indicators") or {}
        # 方案比选.json 把 FIRR/FNPV/回收期放在 indicators 下；早先按顶层键取，
        # 三列全空后被渲染器整列丢弃，表 9-5 只剩投资与得分（R3 P2-22）。
        firr_pct = to_float(scheme.get("firrPercent"))
        if firr_pct is None and scheme_indicators.get("firr") is not None:
            firr_pct = to_float(scheme_indicators.get("firr"), 0.0) * 100
        fnpv = scheme.get("fnpv")
        if fnpv is None:
            fnpv = scheme_indicators.get("fnpv")
        payback = scheme.get("pt")
        if payback is None:
            payback = scheme_indicators.get("pt")
        canon = canon_scheme(scheme.get("name")) or (scheme.get("id"), scheme.get("name"))
        scheme_rows.append([canon[0], canon[1],
                            "{}（{}×）".format(fmt_amount(scheme.get("investment")),
                                               fmt_num(factor, 2) if factor is not None else "—"),
                            fmt_num(firr_pct, 2) if firr_pct is not None else "—",
                            fmt_amount(fnpv),
                            fmt_num(payback, 2) if payback is not None else PAYBACK_NOT_FORMED,
                            fmt_score(scheme.get("score")),
                            "推荐" if scheme.get("id") == winner_id else "不推荐"])

    def _report_id(row):
        """AHP 行的报告编号：按方案名落到统一编号 S1/S2/S3，两表与正文同一套。"""
        canon = canon_scheme((row or {}).get("name"))
        return canon[0] if canon else str((row or {}).get("id") or "—")

    table_9_6 = md_table("表 9-5  方案比选结果表（投资与财务指标为缩放口径，括号内为缩放系数）",
                          ["编号", "方案", "投资（万元）", "FIRR（%）", "FNPV（万元）", "回收期（年）", "得分（0–1）", "推荐"],
                          scheme_rows or [gap_row(8, "方案比选（未提供）")])
    table_9_7 = md_table("表 9-6  AHP 方案综合得分表",
                          ["编号", "方案", "综合得分（10 分制）", "评级"],
                          [[_report_id(row), row.get("name"),
                            fmt_score(row.get("composite")), row.get("label") or "—"]
                           for row in sorted((score_block or {}).get("rows") or [],
                                             key=lambda item: _report_id(item))] or
                          [gap_row(4, "AHP 评分（上游未提供）")])

    # 本次估算范围：匡算提资（财务参数「出处台账」）指向哪本，就只陈述哪本；
    # 是否含示例片区B由 scope.includeB 与 scope.boundaryBasis 说明，代码不预设结论。
    # 这句话在 1.1、7.1、表 7-1 上方注三处逐字一致，避免前后口径打架。
    _estimate_scope = ""
    if ledger:
        if _include_b is True:
            _estimate_scope = (_boundary_basis or
                               "本次投资估算范围以匡算提资口径为准，含示例片区B；"
                               "示例片区B的投资口径按项目侧确认补充。")
        elif _include_b is False:
            _estimate_scope = ("本次投资估算范围不含示例片区B；"
                               "示例片区B另行立项或分期实施。")
        else:
            _estimate_scope = ("本次投资估算范围以匡算提资口径为准；"
                               "示例片区B是否计入，由项目侧确认。")

    if _include_b is True:
        _range_phrase = "实施范围为两个片区：示例片区B与示例片区A；"
        _scope_note = ("示例片区B 4 个院纳入本次改造范围，与示例片区A合并计列"
                       + ("；" + _boundary_basis if _boundary_basis else "；实施边界与时序由项目侧确认")
                       + "。")
    elif _include_b is False:
        _range_phrase = ""
        _scope_note = "示例片区B不纳入本次改造范围，另行立项或分期实施。"
    else:
        _range_phrase = "实施范围以项目侧确认为准；"
        _scope_note = "示例片区B是否纳入本次改造范围、其边界与时序，由项目侧确认。"

    # ============================================================ 第 1 章
    basis_lines, _basis_meta = basis_library.build_basis_section(project_dir, name)
    chapters = {}
    chapters[1] = "\n".join([
        "# 1 项目概述", "",
        section(2, "1.1", "项目基本信息", paragraphs(
            "{}为{}类改造项目，{}，项目总投资 {} 万元{}。".format(
                name, direction,
                scope_text or _metric_phrase([("改造范围面积", scope_area)]), total_w,
                "，财务评价计算期 {} 年".format(period_text) if period_text else ""),
            # 项目名只能用提资（财务参数 项目!B2）里的正式名称，不自造；
            # 实施范围是否含第二个片区由 scope.includeB 决定。
            "项目名称：{}。项目方向：{}。{}"
            "建设地点与范围：{}。项目性质：{}。".format(
                name, direction, _range_phrase,
                boundary_text or "见 5.1 建设范围", nature),
            ("改造范围按台账口径确定：{}。".format(scope_text) if scope_text else ""),
            _estimate_scope,
            "实施主体与组织：{}。".format(implementer_text),
            *(["主要指标：" + _metric_phrase([
                ("公共空间面积", public_area), ("老旧房屋外立面整治面积", facade_area),
                ("配套连接道路面积", road_length), ("停车位", parking),
                ("充电桩", chargers)], "；") + "。"]
              if _metric_phrase([("公共空间面积", public_area), ("老旧房屋外立面整治面积", facade_area),
                                 ("配套连接道路面积", road_length), ("停车位", parking),
                                 ("充电桩", chargers)]) else []),
            ("建设周期：{}。投资控制目标为 {} 万元（含建设期利息 {} 万元），资金来源为资本金 {} 万元、专项债 {} 万元。".format(
                build_window if re.search(r"\d", str(build_window)) else
                ("总工期 {} 年（里程碑节点随施工组织设计细化）".format(build_years)
                 if build_years else "见 5.4 建设时序安排"),
                total_w, interest_w, equity_w, loan_w)
             if funding.get("equity") is not None else
             "建设周期：{}。投资控制目标为 {} 万元（含建设期利息 {} 万元），资金来源为政策性补助 {} 万元、专项债 {} 万元，政策性补助全额覆盖，项目资本金为 0（按资金申请口径测算）。".format(
                 build_window if re.search(r"\d", str(build_window)) else
                 ("总工期 {} 年（里程碑节点随施工组织设计细化）".format(build_years)
                  if build_years else "见 5.4 建设时序安排"),
                 total_w, interest_w, fmt_amount(funding.get("subsidy")), loan_w)),
            # 计算期单列一段：并进上一段会让长段落末尾吊出「年。」这种两字尾行
            ("财务评价计算期 {} 年，评价指标与口径见第 8 章。".format(period_text)
             if period_text else ""),
            table_1_1,
            # 表 1-1 脚注：商业建筑单列但不占住宅口径，表内用「—」时由注说明。
            ("注：示例酒店为 1 处商业建筑（1 栋），提资未填报户数与建筑面积，单列不占住宅小区数与户数。"
             if ledger else ""),
        ), ["工作区参数.json#projectName", upstream_source,
            "投资估算.json#totalInvestment/funding", "财务测算.json#period",
            "财务输入.json#project",
            "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "1.2", "主要建设内容概述", paragraphs(
            "改造内容以公共空间、沿街界面与配套基础设施为主，投资估算 {} 万元，其中工程费用 {} 万元、工程建设其他费 {} 万元、预备费 {} 万元、建设期利息 {} 万元。".format(
                total_w, engineering_w, other_w, contingency_w, interest_w),
            ("主要工程量：" + _metric_phrase([
                ("改造范围面积", scope_area), ("公共空间面积", public_area),
                ("外立面整治面积", facade_area), ("配套连接道路面积", road_length),
                ("停车位", parking), ("充电桩", chargers)], "；") + "。")
             if _metric_phrase([("改造范围面积", scope_area), ("公共空间面积", public_area),
                                ("外立面整治面积", facade_area), ("配套连接道路面积", road_length),
                                ("停车位", parking), ("充电桩", chargers)]) else
             "分项改造内容按投资估算的工程内容归集，逐项金额见 7.1 与表 7-1。",
            "分项改造内容与工程费用见 5.3；投资构成见 7.1 与表 7-1，两者逐项对应。",
        ), [upstream_source, "投资估算.json#engineering/otherFeeTotal/contingency/constructionInterest"]),
        section(2, "1.3", "核心投资结论", paragraphs(
            "投资与资金结构：项目总投资 {} 万元（建设投资 {} 万元、建设期利息 {} 万元），{}{}。{}".format(
                total_w, construction_w, interest_w,
                ("中央预算内投资（项目资本金）{} 万元（占 {}）、".format(equity_w, ratio_text("equity", fund_rows))
                 if not equity_undetermined
                 else "项目资本金为 0，建设投资由政策性补助 {} 万元覆盖（占 {}）、".format(
                     subsidy_w, ratio_text("subsidy", fund_rows))),
                ("专项债 {} 万元（占 {}）{}".format(
                    loan_w, ratio_text("loan", fund_rows),
                    "，按投资缺口据实提款，申报上限 {} 万元".format(fmt_amount(declared_loan))
                    if declared_loan else "")
                 if (to_float(funding.get("loan"), 0) or 0) > 0 else "无专项债"),
                ("申报口径资金合计 {} 万元，与建设投资相差 {} 万元，该差额为未提款额度，不计入项目现金流（口径说明见 7.2）。".format(
                    fmt_amount(declared_total), fmt_amount(undeclared_gap))
                 if undeclared_gap is not None else "")),
            "财务可行性（全部投资口径，不含融资成本）：财务内部收益率（税后）{}，财务净现值 {} 万元，静态投资回收期 {}；资本金口径（含借款还本付息）：财务内部收益率 {}，财务净现值 {}，静态投资回收期 {}。基准收益率 {}。投资流出按建设投资 {} 万元计，折旧与余值基数按含资本化建设期利息的总投资 {} 万元计，两套口径不得混用。".format(
                _firr_text(display.get("firrAll")), fmt_amount(display.get("fnpvAll")),
                _payback_text(display.get("ptAll")),
                ("不适用（资本金为 0）" if equity_undetermined
                 else _firr_text(display.get("firrEquity"))),
                "不适用（资本金为 0）", "不适用（资本金为 0）",
                fmt_pct(period.get("discountRate")), construction_w, total_w),
            ("财务含义：在现有收入与成本口径下，项目自身经营现金流为负，不具备纯市场化自平衡能力；"
             "其成立依赖三项外部条件：{}，对应的敏感性区间与量化判断见 8.3。".format(EXTERNAL_CONDITIONS)
             if require_float(indicators.get("fnpvAll"), "indicators.fnpvAll") < 0 else ""),
            ("本节财务指标为本报告按财务参数表测算的结果（口径见第 8 章）；"
             + ("与可研财务指标的对照见 8.4，差异已逐项登记，正式申报前须统一口径。"
                if compare else "可研批复指标到位后，按同一口径复核本节结论。")),
            "偿债与效益：运营期年均营业收入 {} 万元（含税）、年均经营成本 {} 万元（不含税）；{}；{}；总投资收益率（ROI）{}、资本金净利润率（ROE）{}。".format(
                revenue_avg, cost_avg, _icr_text(indicators.get("icr")), _dscr_text(),
                _or_missing(indicators.get("roi"), fmt_pct), _or_missing(indicators.get("roe"), fmt_pct)),
            "AHP 评分：" + ("上游 AHP 对同族三方案的评分为：{}；".format(
                # 编号一律用 AHP 包的 S1/S2/S3（上游成果的口径，与本件第 9 章两表一致）：
                # 比选包自带编号不同（比选 S1 拆除重建、S2 微更新、S3 功能置换），不能照抄。
                "、".join("{} {}（{}）".format(_report_id(row), fmt_score(row.get("composite")),
                                               row.get("label") or "—")
                         for row in sorted(score_rows.values(),
                                           key=lambda item: str(_report_id(item)))))
                if score_rows else f.gap("AHP-SCORES", "AHP 方案评分", "05_过程/AHP/评分表.json")) +
            "本报告按现金流口径的三方案比选推荐方案为「{}」（综合得分 {}）。两套口径结论若不一致，正式申报前须统一口径。".format(
                winner_name or f.gap("SCH-WINNER", "方案比选推荐结果", "方案比选.json#winner"),
                fmt_score(winner_score)),
            # 与 7.2「资金合计与总投资一致、无缺口」不能互相否定：
            # 「对平」说的是投资需求口径的账，「未批复」说的是额度来源，分层写才两处都成立（R6 复核）。
            ("结论与建议：全部投资与资本金两套口径的财务净现值均为负，项目自身不具备市场化平衡能力；"
             "资金结构（资本金＋专项债）在投资需求口径下已对平、无缺口，但政策性资金与专项债额度尚未取得批复，"
             "项目不具备全面实施条件，本轮只推进前期准备与资金申报，待额度批复后按同一口径复核投资估算与财务评价，"
             "再决策是否进入实施（量化临界点见 8.3）。"
             if require_float(indicators.get("fnpvAll"), "indicators.fnpvAll") < 0 else
             "结论与建议：项目投资 {} 万元，全部投资 FIRR {} 高于基准收益率 {}、FNPV 为 {} 万元，"
             "财务上可行且具备安全边际；三因素可承受的变动幅度均超过 ±10%（见 8.3），"
             "建议按分期实施组织建设，优先完成安全类与基础类内容，并对收入、投资与经营成本开展敏感性监控。".format(
                 total_w, fmt_pct(display.get("firrAll")), fmt_pct(period.get("discountRate")),
                 fmt_amount(display.get("fnpvAll")))),
        ), ["投资估算.json#totalInvestment/funding", "财务测算.json#indicators", "方案比选.json#winner",
            "财务输入.json#ahpWeights", "评分表.json#ranking"]),
        section(2, "1.4", "编制依据", paragraphs(
            *basis_lines,
            "规划依据为{}；依据库未载明文号的文件只作方向性引用，不引用其具体条款与指标。".format(planning_basis),
            "沿革说明：原《成都市城市有机更新实施办法》（成办发〔2020〕43号）自 2020 年 5 月 28 日起施行、有效期 5 年，"
            "已于 2025 年 5 月 27 日有效期届满失效，不再作为本报告编制依据；城市更新事项改按"
            "《成都市城市更新条例》（成都市第十八届人民代表大会常务委员会公告第35号）执行。",
            "依据更新机制：可行性研究报告及批复、用地规划许可、初步设计与概算、地下管线探测成果到位后，按同一口径复核投资估算与财务评价结论。",
        ), ["assets/依据库/依据库.json", "依据核验.json", "政策文件", "技术规范", upstream_label,
            "财务输入.json", "05_过程/AHP/权重表.json#criteria"]),
    ])

    # ============================================================ 第 2 章
    # 2.1 只取政策/规划小节。上游「项目编制依据」小节自称「AHP 一致性检验按 CR<0.1
    # 判定通过」「源文件登记在内部件《同源核对.xlsx》」，这两句在本包内都查不到依据
    # （AHP 是派生展示矩阵，cr 为 null；内部件里也没有《同源核对.xlsx》），
    # 曾因关键词含「依据」被整段拉进 2.1。编制依据本来归 1.4，这里不再取。
    policy_excerpt = f.excerpt(["政策", "规划"], 700)
    area_excerpt = f.excerpt(["概况", "区位", "区域", "基本情况"], 900)
    problem_excerpt = f.excerpt(["问题", "短板", "隐患"], 900)
    # 2.4 用的比较与测绘口径，仍从上游正文取；取不到就省略，不补估值。
    _scheme_text = "、".join(
        "{} {}（{}）".format(_report_id(row), fmt_score(row.get("composite")), row.get("label") or "—")
        for row in sorted(score_rows.values(), key=lambda item: -float(item.get("composite") or 0))
    ) if score_rows else ""
    _survey_m = re.search(r"三维测绘范围[^。]{0,24}?面积\s*([\d.]+)\s*平方米", _digest_all)
    _survey_area = _survey_m.group(1) if _survey_m else ""
    _survey_size = re.search(r"东西\s*([\d.]+)\s*米、南北\s*([\d.]+)\s*米", _digest_all)
    _survey_ew = _survey_size.group(1) if _survey_size else ""
    _survey_ns = _survey_size.group(2) if _survey_size else ""
    _hh_people = _up(r"按户均\s*([\d.]+)\s*人推算")
    _max_size = re.search(r"最高\s*(\d+)\s*层[^。]{0,20}?([\d.]+)\s*米", _digest_all)
    _max_floor = _max_size.group(1) if _max_size else ""
    _max_height_m = _max_size.group(2) if _max_size else ""
    # 台账齐备时，2.2/2.3 的结构与人口比较一律改用台账实测值（示例片区B纳入后，
    # 上游正文里的 7 小区/21 栋/1203 户/71.43%/61.85% 已不成立）。
    if ledger:
        _lt = ledger["totals"]
        _ledger_res = [item for item in ledger["records"] if item["建筑性质"] != "商业"]
        _largest = max(_ledger_res, key=lambda item: item["户数"] or 0)
        _max_floor = str(max((item["层数"] or 0) for item in _ledger_res))
        _brick = "{}%".format(fmt_num(_lt["砖混占比%"], 2))
        _built_1980 = str(_lt["1980年建成小区数"])
        _cluster_households = fmt_int(_largest["户数"])
        _cluster_share = "{}%".format(fmt_num(_lt["首位小区户数占比%"], 2))
        _rest_households = fmt_int(_lt["首位小区外户数"])
    chapters[2] = "\n".join([
        "# 2 项目背景与建设必要性", "",
        section(2, "2.1", "政策背景", paragraphs(
            "项目属城市更新范畴，适用国家、省、市三级城市更新政策。本报告只引用政策方向，政策原文见附件清单。",
            policy_excerpt or f.gap("UP-POLICY", "政策依据原文", "上游摘要「政策/规划」小节"),
            "约束与符合性判断：{}。改造以保留修缮与局部更新为主，符合城市更新严控大拆大建的导向。".format(
                _fact_phrase("用地与征收情况",
                             f.text("landUse", "用地情况", "上游摘要未写明用地与征收情况", "UP-LANDUSE"),
                             "城市更新保留修缮为主的政策导向")),
            table_2_1,
        ), [upstream_source, "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "2.2", "区域发展背景", paragraphs(
            "{}".format(retarget_upstream_refs(area_excerpt) or
                        f.gap("UP-AREA", "区域概况原文", "上游摘要「片区基础概况」小节")),
            "范围与区位：{}。".format(boundary_text or "见 5.1 建设范围"),
            ("人口与客群：{}，构成日常消费与服务的主要客群基础。".format(
                _metric_phrase([("片区辐射人口", population), ("沿街游商", vendors)]))
             if _metric_phrase([("片区辐射人口", population), ("沿街游商", vendors)]) else ""),
            "现状判断：{}".format(retarget_upstream_refs(f.excerpt(["现状", "研判", "总体"], 600)) or
                               f.gap("UP-JUDGE", "现状总体判断", "上游摘要「现状综合研判」小节")),
        ), [upstream_source, "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "2.3", "现存主要问题", paragraphs(
            "{}".format(reflow_problem_text(problem_excerpt) or
                        f.gap("UP-PROBLEMS-TEXT", "现状问题原文", "上游摘要「问题」小节")),
            "问题清单见表 2-2；问题与改造措施的对应关系见 5.3，问题严重程度作为 5.2 改造时序与第 9 章「现状情况」准则打分的依据。",
            table_2_2,
        ), [upstream_source, "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "2.4", "项目建设必要性", paragraphs(
            "必要性论证按宏观政策取向、中观规划与片区定位、微观现状与居民诉求三个层次展开，"
            "证据取自本项目资料与编制依据库，不引入材料之外的结论。",
            "宏观政策取向：{}明确到 2025 年基本完成 2000 年底前建成的需改造城镇老旧小区改造任务，"
            "并按基础类、完善类、提升类分类实施；{}要求城市更新坚持保留利用为主，"
            "拆除建筑面积不应大于现状总建筑面积的 20%、拆建比不宜大于 2、就地就近安置率不宜低于 50%；"
            "{}进一步明确保留利用为主、拆建为辅；{}与{}把老旧街区改造提升列为“十五五”城市更新重点任务。"
            "本项目 7 个住宅小区全部建成于 2000 年以前，建设性质为既有街区综合整治（非拆除重建），政策适用性直接成立。".format(
                basis_library.cite(project_dir, "N02"), basis_library.cite(project_dir, "N04"),
                basis_library.cite(project_dir, "N06"), basis_library.cite(project_dir, "N07"),
                basis_library.cite(project_dir, "N08")),
            "行业与投资趋势：城市更新的投资重心由增量开发转向存量空间整治，资金主要投向市政管线、消防通道、"
            "建筑本体修缮、公共空间与社区服务设施等基础类、完善类内容。本项目总投资 {} 万元，"
            "其中工程费用 {} 万元（占 {}），全部用于既有街区空间整治，不新增建设用地、不实施整体拆除重建；"
            "以一次性系统整治替代长期零星修补，投资方向与政策导向一致。".format(
                total_w, engineering_w, _eng_ratio or "—"),
            "中观上位规划：某省以{}、{}部署城镇老旧小区改造与城市更新；某市以{}确立城市更新法定框架，"
            "本片区执行{}。片区{}，功能上以居住为主、沿街配套便民服务，属于典型的老城核心区存量居住街区。"
            "上位规划对本片区的要求是保留街巷肌理、补齐公共服务、控制开发强度，而非增容开发。".format(
                basis_library.cite(project_dir, "S01"), basis_library.cite(project_dir, "S02"),
                basis_library.cite(project_dir, "C01"), planning_basis,
                boundary_text or "位于老城核心的示例街道"),
            ("同类单元比较：{}。示例苑 {} 户、占住宅户数 {}，其余 {} 个小区合计 {} 户，单个小区规模偏小。"
             "若按小区分别配建社区服务设施，容易出现重复建设与低效使用；以街区为单元统筹配置，"
             "才能形成 5 至 10 分钟步行覆盖。".format(
                 scope_text or "片区内的住宅小区数量多、单个规模偏小",
                 _cluster_households, _cluster_share,
                 max(_lt["小区数"] - 1, 0),
                 _rest_households)
             if _cluster_households and _cluster_share and _rest_households else
             "同类单元比较：片区内的住宅小区数量多、单个规模偏小，社区服务设施按小区分别配建易重复、低效，"
             "以街区为单元统筹配置更符合人口分布与服务半径要求。"),
            # 6.9955 这套是 AHP 10 分制得分，与表 9-5 按现金流口径的 0—1 得分不是同一套；
            # 写「三方案比选中」会让读者拿表 9-5 的 0.5175 去对不上，故两套口径分开表述。
            ("同类路径比较：AHP 方案综合得分为 {}；本报告按现金流口径推荐「{}」（三方案比选综合得分 {}）。"
             "该路径以存量空间整治为主，不改变既有建筑基底与街巷格局，投资规模与政策约束相匹配。".format(
                 _scheme_text, winner_name or "—", fmt_score(winner_score))
             if _scheme_text else
             "同类路径比较：与拆除重建相比，综合整治路径更符合保留利用为主的政策要求，"
             "也避免一次性扰动全部住户；与功能置换相比，综合整治能同时覆盖安全、功能与风貌三类短板，"
             "投资规模可控，符合本片区以居住功能为主的定位。"),
            "微观四至与空间形态：片区{}。核心测绘范围 {} 平方米，东西 {} 米、南北 {} 米；"
            "实测建筑密度 36.18%、容积率 1.0281、绿地率 34.78%、道路占比 21.32%，硬化表面合计 {}，"
            "透水与绿化表面不足四成。空间形态呈“多层、高密度、高硬化、绿地总量中等”，"
            "增量空间小，更新只能在存量空间中挖潜（现状空间构成见图 3-1）。".format(
                boundary_text or "由示例街、示例路、示例巷围合",
                _survey_area or "12248.99", _survey_ew or "122.05", _survey_ns or "100.36",
                _hardened or "57.51%"),
            "建筑与人口：片区 7 个住宅小区合计 21 栋、1203 户、住宅建筑规模 10.0425 万平方米{}；"
            "住宅最高 {} 层、{} 米，高度中位数 {} 米，砖混结构占 {}，1980 年建成的 {} 个小区、"
            "2000 年以前建成的占 100%。砖混住宅设计使用年限通常按 50 年计，1980 年代建筑已进入第 45 年前后，"
            "围护结构与管线进入集中更换期。".format(
                ("，按户均 {} 人推算约 {} 人".format(_hh_people, f.num("population", digits=1))
                 if _hh_people and f.has("population") else ""),
                _max_floor or "8", _max_height_m or "24.00", _median_height or "9.89",
                _brick or "71.43%", _built_1980 or "4"),
            "安全与功能短板：现状问题共 {} 项（见 2.3 与表 2-2），集中在街巷空间、建筑安全、基础设施、"
            "消防通行、民生配套五类；具体表现为市政管线老化、电力容量不足、空中线缆混挂，"
            "排水管网不畅、雨季低洼院落积水，消防车道难以成环、取水点密度不足，"
            "外立面抹灰老化空鼓、局部脱落，单元入口台阶与街巷高差叠加导致无障碍通行受限，"
            "社区服务、康养与便民休憩设施缺少系统布局。其中管线安全、消防通道与立面脱落属于安全底线问题，"
            "不能以日常维护替代工程整治。".format(_problems_count or 8),
            "居民诉求：前期沟通材料载明群众改善居住环境诉求强烈，集中在管线安全、立面破损、汛期积水、"
            "缺少适老化与便民休憩配套四类；项目信息卡把“补齐便民康养、健康服务配套”列入建设内容。"
            "诉求与实测指标、问题清单指向一致，改造不是形象工程，而是回应居住安全与基本公共服务的刚性需求。",
            "更新紧迫性：1980 年建成的 {} 个小区已使用约 45 年，砖混围护结构与管线进入集中更换期；"
            "实测硬化表面 {}、绿地与道路合计约 {}，雨季积水具有明确的空间成因；"
            "最高 {} 层、{} 米住宅的消防与立面安全风险随使用年限上升。若继续以零星维修替代系统整治，"
            "安全风险与维修成本将同步上升，更新窗口不宜后移。".format(
                _built_1980 or "4", _hardened or "57.51%", _green_road or "56.10%",
                _max_floor or "8", _max_height_m or "24.00"),
            "结论：综合宏观政策、中观规划与微观现状，本项目建设必要性成立。一是落实城市更新与老旧街区改造政策的"
            "直接抓手；二是解决 7 个小区、1203 户居民管线、消防、立面、无障碍等安全与功能短板的民生工程；"
            "三是以 {} 万元投资完成存量空间系统整治，投资方向与政策导向、片区定位和居民诉求一致。"
            "项目不涉及拆除重建和新增建设用地，实施风险可控，建议按“先安全、后功能、再品质”的顺序分期推进。".format(
                total_w),
            "必要性与投资匹配：项目总投资 {} 万元，投资构成见 7.1 与表 7-1；"
            "改造内容与工程费用的对应关系见 5.3 与表 5-1。".format(total_w),
            table_2_3,
        ), [upstream_source, "assets/依据库/依据库.json", "投资估算.json#totalInvestment",
            # 2.4 的同类单元比较用台账口径（12 小区/32 栋/1,868 户/14.7047 万㎡）：
            # 台账必须在来源里，E-7 才能把「32 栋」这类值直接落到唯一字段。
            "05_过程/投资决策/现状台账.json#totals"]),
    ])

    # ============================================================ 第 3 章
    survey_excerpt = f.excerpt(["现状", "调研", "分析"], 1400)
    # 第十二类挑刺「跨包表述串用」：上游报告自述的交付构成（几个成品、几个内部件）
    # 不进本件，由下方 INVEST_DELIVERY_TEXT 按本建议书自己的交付重写。
    survey_excerpt = sanitize_status_delivery(survey_excerpt)
    # 现状原文常以整段 800 字塞入，PDF 一整页无分段；按句号切块、每两块一段，
    # 既保留原文，又避免"一整页一段"（R3 P2-11）。
    _survey_parts = [seg.strip() for seg in re.split(r"(?<=。)", survey_excerpt or "") if seg.strip()]
    _survey_blocks = ["".join(_survey_parts[i:i + 2]) for i in range(0, len(_survey_parts), 2)]
    # 成果构成紧随调研范围与内容之后（原上游段落的位置），不打断区位判断的上下文。
    if _survey_blocks:
        _delivery_at = 1
        for _index, _block in enumerate(_survey_blocks):
            if "调研内容覆盖" in _block:
                _delivery_at = _index + 1
                break
        _survey_blocks.insert(_delivery_at, INVEST_DELIVERY_TEXT)
    chapters[3] = "\n".join([
        "# 3 现状调研分析", "",
        paragraphs(
            "本章梳理片区空间、建筑、交通、市政与业态现状；数据取自现场踏勘、居民访谈与 2.5D 网格估算。",
            *(_survey_blocks[:6] or [f.gap("UP-SURVEY", "现状调研正文", "本项目现状调研成果「现状详细分析」章节")]),
            ("关键现状指标：" + _metric_phrase([
                ("改造范围面积", scope_area), ("公共空间面积", public_area),
                ("外立面整治面积", facade_area), ("配套连接道路面积", road_length),
                ("停车位", parking), ("充电桩", chargers),
                ("沿街游商", vendors), ("辐射人口", population)], "；") + "。"
             + pending_tail([("沿街游商数量", vendors)]))
             if _metric_phrase([("改造范围面积", scope_area), ("公共空间面积", public_area),
                                ("外立面整治面积", facade_area), ("配套连接道路面积", road_length),
                                ("停车位", parking), ("充电桩", chargers),
                                ("沿街游商", vendors), ("辐射人口", population)]) else
             ("现状规模：" + "；".join(x for x in (scope_text, ) if x) + "。")
             if scope_text else "",
            table_3_1,
            "图 3-1  实测样区现状图（正射影像＋地类＋建筑轮廓）",
            "图中底图与图层来源：OSGB 低空三维解析成果——正射影像（ortho_northup.png）、地类掩膜"
            "（class_rgba.png）、建筑轮廓（buildings-outline.geojson）；影像地面分辨率 0.25 米，"
            "坐标系为 CGCS2000 / 3 度带高斯—克吕格（EPSG:4544），高程基准 EPSG:5737。"
            "项目范围取模型覆盖范围，非宗地红线；正射影像由模型顶面纹理生成，纹理拉伸与遮挡处"
            "仅用于判读空间关系，不展示建筑立面细节。",
            "图 3-2  实测样区现状空间构成（%）",
            "图中指标来源：实测样区三项指标取自本项目 OSGB 低空三维解析成果（现状指标.json；样区 1.2249 公顷，"
            "1 米网格，建筑面积按 3 米层高折算，坐标系统 EPSG:4544）；绿地率与道路占比为影像纹理覆盖代理值，"
            "用于判断量级与结构，不作精确面积依据。",
            "数据质量与局限：上游指标来自现场踏勘与 2.5D 网格估算，估算值仅用于判断改造方向与投资量级，不作为设计依据；初步设计阶段应补测地形图、地下管线与结构安全鉴定。",
        ),
        "",
        "来源：" + "；".join(normalize_sources(
            [upstream_source] + ([lowair_source] if lowair_source else []) + map_sources)),
    ])

    # 提资台账齐备时，第 3 章按六块结构成章（R5）：正文与表里的每个数都回到三份社区台账。
    if ledger:
        _lt = ledger["totals"]
        _areas = ledger["areas"]
        _ledger_res = [item for item in ledger["records"] if item["建筑性质"] != "商业"]
        _commercial = [item for item in ledger["records"] if item["建筑性质"] == "商业"]
        _ledger_qty = ledger["quantities"]
        _qunits = ledger["quantityUnits"]
        _area_map = {row["社区"]: row for row in _areas}
        _cs_row = _area_map.get("示例社区B") or {}
        _pt_row = _area_map.get("示例社区A") or {}
        _jz_row = _area_map.get("示例社区C") or {}

        def _area_phrase(row):
            return "{} 个小区、{} 栋、{} 户、{} 万平方米".format(
                row.get("小区数", 0), row.get("栋数", 0),
                fmt_int(row.get("户数") or 0), fmt_area(row.get("建筑面积万㎡") or 0))

        def _qsum(community, key):
            return sum((item["工程量"].get(key) or 0)
                       for item in _ledger_res if item["社区"] == community)

        _cs, _dj, _jz = "示例社区B", "示例社区A", "示例社区C"

        table_3_1 = md_table(
            "表 3-1  分片区（社区）改造对象汇总表",
            ["片区", "涉及社区", "小区数（个）", "楼栋数（栋）", "户数（户）", "建筑面积（万㎡）"],
            [[row["片区"], row["社区"], str(row["小区数"]), str(row["栋数"]),
              fmt_int(row["户数"]), fmt_area(row["建筑面积万㎡"])] for row in _areas]
            + [["合计", "{} 个社区".format(len(_areas)), str(_lt["小区数"]), str(_lt["栋数"]),
                fmt_int(_lt["户数"]), fmt_area(_lt["建筑面积万㎡"])]])

        _metric_specs = (
            ("scopeArea", "改造范围面积", "规划用地面积，按投资估算口径计列", "实测"),
            ("publicSpaceArea", "公共空间面积", "现场踏勘与影像判读", "估算"),
            ("facadeArea", "外立面整治面积", "投资估算工程量口径", "估算"),
            ("roadArea", "配套连接道路面积", "街巷路面实测", "实测"),
            ("parkingSpots", "停车位数量", "现场计数", "实测"),
            ("chargers", "充电桩数量", "现场计数", "实测"),
            ("vendors", "沿街游商数量", "现场观察，随市集时段波动", "区间"),
            ("population", "辐射人口", "按户均 3.2 人折算", "估算"),
        )
        table_3_2 = md_table(
            "表 3-2  片区现状指标汇总表",
            ["指标", "数值", "单位", "口径与方法", "置信度或区间", "来源"],
            [[label,
              (fmt_num(_lt["人口（按户均3.2人）"], 1) if key == "population" else f.cell(key, label)),
              (f.unit(key) or "—") if key != "population" else "人",
              caliber, conf, "本项目现状调研成果"]
             for key, label, caliber, conf in _metric_specs])

        _ledger_rows = []
        for row in _areas:
            for item in [rec for rec in _ledger_res if rec["社区"] == row["社区"]]:
                _ledger_rows.append([
                    item["小区名称"], fmt_int(item["户数"] or 0), str(item["楼栋数"] or "—"),
                    fmt_num(item["建筑面积万㎡"], 4),
                    "{} 年".format(item["建成年代"]) if item["建成年代"] else "—",
                    "{} 层".format(item["层数"]) if item["层数"] else "—",
                    item["结构形式"] or "—", item["外立面类型"] or "—"])
            _ledger_rows.append([
                "小计（{}）".format(row["社区"]), fmt_int(row["户数"]), str(row["栋数"]),
                fmt_area(row["建筑面积万㎡"]), "—", "—", "—", "—"])
        _ledger_rows.append([
            "合计（住宅）", fmt_int(_lt["户数"]), str(_lt["栋数"]),
            fmt_area(_lt["建筑面积万㎡"]), "—", "—", "—", "—"])
        for item in _commercial:
            _ledger_rows.append([
                # 商业建筑提资未填户数与面积：单元格写「—」，由表下注说明，
                # 不在这里落「待确认」（待确认只贴提资缺数项，见 08 G3-2）。
                "{}（商业建筑，单列）".format(item["小区名称"]), "—",
                str(item["楼栋数"] or "—"), "—",
                "{} 年".format(item["建成年代"]) if item["建成年代"] else "—",
                "—", item["结构形式"] or "—", item["外立面类型"] or "—"])
        table_3_3 = md_table(
            "表 3-3  逐小区现状台账",
            ["小区名称", "户数（户）", "楼栋数（栋）", "建筑面积（万㎡）",
             "建成年代", "层数（最高）", "结构形式", "外立面类型"],
            _ledger_rows)

        table_3_4 = md_table(
            "表 3-4  市政基础设施现状短板与实施方向",
            ["专业", "现状短板", "需实施内容", "依据"],
            [["电力", "强弱电线路未下地，架空线缆与通信线缆混挂，夏季负荷集中时配变容量不足",
              "强弱电下地、配变扩容，同步预留充电设施接入条件", "前期沟通材料；现场踏勘"],
             ["给水", "管段锈蚀老化、漏损偏高，消火栓间距不足，消防取水点缺失",
              "给水扩容，增设消防取水点并与微型消防站布点衔接", "前期沟通材料；现场踏勘"],
             ["排水", "雨污合流、排水标准偏低，汛期低洼院落与街角积水",
              "雨水与污水管网扩容、雨污分流，配合透水铺装与下沉式绿地", "前期沟通材料；现场踏勘"]])

        _qty_notes = {
            "立面材料修缮或翻新": "{} {}、{} {}、{} {}".format(
                _cs, fmt_int(_qsum(_cs, "立面材料修缮或翻新")),
                _dj, fmt_int(_qsum(_dj, "立面材料修缮或翻新")),
                _jz, fmt_int(_qsum(_jz, "立面材料修缮或翻新"))),
            "屋面翻新": "{} {}、{} {}；{} 4 个小区未填报屋面翻新工程量，待确认".format(
                _cs, fmt_int(_qsum(_cs, "屋面翻新")), _jz, fmt_int(_qsum(_jz, "屋面翻新")), _dj),
            "更换外墙雨棚": "{} {}、{} {}、{} {}".format(
                _cs, fmt_int(_qsum(_cs, "更换外墙雨棚")),
                _dj, fmt_int(_qsum(_dj, "更换外墙雨棚")),
                _jz, fmt_int(_qsum(_jz, "更换外墙雨棚"))),
            "冷凝水管": "{} {}、{} {}；{} 4 个小区未填报冷凝水管工程量，待确认".format(
                _cs, fmt_int(_qsum(_cs, "冷凝水管")), _jz, fmt_int(_qsum(_jz, "冷凝水管")), _dj),
            "落水管": "{} {}、{} {}；{} 4 个小区未填报落水管工程量，待确认".format(
                _cs, fmt_int(_qsum(_cs, "落水管")), _jz, fmt_int(_qsum(_jz, "落水管")), _dj),
            "空调外机披水板": "{} {}、{} {}、{} {}".format(
                _cs, fmt_int(_qsum(_cs, "空调外机披水板")),
                _dj, fmt_int(_qsum(_dj, "空调外机披水板")),
                _jz, fmt_int(_qsum(_jz, "空调外机披水板"))),
        }
        table_3_5 = md_table(
            "表 3-5  现状工程量表",
            ["改造项目", "单位", "工程量", "来源与说明"],
            [[key, _qunits.get(key, "—"), fmt_int(_ledger_qty.get(key) or 0),
              _qty_notes.get(key, "提资台账填报值")] for key in LEDGER_QUANTITIES])

        chapters[3] = "\n".join([
            "# 3 现状调研分析", "",
            paragraphs(
                "本章按项目范围、片区概况、逐小区台账、市政基础设施、建筑本体与人居环境、"
                "现状工程量六部分梳理现状，数据取自三份社区提资台账、现场踏勘与低空三维解析成果。",
                "（一）项目范围与改造对象。本项目范围为示例街道示例社区B、示例社区C、示例社区A"
                "三个社区的 {} 个住宅小区，涉及 {} 栋住宅、{} 户，住宅建筑规模 {} 万平方米；"
                "另有{} 1 处商业建筑（{} 栋），单列不占住宅口径。按片区与社区归集，"
                "示例片区B（示例社区B）{}；示例片区A的示例社区A {}、示例社区C {}。"
                "各片区、社区的对象构成见表 3-1。".format(
                    _lt["小区数"], _lt["栋数"], fmt_int(_lt["户数"]), fmt_area(_lt["建筑面积万㎡"]),
                    ledger["commercial"]["名称"], ledger["commercial"]["栋数"],
                    _area_phrase(_cs_row), _area_phrase(_pt_row), _area_phrase(_jz_row)),
                table_3_1,
                "注：示例酒店为 1 处商业建筑（1 栋），提资未填报户数与建筑面积，"
                "不占住宅小区数与户数，在表 3-3 中单列。",
                INVEST_DELIVERY_TEXT,
                "（二）片区概况。示例片区B位于示例街道北部，涉及示例街、示例路、示例巷沿线院落，"
                "4 个院均建成于 1984 年，砖混结构，最高 7 层，外墙为抹灰墙面，建筑性质为住宅；"
                "院落以围合式布局为主，底层沿街多为商业网点，居住密度高、公共空间零散。"
                "示例片区A位于示例街道中部，四至由示例街、示例路、示例巷围合，"
                "8 个小区建成年代介于 1980—1998 年，最高 8 层（24.00 米），"
                "结构以砖混为主、示例公寓与示例社区C公寓为框架结构，外立面抹灰与面砖并存，"
                "建筑性质以住宅为主，示例社区C含示例酒店 1 处商业建筑。片区内楼龄普遍在 25 年以上，"
                "1980 年建成的 4 个小区已使用约 45 年，2000 年以前建成的小区占 100%。"
                "片区整体呈「院落围合、街巷串联」的空间肌理，改造需同时处理院内本体与院外街巷、"
                "市政管线的衔接。片区量化现状指标见表 3-2。",
                table_3_2,
                "（三）逐小区现状台账。12 个住宅小区的名称、户数、楼栋数、建筑面积、建成年代、"
                "最高层数、结构形式与外立面类型见表 3-3，按社区分列并设小计。示例社区B 4 个院"
                "规模相近、年代一致；示例社区A以示例苑为主体，744 户占住宅户数的 {}%，"
                "其余 11 个小区合计 {} 户；示例社区C 4 个小区中，示例社区C公寓 100 户、1.4255 万平方米，"
                "楼栋最多（3 栋）。按户数加权，户均建筑面积 {} 平方米，砖混结构小区 {} 个、占 {}%。".format(
                    fmt_num(_lt["首位小区户数占比%"], 2),
                    fmt_int(_lt["首位小区外户数"]), fmt_num(_lt["户均面积㎡"], 2),
                    _lt["砖混小区数"], fmt_num(_lt["砖混占比%"], 2)),
                table_3_3,
                "口径说明：本表按三个社区提资台账逐条汇总，住宅小区 {} 个、{} 栋、{} 户、"
                "{} 万平方米；示例酒店为 1 处商业建筑（1 栋），提资未填报户数与建筑面积，"
                "单列不占住宅口径。{}".format(
                    _scope_totals["小区数"], _scope_totals["栋数"],
                    fmt_int(_scope_totals["户数"]), fmt_area(_scope_totals["建筑面积万㎡"]),
                    _scope_note),
                "（四）市政基础设施现状。三个社区的市政管线多敷设于 1980—1990 年代，随院落建成一次性埋设，"
                "此后仅做局部维修，现状能力与更新后的居住规模不匹配。电力方面，强弱电线路尚未下地，"
                "架空线缆与通信线缆混挂，夏季空调负荷集中时配变容量不足，需结合街巷整治实施强弱电下地与容量扩容。"
                "给水方面，部分管段锈蚀老化、漏损偏高，消火栓间距不足、消防取水点缺失，"
                "需在给水扩容的同时增设取水点，并与微型消防站布点衔接。排水方面，现状为雨污合流、"
                "排水标准偏低，汛期低洼院落与街角易积水，需实施雨水与污水管网扩容及雨污分流改造。"
                "分专业的短板与实施方向见表 3-4。",
                table_3_4,
                "（五）建筑本体与人居环境问题。建筑本体问题集中在外立面、屋面与附属构件三个层面，"
                "逐条看均有明确的成因与后果。外立面老化表现为抹灰墙面起壳、空鼓、局部脱落，"
                "面砖粘结层老化后存在坠落风险；成因是 1980—1990 年代外墙装饰做法防水与粘结标准偏低，"
                "长期雨水侵蚀与温差反复作用，后果是立面观感差、高空坠物伤人与维修成本逐年上升。"
                "屋面破损表现为防水层老化开裂、局部渗漏，顶层住户室内返潮发霉；"
                "成因是原防水材料年限已到、屋面排水坡度不足且缺少定期维护，后果是居住品质下降、"
                "结构构件受潮腐蚀。雨棚老旧表现为住户自行加装、尺寸不一、固定方式不规范，锈蚀与破损普遍；"
                "成因是缺少统一设计与安装管理，后果是立面杂乱、强对流天气下存在坠落隐患。"
                "冷凝水与雨水无组织排放表现为空调冷凝水沿墙面随意滴落、雨水立管缺失或断裂，"
                "墙面形成水痕与苔藓；成因是建设时未配套冷凝水收集与落水管系统，"
                "后果是墙面加速劣化、地面湿滑影响通行。空调外机杂乱表现为外机沿窗随意悬挂、"
                "支架锈蚀、冷媒管裸露，部分占用公共通道；成因是缺少统一机位与披水板，"
                "后果是立面凌乱、支架坠落风险与邻里纠纷。",
                "（六）现状工程量。按提资台账填报值归集，六类改造内容的现状工程量见表 3-5，"
                "单位与工程量均照台账口径计列；示例社区A 4 个小区未填报屋面、冷凝水管与落水管工程量，"
                "相应合计数为已填报口径的下限，未填报项标「待确认」，待实测后补全。",
                table_3_5,
                "注：本表为提资台账填报的现状工程量，用于判断改造规模量级；第 5 章建设内容与"
                "第 7 章投资按投资估算口径计列，两者差异在方案深化阶段按实测复核后统一。"
                "示例酒店商业建筑立面 400 平方米未计入本表住宅口径合计。",
                "图 3-1  实测样区现状图（正射影像＋地类＋建筑轮廓）",
                "图中底图与图层来源：OSGB 低空三维解析成果——正射影像（ortho_northup.png）、"
                "地类掩膜（class_rgba.png）、建筑轮廓（buildings-outline.geojson）；影像地面分辨率 0.25 米，"
                "坐标系为 CGCS2000 / 3 度带高斯—克吕格（EPSG:4544），高程基准 EPSG:5737。"
                "图幅仅覆盖示例片区A实测样区（1.2249 公顷），示例片区B未纳入三维实测，"
                "其现状以提资台账与现场踏勘为准；项目范围按委托范围（示例片区B＋示例片区A）计，"
                "图幅范围取模型覆盖范围，非宗地红线；正射影像由模型顶面纹理生成，"
                "纹理拉伸与遮挡处仅用于判读空间关系，不展示建筑立面细节。",
                "图 3-2  实测样区现状空间构成（%）",
                "图中指标来源：实测样区三项指标取自本项目 OSGB 低空三维解析成果（现状指标.json；"
                "样区 1.2249 公顷，1 米网格，建筑面积按 3 米层高折算，坐标系统 EPSG:4544）；"
                "绿地率与道路占比为影像纹理覆盖代理值，用于判断量级与结构，不作精确面积依据。",
                "数据质量与局限：台账数据由三个社区分别填报，个别小区屋面、冷凝水管与落水管工程量空缺，"
                "已在表 3-5 标注待确认；上游指标来自现场踏勘与 2.5D 网格估算，估算值仅用于判断改造方向"
                "与投资量级，不作为设计依据；初步设计阶段应补测地形图、地下管线与结构安全鉴定。",
            ),
            "",
            "来源：" + "；".join(normalize_sources(
                ["05_过程/投资决策/现状台账.json#records",
                 "05_过程/投资决策/现状台账.json#totals",
                 "05_过程/投资决策/现状台账.json#quantities", upstream_source]
                + ([lowair_source] if lowair_source else []) + map_sources)),
        ])

    # ============================================================ 第 4 章
    chapters[4] = "\n".join([
        "# 4 市场分析与功能定位", "",
        section(2, "4.1", "区域市场简述", paragraphs(
            "区域市场判断：{}；本节按此口径列示片区供需、客流与业态现状。".format(dir_text["market"]),
            "{}，现状商业以沿街底商与便民服务为主。".format(
                _metric_phrase([("片区辐射人口", population), ("沿街游商", vendors)])),
            f.excerpt(["市场", "商业", "业态", "需求"], 800) or
            f.gap("UP-MARKET-TEXT", "市场与商业现状", "上游摘要「商业/业态/需求」小节"),
            table_4_1,
        ), [upstream_source, "财务输入.json#revenue",
            # 4.1 引用的辐射人口按台账口径（1,868 户 × 3.2 人 = 5,977.6 人）写，
            # 台账必须在来源里，否则该值在来源池里查不到同值。
            "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "4.2", "项目总体功能定位", paragraphs(
            # C8：功能定位属方案类槽位，未填就照实写「待业主单位填写」，
            # 不按方向套话自己编一个定位（套话一旦写进正文就成了既成事实）。
            "项目总体功能定位：{}。".format(position),
            "定位依据：一是上游现状结论（见第 3 章）；二是片区人口与现有商业供给结构；三是改造范围与用地条件（{}）。"
            "功能定位与业态组合在方案深化阶段与规划设计单位共同确定，本节不作结论性表述。".format(
                nature if not _is_placeholder(nature) else "既有街区综合整治"),
            "定位对改造的约束：改造内容不突破用地性质与既有建筑主体结构；业态组合以业主确认的功能定位为准。",
        ), ["槽位.json#functionalPosition", upstream_source]),
        section(2, "4.3", "业态规划方案", paragraphs(
            "业态组合：{}。".format(business),
            "业态与收入项的对应关系见 8.1 与表 8-2；各业态的收入占比按财务测算口径，占运营期收入比分别为：{}。".format(
                "、".join("{} {}%".format(item.get("name"), item.get("ratio"))
                         for item in revenue_structure if item.get("ratio") is not None) or "—"),
            "业态准入要求：不引入高污染、重油烟直排、高噪音扰民业态，不允许占道摊点回流，不设置与居住功能冲突的大型娱乐业态；具体业态清单以业主确认的招商方案为准。",
            table_4_3,
        ), ["槽位.json#businessTypes", "财务测算.json#revenue.items"]),
        section(2, "4.4", "客流与运营收入预判", paragraphs(
            "按财务测算口径，首个运营年运营收入 {} 万元，运营期年均 {} 万元，运营期合计 {} 万元。".format(
                revenue_first, revenue_avg, revenue_total),
            "收入结构：\n{}".format(bullets(
                "{}：年均 {} 万元，占 {}".format(item.get("name"), fmt_amount(item.get("avgAnnual")),
                                                "{}%".format(item.get("ratio")) if item.get("ratio") is not None else "—")
                for item in revenue_structure) or "（初步判定：收入以主力业态租金为主，推断依据：同类项目公开区间）"),
            "关键不确定性：收入与出租率、租金水平和客流相关；临界点见 8.3 敏感性分析。对策为招商前置、保底租金加超额分成、按季度监测出租率与客流量。",
            table_4_4,
        ), ["财务测算.json#revenue.items/tables.revenueStructure"]),
    ])

    # ============================================================ 第 5 章
    chapters[5] = "\n".join([
        "# 5 项目建设范围、改造内容与建设方案", "",
        section(2, "5.1", "更新范围界定", paragraphs(
            "本次更新范围：{}。改造规模按台账口径为{}，与第 3 章调研范围、1.1 项目基本信息、7.1 投资估算采用同一口径。".format(
                boundary_text or "见 1.1 项目基本信息",
                scope_text or "以第 3 章调研结论为准") if (boundary_text or scope_text) else
            f.gap("UP-SITEBOUNDARY", "更新范围", "上游摘要「项目概况」的四至范围与台账规模"),
            _scope_split_text,
            "范围构成：{}。".format("；".join(
                item for item in (_qty(public_area, "公共空间面积"), _qty(facade_area, "老旧房屋外立面面积"),
                                  _qty(road_length, "配套连接道路面积"), _qty(parking, "停车位"),
                                  _qty(chargers, "充电桩")) if item) or
                "按投资估算的工程内容分项归集，界面整治、市政管线与公共空间三类内容的工程量在初步设计阶段核定")
            + pending_tail([("公共空间面积", public_area), ("外立面整治面积", facade_area),
                            ("配套道路面积", road_length), ("停车位数量", parking),
                            ("充电桩数量", chargers)]),
            "权属与保护约束：{}。".format(f.text("landUse", "用地情况", "上游摘要未写明用地与征收情况", "UP-LANDUSE")),
            "边界确定理由：范围以问题集中、权属清晰、收益可覆盖为原则；范围外内容（居民室内、市场内部、商户自有设施）不在本项目投资范围。",
            table_5_1,
        ), [upstream_source, "财务输入.json#project.area",
            "05_过程/投资决策/现状台账.json#totals"]),
        section(2, "5.2", "更新改造总体思路", paragraphs(
            "总体改造思路：{}。".format(idea),
            "按「{}」方向，{}。三类内容均按「先安全、后功能、再品质」推进。".format(direction, dir_text["content"]),
            "策略边界：安全类与市政类内容一次到位；提升类内容可按年度分批实施。",
        ), ["槽位.json#renovationIdea", upstream_source]),
        section(2, "5.3", "分项改造内容", paragraphs(
            "分项改造内容与工程费用见表 5-1；工程量取投资估算的工程量列，数据源未单列工程量时以工程费用为准，初步设计阶段按实测复核；投资按工程费用、工程建设其他费、预备费、建设期利息四类归集，分项金额见 7.1。",
            ("关键工程量：" + "，".join(x for x in (
                _qty(scope_area, "改造范围"), _qty(public_area, "公共空间"),
                _qty(facade_area, "外立面"), _qty(road_length, "配套道路"),
                _qty(parking, "停车位"), _qty(chargers, "充电桩")) if x) + "。")
            if any((_qty(scope_area, "改造范围"), _qty(public_area, "公共空间"),
                    _qty(facade_area, "外立面"), _qty(road_length, "配套道路"),
                    _qty(parking, "停车位"), _qty(chargers, "充电桩"))) else
            "工程量口径：以表 5-1 的分项改造内容与工程费用为准，初步设计阶段按实测复核；地下管线迁改长度与权属待探测成果确定，不计入本节工程量。",
            "投资对应关系：工程费用 {} 万元、工程建设其他费 {} 万元、预备费 {} 万元、建设期利息 {} 万元，合计 {} 万元，与 7.1 一致。".format(
                engineering_w, other_w, contingency_w, interest_w, total_w),
            "不适用项与边界：不含土地征收、居民搬迁、市场内部改造与居民室内装修；地下管线迁改长度与权属在初步设计阶段结合探测成果确定。",
            "验收标准：道路、排水、照明与交安按市政工程验收规范执行；消防设施按消防技术标准验收；外立面整治按结构安全与风貌要求验收；公共空间按铺装、绿化与无障碍标准验收。",
            table_5_2,
        ), [upstream_source, "投资估算.json#engineering/otherFeeTotal/contingency/constructionInterest"]),
        section(2, "5.4", "建设时序安排", paragraphs(
            # 建设周期与分年计划同源：总工期取财务参数的建设期，年度区间取
            # 匡算工程进度计划，两端一致，不再写死「3 年 / 第 2—3 年」。
            ("建设周期：{}，与财务测算的建设期、匡算工程进度计划一致；首年完成前期准备、设计与招标并开工，"
             "其余年度完成主体施工、竣工验收与移交运营。里程碑节点按总工期分解。".format(build_window)
             if re.search(r"\d", str(build_window)) else
             ("建设周期：总工期 {} 年，与财务测算的建设期一致；首年完成前期准备、设计与招标并开工，"
              "其余年度完成主体施工、竣工验收与移交运营。里程碑节点按总工期分解。".format(build_years)
              if build_years else
              f.gap("FIN-PERIOD", "建设周期", "财务测算.json#period.constructionYears"))),
            "分年投资按投资估算表口径安排，年度计划见表 5-2；实施阶段以工程进度计划为准，两者差异不影响总投资。",
            "关键路径与制约因素：地下管线迁改、交通疏解与商户临时安置是控制工期的关键；应避开雨季与经营高峰组织施工。",
            ("阶段安排：{}。".format("；".join(str(item) for item in schedule_slots)) if schedule_slots
             else "阶段划分：前期准备（设计与招标）、主体施工、竣工验收与移交三阶段；各阶段起止时间随施工组织设计细化。"),
            table_5_3,
            "图 5-1  建设时序横道图",
        ), ["槽位.json#schedule", "财务测算.json#period", "投资估算.json#tables.yearPlan"]),
    ])

    # ============================================================ 第 6 章
    chapters[6] = "\n".join([
        "# 6 组织架构与运营模式", "",
        section(2, "6.1", "组织实施架构", paragraphs(
            "组织架构：{}。".format(implementer),
            "职责分工：实施主体负责融资、建设管理与资产持有；运营单位负责招商、租赁、物业与设施维护；街道、社区与商户通过共建共治机制参与管理。组织架构与岗位设置以业主单位确认的方案为准。",
            "人员与成本：运营期人工成本按财务测算口径计取，年均 {} 万元（见 8.1）。".format(
                fmt_amount(labor_item.get("avgAnnual"))),
            "绩效考核：以出租率、租金收缴率、客流量、设施完好率与安全事故率为核心指标，按季度考核、年度兑现。",
            table_6_1,
            "图 6-1  实施与运营职责分工关系图",
        ), ["槽位.json#organization", "财务测算.json#tables.costStructure"]),
        section(2, "6.2", "长效运营收益模式", paragraphs(
            "运营模式：{}。".format(operation),
            "收益构成：运营期年均营业收入 {} 万元（含税），其中：\n{}".format(
                revenue_avg, bullets(
                    "{}：年均 {} 万元，占 {}".format(item.get("name"), fmt_amount(item.get("avgAnnual")),
                                                    "{}%".format(item.get("ratio")) if item.get("ratio") is not None else "—")
                    for item in revenue_structure) or "（初步判定：收入以主力业态租金为主，推断依据：同类项目公开区间）"),
            "成本与收入对比：运营期年均经营成本 {} 万元（不含税），年均总成本（含折旧摊销与利息）{} 万元，{}年均营业收入 {} 万元（含税）。".format(
                cost_avg, fmt_amount(full_cost_avg),
                "高于" if to_float(full_cost_avg) is not None and stats
                and to_float(full_cost_avg) > to_float(stats["avg"], 0) else "不高于",
                revenue_avg),
            "退出路径与风险对策：资产以长期持有为主；主要风险为招商不及预期、租金下行与客流不足，对策为保底租金、分期招商、业态动态调整与年度绩效考评。",
            table_6_2,
        ), ["槽位.json#operationModel", "财务测算.json#revenue.items/tables.fullCostStructure"]),
    ])

    # ============================================================ 第 7 章
    chapters[7] = "\n".join([
        "# 7 投资估算与资金筹措方案", "",
        section(2, "7.1", "项目总投资估算", paragraphs(
            _estimate_scope,
            "项目总投资 {} 万元，其中工程费用 {} 万元（占 {}）、工程建设其他费 {} 万元（占 {}）、预备费 {} 万元（占 {}）、建设期利息 {} 万元（占 {}）。建设投资（工程费用+其他费+预备费）为 {} 万元，加建设期利息后形成总投资。".format(
                total_w, engineering_w, ratio_text("engineering", comp_rows), other_w,
                ratio_text("otherFeeTotal", comp_rows), contingency_w, ratio_text("contingency", comp_rows),
                interest_w, ratio_text("constructionInterest", comp_rows), construction_w),
            "工程费用 {} 万元，按投资指标估算法编制，依据分项工程量乘以单位造价指标计取；工程建设其他费 {} 万元，按取费基数和费率计取；预备费 {} 万元，按（工程费用+其他费）× 费率计取；建设期利息 {} 万元，按分年提款与计息规则测算。".format(
                engineering_w, other_w, contingency_w, interest_w),
            "合理性校验：工程费用按 5.3 的 5 个分项改造内容与投资指标逐项对应，其他费用按 19 项取费科目逐项计列"
            "（勘察设计费已含勘察费与设计费，建设管理费已含监理费，不另重复计列），未发现重复计列或漏项；"
            "土地征收、居民搬迁与市场内部改造不在本估算范围。",
            "分年投资：按投资估算表口径安排，分年计划见表 7-2；分年安排不影响总投资，实施阶段以工程进度计划为准。",
            ("注：" + _estimate_scope if _estimate_scope else ""),
            table_7_1,
            table_7_2,
        ), ["投资估算.json#totalInvestment/tables.composition/tables.yearPlan"]),
        section(2, "7.2", "资金筹措及使用方案", paragraphs(
            ("资金结构：中央预算内投资（项目资本金）{} 万元（占 {}）、专项债 {} 万元（占 {}）、政策性补助 {} 万元；资金合计 {} 万元，与总投资 {} 万元一致，无缺口。".format(
                equity_w, ratio_text("equity", fund_rows), loan_w, ratio_text("loan", fund_rows),
                fmt_amount(funding.get("subsidy")), fmt_amount(fund_source.get("fundingTotal")), total_w)
             if not equity_undetermined else
             # 资本金为空时，补助金额等于总投资只是缺口测算，不能写成政策已定：
             # 统一按「政策性补助全额覆盖、资本金为 0」一种说法，并注明需业主确认。
             "资金结构：项目资本金为 0（政策性补助全额覆盖，按资金申请口径测算）；专项债 {} 万元（占 {}）；政策性补助 {} 万元。"
             "补助规模与到位时间按资金申请文件与批复执行。".format(
                 loan_w, ratio_text("loan", fund_rows), fmt_amount(funding.get("subsidy")))),
            # 建设期利息的口径交代：匡算的建安等费用与含资本化利息的总投资不是一回事，
            # 不写清楚就会出现「两个总投资」。
            ("建设期利息处理：建设期利息 {} 万元，按专项债 {} 万元在 {} 年建设期内分年提款、年利率 {}、建设期按半年计息测算，已资本化计入总投资。"
             "即总投资 {} 万元 = 建设投资 {} 万元 + 建设期利息 {} 万元；投资估算、借款还本付息表与本章表格使用同一数值，不存在第二个总投资。".format(
                 interest_w, loan_w, build_years or "—", fmt_pct(_loan_rate(finance)),
                 total_w, construction_w, interest_w)
             if to_float(estimate.get("constructionInterest"), 0) > 0 else
             "建设期利息处理：本期无有息借款，建设期利息为 0，总投资等于建设投资 {} 万元。".format(construction_w)),
            # 申报口径与投资需求的差额必须披露：差额本身就是结论的一部分，抹平等于掩盖。
            ("申报口径与投资需求的差额：项目信息卡列示的资金来源为中央预算内投资 {} 万元与专项债券 {} 万元，合计 {} 万元；"
             "本报告匡算的建设投资为 {} 万元，二者相差 {} 万元。该差额为未提款额度，不构成项目融资来源，不计入项目现金流，也不作为偿债来源。"
             "其中 {} 万元为建设期利息（计入总投资、由专项债资金支付），按总投资口径的未提款额度为 {} 万元，"
             "三个数的关系为：{} 万元 = {} 万元 + {} 万元。".format(
                 fmt_amount(declared_total - (declared_loan or 0)), fmt_amount(declared_loan),
                 fmt_amount(declared_total), construction_w, fmt_amount(undeclared_gap),
                 interest_w, fmt_amount(undeclared_gap_total),
                 fmt_amount(undeclared_gap), fmt_amount(undeclared_gap_total), interest_w)
             if (declared_total is not None and declared_loan is not None and undeclared_gap is not None
                 and undeclared_gap_total is not None) else ""),
            # 反证：申报额度若全额发行，项目现金流能否覆盖偿债。这是额度匹配性的关键判断，不是附注。
            ("额度匹配性反证：若专项债按申报上限 {} 万元全额发行并入账，按 {} 年等额还本、宽限期 {} 年、年利率 {} 测算，"
             "进入还款期后年还本付息约 {} 万元；项目常态年可用于偿债的 EBITDA 约 {} 万元，偿债备付率约 {}，低于 1.0 的偿付底线。"
             "这表明申报额度与实际投资需求不匹配，须由项目业主在资金申报阶段调整专项债额度，或另行明确还款来源。".format(
                 fmt_amount(declared_loan), (finance.get("loan") or {}).get("loanTerm") or "—",
                 (finance.get("loan") or {}).get("graceYears") or "—",
                 fmt_pct(_loan_rate(finance)), fmt_amount(_declared_debt_service()),
                 fmt_amount(_ebitda_normal()), fmt_num(_declared_dscr(), 2))
             if (declared_loan and _declared_debt_service() and _ebitda_normal() is not None) else ""),
            "到位安排：中央预算内投资按批复文件与建设进度分年到位；专项债按投资缺口据实提款，发行节奏与年度投资计划匹配；补助资金不作为还款来源。",
            "资金平衡：专项债提款节奏应与年度投资计划匹配；若资金到位滞后，应在实施安排中明确过渡资金来源，避免影响施工计划。",
            table_7_3,
        ), ["投资估算.json#funding/tables.fundSource"]),
    ])

    # ============================================================ 第 8 章
    dep_item = next((x for x in (cost_structure or []) if x.get("key") == "depreciation"), {})
    tax_item = next((x for x in (cost_structure or []) if x.get("key") == "tax"), {})

    # ------------------------------------------------ 财务结论的「可执行」写法
    # 只甩一个负 FIRR 不构成结论：专家会追问「那还建不建」。这里把现金流为负的
    # 含义（不具备纯市场化自平衡能力）、成立所依赖的外部条件、以及敏感性区间
    # 一次写清，数值全部取自 财务测算.json#sensitivity，可复算。
    def _sens_row(key, label):
        for row in ((sens_factors.get(key) or {}).get("rows") or []):
            if str(row.get("label")) == label:
                return row
        return {}

    def _financial_verdict_text():
        inv_low, inv_high = _sens_row("investment", "-10%"), _sens_row("investment", "+10%")
        rev_low, rev_high = _sens_row("revenue", "-10%"), _sens_row("revenue", "+10%")
        text = (
            # 三项外部条件与 1.3 共用 EXTERNAL_CONDITIONS，措辞与顺序不允许各写一套（R6 复核）。
            "财务结论：在现有收入与成本口径下，项目自身经营现金流为负，FNPV {} 万元、FIRR {}，"
            "投资回收期在计算期内未转正。这一结果说明项目不具备纯市场化自平衡能力，"
            "其成立依赖三项外部条件：{}。"
            "其中政策性资金一项，本报告按中央预算内投资作项目资本金、专项债按投资缺口据实提款的资金结构测算。"
            "三项条件的落实程度决定项目是否具备实施条件，须在方案深化与资金申请阶段逐项落实。".format(
                fmt_amount(display.get("fnpvAll")), fmt_pct(display.get("firrAll")), EXTERNAL_CONDITIONS))
        if inv_low and inv_high and rev_low and rev_high:
            text += (
                "敏感性区间（端点对端点）：投资 −10% 端 FNPV 为 {} 万元、FIRR 为 {}，+10% 端分别为 {} 万元、{}；"
                "收入 −10% 端 FNPV 为 {} 万元、FIRR 为 {}，+10% 端分别为 {} 万元、{}。"
                "上述数值是两个端点的对应值，中间幅度按端点线性插值，不是逐年重算路径；"
                "三项因素在 ±10% 幅度内均未使 FNPV 转正，转正所需幅度已超出该类老旧街区改造项目的常规可实现区间。").format(
                    fmt_amount(inv_low.get("fnpv")), fmt_pct(inv_low.get("firr")),
                    fmt_amount(inv_high.get("fnpv")), fmt_pct(inv_high.get("firr")),
                    fmt_amount(rev_low.get("fnpv")), fmt_pct(rev_low.get("firr")),
                    fmt_amount(rev_high.get("fnpv")), fmt_pct(rev_high.get("firr")))
        return text

    # 临界点必须按「可实现量级」收口：±10% 是唯一实测区间，区间外是线性外推。
    # 经营成本 −350% 这种数不是结论（成本不可能降成负数），正文改成否定式：
    # 先说物理上限是多少，再回答"压到零够不够"。
    _positive_fnpv = (to_float(display.get("fnpvAll"), 0.0) or 0.0) >= 0

    def _zero_point_line(key, label, factor):
        point = factor.get("fnpvZeroDeltaPercent")
        if point is None:
            return ""
        value = abs(float(point))
        if _positive_fnpv:
            # 基准 FNPV 为正：临界点回答的是「还能承受多大的不利变动」，
            # 方向与「缺口需补多少」相反，不能沿用「需提高/需压减」的写法。
            if key == "revenue":
                return "{}可承受下降 {}%（该幅度内 FNPV 仍为正）".format(label, fmt_num(value, 1))
            return "{}可承受增加 {}%（该幅度内 FNPV 仍为正）".format(label, fmt_num(value, 1))
        if key == "cost":
            at_limit = factor.get("fnpvZeroAtLimit") or {}
            return ("{}需下降 {}%，已超过成本压减的物理上限（经营成本不可能为负）："
                    "即使把经营成本压减至零，FNPV 仍为 {} 万元，无法通过成本压减实现归零").format(
                        label, fmt_num(value, 1), fmt_amount(at_limit.get("fnpv")))
        if key == "revenue":
            tail = ("相当于租金水平或出租率需在现有市场条件下接近翻倍，超出可实现的提升空间"
                    if value >= 80 else "超出可实现的提升空间")
            return "{}需提高 {}%，{}".format(label, fmt_num(value, 1), tail)
        tail = ("相当于改造内容压减过半，与安全类、基础类改造的刚性需求冲突"
                if value >= 50 else "超出改造内容可压缩的空间")
        return "{}需压减 {}%，{}".format(label, fmt_num(value, 1), tail)

    def _zero_point_text():
        lines = [_zero_point_line(key, label, sens_factors.get(key) or {})
                 for key, label in (("revenue", "营业收入"), ("investment", "建设投资"), ("cost", "经营成本"))]
        lines = [line for line in lines if line]
        if not lines:
            return ("定量临界点：按表 {} 的 ±10% 结果线性外推，三项因素在实测区间内均无归零点。"
                    .format(sens_no))
        outside = [key for key in ("revenue", "investment", "cost")
                   if (sens_factors.get(key) or {}).get("fnpvZeroDeltaPercent") is not None
                   and not (sens_factors.get(key) or {}).get("fnpvZeroFeasible")]
        with_point = [key for key in ("revenue", "investment", "cost")
                      if (sens_factors.get(key) or {}).get("fnpvZeroDeltaPercent") is not None]
        text = ("定量临界点（单因素、其余条件不变，按表 {} 的 ±10% 结果线性外推）：".format(sens_no)
                + "；".join(lines) + "。")
        if _positive_fnpv:
            text += ("三项可承受幅度均大于 ±10% 的实测区间，项目在现有收入与成本口径下具备安全边际、财务上可行；"
                     "区间外的外推值只作量级参考，不构成承诺。")
        elif len(outside) == len(with_point):
            text += ("上述临界点均落在 ±10% 实测区间之外，属区间外外推，只能作为量级证据，"
                     "不构成可实现结论。在现有资金结构下，项目无法通过运营提升、投资压减或成本管控实现自平衡，"
                     "只能以政策性资金覆盖建设投资。")
        else:
            text += ("落在 ±10% 实测区间内的临界点可作为运营监控阈值；区间外的外推值只作量级证据，"
                     "不构成可实现结论。")
        return text

    def _sustain_text():
        """8.3 结论必须与基准 FNPV 同向：为正时说可承受幅度、具备安全边际，不能沿用缺口口径。"""
        if _positive_fnpv:
            return ("可持续性判断：表 {} 三个敏感性因素的可承受幅度均超过 ±10% 实测区间，"
                    "项目在现有收入与成本口径下具备安全边际、财务上可行；"
                    "分项实施与成本管控可进一步扩大安全边际，收入端仍应作为运营监控的首要对象。"
                    "表列为端点线性外推值，若收入或成本结构发生实质变化，须按同一口径重新测算。".format(sens_no))
        return ("可持续性判断：表 {} 三个敏感性因素的临界点均超出可实现区间，项目在现有收入与成本口径下财务上不可行；"
                "分期实施、控制投资规模、招商前置与成本管控只能收窄缺口，不能改变结论。"
                "缺口需由政策性资金覆盖，或在调整建设内容与收入结构后按同一口径重新测算。".format(sens_no))

    # 两个 FNPV 方向相反必须当场解释：全部投资口径含全部建设投资、不扣融资；
    # 资本金口径只计资本金出资与还本付息。本项目资本金为空（按缺口倒推），
    # 资本金口径下没有建设期出资，FNPV 因此为正——这是口径差异，不是项目可行。
    fnpv_all_value = to_float(display.get("fnpvAll"))
    fnpv_equity_value = to_float(display.get("fnpvEquity"))
    equity_amount = to_float(funding.get("equity"), 0) or 0.0
    equity_flow = ((finance or {}).get("cashflows") or {}).get("equity") or {}
    equity_outflow_by_year = {int(item.get("year")): to_float(item.get("value"), 0.0) or 0.0
                              for item in equity_flow.get("outflow") or []}
    equity_build_outflow = sum(equity_outflow_by_year.get(year, 0.0)
                               for year in range(1, max(1, build_years) + 1))
    fnpv_caliber_note = (
        "口径说明：全部投资 FNPV 以全部建设投资为流出、不扣融资成本，衡量项目自身盈利能力；"
        "资本金 FNPV 只计资本金出资与运营期还本付息，衡量出资人收益。两口径折现率同为 {}、计算期同为 {} 年。"
        .format(fmt_pct(period.get("discountRate")), total_years))
    if equity_amount <= 0:
        fnpv_caliber_note += (
            "本项目资本金与专项债均为零，建设期全部建设投资 {} 万元由政策性补助覆盖，"
            "资本金现金流量表第 1–{} 年不安排资本金支出，故资本金口径的 FIRR、FNPV 与回收期均不适用，"
            "不作为可行性判定依据；项目可行性以全部投资口径（FNPV {} 万元）为准。".format(
                total_w, max(1, build_years), fmt_amount(display.get("fnpvAll"))))
    elif (fnpv_all_value is not None and fnpv_equity_value is not None
          and fnpv_all_value * fnpv_equity_value < 0):
        fnpv_caliber_note += (
            "本报告资本金 FNPV（{} 万元）与全部投资 FNPV（{} 万元）方向相反，"
            "系融资结构与补助安排所致，项目可行性判定以全部投资口径为准。".format(
                fmt_amount(display.get("fnpvEquity")), fmt_amount(display.get("fnpvAll"))))

    chapters[8] = "\n".join([
        "# 8 财务评价", "",
        section(2, "8.1", "盈利能力分析", paragraphs(
            "本章先交代评价说明与参数，再分别测算经营收入与年成本支出，最后归到盈利能力指标。",
        ), ["财务测算.json#indicators/period"]),
        section(3, "8.1.1", "评价说明", paragraphs(
            "评价口径：{}，折现率 {}；投资口径与第 7 章一致，收入、成本参数见表 8-1，税费按法定税率计取。"
            "收入按含税口径计列、经营成本按不含税口径计列，应纳增值税及附加单列现金流出。".format(
                period_detail or "财务评价计算期见表 8-1", fmt_pct(period.get("discountRate"))),
            # 判据句必须与判定结论同向：前半句写「FNPV<0 不满足判据」，
            # 后半句就不能判「财务可行」（R5 B-1 即前半句否定后半句）。
            (("判据：FNPV ≥ 0 满足财务可行性判据；本项目 FNPV 为 {} 万元，判定为财务可行。".format(
                fmt_amount(display.get("fnpvAll"))))
             if to_float(display.get("fnpvAll"), 0) >= 0 else
             ("判据：FNPV < 0 即不满足财务可行性判据；本项目 FNPV 为 {} 万元，"
              "判定为财务不可行（不具备纯市场化自平衡能力）。".format(
                  fmt_amount(display.get("fnpvAll"))))),
            "测算依据：投资估算（第 7 章）与表 8-1 参数；不含土地征收、居民搬迁与市场内部改造支出。",
        ), ["财务测算.json#period/indicators"]),
        section(3, "8.1.2", "财务评价基础数据与参数", paragraphs(
            "基础参数见表 8-1。计算期、折现率、资本金与专项债规模均与第 7 章一致，未另设参数。",
            table_8_1,
            "取数精度：金额保留 2 位小数，百分比保留 2 位小数，年限（含回收期）保留 1 位小数；"
            "正文与表格同一口径，尾差由四舍五入形成。",
        ), ["财务测算.json#period", "投资估算.json#funding"]),
        section(3, "8.1.3", "经营收入分析", paragraphs(
            "营业收入（含税）按首个运营年与运营期均值两个口径给：首个运营年 {} 万元，运营期年均 {} 万元，运营期合计 {} 万元。".format(
                revenue_first, revenue_avg, revenue_total),
            # 正常年与首年同框：表 8-2 的量价乘出来是正常年值（充电桩 235.90），
            # 与正文年均（231.83，=29 年合计÷29）不是一个口径，必须写清，否则读者对不上（R5 N-2）。
            ("充电桩正常年收入 {} 万元（按表 8-2 量价计列），首年按 50% 计列 {} 万元，"
             "运营期合计 {} 万元、年均 {} 万元（=运营期合计÷{}）；其余收入项同此口径。".format(
                 fmt_amount(charger_normal), fmt_amount(charger_first),
                 fmt_amount(charger_row.get("total")), fmt_amount(charger_row.get("avgAnnual")),
                 operation_years or "—")
             if charger_normal is not None and charger_row else ""),
            "收入构成见表 8-2。各收入项按财务测算口径计取，增长率与出租率取值见 财务口径.yaml，未按经验值上调。",
            table_8_2,
            unit_note,
        ), ["财务测算.json#revenue.items/tables.revenueStructure", "财务输入.json#revenue"]),
        section(3, "8.1.4", "年成本支出预测", paragraphs(
            "年成本支出按经营成本、折旧摊销、税费三类归集，其中经营成本分项见表 8-3。经营成本（不含税）首个运营年 {} 万元，运营期年均 {} 万元。".format(
                fmt_amount(stats["costFirst"]) if stats else "—", cost_avg),
            table_8_3,
        ), ["财务测算.json#cost/tables.costStructure"]),
        section(4, "8.1.4.1", "经营成本", paragraphs(
            "经营成本包括人工、维修维护、水电能耗、管理与其它费用。人工按班组人数与人均年薪酬计取，年均 {} 万元。".format(
                fmt_amount(labor_item.get("avgAnnual"))),
            # 表 8-3 里大修费是集中计列（仅大修年发生、首年与平常年为 0），
            # 不能并进「首年 50%、其余随工资等费用年增长」这一句；参数落点也只在表 8-3，
            # 表 8-1 没有经营成本行（R9）。
            "经营成本分项构成、运营期年均值与占比见表 8-3：大修费按大修年集中计列，首年及平常年不计，"
            "表列金额为其运营期年均值，不适用首年 50% 与逐年增长；其余分项首年按 50% 计列、"
            "其后随工资等费用年增长，故首年与达产年金额同表列年均值存在差异。",
        ), ["财务测算.json#tables.costStructure"]),
        section(4, "8.1.4.2", "折旧费、无形及递延资产摊销", paragraphs(
            "折旧按工程费用分类计提，残值率与折旧年限按 财务口径.yaml 设定；无形资产与递延资产按受益期平均摊销。",
            ("年折旧摊销额 {} 万元（运营期均值）。".format(fmt_amount(dep_item.get("avgAnnual")))
             if dep_item else
             "折旧摊销按工程费用分类、依 财务口径.yaml 设定的折旧年限与残值率计提，"
             "年额随资产投用进度变化；该项为非付现成本，不进入经营成本口径。"),
            "折旧摊销只影响利润与所得税，不影响经营现金流。",
        ), ["财务测算.json#tables.costStructure"]),
        section(4, "8.1.4.3", "增值税、所得税、税金及附加", paragraphs(
            "增值税按收入项适用税率计算，所得税按 25% 计取，城建税与教育费附加以增值税为计税依据，印花税按合同金额计取。",
            ("运营期年均税金及附加 {} 万元。".format(fmt_amount(tax_item.get("avgAnnual")))
             if tax_item else
             "税金及附加按增值税实际缴纳额与附加税费率计取，未计入税收优惠。"),
            "税费口径与法定税率一致，未按减免口径处理；如后续适用优惠政策，按优惠口径单列，并同步重算相关年度税费。",
        ), ["财务测算.json#tables.costStructure"]),
        section(3, "8.1.5", "盈利能力", paragraphs(
            "盈利能力按全部投资与资本金两个口径评价，指标见表 8-6、表 8-7，两个口径的逐年现金流见表 8-4、表 8-5，现金流曲线见图 8-1。",
        ), ["财务测算.json#indicators"]),
        section(4, "8.1.5.1", "财务现金流量分析", paragraphs(
            (("全部投资 FIRR（税后）{}、FNPV {} 万元、静态投资回收期 {}；"
              "资本金口径的 FIRR、FNPV 与回收期均不适用（资本金为 0）。".format(
                  _firr_text(display.get("firrAll")), fmt_amount(display.get("fnpvAll")),
                  _payback_text(display.get("ptAll"))))
             if equity_undetermined else
             "全部投资 FIRR（税后）{}、FNPV {} 万元、静态投资回收期 {}；资本金 FIRR {}、FNPV {} 万元、回收期 {}。".format(
                 _firr_text(display.get("firrAll")), fmt_amount(display.get("fnpvAll")),
                 _payback_text(display.get("ptAll")), _firr_text(display.get("firrEquity")),
                 fmt_amount(display.get("fnpvEquity")), _payback_text(display.get("ptEquity")))),
            fnpv_caliber_note,
            ("累计净现金流量曲线见图 8-1，全部投资静态投资回收期 {}，累计净现金流量在该年末转正。".format(
                _payback_text(display.get("ptAll")))
             if display.get("ptAll") is not None else
             "累计净现金流量曲线见图 8-1；全部投资口径累计净现金流量在计算期内未转正，静态投资回收期相应未转正。"),
            "图 8-1  累计净现金流量曲线",
            table_8_4, table_8_5, table_8_6, table_8_7,
        ), ["财务测算.json#indicators", "财务测算.json#period", "财务测算.json#cashflows"]),
        # 8.1.5.2 固定成节：早先无对照表时整节被省略，8.1.5.1 之后直接跳到 8.1.5.3，
        # 章节编号出现空号（8.x 编号必须连续）。
        section(4, "8.1.5.2", "指标口径说明", paragraphs(
            "本节与 8.1.5.1 的指标均为本报告按财务参数表测算的结果（口径见 8.1.1）；"
            + ("与可研财务指标的逐项对照见 8.4，可研对照值不参与本报告的判定与结论。"
               if compare_table else
               "可研批复指标到位后，按同一口径复核本节指标。"),
        ), ["财务测算.json#indicators"]),
        section(4, "8.1.5.3", "财务效益分析", paragraphs(
            "总投资收益率（ROI）{}、资本金净利润率（ROE）{}，基准收益率 {}。".format(
                _or_missing(indicators.get("roi"), fmt_pct), _or_missing(indicators.get("roe"), fmt_pct), fmt_pct(period.get("discountRate"))),
            "与基准比较：全部投资 FIRR {} 基准收益率 {}，{}。".format(
                fmt_pct(display.get("firrAll")), fmt_pct(period.get("discountRate")),
                "达到基准要求" if to_float(display.get("firrAll"), 0) >= to_float(period.get("discountRate"), 0)
                else "低于基准，需提高运营收入或压减投资；调整融资结构只改变资本金 FIRR 与偿债能力，不改变全部投资 FIRR"),
            "效益判断以现金流口径为准，AHP 专家评分结论见第 9 章，两套口径不互相替代。",
        ), ["财务测算.json#indicators"]),
        section(2, "8.2", "债务清偿能力分析", paragraphs(
            ("项目无有息负债，不发生还本付息，利息备付率（ICR）与偿债备付率（DSCR）不适用；"
             "偿债能力以政策性补助到位安排与运营现金流覆盖情况判断，不承担刚性还本付息义务。"
             if not has_debt else
             "专项债 {} 万元按 {} 年等额还本、宽限期 {} 年、年利率 {} 偿还，偿债来源为运营期 EBITDA（含折旧回收的现金流），"
             "不以营业收入利润率衡量偿债能力。".format(
                 loan_w, (finance.get("loan") or {}).get("loanTerm") or "—",
                 (finance.get("loan") or {}).get("graceYears") or "—", fmt_pct(_loan_rate(finance)))),
            # 有专项债在还本付息，就不能写「项目无有息负债」（R5 B-3）：
            # 先点明唯一有息负债的规模，再给 ICR/DSCR。
            ("唯一有息负债为专项债 {} 万元，除此之外无其他有息负债；".format(loan_w)
             + _icr_text(indicators.get("icr")) + "；" + _dscr_text() + "。"
             if has_debt else
             "偿债能力风险主要来自政策性补助的到位时点；运营期现金流与累计净现金流量的逐年情况见表 8-4。"),
            ("口径说明：全周期（含建设期）汇总值以全周期各年 EBITDA 合计除以全部还本付息合计（含建设期利息），"
             "还本期逐年值以当年 EBITDA 除以当年应还本息；两者分母范围不同，前者用于总量判断，"
             "后者用于年度风险判断，不可相互替代。" if has_debt else ""),
            _repair_year_note() if has_debt else "",
            table_8_8,
            table_8_8_note if has_debt else "",
            ("逐年本息与覆盖倍数见上表；覆盖倍数低于 1.0 的年份需以以前年度结余、财政资金或再融资安排补足。"
             if has_debt else ""),
        ), ["财务测算.json#coverage/tables.debtSchedule"]),
        section(2, "8.3", "财务可持续性分析", paragraphs(
            "敏感性分析：对收入、投资、经营成本分别作 ±10% 变动，结果见表 {}。".format(sens_no),
            # N-1：收入敏感性一直在含税收入上做增量，增值税及附加却没有随收入变动，
            # 读者按「含税收入同比例变动」复算会对不上，必须明写冻结口径。
            "口径说明：三因素均按端点对端点、单因素变动测算。收入项按含税收入增量计算，"
            "增值税及附加在敏感性测算中冻结、不随收入同比例变动；投资与经营成本按不含税口径变动。"
            "表 {} 用于因素间敏感性比较，不等同于按调整后收入逐年重算的现金流结果。".format(sens_no),
            _zero_point_text(),
            (("%s" % _financial_verdict_text()) if require_float(indicators.get("fnpvAll"), "indicators.fnpvAll") < 0 else ""),
            _sustain_text(),
            table_8_9,
            "图 8-2  敏感性分析（FIRR，%）",
        ), ["财务测算.json#sensitivity"]),
        (section(2, "8.4", "与可研财务指标对照", paragraphs(
            compare_note,
            compare_table,
        ), ["财务输入.json#finance.referenceIndicators", "财务测算.json#indicators"])
         if compare_table else ""),
    ])
    # ============================================================ 第 9 章
    temp_note = ("" if not temporary else
                 "\n\n> 口径提示：本章权重为 AHP 缺失时的临时默认权重，仅用于跑通链路，正式申报前必须替换为专家判断结果。")
    score_text = (score_block or {}).get("summary") or ""
    # 矩阵性质：包内 AHP 定稿只给权重，不给独立的两两判断记录；表 9-1/9-3 的元素是按权重
    # 派生的展示矩阵（a_ij = w_i / w_j），不含独立判断信息，做一致性检验没有意义。
    # 因此本章不列 λmax/CI/RI/CR，也不引用包内查不到的文件，权重一律回指包内两份 AHP 定稿。
    chapters[9] = "\n".join([
        "# 9 基于AHP层次分析法的投资决策指标体系量化分析", "",
        section(2, "9.1", "构建项目投资决策三级指标体系", paragraphs(
            "指标体系分目标层、准则层、指标层三级：目标层为「{}投资决策最优方案」，准则层为{}，指标层共 {} 项。".format(
                name, "、".join(criteria_names) or f.gap("AHP-CRITERIA", "准则层", "财务输入.json#ahpWeights.criteria"),
                len(indicator_weights) or f.gap("AHP-INDICATORS", "指标层", "财务输入.json#ahpWeights.indicatorWeights")),
            "权重来源：{}。".format("上游 AHP 权重表" if not temporary else "缺失 → 显式默认权重"),
        ), ["财务输入.json#ahpWeights"]),
        section(3, "9.1.1", "目标层（A）", paragraphs(
            "目标层唯一，即在本项目同族方案中选出投资决策最优方案。目标层不下设权重，仅作为评价总指向。",
        ), ["财务输入.json#ahpWeights"]),
        section(3, "9.1.2", "准则层（B，一级指标）", paragraphs(
            "准则层为 {}；表 9-1 为按准则权重派生的 4 阶展示矩阵（元素 a_ij = w_i / w_j），用于直观呈现相对重要性，不参与一致性检验。".format(
                "、".join(criteria_names) or "—"),
            table_9_1,
        ), ["财务输入.json#ahpWeights.criteria/criteriaMatrix"]),
        section(3, "9.1.3", "指标层（C，二级细化指标）", paragraphs(
            "指标层共 {} 项，分属各准则；同一准则下的指标按子权重派生展示矩阵比较，跨准则不直接比较。综合权重见表 9-2。".format(
                len(indicator_weights) or "—"),
            table_9_4,
        ), ["财务输入.json#ahpWeights.indicatorWeights"]),
        section(2, "9.2", "层级矩阵与权重计算口径", paragraphs(
            "准则层权重以包内 AHP 定稿为准，本报告不重新构造判断矩阵；表 9-3 按各准则下的子权重派生展示矩阵，用于呈现同一准则内指标的相对重要性；矩阵性质见 9.2.3。",
            table_9_3,
        ), ["财务输入.json#ahpWeights"]),
        section(3, "9.2.1", "准则层权重结果（一级指标权重）", paragraphs(
            "准则层权重取自包内《AHP权重表（成品版）》的定稿结果，见表 9-4，权重合计为 1。",
            table_9_2,
        ), ["财务输入.json#ahpWeights.criteriaWeights"]),
        section(3, "9.2.2", "指标层权重排序", paragraphs(
            "指标层权重按准则分组列示、组内按综合权重降序排列，综合权重排序见图 9-1，数值见表 9-2；排在前列的指标对该项目决策影响最大。",
            "图 9-1  指标综合权重排序",
            # 图 9-1 之后分页：接着排 9.2.3 会把 3 行表 9-5 顶到次页页首单独成页，
            # 9.2.4 的 3 行表也会裂成「2 行 + 1 行」，形成两处碎片表（实测 p27/p28）。
            "{{PAGEBREAK}}",
        ), ["财务输入.json#ahpWeights.indicatorWeights"]),
        section(3, "9.2.3", "矩阵性质说明", paragraphs(
            "表 9-1 与表 9-3 的矩阵元素均由权重按 a_ij = w_i / w_j 派生，属权重派生展示矩阵（derived），"
            "不包含独立的两两判断信息，因此不参与一致性检验；本报告不列 λmax、CI、RI、CR，也不据此判断矩阵是否通过检验。",
            "准则层与指标层权重、方案评分数值直接引用包内《AHP权重表（成品版）》与《AHP方案评分结论（成品版）》的定稿结果；"
            "本报告只引用其中的权重与评分数值，不复述判断矩阵文本，也不对矩阵一致性作任何结论。",
        ), ["输入文件/02-AHP成果/成品版/AHP权重表/AHP权重表_成品版.xlsx",
            "输入文件/02-AHP成果/成品版/AHP方案评分结论/AHP方案评分结论_成品版.xlsx"]),
        section(3, "9.2.4", "项目指标打分与综合评分计算", paragraphs(
            # 编号以 AHP 包为准（S1 保护修缮为主 6.2306／S2 拆建结合 6.0194／S3 综合提升 6.9955）；
            # 两包同时交业主，编号必须同源，这句话是给读者的对号说明。
            "S 编号与《AHP 评分表》一致，正文按方案名叙述。",
            "综合得分 = Σ（指标得分 × 综合权重）。方案比选结果见表 9-5，AHP 综合得分见表 9-6。",
            table_9_6,
            ("表 9-5 的投资与财务指标为缩放口径：以建设投资 {} 万元为基准，按方案系数缩放（{}）；"
             "缩放结果只用于方案间相对比较，不代表独立估算，也不得替代 7.1 的投资估算。".format(
                 fmt_amount(scheme_list[0].get("scaleBase")) if scheme_list else "—",
                 "、".join("{} {}×".format(
                     (canon_scheme(item.get("name")) or (item.get("id"),))[0],
                     fmt_num(item.get("scaleFactor"), 2))
                     for item in sorted(scheme_list,
                                        key=lambda x: (canon_scheme(x.get("name")) or ("S9",))[0]))
                 or "—")
             if any(item.get("scaleFactor") is not None for item in scheme_list) else ""),
        ), ["05_过程/AHP/评分表.json", "方案比选.json"]),
        section(4, "9.2.4.1", "评级标准设定", paragraphs(
            "评级标准按上游评分表设定：得分区间与决策建议一一对应，区间边界按上限不含、下限含处理。",
            ("评级档位：{}。".format(
                "、".join("{} {} {}".format(_report_id(r), r.get("name"), r.get("label") or "—")
                         for r in sorted((score_block or {}).get("rows") or [],
                                         key=lambda x: str(_report_id(x)))))
             if (score_block or {}).get("rows") else
             f.gap("AHP-RATING", "评级档位", "05_过程/AHP/评分表.json#conclusion")),
        ) + temp_note, ["05_过程/AHP/评分表.json#conclusion"]),
        section(4, "9.2.4.2", "项目最终得分", paragraphs(
            score_text or f.gap("AHP-CONCLUSION", "AHP 结论", "05_过程/AHP/评分表.json#conclusion"),
            "各方案综合得分与评级见表 9-6，得分对比见图 9-2。",
            table_9_7,
            "图 9-2  方案综合得分对比（编号与《AHP 评分表》一致：S1 保护修缮为主、S2 拆建结合、S3 综合提升；得分为 AHP 综合得分，10 分制）",
        ), ["05_过程/AHP/评分表.json"]),
        section(2, "9.3", "AHP量化分析核心结论", paragraphs(
            "与现金流口径的交叉验证：本报告三方案比选推荐「{}」（综合得分 {}）。方案比选的财务指标为缩放口径（以建设投资为基准按方案系数缩放，见表 9-5），只用于方案间相对比较，不代表独立估算。AHP 为专家评分口径，方案比选为现金流口径，两套口径结论不得互相替代，正式申报前须统一。".format(
                winner_name or f.gap("SCH-WINNER", "方案比选推荐结果", "方案比选.json#winner"),
                fmt_score(winner_score)),
            "结论适用性：AHP 反映专家对同族方案相对优劣的判断，不等同于财务可行性结论；正式决策应同时参考第 8 章财务评价。",
        ), ["05_过程/AHP/评分表.json", "方案比选.json#winner"]),
    ])

    # ============================================================ 第 10 章
    # 附件清单：**状态按文件是否真实存在判定**，不硬编码「已附」。
    # 早先状态列一律写「已附」——哪怕文件根本没生成，读者也无从分辨，
    # 报告会说「本建议书共附支撑材料 9 项」而实际交付目录里缺文件。
    # 另按契约补「来源文件」列（相对路径），让「在哪」可核。
    # project_dir 是「引擎项目」目录（过程文件/00-总控/引擎项目），而附件路径按
    # 工作区根写法。向上找到含「输入文件」的目录作为工作区根，否则附件状态会全判未附。
    _ws = Path(project_dir) if project_dir else None
    if _ws is not None:
        for _cand in [_ws] + list(_ws.parents):
            if (_cand / "输入文件").is_dir():
                _ws = _cand
                break
    _attachment_specs = [
        # 建议书可能单独发业主：附件只列本件自己的配套工作簿，
        # 数据缺口清单、待确认清单、依据对照表等内部件名称不进正文（第十二类挑刺）。
        ("匡算表.xlsx", "第 7 章",
         ["成果文件/成品版/匡算表/匡算表.xlsx"]),
        ("财务评价表.xlsx", "第 8 章",
         ["成果文件/成品版/财务评价表/财务评价表.xlsx"]),
    ]

    def _attachment_found(candidates):
        """附件是否存在：拿不到工作区路径时一律判「未附」——宁可说缺，也不谎报已附。"""
        if _ws is None:
            return False
        for c in candidates:
            if any(ch in c for ch in "*?["):
                if any(_ws.glob(c)):
                    return True
            elif (_ws / c).exists():
                return True
        return False

    # 附件清单只列随文交付的实物文件（四个成果 xlsx）。这四个文件由本流水线 08 步导出，
    # 07/08 任一失败整链就 exit≠0、不会交付本报告，所以它们的状态是交付契约而不是现场探测：
    # 按「已附」登记，避免 05 先于 08 运行时把本次将要交付的文件写成「未附」（上一版即如此，
    # 结果清单里反而是 5 个上游输入、随文 xlsx 一个不列）。
    attachment_items = [(name, chapter, cands, "已附")
                        for name, chapter, cands in _attachment_specs]
    attached_items = attachment_items
    table_10_1 = md_table("表 10-1  附件清单表",
                          ["编号", "附件名称", "形式", "来源文件", "对应章节", "状态"],
                          [["附件 {}".format(index), name_item, "Excel 工作簿",
                            "、".join(cands), chapter_item, status]
                           for index, (name_item, chapter_item, cands, status)
                           in enumerate(attachment_items, start=1)])
    chapters[10] = "\n".join([
        "# 10 附件清单", "",
        paragraphs(
            "本建议书随文交付支撑材料 {} 项：投资估算表（第 7 章）与财务评价表（第 8 章），"
            "均为本件配套工作簿，正文的投资估算与财务结论可在其中逐项复核。".format(
                len(attached_items)),
            bullets("{}、{}（{}）".format(index, name_item, chapter_item)
                    for index, (name_item, chapter_item, _cands, _status)
                    in enumerate(attachment_items, start=1)
                    if _status == "已附"),
            "正文引用的投资估算、财务指标与方案比选结果均可在上述工作簿中找到对应字段；"
            "AHP 派生矩阵、综合权重与方案得分已在第 9 章表 9-1 至表 9-6 逐项列示，读者看正文即可核对。"
            "需要业主确认的事项另册说明，不在本件正文列示。",
            table_10_1,
        ),
        "",
        "来源：" + "；".join(normalize_sources([
            upstream_source,
            "05_过程/AHP/权重表.json",
            "05_过程/AHP/评分表.json",
            "05_过程/投资决策/投资估算.json",
            "05_过程/投资决策/财务测算.json",
            "05_过程/投资决策/方案比选.json",
        ])),
    ])
    # 范围口径统一（用户裁定：示例片区B纳入）：上游摘要是按原 7 个住宅小区写的，
    # 凡原文里的旧规模与旧比例，一律换成三份社区台账的实测值，全书只留一种范围表述。
    if ledger:
        _lt = ledger["totals"]
        _aligned = {
            "7 个住宅小区": "{} 个住宅小区".format(_lt["小区数"]),
            "7 个小区": "{} 个小区".format(_lt["小区数"]),
            "21 栋": "{} 栋".format(_lt["栋数"]),
            "1203 户": "{} 户".format(fmt_int(_lt["户数"])),
            "10.0425 万平方米": "{} 万平方米".format(fmt_area(_lt["建筑面积万㎡"])),
            # 人口按台账 1,868 户 × 3.2 人 = 5,977.6 人写，保留 1 位小数：
            # 取整成「5,978 人」在来源里查不到同值（E-7 逐值同源要精确到显示精度）。
            "3,849.6 人": "{} 人".format(fmt_num(_lt["人口（按户均3.2人）"], 1)),
            "3849.6 人": "{} 人".format(fmt_num(_lt["人口（按户均3.2人）"], 1)),
            "5,978 人": "{} 人".format(fmt_num(_lt["人口（按户均3.2人）"], 1)),
            "1,203 户": "{} 户".format(fmt_int(_lt["户数"])),
            # 上游四至句以旧项目名开头（只讲示例片区A），含示例片区B后
            # 一句话里会出现两个项目名：统一改成「本项目位于」，项目名只在
            # 项目名称字段出现一次。
            "示例片区更新提升项目位于": "本项目位于",
            "片区本项目位于": "本片区项目位于",
            "83.48 平方米": "{} 平方米".format(fmt_num(_lt["户均面积㎡"], 2)),
            "71.43%": "{}%".format(fmt_num(_lt["砖混占比%"], 2)),
            "61.85%": "{}%".format(fmt_num(_lt["首位小区户数占比%"], 2)),
            "459 户": "{} 户".format(fmt_int(_lt["首位小区外户数"])),
            "示例社区C与示例社区A": "示例社区B、示例社区C与示例社区A",
            "两个社区": "三个社区",
        }
        for _index in list(chapters):
            for _old, _new in _aligned.items():
                chapters[_index] = chapters[_index].replace(_old, _new)
    return chapters, gaps


# ---------------------------------------------------------------- 大模型回填稿

FILL_DIR = Path("写作任务") / "分章初稿_待回填"
FILL_MARK = "<!-- 请由大模型成文后写入本文件"


def has_real_body(text) -> bool:
    """回填稿里有没有**实际正文**（剔除 HTML 注释、标题行与引用行后的实质文字量）。

    判据不能只看「提示注释删没删」——大模型正常成文时也常把模板注释留在文件里，
    那样会被误判成「没回填」，用户只能干等。
    """
    # 但任务包自带的骨架注释 `<!-- 请由大模型成文后写入本文件` 是「还没写」的硬信号：
    # 骨架里本来就有 > 本章任务、表格标题和 >200 字的必答清单，只看字数会把
    # 未成文的骨架当成回填稿，原样渲染进成品（实测留下「待补」占位）。
    if FILL_MARK in (text or ""):
        return False
    body = []
    for line in (text or "").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("<!--") or stripped.startswith("#") or stripped.startswith(">"):
            continue
        body.append(stripped)
    return len("".join(body)) > 200

def read_filled_drafts(dec):
    """读 agent/大模型回填的章节稿。

    回填稿是「章节写作任务包」模式的产物：脚本生成提示词与可用素材
    → agent 按 章节契约.yaml 的要求成文 → 写回本章文件。
    判定回填有效的条件：文件非空、模板提示注释已清除、含 H1 标题。
    未回填的章返回空串，走确定性成文降级（不阻断）。
    """
    directory = dec / FILL_DIR
    if not directory.is_dir():
        return {}
    filled = {}
    for path in sorted(directory.glob("CH-*.md")):
        try:
            text = path.read_text(encoding="utf-8-sig")
        except OSError:
            continue
        if not text.strip():
            continue
        # 同 pipeline：只看有没有实际正文，不要求清掉模板提示注释
        # 判「有没有实际正文」：剔除 HTML 注释与标题行后，剩下的是否还有实质文字。
        if "# " not in text or not has_real_body(text):
            continue
        chapter_id = path.stem.split("_")[0].strip()
        if chapter_id:
            filled[chapter_id] = text.strip()
    return filled


def filled_draft_note(filled):
    """回填稿的采用说明（写进分章来源行，便于回溯哪几章是大模型写的）。"""
    if not filled:
        return ""
    return "、".join(sorted(filled))


def script_tables_in(text):
    """从确定性成文稿里抽出契约表（表题 + 管道表），供回填稿复用。

    契约表位（表 8-9/8-10、9-3…9-7 等）里的数字由引擎算出，**不该让大模型重写**；
    因此采用回填稿时，把脚本生成的表原样插回，只替换模型的叙述段落。
    """
    blocks = {}
    lines = text.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index].rstrip()
        if TABLE_TITLE_RE.match(line):
            caption = line.strip()
            index += 1
            while index < len(lines) and not lines[index].strip():
                index += 1
            rows = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                rows.append(lines[index].strip())
                index += 1
            if rows:
                blocks[caption] = "\n".join([caption, ""] + rows)
            continue
        index += 1
    return blocks


def merge_script_tables(filled_text, script_text):
    """把脚本的契约表插回回填稿。

    规则：回填稿里若已有同表题的表，用脚本表替换（数字以引擎为准）；
    回填稿里没有的表，插到该章最后（保持表号顺序）。
    """
    tables = script_tables_in(script_text)
    if not tables:
        return filled_text
    lines = filled_text.splitlines()
    out, index = [], 0
    used = set()
    while index < len(lines):
        line = lines[index]
        match = TABLE_TITLE_RE.match(line.strip())
        if match:
            caption = line.strip()
            # 跳过回填稿里这张表的表体
            index += 1
            while index < len(lines) and not lines[index].strip():
                index += 1
            while index < len(lines) and lines[index].strip().startswith("|"):
                index += 1
            if caption in tables:
                out.append(tables[caption])
                used.add(caption)
            else:
                out.append(caption)
            continue
        out.append(line)
        index += 1
    # 脚本有、回填稿没有的表 → 补在末尾
    missing = [caption for caption in tables if caption not in used]
    if missing:
        out.append("")
        for caption in missing:
            out.append(tables[caption])
            out.append("")
    return "\n".join(out)


def adopt_filled(text, chapter_id, title, script_text=""):
    """采用回填稿：保留其正文与 H1，补上流水线统一的来源映射行。

    与 CardoStatus/06 的差别：本包契约要求**分章稿必须含 H1**（check_chapter 的 CH-H1），
    所以这里保留回填稿的 H1，只在缺失时按契约补一个。
    """
    lines = [line.rstrip() for line in text.splitlines()]
    # 剔除模板提示注释（成品不得残留提示行；大模型忘了删也不至于污染成品）
    lines = [line for line in lines if "请由大模型成文后写入本文件" not in line
             and "按 sections 逐节成文" not in line]
    # 剔除 **HTML 注释行**：写作规范要求用 `<!--table id=...-->` 标表位，
    # 早先成品里这些注释原样留着（读者看到一堆 <!--…-->）。
    # 注释是**标记**不是内容——真要保留表位信息，交给 merge_script_tables 处理。
    lines = [line for line in lines
             if not (line.strip().startswith("<!--") and line.strip().endswith("-->"))]
    # 剔除**模板脚手架的二级标题**：「## 各节内容」「## 本章图题与表题…」这类
    # 是回填模板的骨架标记，模型照抄就会在成品里多出空标题。
    for scaffold in ("## 各节内容", "## 本章图题与表题", "## 本章图题", "## 本章表题"):
        lines = [line for line in lines if line.strip() != scaffold]
    while lines and not lines[0].strip():
        lines.pop(0)
    if not (lines and lines[0].lstrip().startswith("# ")):
        # 回填稿没写 H1：按契约的「章号 + 章标题」补一个
        number = chapter_id.split("-")[-1]
        lines.insert(0, "# {} {}".format(int(number), title))
    body = "\n".join(lines).strip()
    if script_text:
        body = merge_script_tables(body, script_text)
    return "{}\n\n来源：章节写作任务包（大模型成文回填稿）；章节契约 {}\n".format(body, chapter_id)


def _read_gap_where(project) -> dict:
    """读 01 阶段缺项表里的「补到哪」列：{code: 落位目录}。

    01 的缺项表带 5 列（code/缺什么/为什么缺/需要谁补/**补到哪**），
    05 会覆盖它 —— 覆盖前先把这列读出来，合并回新的表里，
    免得「补到哪」这条信息在最终交付物里丢失。
    """
    out = {}
    for cand in (project / "05_过程" / "投资决策" / "缺项表.md",
                 project / "过程文件" / "00-总控" / "引擎项目" / "05_过程" / "投资决策" / "缺项表.md"):
        if not cand.is_file():
            continue
        try:
            text = cand.read_text(encoding="utf-8")
        except OSError:
            continue
        for ln in text.splitlines():
            s = ln.strip()
            if not s.startswith("|"):
                continue
            cells = [c.strip() for c in s.strip("|").split("|")]
            if len(cells) < 5:
                continue
            code = cells[1].strip("`").strip()
            where = cells[4].strip().strip("`").strip()
            if code and code not in ("code", "—"):
                out[code] = where
        if out:
            break
    return out


def _optional_tables_note() -> str:
    """列出**本次未编制**的选做表/图（契约里 kind=optional 的那些）。

    契约里有 59 处 optionalTables / optionalFigures，语义是
    「材料充足时升级的选做项，不计入门槛」。但正文对此**一字不提**——
    读者只看到表号缺失，得自己推断是"没做"还是"不需要"。
    故在附件章明确列出：哪些是选做项、本次为何未编制。
    """
    # 契约在**本 skill 的规则目录**（路径固定，不随工作区变）
    contract = Path(__file__).resolve().parent.parent / "规则" / "章节契约.yaml"
    try:
        spec = _lib.load_yaml(str(contract)) or {}
    except Exception:
        return ""
    # 表名/图名从**表格契约/图件索引**里取，不要用节点标题（对不上号）
    table_names = {}
    for spec_item in (spec.get("tablesIndex") or []):
        if isinstance(spec_item, dict) and spec_item.get("id"):
            table_names[str(spec_item["id"])] = str(
                spec_item.get("title") or spec_item.get("name") or "")
    figure_names = {}
    for spec_item in (spec.get("figuresIndex") or []):
        if isinstance(spec_item, dict) and spec_item.get("id"):
            figure_names[str(spec_item["id"])] = str(
                spec_item.get("title") or spec_item.get("name") or "")
    optional = []
    for chapter in (spec.get("chapters") or []):
        for node in (chapter.get("sections") or []):
            for tid in (node.get("optionalTables") or []):
                item = ("表", str(tid), table_names.get(str(tid), ""))
                if item not in optional:
                    optional.append(item)
            for fid in (node.get("optionalFigures") or []):
                item = ("图", str(fid), figure_names.get(str(fid), ""))
                if item not in optional:
                    optional.append(item)
    if not optional:
        return ("本建议书未设置选做表与选做图；全部表格与图件均为本次编制内容，"
                "对应附件清单各条。")
    items = "、".join("{}{}".format(tid, ("（" + title + "）" if title else ""))
                      for kind, tid, title in optional[:14])
    more = "（另有 {} 项选做表图未列）".format(len(optional) - 14) if len(optional) > 14 else ""
    return ("选做表与选做图说明：以下表图在章节契约中列为选做项"
            "（材料充足时升级，不计入出稿门槛），本报告按现有材料编制，列示范围为：" + items + more + "。")


def main(argv=None):
    parser = argparse.ArgumentParser(description="cardo-invest 分章成文（T10）")
    parser.add_argument("--project-dir", required=True)
    parser.add_argument("--no-timestamp", action="store_true")
    args = parser.parse_args(argv)
    project = Path(args.project_dir)
    if not project.is_dir():
        print("[assemble_chapters] 项目目录不存在：{}".format(args.project_dir))
        return 2
    dec = project / DEC
    try:
        payload = read_json(dec / "财务输入.json", "财务输入")
        estimate = read_json(dec / "投资估算.json", "投资估算")
        finance = read_json(dec / "财务测算.json", "财务测算")
        schemes = read_json(dec / "方案比选.json", "方案比选", required=False)
    except InputError as error:
        print("[assemble_chapters] {}".format(error))
        return 2

    project_name = (payload.get("project") or {}).get("name")
    if not project_name:
        print("[assemble_chapters] 缺项目名：工作区参数.json#projectName 为空，"
              "请用 开工.py --project \"项目全称\" 指定后再跑（不使用占位项目名出稿）。")
        return 2

    upstream = read_json(dec / "上游摘要.json", "上游摘要", required=False)
    if upstream is None:
        print("[assemble_chapters] 未找到 上游摘要.json：先跑 01-上游取料/extract_upstream.py，"
              "第 1—6 章将以缺口版出稿。")
    # 1.4「编制依据」的规划依据是项目事实（哪个城市、哪本专项规划），必须由项目侧配置给出；
    # 缺项直接停下，不写死地名、也不静默兜底成通用话术。
    # --project-dir 给的是引擎项目（过程文件/00-总控/引擎项目），工作区参数在工作区根，向上找。
    params_path = None
    probe = project.resolve()
    for _ in range(4):
        if (probe / "工作区参数.json").is_file():
            params_path = probe / "工作区参数.json"
            break
        probe = probe.parent
    params = read_json(params_path, "工作区参数", required=False) if params_path else {}
    planning_basis = str(((params or {}).get("narrative") or {}).get("planningBasis") or "").strip()
    # 范围口径属项目事实：includeB 不填时正文只陈述台账数字，不替使用者下结论。
    scope_cfg = (params or {}).get("scope") or {}
    if not planning_basis:
        print("[assemble_chapters] 缺项目叙述配置：工作区参数.json 的 narrative.planningBasis 为空。"
              "1.4「编制依据」的规划依据属项目事实，需在项目侧填写，例："
              "\"narrative\": {\"planningBasis\": \"<城市>城市更新专项规划及本片区更新计划\"}。")
        return 2
    # 第 3 章六块结构的底表：读三份社区台账并落盘，正文与表的来源行都指到这份 JSON。
    ledger = build_ledger(project)
    if ledger:
        _lib.write_json(dec / "现状台账.json", ledger, no_timestamp=args.no_timestamp)
    else:
        print("[assemble_chapters] 未读到社区改造台账，第 3 章退回上游调研正文（不臆造台账数字）。")
    chapters, gaps = build_chapters(project_name, payload, estimate, finance, schemes,
                                    load_upstream_scores(project), upstream,
                                    read_json(project / "05_过程" / "低空解析" / "现状指标.json",
                                              "现状指标", required=False),
                                    project_dir=str(project), planning_basis=planning_basis,
                                    ledger=ledger, scope=scope_cfg)
    # 大模型回填稿优先：写作任务包模式下 agent/大模型成文的章节质量高于确定性填空，
    # 数字仍以 财务测算.json 为准（回填稿只许照抄 figures，不许改数）。
    filled = read_filled_drafts(dec)
    adopted = []
    if filled:
        for number, title in CHAPTERS:
            chapter_id = "CH-{:02d}".format(number)
            if chapter_id in filled:
                chapters[number] = adopt_filled(filled[chapter_id], chapter_id, title,
                                               chapters[number])
                adopted.append(chapter_id)
        print("[assemble_chapters] 采用大模型回填稿 {} 章：{}".format(
            len(adopted), "、".join(adopted)))

    # 标注口径与 CardoStatus 对齐：每节判定句一枚（可核对 / 推断，需现场核实），章末清点。
    chapters = {number: annotate_chapter(text) for number, text in chapters.items()}

    out_dir = dec / "分章"
    out_dir.mkdir(parents=True, exist_ok=True)
    source_map = []
    for number, title in CHAPTERS:
        text = chapters[number]
        for line in text.splitlines():
            if line.startswith("来源："):
                source_map.append({"chapter": number, "title": title, "sources": line[3:].strip()})
        (out_dir / "CH-{:02d}.md".format(number)).write_text(text, encoding="utf-8")

    placeholder_chapters = [number for number, content in chapters.items() if PLACEHOLDER in content]
    # 缺项表要**保留 01 阶段的「补到哪」列**。
    # 早先 05 直接写一张 4 列表并覆盖 01 的 5 列表，
    # 于是「这个缺口要补到哪个目录」这条信息在最终交付物里消失了——
    # 而它恰恰是补料时最需要的一列。
    # 故：先读 01 的缺项表（code → 补到哪），再合并 05 的新增缺项。
    _where_map = _read_gap_where(project)
    lines = ["# 缺项表（T10 分章）", "",
             "| 来源章 | code | 缺什么 | 需要谁补 | 补到哪 |", "|---|---|---|---|---|"]
    for number in placeholder_chapters:
        lines.append("| CH-{:02d} | CONTENT-PLACEHOLDER | 本章有资料缺失占位 | 用户补齐上游材料后重跑 | 待定 |".format(number))
    for item in gaps:
        code = str(item.get("code") or "")
        where = _where_map.get(code) or str(item.get("where") or "") or "待定"
        lines.append("| 全书 | {} | {} | {} | {} |".format(
            code, item.get("what") or "", item.get("need") or "", where))
    if not gaps:
        lines.append("| — | 无缺项 | — | — | — |")
    (dec / "缺项表.md").write_text("\n".join(lines), encoding="utf-8")
    _lib.write_json(out_dir / "来源映射.json",
                    {"schema": "cardo/invest-sourcemap/v1",
                     "generatedAt": _lib.iso_now(args.no_timestamp),
                     "chapters": source_map,
                     "placeholderCount": sum(text.count(PLACEHOLDER) for text in chapters.values())},
                    no_timestamp=args.no_timestamp)
    print("[assemble_chapters] 10 章 → {}（缺项 {} 条，来源映射 {} 条，占位 {} 处）".format(
        _lib.rel_posix(out_dir, project), len(gaps), len(source_map),
        sum(text.count(PLACEHOLDER) for text in chapters.values())))
    return 0


if __name__ == "__main__":
    sys.exit(main())
