#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_evidence_register.py -- 把分章稿的「来源：」行整理成《依据对照表》。

分章稿（CH-01…CH-10.md）是**标注源**：正文段落/表格后面跟一行或多行
`来源：<文件>#<表/字段>；<文件>#<键>`。本脚本逐块登记：
  章 / 节 / 段序 / 类型（正文·表格·图）/ 数值 / 来源引用 / 解析状态。
并核对每条引用能不能真落到文件与 JSON 指针上——"写得出处"与"查得到处"是同一件事。

输出：
  过程文件/05-章节成文/分章/依据对照表.json   （schema cardo/evidence-register/v1）
退出码：0 登记完成（有问题也返回 0，由 07 回读按覆盖率/未解析数判定）；
        2 分章目录不存在。
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

NUM_RE = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(万元|亿元|万㎡|平方米|㎡|%|年|户|人|个|处|家|栋|米|倍)?")
# 交叉引用与标准号不是数据：表 7-1 / 第 7 章 / 见 5.3、4.2 / GB 50500—2024 / 〔2016〕1 号。
# 先剔掉再抽数值，否则对照表会把「见 5.3」登记成一个查不到出处的数（07 会误拦）。
CROSS_RE = re.compile(
    r"(?:表|图)\s*\d+\s*[-–]\s*\d+"
    r"|第\s*\d+(?:\.\d+)?\s*[章节条款项]"
    # 「见 9 章」「见 5.3 节」这类不带「第」的章号也是交叉引用，不是数据。
    r"|(?<![\d.,])\d{1,2}(?:\.\d+)?\s*[章节](?!\s*[、,，])"
    r"|\d+\.\d+(?=\s*(?:节|章|条))"
    r"|(?:见|详见|参见|对照|与|据|按|向|作为)\s*\d{1,2}\.\d{1,2}(?:\s*[、,，]\s*\d{1,2}\.\d{1,2})*"
    # 「1.1 项目基本信息、7.1 投资估算」这类节号列表：一位小数、两位以内整数，
    # 且后面不跟单位或百分号——数据值一般带单位或两位小数，不会被误伤。
    r"|(?<![\d.,])(?:[1-9]|1[0-2])\.\d(?![\d%])(?!\s*(?:万|亿|元|㎡|个|处|人|户|家|栋|米|倍|年))"
    r"|(?:按|据)\s*\d+(?:\.\d+)?\s*%"
    r"|(?:GB|GB/T|CJJ|JGJ|CJ/T|DG|DB|JG)\s*[/T]*\s*\d+(?:\s*[—–-]\s*\d{4})?"
    r"|〔\s*\d{4}\s*〕\s*\d+\s*号"
    r"|(?:19|20)\d{2}\s*年"
    # 坐标系编号与摘要算法不是数据：「EPSG:4544」「sha256」曾被当成无源数值。
    r"|EPSG\s*[:：]?\s*\d{4,5}"
    r"|sha-?\d+"
    # 月日与「公告第N号」也是引用不是数据（与 07 回读同一口径）。
    r"|\d{1,2}\s*月\s*\d{1,2}\s*日"
    r"|公告\s*第\s*\d+\s*号")
SRC_RE = re.compile(r"^来源\s*[:：]\s*(.+)$")
HEAD_RE = re.compile(r"^(#{1,4})\s+(.+)$")
TABLE_RE = re.compile(r"^\|")
FIG_RE = re.compile(r"^(图|表)\s*\d+\s*[-–]\s*\d+")


def read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def numbers_in(text):
    text = CROSS_RE.sub(" ", text or "")
    out = []
    for m in NUM_RE.finditer(text):
        raw = m.group(1)
        start = m.start(1)
        sign = ""
        if start > 0 and text[start - 1] in "-−－":
            before = text[start - 2] if start >= 2 else ""
            if not before.isdigit():      # 「2-3 年」是区间，不是负数
                sign = "-"
        value = sign + raw.replace(",", "")
        try:
            float(value)
        except ValueError:
            continue
        out.append(value + (m.group(2) or ""))
    return out


def split_refs(text):
    return [part.strip() for part in re.split(r"[；;]", text or "") if part.strip()]


def resolve_file(ref_path, workspace):
    """把来源里的文件路径落到真实文件：先按工作区相对路径，再按引擎项目暂存目录，
    最后按文件名在工作区内找（会标 fallback，读的人知道不是原路径）。"""
    if os.path.isabs(ref_path):
        return ref_path if os.path.isfile(ref_path) else ""
    candidates = [os.path.join(workspace, ref_path),
                  os.path.join(workspace, "过程文件/00-总控/引擎项目", ref_path)]
    for cand in candidates:
        if os.path.isfile(cand):
            return cand
    base = os.path.basename(ref_path)
    for root, _dirs, files in os.walk(workspace):
        if base in files:
            return os.path.join(root, base)
    return ""


def json_at(path, pointer):
    """按 a.b.c 取 JSON 节点；取不到返回 (None, False)。"""
    try:
        with open(path, encoding="utf-8") as f:
            node = json.load(f)
    except (OSError, ValueError):
        return None, False
    for key in [k for k in pointer.split(".") if k]:
        if isinstance(node, list):
            try:
                node = node[int(key)]
            except (ValueError, IndexError):
                return None, False
        elif isinstance(node, dict) and key in node:
            node = node[key]
        else:
            return None, False
    return node, True


def find_by_id(node, target, depth=0):
    """按 id 字段深搜（上游摘要用 sections[].id = S11 这类编号做定位）。"""
    if depth > 6:
        return None
    if isinstance(node, dict):
        if str(node.get("id") or "") == target:
            return node
        for value in node.values():
            hit = find_by_id(value, target, depth + 1)
            if hit is not None:
                return hit
    elif isinstance(node, list):
        for item in node:
            hit = find_by_id(item, target, depth + 1)
            if hit is not None:
                return hit
    return None


def check_ref(ref, workspace):
    """校验一条引用，返回 (状态, 说明, 解析到的文件)。"""
    path_part, _, pointer = ref.partition("#")
    path_part, pointer = path_part.strip(), pointer.strip()
    if not path_part:
        return "无效", "引用为空", ""
    if "/" not in path_part and not re.search(r"\.\w{2,5}$", path_part):
        # 「章节写作任务包」「章节契约 CH-01」这类是成文来历说明，不是文件引用
        return "说明", path_part, ""
    pointer = pointer.split()[0] if pointer else ""   # 允许写成 #S11 标题，尾部标题不参与定位
    real = resolve_file(path_part, workspace)
    if not real:
        return "文件缺失", path_part, ""
    rel = os.path.relpath(real, workspace)
    fallback = "" if os.path.abspath(os.path.join(workspace, path_part)) == os.path.abspath(real) \
        else "（原路径无此文件，按文件名解析到 {}）".format(rel)
    if pointer and path_part.lower().endswith(".json"):
        node, ok = json_at(real, pointer)
        if not ok:
            try:
                with open(real, encoding="utf-8") as f:
                    ok = find_by_id(json.load(f), pointer) is not None
            except (OSError, ValueError):
                ok = False
            if not ok:
                return "指针缺失", "{}#{}{}".format(rel, pointer, fallback), real
            return "可解析", "{}#{}{}（按 id 定位）".format(rel, pointer, fallback), real
        return "可解析", "{}#{}{}".format(rel, pointer, fallback), real
    return "可解析", rel + fallback, real


# ---------------------------------------------------------------- 逐值绑定
# 一个数值只能挂一个「文件#字段」：早先按节共用一个来源池，CH-01 1.3 的四个 AHP 分数
# 全被记到 投资估算.json#totalInvestment（值是 23772.09），打开 JSON 就能看穿。
VALUE_FILES = [
    "过程文件/02-匡算/投资估算.json",
    "过程文件/03-财务评价/财务测算.json",
    "过程文件/04-三方案比选/方案比选.json",
    "过程文件/01-上游取料/财务输入.json",
    "过程文件/00-总控/引擎项目/05_过程/AHP/权重表.json",
    "过程文件/00-总控/引擎项目/05_过程/AHP/评分表.json",
    "过程文件/01-上游取料/上游摘要.json",
    # OSGB 实测指标：现状分析定稿的唯一数值出处，不索引就只能靠引述原文核对。
    "过程文件/00-总控/引擎项目/05_过程/低空解析/现状指标.json",
]
TOKEN_RE = re.compile(r"^(-?[\d,]+(?:\.\d+)?)(万元|亿元|万㎡|平方米|㎡|%|年|户|人|个|处|家|栋|米|倍)?$")
# 幅度方向词：正文写「压减 57.7%」而上游存 -57.69 时，按绝对值绑定的触发条件。
DIRECTION_RE = re.compile(r"提高|提升|压减|压缩|下降|降低|上升|增加|减少|减幅|降幅|增幅")
# 上游 JSON 的数值多数以字符串存（"13074.65" / "-0.009" / "3.5%"），只摊平 int/float 会漏掉
# 整份方案比选与财务测算，导致正文值一个都绑不上。
STR_NUM_RE = re.compile(
    r"(-?\d[\d,]*(?:\.\d+)?)\s*(万元|亿元|万㎡|平方米|㎡|%|年|户|人|个|处|家|栋|米|倍)?")
# 元数据键不是可核对字段：schemaVersion=1 曾被当成「100%」「1年」「1项」的出处，
# 逐数明细里 7 行因此指到版本号上。
# 必须按「整段相等」判定：早先写成子串匹配，「engineering」里的 engine 被误伤，
# 整张分年投资表（engineering/otherFees/contingency/interest）因此进不了索引，
# 正文 6,244.15 这类值全部落成「未绑定」。
META_POINTER_RE = re.compile(
    r"(?:^|\.)(schema|schemaversion|version|generatedat|producer|engine|timestamp)$", re.I)

# ---------------------------------------------------------------- 语义域白名单
# 只按「值相等」反查字段会把语义完全不同的字段绑到一起：正文「1 年」曾绑到
# trace.toleranceWan（值恰好=1）、「30 年」绑到 otherFees[21].amount（30 万元）、
# 「0.9」按一位小数四舍五入绑到 ratioOfParent=0.86。读者点开 JSON 就露馅。
# 因此先按 token 的单位确定语义域，再在域内找值：
#   年/个月 → 年份、期限字段；% → 比率字段且精确相等；其余 → 普通数值域。
YEAR_TAILS = {"year", "years", "term", "duration", "ptall", "ptequity",
              "payback", "paybackperiod"}
RATIO_TAILS = {"ratio", "ratioofparent", "ratiooftotal", "pct", "percent", "share",
               "weight", "weights", "firr", "irr", "roi", "roe", "cr", "ci",
               # 两个口径的收益率：firrEquity 的尾巴是 equity，不在 rate/ratio 词根里，
               # 不补进来「-3.05%」就永远绑不上（同一段里 -2.55% 却能绑，评审一眼能看出）。
               "firrequity", "firrbeforetax", "firraftertax", "irrequity",
               "deltapercent", "interestrate"}
MONTH_TAILS = {"month", "months", "monthcount"}


def _pointer_tail(pointer):
    return re.split(r"[.]", pointer or "")[-1].lower()


def _in_year_domain(pointer):
    tail = _pointer_tail(pointer)
    return tail in YEAR_TAILS or tail.endswith(("years", "term", "period")) \
        and tail not in ("period", "phase")


def _in_ratio_domain(pointer):
    tail = _pointer_tail(pointer)
    return tail in RATIO_TAILS or tail.endswith(("rate", "ratio", "share", "weight",
                                                 "percent", "pct"))


def _in_month_domain(pointer):
    return _pointer_tail(pointer) in MONTH_TAILS


def _in_meter_domain(pointer):
    """长度字段：指针尾部以 m 结尾（cellM / floorHeightM / thresholdM / radiusM…）。

    「1 米网格」曾绑到 trace.toleranceWan（值恰好=1），打开 JSON 就是笑话。
    「weightSum」也以 m 结尾，但它不是长度字段——按词尾先排除，否则「1 米」
    会被绑到 AHP 权重合计上。
    """
    tail = _pointer_tail(pointer)
    if tail in METER_FALSE_TAILS:
        return False
    return tail.endswith("m") or tail in ("height", "radius", "width", "depth", "length")


METER_FALSE_TAILS = {"weightsum", "sum", "amount", "item", "minimum", "maximum",
                     "random", "column", "freedom", "from"}


# 计数单位：个/处/栋/户/家/人/项。这些数只能绑到计数字段，
# 不能因为「值相等」就绑到第 N 年、期限或权重上（「7 个小区」曾绑成「第 7 年」，
# 「21 栋」曾绑成「第 21 年」，都是按正则抓数字、不看字段语义造成的）。
COUNT_TAILS = {"count", "counts", "num", "number", "numbers", "size", "quantity",
               "total", "totals", "items", "unitcount", "buildingcount", "buildings",
               "households", "householdcount", "population", "staff", "persons",
               "unitnum", "shopcount", "unitcount"}


def _in_count_domain(pointer):
    tail = _pointer_tail(pointer)
    return tail in COUNT_TAILS or tail.endswith(("count", "num", "quantity"))


def standard_value_ref(token, text):
    """AHP 查表标准值不是本项目测算值，单独登「标准值」来源，不绑任何 JSON 字段。

    RI=0.9 是 Saaty 平均随机一致性指标表 n=4 的查表值；0.1 是 CR<0.1 的判据阈值。
    两者曾分别绑到投资估算 composition 的 ratioOfParent(0.86) 与 ratio(0.05)，
    值靠着四舍五入"碰巧"对上，出处却完全不是一回事。
    """
    text = str(text or "")
    if token == "0.9" and re.search(r"(?<![A-Za-z])RI(?![A-Za-z])", text):
        return "标准值:Saaty 平均随机一致性指标 RI 表（n=4）"
    if token == "0.1" and re.search(r"CR\s*[<＜≤=＝]\s*0?\.1", text):
        return "标准值:AHP 一致性判据 CR<0.1"
    return ""


def _leaves(node, prefix=""):
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _leaves(value, "{}.{}".format(prefix, key) if prefix else str(key))
    elif isinstance(node, list):
        for index, value in enumerate(node):
            yield from _leaves(value, "{}.{}".format(prefix, index) if prefix else str(index))
    else:
        yield prefix, node


def build_value_index(workspace):
    """把上游 JSON 的所有数值叶子摊平成 {数值: [(文件, 指针, 原值)]}。"""
    index = {}
    for rel in VALUE_FILES:
        real = resolve_file(rel, workspace)
        if not real:
            continue
        try:
            with open(real, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            continue
        for pointer, leaf in _leaves(data):
            if META_POINTER_RE.search(pointer):
                continue                     # 版本号/生成时间不是可核对字段
            if leaf is None or isinstance(leaf, bool):
                continue
            if isinstance(leaf, str):
                # 字符串型数值必须整串就是一个数（"-10973"、"0.5922"、"3.5%"）；
                # 版本号 "cardo-invest-estimate/0.3.0" 里抽出来的 0.3 曾被当成字段值，
                # 让「专项债 0.00 万元」绑到 producer 上——打开 JSON 就是笑话。
                matched = STR_NUM_RE.fullmatch(leaf.strip())
                if not matched:
                    continue
                try:
                    number = float(matched.group(1).replace(",", ""))
                except ValueError:
                    continue
                if (matched.group(2) or "") == "%":
                    number /= 100.0
                index.setdefault(round(number, 6), []).append((rel, pointer, number))
                continue
            try:
                number = float(leaf)
            except (TypeError, ValueError):
                continue
            index.setdefault(round(number, 6), []).append((rel, pointer, number))
    return index


def bind_value(token, index, context=""):
    """把一个正文数值绑到语义域内的唯一字段。

    先按单位划域：年/个月只许年份与期限字段，% 只许比率字段且精确相等；
    再按「投资估算 → 财务测算 → 方案比选」优先、同文件取最短指针。
    上下文用于消除同域歧义：「回收期 3 年」取 ptEquity，不取贷款宽限期；
    「第 10 年」取 period.years[9].year，不取 loanTerm。
    """
    match = TOKEN_RE.match(token)
    if not match:
        return None
    raw, unit = match.group(1), match.group(2) or ""
    sign = -1.0 if raw.startswith("-") else 1.0
    body = raw.lstrip("-")
    digits = len(body.split(".")[1]) if "." in body else 0
    value = sign * float(body.replace(",", ""))
    # 不带单位的小整数是序号/条数（「4 项」「3 类」），不是可核对的指标值
    if not unit and digits == 0 and abs(value) < 100:
        return None
    # 表格把单位提到表头（「FIRR（%）」「占比（%）」）时，单元格里只剩 -2.55；
    # 不补单位就永远绑不上比率字段。只对表格补：正文段落里只要出现一处「%」，
    # 整段的 DSCR「1.17」「-0.88」都会被当成百分数去比率域里找，必然落空。
    if not unit and digits and abs(value) < 100 and "|" in str(context or "") \
            and re.search(r"百分比|占比|%", str(context or "")):
        unit = "%"
    if unit == "%":
        domain = _in_ratio_domain
    elif unit == "个月":
        domain = _in_month_domain
    elif unit == "年":
        domain = _in_year_domain
    elif unit == "米":
        domain = _in_meter_domain
    elif unit in ("个", "处", "栋", "户", "家", "人", "项", "台", "套"):
        domain = _in_count_domain
    else:
        domain = None
    # 匹配分两档：精确相等优先；没有再按「正文显示位数」放四舍五入容差。
    # 报告写 95.6%、上游存 95.61%；写 6,244.15、上游存 6244.15；
    # 不放容差这些关键结论值全部落成「未绑定」，xlsx 就不能叫逐数有源。
    tolerance = (0.5 * 10 ** -digits + 1e-9) if digits else 1e-9

    def _collect(match_value, negative_only=False):
        exact, near = [], []
        for number, entries in index.items():
            if negative_only and number >= 0:
                continue
            basis = abs(number) if negative_only else number
            if unit == "%":
                # 比率来源既可能存小数（0.9561），也可能存百分数（95.61）
                distance = min(abs(basis - match_value), abs(basis * 100 - match_value))
            else:
                distance = abs(basis - match_value)
            if distance > tolerance:
                continue
            bucket = exact if distance <= 1e-9 else near
            for entry in entries:
                if domain is None or domain(entry[1]):
                    bucket.append((distance, entry))
        return exact, near

    exact, near = _collect(value)
    # 正文常用方向词写幅度（「压减 57.7%」「下降 350.1%」），上游存的是带负号的
    # 临界点幅度（-57.69 / -350.06）。符号不一致时按绝对值再匹配一档，只认负值来源，
    # 否则这三个临界点永远落成「人工核」，xlsx 就不能叫逐数有源。
    if not exact and not near and unit == "%" and value > 0 \
            and DIRECTION_RE.search(str(context or "")):
        exact, near = _collect(value, negative_only=True)
    if exact:
        hits = [entry for _, entry in exact]
    elif near:
        # 容差内可能有多个邻近值（6244.15 / 6244.16），只取最近的一档；
        # 同一距离仍有多个字段时交给 _prefer_by_context 按上下文收敛。
        best = min(distance for distance, _ in near)
        hits = [entry for distance, entry in near if distance <= best + 1e-9]
    else:
        return None
    hits = _prefer_by_context(hits, value, unit, context)
    order = {rel: pos for pos, rel in enumerate(VALUE_FILES)}
    hits.sort(key=lambda item: (order.get(item[0], 99), len(item[1]), item[1]))
    rel, pointer, number = hits[0]
    return {"number": token, "ref": "{}#{}".format(rel, pointer),
            "path": rel, "value": number}


def _prefer_by_context(hits, value, unit, context):
    """同值多字段时，用正文上下文挑语义最贴的那个。

    白名单只保证「落在同一语义域」，域内仍可能同值：贷款宽限期 3 年、回收期 3 年、
    第 3 年都等于 3。这里按句中的限定词收窄，收窄后为空则退回全量。
    """
    text = str(context or "")
    # 投资口径：总投资与建设投资同段出现时，按词义钉到各自的字段，
    # 不让「总投资 23,823.72」绑到 fundingTotal、建设投资绑到 trace 步骤上。
    if "总投资" in text:
        picked = [item for item in hits if _pointer_tail(item[1]) == "totalinvestment"]
        if picked:
            return picked
    if "建设投资" in text:
        picked = [item for item in hits if _pointer_tail(item[1]) == "constructioninvestment"]
        if picked:
            return picked
    if unit == "年":
        def year_sel(pattern):
            return [item for item in hits if re.search(pattern, _pointer_tail(item[1]))]
        if "回收期" in text:
            picked = year_sel(r"^(ptall|ptequity|payback|paybackperiod|paybackyears)$")
            if picked:
                return picked
        if "计算期" in text or "评价期" in text:
            picked = year_sel(r"^(totalyears|totals)$")
            if picked:
                return picked
        if "宽限" in text:
            picked = [item for item in hits if "grace" in _pointer_tail(item[1])]
            if picked:
                return picked
        if "建设期" in text:
            picked = [item for item in hits if "construction" in item[1].lower()]
            if picked:
                return picked
        if "运营期" in text:
            picked = [item for item in hits if "operation" in item[1].lower()]
            if picked:
                return picked
        if value == int(value) and re.search(r"第\s*{}(?:\s*年)".format(int(value)), text):
            picked = year_sel(r"^year$")
            if picked:
                return picked
    if unit == "%":
        if any(word in text for word in ("资金来源", "资本金", "补助", "筹措")):
            picked = [item for item in hits if "fund" in item[1].lower()]
            if picked:
                return picked
        if any(word in text for word in ("费用类别", "投资构成", "工程费用")):
            picked = [item for item in hits if "composition" in item[1].lower()]
            if picked:
                return picked
    return hits


def split_blocks(lines):
    """切块：标题另计；空行分段；表格整块；图/表题注单独成块（带 1 起始行号）。"""
    blocks, buf, section = [], [], ""
    buf_start = [1]

    def flush(kind="p"):
        if buf:
            blocks.append({"kind": kind, "text": "\n".join(buf), "section": section,
                           "line": buf_start[0]})
            buf.clear()

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            flush()
            i += 1
            continue
        m = HEAD_RE.match(line)
        if m:
            flush()
            if len(m.group(1)) <= 2:
                section = m.group(2).strip()
            blocks.append({"kind": "h", "text": m.group(2).strip(), "section": section, "line": i + 1})
            i += 1
            continue
        src = SRC_RE.match(line.strip())
        if src:
            flush()
            blocks.append({"kind": "s", "text": src.group(1).strip(), "section": section, "line": i + 1})
            i += 1
            continue
        if TABLE_RE.match(line.strip()):
            flush()
            start = i + 1
            rows = []
            while i < len(lines) and TABLE_RE.match(lines[i].strip()):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append({"kind": "t", "text": "\n".join("|".join(r) for r in rows),
                           "rows": rows, "section": section, "line": start})
            continue
        if FIG_RE.match(line.strip()):
            flush()
            blocks.append({"kind": "f", "text": line.strip(), "section": section, "line": i + 1})
            i += 1
            continue
        if not buf:
            buf_start[0] = i + 1
        buf.append(line)
        i += 1
    flush()
    return blocks


def is_structural(token):
    """序号/编号类的小整数（S1、附件 5、第 3 类）不是数据事实。

    这类 token 本来就不该有字段出处，但也不能算进「未绑定」——否则 29% 的缺口里
    混着一堆「1」「2」「5」，真实缺口被稀释，评审没法判断到底哪些数没源。
    """
    return bool(re.fullmatch(r"-?\d{1,2}", str(token or "").strip()))


def build(chapter_dir, workspace):
    rows, problems, manual = [], [], []
    stats = {"chapters": 0, "blocks": 0, "withSource": 0, "noSource": 0,
             "valueBound": 0, "valueUnbound": 0, "valueStructural": 0, "standardValue": 0,
             "blocksFullBound": 0, "blocksPartial": 0, "blocksManual": 0}
    # 数值索引只建一次：正文每个数值按「值」反查唯一字段，不再按节共用一个来源池。
    index = build_value_index(workspace)
    for path in sorted(Path(chapter_dir).glob("CH-*.md")):
        stats["chapters"] += 1
        chapter = path.stem.replace("CH-", "")
        blocks = split_blocks(read_text(path).splitlines())
        # Invest 分章稿的写法是「一节末尾一条来源行」，语义上覆盖本节全部段落、表格与图，
        # 因此按节归属：本节所有来源行都记到本节每个内容块上（含来源行之后的表）。
        seq = 0
        section_refs = {}
        for block in blocks:
            refs = section_refs.setdefault(block["section"], [])
            if block["kind"] == "s":
                refs.extend((ref, block.get("line")) for ref in split_refs(block["text"]))
        for block in blocks:
            if block["kind"] in ("s", "h"):
                continue
            pairs = section_refs.get(block["section"], [])
            lines_by_ref = {}
            for ref, line in pairs:
                lines_by_ref.setdefault(ref, line)
            block["sources"] = list(lines_by_ref)
            block["sourceLines"] = lines_by_ref
        for block in blocks:
            if block["kind"] in ("h", "s"):
                continue
            seq += 1
            refs = block.get("sources") or []
            nums = numbers_in(block["text"])
            detail, bad, value_refs = [], 0, []
            # 节级来源先全部校验一遍：即使该块最终用逐值绑定、不再并列这些来源，
            # 分章稿里写错的指针也必须暴露（坏样例 B 就是把 7.1 指针改成不存在的键，
            # 若只在「没有数值」时才校验，改完仍会 exit=0）。
            checked = []
            for ref in refs:
                status, note, real = check_ref(ref, workspace)
                if status not in ("可解析", "说明"):
                    bad += 1
                    line = block.get("sourceLines", {}).get(ref) or block.get("line") or ""
                    msg = "CH-{}.md:{} {}：{}（{}）".format(
                        chapter, line, block["section"], ref, status)
                    if msg not in problems:
                        problems.append(msg)
                checked.append({"ref": ref, "status": status, "note": note,
                                "path": os.path.relpath(real, workspace).replace("\\", "/")
                                if real else ""})
            if nums:
                # 一个值一个指针：逐值反查唯一字段，登记进 valueRefs。
                # 同值多字段时按 VALUE_FILES 顺序 + 最短指针裁决，不再把整节来源并成比对池。
                for token in nums:
                    # AHP 查表标准值（RI 0.9 / CR 0.1）不是本项目测算值，单独登来源。
                    standard = standard_value_ref(token, block["text"])
                    if standard:
                        value_refs.append({"number": token, "ref": standard, "path": "",
                                           "value": None, "status": "标准值"})
                        stats["standardValue"] += 1
                        continue
                    bound = bind_value(token, index, block["text"])
                    if bound:
                        value_refs.append({"number": token, "ref": bound["ref"],
                                           "path": bound["path"], "value": bound["value"],
                                           "status": "可解析"})
                        stats["valueBound"] += 1
                    elif is_structural(token):
                        value_refs.append({"number": token, "ref": "", "path": "",
                                           "value": None, "status": "结构值"})
                        stats["valueStructural"] += 1
                    else:
                        value_refs.append({"number": token, "ref": "", "path": "",
                                           "value": None, "status": "未定位"})
                        stats["valueUnbound"] += 1
                        manual.append({"chapter": "CH-{}".format(chapter),
                                       "section": block["section"], "seq": seq,
                                       "line": block.get("line"),
                                       "kind": {"p": "正文", "t": "表格",
                                                "f": "图表题注"}.get(block["kind"], block["kind"]),
                                       "excerpt": re.sub(r"\s+", " ", block["text"])[:60],
                                       "number": token})
                bound_seen = set()
                for item in value_refs:
                    if not item["ref"] or item["ref"].startswith("标准值:"):
                        continue
                    if item["ref"] in bound_seen:
                        continue
                    bound_seen.add(item["ref"])
                    detail.append({"ref": item["ref"], "status": "可解析",
                                   "note": item["ref"], "path": item["path"]})
            has_binding = any(item["ref"] and not item["ref"].startswith("标准值:")
                              for item in value_refs)
            unbound_n = sum(1 for item in value_refs if item["status"] == "未定位")
            bound_n = sum(1 for item in value_refs if item["status"] in ("可解析", "标准值"))
            structural_n = sum(1 for item in value_refs if item["status"] == "结构值")
            unbound = unbound_n > 0
            for item in checked:
                # 全部数值都绑上了精确字段时，节级来源只做校验、不并列进 sources；
                # 但只要还有数值没绑上，本节来源就必须保留——正文引述上游报告的句子，
                # 这些值只在报告原文里、不在任何 JSON 字段里，节级来源是唯一可核出处。
                if has_binding and not unbound:
                    break
                if any(existing["ref"] == item["ref"] for existing in detail):
                    continue
                # path 是给下游机器读的落地路径；note 里带的人类可读说明不能当路径。
                detail.append(item)
            stats["blocks"] += 1
            if detail:
                stats["withSource"] += 1
            else:
                stats["noSource"] += 1
            if nums and (not detail or bad):
                stats["numericNoSource"] = stats.get("numericNoSource", 0) + 1
            # 状态只认三档，不把「部分绑定」混进「有源」：
            # 早先只要块里有任意一个来源就标「有源」，未绑定值全藏在「有源」里。
            if bad:
                status = "引用不实"
            elif unbound_n and bound_n:
                status = "部分有源（{} 个值待人工核）".format(unbound_n)
                stats["blocksPartial"] += 1
            elif unbound_n:
                status = "人工核（{} 个值未绑定）".format(unbound_n)
                stats["blocksManual"] += 1
            elif detail:
                status = "有源"
                stats["blocksFullBound"] += 1
            else:
                status = "无源"
            rows.append({
                "chapter": "CH-{}".format(chapter),
                "section": block["section"],
                "seq": seq,
                "line": block.get("line"),
                "kind": {"p": "正文", "t": "表格", "f": "图表题注"}.get(block["kind"], block["kind"]),
                "excerpt": re.sub(r"\s+", " ", block["text"])[:60],
                "numbers": nums,
                "valueRefs": value_refs,
                "sources": detail,
                "boundCount": bound_n,
                "unboundCount": unbound_n,
                "structuralCount": structural_n,
                "status": status,
            })
    return {"schema": "cardo/evidence-register/v1",
            "generatedAt": __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M"),
            "stats": stats, "rows": rows, "manual": manual, "problems": problems}


def _selftest(workspace):
    """反例验证：把历史错配样本喂给 bind_value，必须被拦下。

    样本来自 R2 挑刺：586 行依据对照表里 82 行错配，典型是
    「7 个」→ period.years[6].year（第 7 年）、「21 栋」→ 第 21 年、「1 米」→ weightSum。
    判据：要么不绑定（None），要么绑到语义域内字段；绑到年份/期限/权重即失败。
    """
    index = build_value_index(workspace)
    banned = re.compile(r"(?:^|\.)(year|years|term|duration|weightsum|weight|weights)$", re.I)
    cases = [("7 个", "共 7 个小区"), ("21 栋", "改造 21 栋住宅"),
             ("1 米", "按 1 米网格采样"), ("1,203 户", "惠及 1,203 户")]
    failures = 0
    for token, context in cases:
        bound = bind_value(token, index, context)
        ok = True
        if bound:
            tail = _pointer_tail(bound["ref"].split("#", 1)[-1])
            ok = not banned.search(tail)
        print("  {} {}（{}）→ {}".format(
            "PASS" if ok else "FAIL", token, context,
            (bound["ref"] if bound else "未绑定（拦下，交人工核）")))
        failures += 0 if ok else 1
    print("[依据对照表·selftest] {}：{} 例，失败 {}".format(
        "PASS" if not failures else "FAIL", len(cases), failures))
    return 0 if not failures else 1


def main(argv=None):
    ap = argparse.ArgumentParser(description="Invest 依据对照表（分章来源行 → 逐块登记）")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--chapters", default="")
    ap.add_argument("--out", default="")
    ap.add_argument("--no-timestamp", action="store_true")
    ap.add_argument("--selftest", action="store_true",
                    help="用历史错配样本验证取值绑定：错配必须被拦（返回 None 或域内正确字段）")
    args = ap.parse_args(argv)
    workspace = os.path.abspath(args.project_dir)
    if args.selftest:
        return _selftest(workspace)
    chapter_dir = args.chapters or os.path.join(workspace, "过程文件/05-章节成文/分章")
    if not os.path.isdir(chapter_dir):
        print("[依据对照表] 找不到分章目录：{}\n请先跑 05 分章成文。".format(chapter_dir))
        return 2
    payload = build(chapter_dir, workspace)
    out = args.out or os.path.join(chapter_dir, "依据对照表.json")
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    stats = payload["stats"]
    fact = stats.get("valueBound", 0) + stats.get("valueUnbound", 0)
    print("[依据对照表] {} 章 / {} 块：有源 {}、无源 {}、引用不实 {} → {}".format(
        stats["chapters"], stats["blocks"], stats["withSource"], stats["noSource"],
        len(payload["problems"]), os.path.relpath(out, workspace)))
    print("  数值绑定 {}/{}（{:.1f}%）；待人工核 {}（{:.1f}%）；结构值 {}；标准值 {}".format(
        stats.get("valueBound", 0), fact,
        100.0 * stats.get("valueBound", 0) / fact if fact else 0.0,
        stats.get("valueUnbound", 0),
        100.0 * stats.get("valueUnbound", 0) / fact if fact else 0.0,
        stats.get("valueStructural", 0), stats.get("standardValue", 0)))
    for problem in payload["problems"][:8]:
        print("   -", problem)
    return 0


if __name__ == "__main__":
    sys.exit(main())
