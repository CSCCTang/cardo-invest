"""可研 PDF → 财务参数抽取（T6 预填器核心库）。

只从可研原文抽数，不补数、不用经验值顶替。
每条抽到的值都带 {file, page, quote} 三要素出处。

设计要点：
  * 附表（项目总投资估算表）是多栏无框线表格，pymupdf 的 find_tables 不可靠，
    因此主路径走「文本行 + 数值列解析」，find_tables 只作辅助。
  * 金额单位统一折算为万元：原文若为「元」需按 10000 折算并记录原单位。
  * 抽不到的字段一律返回 None（不猜、不补），由上层登记缺项。
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

# ---------------------------------------------------------------- 数值工具

# 金额数字：支持千分位与 4 位以上整数。
# 注意不能写成 `\d{1,3}(?:,\d{3})*` 开头 —— 那样 2821.07 会被截成 282.07（实测踩过）。
_NUM_RE = re.compile(r"-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+(?:\.\d+)?")


def to_decimal(text):
    """把文本里的第一个数字转成 Decimal；带千分位逗号；无数字返回 None。"""
    if text is None:
        return None
    if isinstance(text, (int, float, Decimal)):
        try:
            return Decimal(str(text))
        except InvalidOperation:
            return None
    match = _NUM_RE.search(str(text).replace(" ", ""))
    if not match:
        return None
    try:
        return Decimal(match.group(0).replace(",", ""))
    except InvalidOperation:
        return None


def clean_line(line):
    """压掉 PDF 抽取出来的换行/全角空格干扰。"""
    return re.sub(r"\s+", " ", (line or "").replace("　", " ")).strip()


def fmt(value, digits=2):
    """格式化为不带多余尾零的字符串；None → 空串。"""
    if value is None:
        return ""
    quant = Decimal(1).scaleb(-digits)
    try:
        out = Decimal(value).quantize(quant)
    except (InvalidOperation, TypeError):
        return ""
    text = format(out, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


# ---------------------------------------------------------------- 附表解析

# 可研投资估算附表里的层级标题（用于切分大类）
_SECTION_MARKERS = [
    ("engineering", ("一", "工程费用")),
    ("otherFees", ("二", "工程建设其他费用")),
    ("contingency", ("三", "预备费")),
    ("total", ("四", "项目总投资")),
    ("total", ("五", "项目总投资")),
]

# 噪声行：这些是表头、说明、比例列的表注，不是费用项
_NOISE_KEYS = (
    "序号", "工程或费用名称", "建筑工程费", "设备购置", "合计", "技术经济指标",
    "占总投资", "备注", "单位", "数量", "单价", "工程或费用", "项目总投资估算表",
    "附表", "总投资估算表", "估表",
)

# 合计/小计/总计行：不进明细（明细由子项加总，合计行会重复计数）
_AGGREGATE_NAMES = ("合计", "小计", "总计", "共计")


def _is_noise(name):
    if not name:
        return True
    if name in _NOISE_KEYS:
        return True
    return any(key in name for key in _NOISE_KEYS)


def _is_aggregate(name):
    return name in _AGGREGATE_NAMES or name.startswith(("合计", "小计", "总计"))


def _norm_name(name):
    """去掉名称里的编号前缀与空白，用于聚合判断。"""
    name = clean_line(name)
    name = re.sub(r"^[（(]?[一二三四五六七八九十\d]+[）)、.．]?", "", name)
    return name.strip("．.、 ")


def parse_estimate_table(lines, page_no, file_label):
    """解析可研「项目总投资估算表」文本块。

    lines: [(page, text)] 或 [text]；返回 (categories, evidence)
    categories = {"engineering": [...], "otherFees": [...], ...}
    每项 = {"name", "amount", "source": {file, page, quote}}
    """
    entries = []
    current = None
    for item in lines:
        if isinstance(item, tuple):
            page, text = item
        else:
            page, text = page_no, item
        text = clean_line(text)
        if not text:
            continue
        # 大类切换：一、工程费用 / 二、工程建设其他费用 ...
        for key, marker in _SECTION_MARKERS:
            if text.startswith(marker[0]) and marker[1] in text:
                current = key
                break
        entries.append((page, text, current))

    # 逐行找「名称 + 金额」：金额取该行中第一个能独立成词的数字
    out = {"engineering": [], "otherFees": [], "contingency": [], "total": []}
    for page, text, section in entries:
        if section is None:
            continue
        if _is_noise(text):
            continue
        parts = [clean_line(p) for p in re.split(r"\s{2,}|\t", text) if clean_line(p)]
        if len(parts) < 2:
            continue
        # 名称 = 首个非纯数字片段；金额 = 其后的数字
        name = None
        amount = None
        for part in parts:
            if amount is None:
                if re.fullmatch(r"[\d,.\s%]+", part) and re.search(r"\d", part):
                    candidate = to_decimal(part)
                    if candidate is not None:
                        amount = candidate
                else:
                    if name is None and not _is_noise(part):
                        name = part
            else:
                break
        if not name or amount is None:
            continue
        if _is_aggregate(name):
            continue
        # 跳过纯比例行（如 "0.34" 跟在备注里）——名称里含数字百分比特征
        if re.search(r"\d+(\.\d+)?\s*%$", text) and not re.search(r"万元|元|万", text):
            continue
        out[section].append({
            "name": name,
            "amount": Decimal(amount),
            "source": {"file": file_label, "page": page, "quote": text[:120]},
        })
    return out


def merge_duplicate_names(items):
    """同名项目合并（同名不同页的重复列示），金额相加，出处合并。"""
    merged = {}
    order = []
    for item in items:
        key = _norm_name(item["name"])
        if not key:
            continue
        if key not in merged:
            merged[key] = {
                "name": item["name"],
                "amount": Decimal(item["amount"]),
                "source": item["source"],
                "sources": [item["source"]],
            }
            order.append(key)
        else:
            merged[key]["amount"] += Decimal(item["amount"])
            merged[key]["sources"].append(item["source"])
    return [merged[key] for key in order]


# ---------------------------------------------------------------- 章节定位

_INVEST_KEYS = ("投资估算", "投资概算", "工程费用", "工程费", "总概算", "总投资估算")
_FUND_KEYS = ("资金筹措", "资金来源", "资金运用", "融资方案")
_REVENUE_KEYS = ("营业收入", "运营收入", "经营收入", "年收入", "收入预测")
_COST_KEYS = ("运营成本", "经营成本", "总成本", "运营费用")


def find_pages(doc, keys, limit=None):
    """返回命中任一关键词的页码列表（1-based）。"""
    hits = []
    for index in range(doc.page_count):
        text = doc[index].get_text()
        if any(key in text for key in keys):
            hits.append(index + 1)
            if limit and len(hits) >= limit:
                break
    return hits


def find_number_near(doc, pages, patterns, window=400):
    """在指定页范围内，按正则找「关键词 + 数字」的句子，返回 (值, 出处)。

    patterns: [(正则, 取第几个捕获组)]
    """
    for page in pages:
        text = doc[page - 1].get_text()
        flat = clean_line(text)
        for pattern, group in patterns:
            match = re.search(pattern, flat)
            if not match:
                continue
            raw = match.group(group)
            value = to_decimal(raw)
            if value is None:
                continue
            start = max(0, match.start() - 60)
            return Decimal(value), {
                "file": None,  # 由调用方补
                "page": page,
                "quote": flat[start:match.end() + 60][:160],
            }
    return None, None


def extract_funding(doc, pages, file_label):
    """抽资金来源：资本金 / 借款 / 补助（万元）。只认明确的表述。"""
    result = {}
    for page in pages:
        text = doc[page - 1].get_text()
        flat = clean_line(text)
        patterns = [
            (r"(?:项目)?资本金[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元", "equity"),
            (r"(?:项目)?资本金[^。]{0,12}?([\d,]+(?:\.\d+)?)\s*万元", "equity"),
            (r"(?:银行)?借款(?:资金)?[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元", "loan"),
            (r"(?:政策性)?补助(?:资金)?[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元", "subsidy"),
            (r"(?:财政资金|政府投资)[^。]{0,12}?([\d,]+(?:\.\d+)?)\s*万元", "subsidy"),
        ]
        for pattern, key in patterns:
            if key in result:
                continue
            match = re.search(pattern, flat)
            if match:
                value = to_decimal(match.group(1))
                if value is not None:
                    result[key] = {"amount": value, "source": {
                        "file": file_label, "page": page,
                        "quote": flat[max(0, match.start() - 40):match.end() + 40][:160]}}
    return result


def extract_headline_investment(doc, pages, file_label):
    """抽正文结论句里的总投资/工程费/其他费/预备费（万元）。"""
    out = {}
    patterns = [
        ("totalInvestment", r"(?:估算|概算)?总投资[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元"),
        ("engineering", r"工程费(?:用)?[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元"),
        ("otherFees", r"工程建设其他费[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元"),
        ("contingency", r"预备费[为约]?\s*([\d,]+(?:\.\d+)?)\s*万元"),
    ]
    for page in pages:
        text = doc[page - 1].get_text()
        flat = clean_line(text)
        for key, pattern in patterns:
            if key in out:
                continue
            match = re.search(pattern, flat)
            if match:
                value = to_decimal(match.group(1))
                if value is not None:
                    out[key] = {"amount": value, "source": {
                        "file": file_label, "page": page,
                        "quote": flat[max(0, match.start() - 50):match.end() + 50][:160]}}
    return out


def extract_project_meta(doc, file_label):
    """抽项目级字段：名称/方向/面积/建设期/运营期。"""
    meta = {}
    # 项目名称：取封面或页眉出现频次最高的项目全称
    head = ""
    for index in range(min(doc.page_count, 6)):
        text = doc[index].get_text()
        if text.strip():
            head = text
            break
    # 项目名里常含空格与年份（「某项目·1958 园区…」「某区2025 年老旧小区…」）。
    # 旧正则的字符类不含空格 → 匹配在空格处断开，
    # 抽出来的是「年老旧小区配套基础设施改造项目」这种半截名（实测）。
    # 现在允许空格/连字符，先按原样收集候选，再按「去掉空白后的长度」择优。
    name_pattern = re.compile(
        r"([\u4e00-\u9fa5A-Za-z0-9·][\u4e00-\u9fa5A-Za-z0-9·\s\-]{3,58}?项目)")
    for index in range(doc.page_count):
        text = doc[index].get_text()
        for match in name_pattern.finditer(text):
            raw = match.group(1)
            candidate = re.sub(r"\s+", "", raw).strip(" -·")
            if 6 <= len(candidate) <= 50:
                meta.setdefault("_name_candidates", {})
                meta["_name_candidates"][candidate] = \
                    meta["_name_candidates"].get(candidate, 0) + 1
    # **文件名优先**：提资方的文件命名就是项目正式名称，最权威。
    # 正文里的候选会混进多个项目的名字与长句噪声
    # （实测某可研打包多项目：「××年院改项目」出现 102 次，
    #   而本项目名只出现 3 次；另有「水泥空地…审核校核设计项目」这类跨栏噪声）。
    stem = re.sub(r"\.pdf$", "", file_label, flags=re.IGNORECASE)
    stem = re.sub(r"^\s*\d+(?:[\.\-]\d+)*[\.\-]?\s*", "", stem)
    stem = re.sub(r"^附件\d*\s*[\.\-]?\s*", "", stem)
    stem = re.sub(r"[（(][^）)]*(可行性研究报告|项目建议书|代项目建议书|附专家|盖章|全套)[^）)]*[）)]", "", stem)
    stem = re.sub(r"[-—]?可研报告合并$", "", stem)
    stem = re.sub(r"(可行性研究报告|可研报告合并|可研报告|项目建议书|代项目建议书)$", "", stem)
    stem = stem.strip(" -·")
    if 6 <= len(stem) <= 50:
        meta["projectName"] = {"value": stem, "source": {
            "file": file_label, "page": 0, "quote": "文件名（提资方命名）"}}

    # 正文候选兜底：只在文件名不可用时用，且要严格过滤噪声
    # （含动作词、长度适中、不以「本项目」结尾、不跨栏拼接）。
    if not meta.get("projectName") and meta.get("_name_candidates"):
        action_words = ("改造", "更新", "提升", "建设", "整治", "修复", "重建", "配套")
        ranked = []
        for name, count in meta["_name_candidates"].items():
            if len(name) > 24:
                continue
            if not any(word in name for word in action_words):
                continue
            if name.endswith("本项目") or "指标项目" in name:
                continue
            ranked.append((count, len(name), name))
        if ranked:
            best = max(ranked)[2]
            meta["projectName"] = {"value": best, "source": {
                "file": file_label, "page": 1, "quote": best}}

    # 改造面积：m²/ 平方米 前的数字
    for index in range(doc.page_count):
        text = doc[index].get_text()
        flat = clean_line(text)
        match = re.search(r"改造(?:面积|规模)[^\d]{0,20}?([\d,]+(?:\.\d+)?)\s*(?:m2|m²|平方米|㎡)", flat)
        if match:
            value = to_decimal(match.group(1))
            if value:
                meta["rebuildArea"] = {"value": value, "source": {
                    "file": file_label, "page": index + 1,
                    "quote": flat[max(0, match.start() - 40):match.end() + 40][:160]}}
                break
    # 建设期/运营期：N 年
    for index in range(doc.page_count):
        text = doc[index].get_text()
        flat = clean_line(text)
        for key, pattern in (("constructionYears", r"建设期[^\d]{0,12}?(\d+)\s*年"),
                             ("operationYears", r"运营期[^\d]{0,12}?(\d+)\s*年")):
            if key in meta:
                continue
            match = re.search(pattern, flat)
            if match:
                meta[key] = {"value": int(match.group(1)), "source": {
                    "file": file_label, "page": index + 1,
                    "quote": flat[max(0, match.start() - 30):match.end() + 30][:160]}}
    # 建设期：也可能写作「建设周期 N 个月」→ 折算为年（记录原口径）
    if "constructionYears" not in meta:
        for index in range(doc.page_count):
            flat = clean_line(doc[index].get_text())
            match = re.search(r"建设周期[^\d]{0,12}?(\d+)\s*个月", flat)
            if match:
                months = int(match.group(1))
                meta["constructionYears"] = {"value": max(1, round(months / 12)), "derived": True,
                                             "source": {
                    "file": file_label, "page": index + 1,
                    "quote": flat[max(0, match.start() - 30):match.end() + 30][:160]}}
                break
    return meta


def extract_economic_indicators(doc, file_label):
    """抽可研正文里的「经济指标」小节（结构：名称 / 单位 / 数值）。

    实测某企业可研 p11：
        二 经济指标
          1 项目总投资   万元  5300.00
          1.1 工程费用   万元  4301.76   81.17%
          1.2 工程建设其他费用 万元 605.70  11.43%
          1.3 预备费     万元  392.54    7.41%
          2 资金筹措     万元  5300.00  财政资金
    这张表比投资估算附表更好抽（列少、无合并单元格），且带**资金来源**文字
    （「财政资金」/「企业自筹」/「银行贷款」…）——附表里没有资金筹措明细。

    返回 {"items": {名称: {"value", "source"}}, "funding": {...}}
    """
    items = {}
    funding = {}
    pages = find_pages(doc, ("经济指标", "投资规模", "资金来源"))
    for page in pages:
        text = doc[page - 1].get_text()
        lines = [clean_line(line) for line in text.split("\n")]
        for index, line in enumerate(lines):
            if not line:
                continue
            # 只认**已知指标名**（白名单）——否则整页正文都会被当成指标
            # （实测某企业 p10/p11 把「m² 61450」「灭火器箱…」都抽成了指标）。
            if line not in _ECONOMIC_ITEMS:
                continue
            value = None
            unit = None
            for follow in lines[index + 1: index + 4]:
                if follow in ("万元", "元", "亿元"):
                    unit = follow
                    continue
                if re.fullmatch(r"[\d,.]+%?", follow):
                    value = to_decimal(follow)
                    break
                if follow:
                    break
            if value is None:
                continue
            if line in items:
                continue
            items[line] = {"value": value, "source": {
                "file": file_label, "page": page,
                "quote": "{} {} {}".format(line, unit or "", value)}}
            # 资金来源文字（「财政资金」「企业自筹」…）
            if line in ("资金筹措", "资金来源"):
                for follow in lines[index + 1: index + 12]:
                    if follow and not re.fullmatch(r"[\d,.]+%?", follow) and follow not in ("万元", "元"):
                        funding[follow] = {"source": {
                            "file": file_label, "page": page,
                            "quote": "{} {}".format(line, follow)}}
                        break
    return {"items": items, "funding": funding}


# 可研正文「经济指标」小节里的已知指标名（白名单）
_ECONOMIC_ITEMS = (
    "项目总投资", "工程费用", "工程建设其他费用", "预备费",
    "建设期利息", "静态投资", "资金筹措", "资金来源",
    "建设投资", "铺底流动资金",
)


def extract_contingency_rate(doc, file_label):
    """抽预备费率：可研正文「按…计 X%」的说明句。

    实测某企业可研 p78：「(1) 基本预备费：按工程费用与工程建设其他费用之和
    减去建设用地费的8%估算。」；附表备注写「按（一+二-土地费）*8%计」。
    """
    pages = find_pages(doc, ("预备费", "基本预备费"))
    # 必须在「基本预备费 / 预备费」的**同一句**里找费率。
    # 否则会抓到别的费率（实测某企业 p78 同一页有「场地准备费 0.8%」，
    # 全文搜索会先命中它 → 费率变成 0.008 而不是 0.08）。
    sentence = re.compile(r"[^。；\n]{0,60}?(?:基本预备费|预备费)[^。；\n]{0,60}?(\d+(?:\.\d+)?)\s*%")
    for page in pages:
        text = doc[page - 1].get_text()
        flat = clean_line(text)
        for match in sentence.finditer(flat):
            value = to_decimal(match.group(1))
            if value is None or value <= 0 or value > 100:
                continue
            start = max(0, match.start() - 20)
            return {"rate": value / Decimal(100), "source": {
                "file": file_label, "page": page,
                "quote": flat[start:match.end() + 30][:160]}}
    return None


def detect_project_direction(doc):
    """项目方向：老旧小区/老旧厂区/老旧街区，按出现次数判定。"""
    counts = {"老旧小区": 0, "老旧厂区": 0, "老旧街区": 0}
    for index in range(doc.page_count):
        text = doc[index].get_text()
        for key in counts:
            counts[key] += text.count(key)
    best = max(counts.items(), key=lambda kv: kv[1])
    return best[0] if best[1] > 0 else None


def has_revenue_section(doc):
    """可研里是否存在经营收入/运营成本章节 —— 决定是否属经营性项目。"""
    revenue_pages = find_pages(doc, _REVENUE_KEYS)
    cost_pages = find_pages(doc, _COST_KEYS)
    return bool(revenue_pages), bool(cost_pages)
