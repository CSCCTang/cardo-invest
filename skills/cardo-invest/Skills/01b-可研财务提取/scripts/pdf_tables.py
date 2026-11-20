"""可研 PDF 附表解析 —— 以 pymupdf find_tables() 为主，坐标法为辅。

选型经过（实测三种做法）：
  1. get_text() 按行拆 → 列名与数值错位，0 条。不可用。
  2. 自己按坐标聚类分列 → 能取到数值，但跨行名称（「改造绿化」单独一行）
     与续页表头缺失会持续出错，维护成本高。
  3. pymupdf find_tables() → 自动处理无框线表格的跨行单元格、列对齐，
     名称与序号都正确（实测 112/113/114 页分别 10/21/17 行）。
     唯一要处理的是数字被拆成「385336\\n.」（去小数点还原 + 单位是万元）。

因此主路径用 find_tables；表格识别失败时才退回坐标法。

数字还原规则：原文是「万元」，附表里 4301.76 万写作 "430176." —— 即
去掉换行与小数点还原为 4301.76；「0.00」写作 "0"；整数项写作 "10589"。
"""

from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, str(Path(__file__).resolve().parent))

from extract_feasibility import clean_line, to_decimal  # noqa: E402

# 大类 → 语义分类（附表用中文序号标大类）
_CATEGORY = {
    "一": "engineering", "二": "otherFees", "三": "contingency",
    "四": "total", "五": "total", "六": "total",
}

# 附表里出现这些词 → 归到哪个大类。
# 顺序敏感：长词必须在前（否则「工程建设其他费用」会被「工程费用」抢先命中）。
_CONTENT_HINTS = (
    ("工程建设其他费用", "otherFees"),
    ("建设期利息", "constructionInterest"),
    ("建设期融资费", "constructionInterest"),
    ("预备费", "contingency"),
    ("工程费用", "engineering"),
)

# 大类行（序号为中文数字）不进明细，但要单独收集作为权威合计用于校验
_CATEGORY_ROWS = ("一", "二", "三", "四", "五", "六")

# 噪声行：不进明细
_AGGREGATE = ("合计", "小计", "总计", "工程费用", "工程建设其他费用",
              "预备费", "静态投资", "项目总投资", "建设期融资费")


def _clean_cell(value):
    """单元格 → 干净文本（去掉换行与全角空格）。"""
    if value is None:
        return ""
    text = str(value).replace("\n", "").replace("　", " ")
    return clean_line(text)


def _restore_number(text):
    """把表格里被拆开的数字还原。

    附表单位是**万元**，但表内金额按「元」写成整数并把小数点甩到行尾：
        4301.76 万元 → '430176\\n.'   34.41 万元 → '3441.'
        392.54 万元 → '39254.'        暂估 30 万元 → '30'（无小数点）
    因此规则是：带末尾小数点的 → 去点后 ÷100；不带小数点的 → 原值。
    """
    text = _clean_cell(text).replace(",", "")
    if not text:
        return None
    if not re.fullmatch(r"-?\d+\.?", text):
        return to_decimal(text)
    if text.endswith("."):
        digits = text[:-1]
        if not digits:
            return None
        try:
            return (Decimal(digits) / Decimal(100)).quantize(Decimal("0.01"))
        except Exception:
            return None
    return to_decimal(text)


def _classify(seq, row_text, current):
    """判断该行属于哪个大类；返回 (category, new_current, by_keyword)。

    优先级：
      1. 名称/行文本命中大类关键词（如「基本预备费」「工程建设其他费用」）
         → 关键词比列位置可靠（预备费行也会落在「其他费」列上）。
      2. 序号列是中文大类序号（一/二/三…）→ 大类行。
      3. 分类列（建筑/设备/其他费）→ 由调用方按列补。
    """
    for word, key in _CONTENT_HINTS:
        if word in row_text:
            return key, key, True
    if seq:
        key = seq.strip()
        if key in _CATEGORY:
            return _CATEGORY[key], _CATEGORY[key], True
    return current, current, False


def _normalize_seq(text):
    """归一化序号列。

    PDF 表格抽取会把序号拆行成「1\\n2\\n1」或「26\\n1\\n. .」，
    直接去换行会得到「121..」这种垃圾（层级全丢，父级剔除随之失效）。
    做法：把**所有非数字非点字符**（含空白与括号残迹）统一换成点，
    再折叠连续点、去掉首尾点：
        '1\\n2\\n1'      → '1.2.1'
        '26\\n1\\n. .'  → '26.1'
        '（\\n）\\n一'   → '（一）'
    """
    if text is None:
        return ""
    raw = str(text).replace("　", " ")
    compact = re.sub(r"\s+", "", raw)
    match = re.match(r"^[（(][）)]?([一二三四五六七八九十]+)$", compact)
    if match:
        return "（%s）" % match.group(1)
    if re.fullmatch(r"[\d.\s]+", compact):
        dotted = re.sub(r"[^\d.]", ".", raw)
        dotted = re.sub(r"\.{2,}", ".", dotted).strip(".")
        return dotted or compact
    return compact


# 分类列 → 大类。附表列序各可研不同（某企业：建筑/设备/其他；某项目：建筑/安装/设备/其它），
# 因此按**表头文字**判定列角色，而不是写死列号。
_COLUMN_ROLE_BY_HEADER = (
    ("建筑工程", "engineering"),
    ("建筑工", "engineering"),
    ("安装工程", "engineering"),
    ("设备", "engineering"),
    ("其它费用", "otherFees"),
    ("其他费用", "otherFees"),
)


def _column_roles(cells, positions):
    """按表头文字给「分类列」标角色；返回 {列索引: 大类}。"""
    roles = {}
    if not positions:
        return roles
    _seq_idx, name_idx, total_idx, _unit_idx = positions
    for index in range(name_idx + 1, total_idx):
        if index >= len(cells):
            continue
        text = _clean_cell(cells[index]).replace(" ", "")
        for token, role in _COLUMN_ROLE_BY_HEADER:
            if token in text:
                roles[index] = role
                break
    return roles


def _header_positions(cells):
    """从前两行表头里定位关键列的索引。

    各可研的附表表头措辞不同、合并单元格也不同：
      * 某企业饮品：序号 | 工程或费用名称 | 建筑工程费 | 设备购置及安装 | 其他费用 | 合计 | …
      * 某项目1958：序号 | 项目和费用名称 | 估算金额（万元） |(合并) | 建筑工程 | 安装工程 |
                  设备及工器具购置 | 其它费用 | 合计 | …
    注意「估算金额（万元）」是**合并单元格**的表头，真正的合计列在后面，
    所以「合计/小计」优先，「估算金额」只作兜底。

    返回 (seq_idx, name_idx, total_idx, unit_idx) 或 None。
    """
    header = " ".join(_clean_cell(cell) for cell in cells)
    seq_idx = name_idx = total_idx = unit_idx = None
    fallback_total = None
    for index, cell in enumerate(cells):
        text = _clean_cell(cell).replace(" ", "")
        if not text:
            continue
        if seq_idx is None and "序号" in text:
            seq_idx = index
        if name_idx is None and ("费用名称" in text or "工程名称" in text
                                 or "项目名称" in text or "费用及名称" in text):
            name_idx = index
        if total_idx is None and ("合计" in text or "小计" in text or "合价" in text):
            total_idx = index
        if fallback_total is None and ("估算金额" in text or "投资估算" in text):
            fallback_total = index
        if unit_idx is None and text in ("单位", "计量单位"):
            unit_idx = index
    if total_idx is None:
        total_idx = fallback_total
    if name_idx is None or total_idx is None:
        return None
    if seq_idx is None:
        seq_idx = 0
    return seq_idx, name_idx, total_idx, unit_idx


def _row_name_and_amount(cells, raw_cells=None, positions=None):
    """从一行里取 (序号, 名称, 金额, 分类列)。

    分类列 vs 金额列要分开看：
      * 分类列 = 建筑费/设备费/其他费 中**第一个非空**的，
        它说明这笔钱属于哪个大类（工程费 vs 工程建设其他费）；
      * 金额列 = 合计列，是这一行该取的数。
    raw_cells: 未经清洗的原始单元格 —— 序号的层级信息只存在于原始换行里
    （'1\\n2\\n1' 表示 1.2.1），清洗后 '121' 就丢了层级，父级剔除随之失效。
    """
    if len(cells) < 3:
        return None, None, None, None
    seq_idx, name_idx, total_idx, _unit_idx = positions or (0, 1, 5, None)
    raw_seq = raw_cells[seq_idx] if raw_cells and seq_idx < len(raw_cells) else None
    seq = _normalize_seq(raw_seq)
    name = _clean_cell(cells[name_idx]) if name_idx < len(cells) else ""
    amount = _restore_number(cells[total_idx]) if total_idx < len(cells) else None
    if not name:
        return seq, None, amount, None
    if amount is None:
        for index in range(total_idx + 1, len(cells)):
            candidate = _restore_number(cells[index])
            if candidate is not None:
                amount = candidate
                break
    # 分类列：在名称列之后、合计列之前，第一个有值的金额列
    category_col = None
    for index in range(name_idx + 1, total_idx):
        if index < len(cells) and _restore_number(cells[index]) is not None:
            category_col = index
            break
    return seq, name, amount, category_col




def _looks_like_estimate_table(rows):
    """这张表是不是「投资估算附表」？

    判据：含**大类行**（序号是中文数字一二三四五六，且名称是费用大类）
    或含「预备费 / 项目总投资 / 静态投资」这类汇总行。
    用来排除同页的其它表：资金筹措表、分街道/分小区工程量明细表、用电计算表等。
    """
    for row in rows:
        if not row:
            continue
        cells = [_clean_cell(cell) for cell in row]
        if len(cells) < 2:
            continue
        seq = _normalize_seq(cells[0])
        name = cells[1]
        if seq in _CATEGORY_ROWS and any(
                key in name for key in ("工程费用", "工程建设其他费用", "预备费",
                                        "建设期利息", "静态投资", "项目总投资")):
            return True
        if any(key in name for key in ("预备费", "项目总投资", "估算总投资", "静态投资")):
            return True
    return False


# 资金筹措类表的行名（这类表不是投资估算附表，混进来会让工程费虚高）
_FUNDING_ROW_NAMES = ("资金使用情况", "资金筹措", "资金来源", "资金筹措及使用计划")


def _is_funding_row(name):
    return any(key in (name or "") for key in _FUNDING_ROW_NAMES)


def parse_estimate_page(page, file_label, page_no, current=None, layout=None):
    """用 find_tables 解析一页投资估算附表。

    current: 传入上一个分类（大类跨页延续 —— 续页没有大类序号行，
    实测 114 页的「场地准备及临时设施费」等属于其他费，靠上一页的分类延续）。

    layout: 跨页继承的表头定位结果 {"positions":…, "roles":…}。
    附表续页**没有表头行**（实测某企业 113/114 页首行就是数据），
    若每页重新定位，续页会整页被跳过（某企业曾因此只解析出 1 项）。

    返回 ({"engineering": [...], "otherFees": [...], "contingency": [...], "total": [...]}, current)
    每项 = {"name", "amount", "source"}
    """
    try:
        tables = page.find_tables()
    except Exception:
        return {}, current, layout
    result = {"engineering": [], "otherFees": [], "contingency": [], "total": []}
    # 表头定位结果**跨行、跨表继承**：
    #   * 附表续页没有表头行（实测某企业 113/114 页首行就是数据）；
    #   * 表头可能分两行（实测某项目：r0 是合并表头、r1 才是分列名）。
    # 若每行重新定位，续页会被整页跳过（某企业曾因此只解析出 1 项）。
    positions = (layout or {}).get("positions")
    header_roles = dict((layout or {}).get("roles") or {})
    for table in tables.tables:
        try:
            rows = table.extract()
        except Exception:
            continue
        # 表级甄别：投资估算附表必须含「工程费用 / 工程建设其他费用 / 预备费 /
        # 项目总投资」之一的大类行，否则是别的表（资金筹措表、分街道/分小区明细表、
        # 用电计算表…）。混进来会让工程费虚高、预备费串项。
        if not _looks_like_estimate_table(rows):
            continue
        for row_index, row in enumerate(rows):
            if not row:
                continue
            row_text = " ".join(_clean_cell(cell) for cell in row)
            cells = [_clean_cell(cell) for cell in row]
            # 表头行：第一列是「序号」，或含「费用名称/项目名称」。
            # 表头可能分两行（某项目：r0 合并表头、r1 才列出分列名），
            # 因此把连续的表头候选行合并后再定位。
            if cells and (cells[0] in ("序号", "序号工程或费用名称")
                           or "费用名称" in row_text.replace(" ", "")
                           or "项目名称" in row_text.replace(" ", "")):
                merged = list(cells)
                for follow in rows[row_index + 1: row_index + 3]:
                    follow_cells = [_clean_cell(cell) for cell in follow]
                    for index, value in enumerate(follow_cells):
                        if index < len(merged) and not merged[index]:
                            merged[index] = value
                found = _header_positions(merged)
                if found:
                    positions = found
                    header_roles = _column_roles(merged, positions)
                continue
            if "项目总投资估算表" in row_text or "投资估算表" in row_text:
                continue
            if positions is None:
                positions = _header_positions(cells)
                if positions is None:
                    continue
                header_roles = _column_roles(cells, positions)
            roles = header_roles or _column_roles(cells, positions)
            seq, name, amount, category_col = _row_name_and_amount(cells, row, positions)
            category, current, by_keyword = _classify(seq, row_text, current)
            # 分类列兜底（关键词与序号都没命中时）：比序号可靠 ——
            # 续页没有大类序号行时只能靠它。
            if not by_keyword and category_col is not None:
                column_category = roles.get(category_col)
                if column_category:
                    category = column_category
                    current = column_category
            if not name or amount is None:
                continue
            # 资金筹措类表的行不是投资估算明细，直接排除
            if _is_funding_row(name):
                continue
            # 大类行（一/二/三…）：单独收集为权威合计，不进明细
            if seq in _CATEGORY_ROWS and category in result:
                result[category].append({
                    "name": name, "amount": amount, "seq": seq, "isCategory": True,
                    "note": " ".join(_clean_cell(cell) for cell in cells[6:7]).strip(),
                    "source": {"file": file_label, "page": page_no, "quote": row_text[:120]},
                })
                continue
            if name in _AGGREGATE:
                continue
            # 分组标题（「（一）基础设施改造」）单独收集
            if _is_bracket_seq(seq or ""):
                if category in result:
                    result[category].append({
                        "name": name, "amount": amount, "seq": seq, "isGroup": True,
                        "source": {"file": file_label, "page": page_no, "quote": row_text[:120]},
                    })
                continue
            if category not in result:
                continue
            result[category].append({
                "name": name,
                "amount": amount,
                "seq": seq,
                # 备注列常写取费口径（如「按（一+二-土地费）*8%计」），要留住
                "note": " ".join(_clean_cell(cell) for cell in cells[6:7]).strip(),
                "source": {"file": file_label, "page": page_no, "quote": row_text[:120]},
            })
    return result, current, {"positions": positions, "roles": header_roles}


def _seq_level(seq):
    """由序号推层级。

    PDF 表格抽取会吃掉序号里的小数点（'1.2.1' → '121'、'2.6.1' → '261'），
    层级信息在源 PDF 就已丢了一半，但**位数还在**：一位=1 级、两位=2 级、
    三位=3 级（实测 '1' / '11' / '121' 分别对应 1 / 1.1 / 1.2.1）。
    """
    if not seq:
        return 0
    if re.fullmatch(r"[（(][一二三四五六七八九十]+[）)]", seq):
        return 1
    if not re.fullmatch(r"\d+", seq):
        return 0
    return len(seq)


def _children_of(items, index):
    """取第 index 项的**直接子级**（层级恰好深一层），返回金额列表。

    只算直接子级：改造道路(1级) 的直接子级是 原路面处理(2级)、道路改造(2级)、
    人行道改造(2级)；1号路(3级) 是道路改造的子级，不该并进改造道路。
    """
    level = _seq_level(items[index].get("seq") or "")
    if not level:
        return []
    children = []
    for follow in range(index + 1, len(items)):
        child_level = _seq_level(items[follow].get("seq") or "")
        if child_level <= level:
            break
        if child_level == level + 1:
            children.append(items[follow]["amount"])
    return children


def drop_parents(items):
    """把附表明细收敛到**分组层**（保留分组、丢弃更细的数字层）。

    附表层级（实测某企业饮品可研附表1）：
        一  工程费用                    4301.76   ← 大类行（isCategory，权威值）
          （一）基础设施改造            2821.07   ← 分组行（isGroup）→ 取这一层
            1 改造道路                 2383.39
              1.1 原路面处理             430.15
              1.2 道路改造              1895.70
                1.2.1 1号路             1597.20

    取分组层的三条理由：
      1. **分组金额 = 其子项之和，附表本身保证**；数字层可能因跨页、
         名称换行、单元格拆散而抽不全（实测「文化体育场地 719.00」的
         子项只抽到 73.13，差 645.87）。
      2. 分组层粒度合适：既是可核查的投资条目（不是「1号路」这种施工段），
         又能拼出 25 个一级科目的对应关系。
      3. 大类行（一/二/三）另作权威合计，与分组层求和交叉校验。

    没有分组层的类别（如工程建设其他费：1土地费/2技术咨询费用/3工程建设管理费…）
    退回取**最浅的数字层**。

    返回 (明细列表, 大类行列表)
    """
    if not items:
        return [], []
    categories = [item for item in items if item.get("isCategory")]
    groups = [item for item in items if item.get("isGroup")]
    plain = [item for item in items
             if not item.get("isCategory") and not item.get("isGroup")]
    if groups:
        detail = groups
    elif plain:
        levels = [_seq_level(item.get("seq") or "") for item in plain]
        numeric = [level for level in levels if level]
        target = min(numeric) if numeric else 0
        detail = [item for item, level in zip(plain, levels)
                  if level == target] if target else plain
    else:
        detail = []
    return detail, categories


def _is_bracket_seq(seq):
    """序号是否是括号型分组标题（如「（一）」）。"""
    return bool(seq) and bool(re.fullmatch(r"[（(][一二三四五六七八九十]+[）)]", seq))


def check_parent_sums(items, kept):
    """校验父级金额 = 其**直接子级**之和；返回不一致的说明列表。"""
    warnings = []
    levels = [_seq_level(item.get("seq") or "") for item in items]
    for index in range(len(items) - 1):
        if not (levels[index] and levels[index + 1] > levels[index]):
            continue
        children = _children_of(items, index)
        total = sum(children, Decimal(0))
        declared = items[index]["amount"]
        if not children:
            continue
        if total != declared:
            warnings.append({
                "name": items[index]["name"],
                "declared": declared,
                "children": total,
                "diff": abs(total - declared),
            })
    return warnings


def find_estimate_pages(doc, max_span=6):
    """定位投资估算附表的全部页（首页 + 连续续页）。

    表头措辞各可研不一，实测：
      * 某企业饮品：含「工程或费用名称」+「项目总投资估算表」
      * 某项目 1958：含「投资估算表」（无「工程或费用名称」）
      * 某区：含「投资估算表」（无「工程或费用名称」）
    因此关键词要放宽，并且用「页内有表格 + 含投资估算类词」作为兜底。
    """
    heads = []
    for index in range(doc.page_count):
        text = doc[index].get_text()
        if ("工程或费用名称" in text or "项目总投资估算表" in text
                or "投资估算表" in text or "总投资估算表" in text
                or "工程总投资估算" in text or "工程总投资匡算" in text):
            heads.append(index + 1)
    if not heads:
        return []
    pages = []
    for head in heads:
        if head in pages:
            continue
        pages.append(head)
        for offset in range(1, max_span):
            page_no = head + offset
            if page_no > doc.page_count:
                break
            page = doc[page_no - 1]
            text = page.get_text()
            if not text.strip():
                break
            if re.search(r"第[一二三四五六七八九十]+章", text[:200]):
                break
            try:
                if not page.find_tables().tables:
                    break
            except Exception:
                break
            pages.append(page_no)
    return sorted(set(pages))
