"""从成果文件里的「投资匡算/财务评价工作簿」搬数据，造财务参数.xlsx 的输入。

定位：**倒推只是造输入材料**，判据是"够用能跑"，不是"完美还原"。
所以这里只搬 9 张参数表真正需要的字段，抽不到就留空（由缺项表登记）。

识别的表（**内容特征优先，表名兜底** —— 各家项目表名不一致）：
  | 目标            | 识别特征                                    |
  |-----------------|---------------------------------------------|
  | 投资估算表      | 表头含「工程或费用名称」/「项目和费用名称」   |
  | 资金筹措表      | 表名或内容含「资金筹措」                     |
  | 收入表          | 内容含「收入合计」或表名含「收入」            |
  | 总成本费用表    | 内容含「总成本费用」                         |
  | 借款还本付息表  | 表名含「还本付息」                           |
  | 建设规模表      | 内容含「规划用地面积」/「建设内容与规模」     |

实测覆盖：某市某企业片区附表（58 表）、某街附表（58 表，同模板）、
某新区附表（10 表，列名不同但同构）、某路基础改造项表。
"""

from __future__ import annotations

import re
from decimal import Decimal

# ---------------------------------------------------------------- 基础工具


def _clean(value):
    if value is None:
        return ""
    text = str(value).replace("\n", "").replace("　", " ")
    return re.sub(r"\s+", " ", text).strip()


def _num(value):
    """单元格 → Decimal（容忍千分位与「万元」后缀）。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = _clean(value).replace(",", "").replace("，", "")
    for suffix in ("万元", "元"):
        if text.endswith(suffix):
            text = text[: -len(suffix)]
            break
    if not re.fullmatch(r"-?\d+(\.\d+)?", text):
        return None
    try:
        return Decimal(text)
    except Exception:
        return None


def _head_text(rows, limit=8):
    parts = []
    for row in rows[:limit]:
        for cell in row:
            text = _clean(cell)
            if text:
                parts.append(text)
    return " ".join(parts)


def _total_column(rows, header_index=0, limit=5):
    """按表头定位「合计」/「小计」列的索引。

    各家附表的列序不同，**不能从右往左抓数字** —— 资金筹措表右边是分年列，
    往右抓会抓到「占比 0.5792」这种比例数（实测某企业专项债 12100 被读成 0.5792）。
    """
    for row in rows[header_index: header_index + limit]:
        for index, cell in enumerate(row):
            if _clean(cell).replace(" ", "") in (
                    "小计", "合计", "小计（万元）", "合计（万元）", "合价", "总投资"):
                return index
    return None


def _value_at(row, index):
    if index is None or index >= len(row):
        return None
    return _num(row[index])


def _find_sheet(workbook, must_text=(), name_hint=(), limit=8):
    """按内容特征（优先）或表名关键词找表；返回 (表名, 行列表)。"""
    fallback = None
    for name in workbook.sheetnames:
        try:
            rows = list(workbook[name].iter_rows(values_only=True))
        except Exception:
            continue
        if not rows:
            continue
        head = _head_text(rows, limit)
        if must_text and all(word in head for word in must_text):
            return name, rows
        if name_hint and any(word in name for word in name_hint) and fallback is None:
            fallback = (name, rows)
    return fallback if fallback else (None, [])


# ---------------------------------------------------------------- 投资估算表

# 大类识别：按**包含**匹配，不写死全名。
# 实测各家写法不一：「工程建设其他费」（某企业，无"用"字）/「工程建设其他费用」/
# 「工程建设其它费」都有 —— 写死全名会整行丢失（实测丢 2198.29 万其他费）。
_CATEGORY_KEYWORDS = (
    ("其他费", "otherFees"),
    ("其它费", "otherFees"),
    ("工程费用", "engineering"),
    ("工程费", "engineering"),
    ("预备费", "contingency"),
    ("建设期利息", "interest"),
    ("建设期融资费", "interest"),
    ("流动资金", "workingCapital"),
    ("项目总投资", "total"),
)


def parse_investment_table(rows):
    """投资估算表 → {engineering/otherFees/contingency/interest/total: [明细]}。

    列序（各家一致，只是表头措辞不同）：
        序号 | 工程或费用名称 | 建筑工程 | 安装工程 | 设备购置 | 其他费用 | 合计(小计)
    层级：中文序号=大类；括号序号=分组；数字序号=明细。
    取**分组层**（如「（一）厂区道路工程」）作明细 —— 它金额=子项之和，附表本身保证；
    没有分组层时取数字层里「有子项的父级」。
    """
    if not rows:
        return {}
    header_index = 0
    for index, row in enumerate(rows[:8]):
        cells = [_clean(cell) for cell in row]
        if any("工程或费用名称" in cell or "项目和费用名称" in cell for cell in cells):
            header_index = index
            break
    # 合计列：表头里的「小计」/「合计」
    total_col = _total_column(rows, header_index)
    if total_col is None:
        counts = {}
        for row in rows[header_index + 1: header_index + 20]:
            for index, cell in enumerate(row):
                if _num(cell) is not None:
                    counts[index] = counts.get(index, 0) + 1
        if not counts:
            return {}
        total_col = max(counts, key=lambda i: (counts[i], -i))

    groups = {}
    current = None
    items = []
    for row in rows[header_index + 1:]:
        if not row or len(row) <= 1:
            continue
        seq = _clean(row[0])
        name = _clean(row[1]) if len(row) > 1 else ""
        if not name:
            continue
        amount = _num(row[total_col]) if total_col < len(row) else None
        if amount is None:
            continue
        category = None
        for keyword, key in _CATEGORY_KEYWORDS:
            if keyword in name:
                category = key
                break
        if category and seq in ("一", "二", "三", "四", "五", "六"):
            current = category
            groups.setdefault(category, []).append(
                {"name": name.strip(), "amount": amount, "level": "category"})
            continue
        if re.fullmatch(r"[（(][一二三四五六七八九十]+[）)]", seq):
            groups.setdefault(current or "engineering", []).append(
                {"name": name.strip(), "amount": amount, "level": "group"})
            continue
        if re.fullmatch(r"\d+(\.\d+)*", seq):
            items.append({"name": name.strip(), "amount": amount,
                          "seq": seq, "category": current or "engineering"})
    # 数字层明细：剔掉「有子项的父级」
    # 父级必须**按大类**判定：序号在不同大类里会重复（工程费的 5/6/7 与
    # 其他费的 5 环评、6 场地准备、7 工程保险），全局去重会把后三者误删。
    parents = set()
    for item in items:
        parts = item["seq"].split(".")
        for index in range(1, len(parts)):
            parents.add((item["category"], ".".join(parts[:index])))
    leaves = [item for item in items
              if (item["category"], item["seq"]) not in parents]
    for item in leaves:
        groups.setdefault(item["category"], []).append(
            {"name": item["name"], "amount": item["amount"], "level": "item"})
    return groups


# ---------------------------------------------------------------- 资金筹措表

_FUNDING_KEYWORDS = (
    ("资本金", "equity"),
    ("自有资金", "equity"),
    ("自筹资金", "equity"),
    ("企业自筹", "equity"),
    ("财政资金", "subsidy"),
    ("政府投资", "subsidy"),
    ("补助", "subsidy"),
    ("专项资金", "subsidy"),
    ("专项债", "loan"),
    ("专项借款", "loan"),
    ("银行贷款", "loan"),
    ("银行借款", "loan"),
    ("借款", "loan"),
)


def parse_funding_table(rows):
    """资金筹措表 → {amounts:{equity/subsidy/loan}, rate}。"""
    if not rows:
        return {}
    amounts = {}
    rate = None
    total_col = _total_column(rows)
    for row in rows:
        cells = [_clean(cell) for cell in row]
        if len(cells) < 3:
            continue
        seq, name = cells[0], cells[1]
        if not name:
            continue
        amount = _value_at(row, total_col)
        if amount is None:
            continue
        if "利率" in name and rate is None:
            rate = amount
            continue
        if not re.fullmatch(r"\d+(\.\d+)*", seq or ""):
            continue
        for keyword, key in _FUNDING_KEYWORDS:
            if keyword in name:
                # 只取顶层（2.1 / 2.2）；子项（2.1.1 用于建设投资）不计
                if seq.count(".") == 1:
                    if key not in amounts or amount > amounts[key]:
                        amounts[key] = amount
                break
    return {"amounts": amounts, "rate": rate}


# ---------------------------------------------------------------- 收入表

def parse_revenue_table(rows):
    """收入表 → [{name, annualAmount, totalAmount}]。

    附表4 的「合计」列是**整个运营期**的合计，不是年均值；财务参数表要的是
    「年均收入」。这里按年列表头（含「上半年/下半年」半年的按 0.5 年）折算：
        年均收入 = 合计 ÷ 运营期年数
    收入项也**不能**靠名称里含「收入」来识别（实测「充电桩」就没有这两个字），
    而是按「一 收入合计 … 二 增值税及附加」之间、带数字序号的明细行来收。
    """
    if not rows:
        return []
    total_col = _total_column(rows)
    if total_col is None:
        return []
    start = end = None
    for index, row in enumerate(rows):
        name = _clean(row[1]) if len(row) > 1 else ""
        seq = _clean(row[0]) if row else ""
        if start is None and "收入合计" in name:
            start = index
        elif start is not None and seq in ("二", "三") and ("增值税" in name or "税金" in name):
            end = index
            break
    if start is None:
        return []
    # 运营期年数：年表头中整年计 1、半年计 0.5
    years = Decimal("0")
    for row in rows[:12]:
        texts = list(row)[total_col + 1:]
        matched = 0
        for cell in texts:
            if isinstance(cell, (int, float)) and 2000 <= float(cell) <= 2100:
                matched += 1
            elif re.fullmatch(r"20\d{2}(上|下)?半年?", _clean(cell)):
                matched += 1
        if matched >= 3:
            for cell in texts:
                # 年表头有时是数字（2029 → 2029.0），有时是文本（2028下半年）
                if isinstance(cell, (int, float)) and 2000 <= float(cell) <= 2100:
                    years += Decimal("1")
                else:
                    text = _clean(cell)
                    if re.fullmatch(r"20\d{2}", text):
                        years += Decimal("1")
                    elif re.fullmatch(r"20\d{2}(上|下)半年", text):
                        years += Decimal("0.5")
            break
    if years <= 0:
        years = Decimal("1")
    items = []
    for row in rows[start + 1:end or len(rows)]:
        if not row or len(row) < 3:
            continue
        seq = _clean(row[0])
        name = _clean(row[1])
        if not name or not re.fullmatch(r"\d+", seq or ""):
            continue
        amount = _num(row[total_col])
        if amount is None or amount == 0:
            continue
        items.append({
            "name": name.strip(),
            "annualAmount": (amount / years).quantize(Decimal("0.01")),
            "totalAmount": amount,
        })
    return items


# ---------------------------------------------------------------- 成本表

def parse_cost_table(rows):
    """总成本费用表 → {laborCost/repair/...}。"""
    if not rows:
        return {}
    result = {}
    for row in rows:
        cells = [_clean(cell) for cell in row]
        if len(cells) < 3:
            continue
        name = cells[1]
        amount = None
        for index in range(len(cells) - 1, 1, -1):
            candidate = _num(cells[index])
            if candidate is not None:
                amount = candidate
                break
        if amount is None:
            continue
        for keyword, key in (("人工费", "laborCost"), ("工资", "laborCost"),
                             ("外购燃料", "fuel"), ("维修费", "repair"),
                             ("管理费", "management"), ("财务费", "finance")):
            if keyword in name:
                result.setdefault(key, amount)
                break
    return result


# ---------------------------------------------------------------- 规模表

_SCOPE_FIELDS = (
    ("规划用地面积", "siteArea"),
    ("规划建设用地面积", "siteArea"),
    ("改造面积", "rebuildArea"),
    ("规划总建筑面积", "buildingArea"),
    ("总建筑面积", "buildingArea"),
    ("地上建筑面积", "groundArea"),
)


def parse_scope_table(rows):
    """建设内容与规模表 → {siteArea/rebuildArea/buildingArea/...}。"""
    if not rows:
        return {}
    result = {}
    for row in rows:
        cells = [_clean(cell) for cell in row]
        for index, cell in enumerate(cells):
            for keyword, key in _SCOPE_FIELDS:
                if cell == keyword:
                    for follow in range(index + 1, min(index + 4, len(cells))):
                        value = _num(cells[follow])
                        if value is not None:
                            result.setdefault(key, value)
                            break
                    break
    return result


# ---------------------------------------------------------------- 总入口

def extract_from_appendix_workbook(path):
    """匡算/财务评价工作簿 → 财务参数字典。

    返回 {"engineering": [...], "otherFees": [...], "contingency": [...],
          "finance": {...}, "revenue": [...], "cost": {...}, "scope": {...}}
    """
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return None

    result = {"detail": {"engineering": [], "otherFees": [], "contingency": [],
                         "interest": [], "total": []},
              "finance": {}, "revenue": [], "cost": {}, "scope": {}}

    # 投资估算表
    _name, rows = _find_sheet(workbook, must_text=("工程或费用名称",))
    if not rows:
        _name, rows = _find_sheet(workbook, must_text=("项目和费用名称",))
    if rows:
        groups = parse_investment_table(rows)
        # 大类行（一/二/三…）单独留一份：02 的预备费反算与平衡自检要用它的权威值，
        # 明细层则不能混进大类（否则重复计数）。
        result["categoryRows"] = {
            key: [item for item in groups.get(key, []) if item["level"] == "category"]
            for key in ("engineering", "otherFees", "contingency", "interest", "total")
        }
        for items_ in result["categoryRows"].values():
            for item in items_:
                item.setdefault("source", {
                    "file": path.name, "sheet": _name, "row": None,
                    "quote": item.get("name", "")})
        for key in ("engineering", "otherFees", "contingency", "interest", "total"):
            # 优先分组层（（一）xxx，金额=子项之和）；没有则用数字层叶子
            picked = [item for item in groups.get(key, []) if item["level"] == "group"]
            if not picked:
                picked = [item for item in groups.get(key, []) if item["level"] == "item"]
            result["detail"][key] = picked

    # 资金筹措表
    _name, rows = _find_sheet(workbook, name_hint=("资金筹措",),
                              must_text=("投资总额",))
    if rows:
        funding = parse_funding_table(rows)
        result["finance"] = funding

    # 收入表
    _name, rows = _find_sheet(workbook, must_text=("收入合计",))
    if not rows:
        _name, rows = _find_sheet(workbook, name_hint=("收入",))
    if rows:
        result["revenue"] = parse_revenue_table(rows)

    # 成本表
    _name, rows = _find_sheet(workbook, must_text=("总成本费用",))
    if rows:
        result["cost"] = parse_cost_table(rows)

    # 规模表
    _name, rows = _find_sheet(workbook, must_text=("规划用地面积",))
    if not rows:
        _name, rows = _find_sheet(workbook, name_hint=("建设内容与规模", "对标"))
    if rows:
        result["scope"] = parse_scope_table(rows)
    return result
